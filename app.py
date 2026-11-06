import io
import json
import os
import re
import sqlite3
import time
from datetime import datetime
from typing import Dict, List, Optional

from bs4 import BeautifulSoup
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor
from google import genai
from google.genai import types
import requests
import streamlit as st

# ==========================================
# 1. DATABASE LAYER (SQLite)
# ==========================================
DB_NAME = "appraisals.db"


def get_connection():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_connection() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS appraisals (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                paper_title TEXT NOT NULL,
                study_design TEXT,
                tool_selected TEXT,
                overall_verdict TEXT,
                source_name TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                full_json TEXT NOT NULL
            )
        """)
        conn.commit()


def save_appraisal(data: dict, source_name: str) -> int:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            """
            INSERT INTO appraisals (
                paper_title, study_design, tool_selected, 
                overall_verdict, source_name, created_at, full_json
            ) VALUES (?, ?, ?, ?, ?, ?, ?)
        """,
            (
                data.get("paper_title", "Untitled Manuscript"),
                data.get("study_design", "Unknown"),
                data.get("appraisal_tool_selected", "Unknown"),
                data.get("overall_verdict", "Unknown"),
                source_name,
                datetime.now().strftime("%Y-%m-%d %H:%M"),
                json.dumps(data),
            ),
        )
        conn.commit()
        return cursor.lastrowid


def get_all_appraisals() -> List[Dict]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("""
            SELECT id, paper_title, study_design, tool_selected, 
                   overall_verdict, source_name, created_at 
            FROM appraisals 
            ORDER BY id DESC
        """)
        return [dict(row) for row in cursor.fetchall()]


def get_appraisal_by_id(appraisal_id: int) -> Optional[dict]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute(
            "SELECT full_json FROM appraisals WHERE id = ?", (appraisal_id,)
        )
        row = cursor.fetchone()
        if row:
            return json.loads(row["full_json"])
        return None


def delete_appraisal(appraisal_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM appraisals WHERE id = ?", (appraisal_id,))
        conn.commit()


# ==========================================
# 2. APPRAISAL FRAMEWORKS & TAXONOMY
# ==========================================
TOOL_TAXONOMY = {
    "Clinical Practice Guideline / Consensus Statement": {
        "tool": "AGREE II",
        "domains": [
            "Domain 1: Scope and Purpose",
            "Domain 2: Stakeholder Involvement",
            "Domain 3: Rigour of Development",
            "Domain 4: Clarity of Presentation",
            "Domain 5: Applicability",
            "Domain 6: Editorial Independence",
        ],
        "scale": "1 (Strongly Disagree / Absent) to 7 (Strongly Agree / Exemplary)",
    },
    "Randomized Controlled Trial (RCT)": {
        "tool": "Cochrane RoB 2",
        "domains": [
            "Domain 1: Bias arising from the randomization process",
            "Domain 2: Bias due to deviations from intended interventions",
            "Domain 3: Bias due to missing outcome data",
            "Domain 4: Bias in measurement of the outcome",
            "Domain 5: Bias in selection of the reported result",
        ],
        "scale": "Low Risk | Some Concerns | High Risk",
    },
    "Non-Randomized Intervention / Quasi-Experimental": {
        "tool": "ROBINS-I",
        "domains": [
            "Domain 1: Confounding",
            "Domain 2: Selection of participants",
            "Domain 3: Classification of interventions",
            "Domain 4: Deviations from intended interventions",
            "Domain 5: Missing data",
            "Domain 6: Measurement of outcomes",
            "Domain 7: Selection of the reported result",
        ],
        "scale": (
            "Low Risk | Moderate Risk | Serious Risk | Critical Risk | No"
            " Information"
        ),
    },
    "Systematic Review / Meta-Analysis": {
        "tool": "AMSTAR 2",
        "domains": [
            "Domain 1: PICO & Protocol Pre-registration (Critical)",
            "Domain 2: Comprehensive Literature Search Strategy (Critical)",
            "Domain 3: Duplicate Study Selection & Extraction",
            "Domain 4: Excluded Studies Justification (Critical)",
            "Domain 5: Risk of Bias Assessment in Primary Studies (Critical)",
            "Domain 6: Meta-Analytic Synthesis Appropriateness (Critical)",
            "Domain 7: Publication Bias Assessment (Critical)",
            "Domain 8: Conflict of Interest & Funding Reporting",
        ],
        "scale": "Yes | Partial Yes | No",
    },
    "Observational Study (Cohort or Case-Control)": {
        "tool": "Newcastle-Ottawa Scale (NOS) / CASP",
        "domains": [
            "Domain 1: Selection of Cohorts / Cases",
            "Domain 2: Comparability of Groups (Control for Confounders)",
            "Domain 3: Ascertainment of Exposure and Outcome",
        ],
        "scale": "Adequate (Star awarded) | Inadequate | Unclear",
    },
    "Diagnostic Accuracy Study": {
        "tool": "QUADAS-2",
        "domains": [
            "Domain 1: Patient Selection (Risk & Applicability)",
            "Domain 2: Index Test Execution & Interpretation",
            "Domain 3: Reference Standard Validity",
            "Domain 4: Flow and Timing of Assessments",
        ],
        "scale": "Low Risk | High Risk | Unclear",
    },
    "Narrative Review": {
        "tool": "SANRA",
        "domains": [
            "Domain 1: Justification of article's importance",
            "Domain 2: Statement of concrete aims",
            "Domain 3: Description of the literature search",
            "Domain 4: Referencing of primary literature",
            "Domain 5: Scientific reasoning and balanced presentation",
            "Domain 6: Presentation of relevant data",
        ],
        "scale": "0 (Score 0) | 1 (Score 1) | 2 (Score 2 - Fully met)",
    },
    "Qualitative Study": {
        "tool": "CASP Qualitative / COREQ",
        "domains": [
            "Domain 1: Research design appropriateness",
            "Domain 2: Sampling and recruitment strategy",
            "Domain 3: Data collection and reflexivity of the researcher",
            "Domain 4: Ethical considerations addressed",
            "Domain 5: Rigour and depth of data analysis",
            "Domain 6: Clear statement of findings and value",
        ],
        "scale": "Yes | Can't Tell | No",
    },
}

# ==========================================
# 3. EXTRACTION & GEMINI ENGINE
# ==========================================
def extract_text_from_url(url: str) -> str:
    headers = {"User-Agent": "Mozilla/5.0"}
    if not url.startswith("http"):
        url = f"https://doi.org/{url}" if "10." in url else f"https://{url}"

    resp = requests.get(url, headers=headers, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    for elem in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        elem.decompose()
    article = soup.find("article") or soup.find("main") or soup.body
    return article.get_text(separator="\n", strip=True)[:45000] if article else ""


def appraise_with_gemini(
    client: genai.Client,
    file_bytes: bytes = None,
    text_content: str = None,
) -> dict:
    taxonomy_json = json.dumps(TOOL_TAXONOMY, indent=2)

    system_instruction = f"""
    You are an expert epidemiological reviewer and methodological appraiser.
    Evaluate the provided research paper or clinical guideline and perform a rigorous critical appraisal.

    Appraisal Tool Taxonomy:
    {taxonomy_json}

    Step-by-Step Instructions:
    1. Study Design Classification:
       - Classify the manuscript into ONE design from the taxonomy.
       - If it is a formal clinical practice guideline, consensus statement, or clinical pathway with graded recommendations, classify it as "Clinical Practice Guideline / Consensus Statement" and apply AGREE II.
    2. Tool Assignment:
       - Apply the corresponding Gold-Standard Tool and evaluate EVERY domain listed for that tool.
    3. Grounded Evidence Extraction:
       - Extract the exact 'verbatim_evidence' (direct quote).
       - Provide the exact 'location_reference' (e.g., "Page 4, Methods §Randomization" or section name).
       - Write a nuanced 'critique' justifying the score.
    4. Provide an overall verdict and comprehensive synthesis.

    CRITICAL: Output ONLY a valid JSON object matching this schema:
    {{
      "paper_title": "Identified title",
      "study_design": "Identified design matching taxonomy",
      "appraisal_tool_selected": "Name of appraisal tool",
      "overall_verdict": "Low Risk / Moderate Risk / High Risk / Recommended / etc.",
      "overall_synthesis": "Comprehensive narrative explaining methodological quality and clinical applicability.",
      "domain_appraisals": [
        {{
          "domain_name": "Domain name as per taxonomy",
          "rating": "Score or Rating",
          "location_reference": "Section or Page number reference",
          "verbatim_evidence": "Direct quote from text",
          "critique": "Analytical rationale for this domain"
        }}
      ]
    }}
    """

    contents = []
    if file_bytes:
        contents.append(
            types.Part.from_bytes(data=file_bytes, mime_type="application/pdf")
        )
    elif text_content:
        contents.append(text_content)

    contents.append(
        "Perform the full critical appraisal on this paper and return the JSON"
        " object."
    )

    candidate_models = ["gemini-3.8-flash", "gemini-2.5-flash", "gemini-1.5-pro"]
    last_error = None

    for model_name in candidate_models:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        system_instruction=system_instruction,
                        response_mime_type="application/json",
                        temperature=0.1,
                    ),
                )
                return json.loads(response.text)
            except Exception as e:
                err_str = str(e)
                last_error = e
                if "503" in err_str or "UNAVAILABLE" in err_str or "429" in err_str:
                    time.sleep(2 * (attempt + 1))
                    continue
                elif "404" in err_str or "NOT_FOUND" in err_str:
                    break
                else:
                    raise e

    raise last_error


# ==========================================
# 4. WORD DOCUMENT GENERATION (.docx)
# ==========================================
def set_cell_background(cell, fill_hex: str):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def build_docx_report(data: dict) -> io.BytesIO:
    doc = Document()

    title = doc.add_heading(level=0)
    run_t = title.add_run("Critical Appraisal & Methodological Quality Report")
    run_t.font.size = Pt(20)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(31, 78, 121)

    p_sub = doc.add_paragraph(
        "Automated evidence synthesis evaluated via Google Gemini."
    )
    p_sub.runs[0].font.size = Pt(9.5)
    p_sub.runs[0].italic = True

    meta_table = doc.add_table(rows=4, cols=2)
    meta_table.style = "Table Grid"
    fields = [
        ("Evaluated Title:", data.get("paper_title", "N/A")),
        ("Study Architecture:", data.get("study_design", "N/A")),
        (
            "Gold-Standard Tool Applied:",
            data.get("appraisal_tool_selected", "N/A"),
        ),
        ("Overall Quality / Risk Verdict:", data.get("overall_verdict", "N/A")),
    ]
    for idx, (label, val) in enumerate(fields):
        row = meta_table.rows[idx]
        set_cell_background(row.cells[0], "F2F5F8")
        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(10)
        p1 = row.cells[1].paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.size = Pt(10)
        if idx == 3:
            r1.bold = True
            r1.font.color.rgb = RGBColor(170, 30, 30)

    doc.add_paragraph("")
    doc.add_heading("1. Executive Methodological Synthesis", level=1)
    p_syn = doc.add_paragraph(data.get("overall_synthesis", ""))
    p_syn.paragraph_format.line_spacing = 1.15

    doc.add_heading(
        f"2. Domain Appraisal ({data.get('appraisal_tool_selected')})", level=1
    )
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = [
        "Domain Evaluated",
        "Appraisal Score",
        "Verbatim Evidence & Citation",
        "Methodological Critique",
    ]
    for i, title_text in enumerate(headers):
        cell = table.rows[0].cells[i]
        set_cell_background(cell, "1F4E79")
        p = cell.paragraphs[0]
        r = p.add_run(title_text)
        r.bold = True
        r.font.size = Pt(9.5)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for item in data.get("domain_appraisals", []):
        row_cells = table.add_row().cells
        row_cells[0].paragraphs[0].add_run(item.get("domain_name", "")).bold = (
            True
        )
        row_cells[1].paragraphs[0].add_run(item.get("rating", "")).bold = True

        p_e = row_cells[2].paragraphs[0]
        loc = item.get("location_reference", "Section")
        r_loc = p_e.add_run(f"[{loc}]\n")
        r_loc.bold = True
        r_loc.font.size = Pt(8)
        r_loc.font.color.rgb = RGBColor(90, 90, 90)
        r_q = p_e.add_run(f'"{item.get("verbatim_evidence", "")}"')
        r_q.italic = True
        r_q.font.size = Pt(8.5)

        row_cells[3].paragraphs[0].add_run(item.get("critique", ""))

    doc.add_paragraph("")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# ==========================================
# 5. STREAMLIT APP & SECURITY GATE
# ==========================================
st.set_page_config(page_title="AppraiseAI", page_icon="📑", layout="wide")
init_db()

st.markdown(
    """
<style>
    .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; }
</style>
""",
    unsafe_allow_html=True,
)

CORRECT_PASSCODE = "324598"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 AppraiseAI: Access Restricted")
    st.write("Please enter the passcode to access the appraisal workspace.")
    passcode_input = st.text_input("Passcode", type="password")
    if st.button("Unlock App"):
        if passcode_input.strip() == CORRECT_PASSCODE:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Invalid passcode. Please try again.")
    st.stop()

# --- Main App (Post Authentication) ---
st.title("📑 AppraiseAI: Research Paper & Guideline Grader")
st.caption("Powered by Gemini | Automated Tool Selection & Word Export")

api_key = st.secrets.get(
    "GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", None)
)

with st.sidebar:
    st.header("⚙️ Settings")
    if not api_key:
        api_key = st.text_input(
            "Gemini API Key",
            type="password",
            help="Enter your Google AI Studio API key",
        )
    else:
        st.success("API Key Active")

    if st.button("Log Out"):
        st.session_state["authenticated"] = False
        st.rerun()

    st.markdown("---")
    st.header("🗄️ Past Appraisals")
    records = get_all_appraisals()
    if not records:
        st.caption("No saved papers yet.")
    else:
        options = {
            r["id"]: (
                f"{r['created_at'][:10]} | {r['tool_selected']} -"
                f" {r['paper_title'][:25]}..."
            )
            for r in records
        }
        selected_id = st.selectbox(
            "Select saved appraisal:",
            list(options.keys()),
            format_func=lambda x: options[x],
        )
        col_load, col_del = st.columns(2)
        with col_load:
            if st.button("Load"):
                st.session_state["appraisal_data"] = get_appraisal_by_id(
                    selected_id
                )
                st.rerun()
        with col_del:
            if st.button("Delete"):
                delete_appraisal(selected_id)
                st.rerun()

tab_pdf, tab_url = st.tabs(
    ["📤 Upload PDF (Mobile / Desktop)", "🔗 Enter Link or DOI"]
)

pdf_bytes = None
url_text = None
filename_tag = "Paper"

with tab_pdf:
    uploaded = st.file_uploader(
        "Upload Manuscript or Clinical Practice Guideline", type=["pdf"]
    )
    if uploaded:
        pdf_bytes = uploaded.read()
        filename_tag = uploaded.name
        st.success(f"Loaded `{uploaded.name}` ({len(pdf_bytes) // 1024} KB)")

with tab_url:
    url_input = st.text_input(
        "Paste Article Link or DOI",
        placeholder="https://doi.org/10.1016/... or https://pubmed...",
    )
    if st.button("Fetch Article Content"):
        if url_input.strip():
            with st.spinner("Retrieving article..."):
                try:
                    url_text = extract_text_from_url(url_input.strip())
                    filename_tag = "Web_Article"
                    st.success("Article successfully retrieved.")
                except Exception as e:
                    st.error(f"Failed to fetch content: {str(e)}")

if st.button("⚡ Run Methodological Appraisal", type="primary"):
    if not api_key:
        st.warning("Please enter your Gemini API key in the sidebar.")
    elif not pdf_bytes and not url_text:
        st.warning("Please upload a PDF or fetch an article URL first.")
    else:
        client = genai.Client(api_key=api_key)
        with st.spinner(
            "Gemini is classifying design, matching the tool, and evaluating"
            " domains..."
        ):
            try:
                data = appraise_with_gemini(
                    client, file_bytes=pdf_bytes, text_content=url_text
                )
                new_id = save_appraisal(data, filename_tag)
                st.session_state["appraisal_data"] = data
                st.success(
                    f"Appraisal finished and saved to library (ID: #{new_id})!"
                )
                st.rerun()
            except Exception as e:
                st.error(f"Appraisal failed: {str(e)}")

# Display Results
if "appraisal_data" in st.session_state:
    data = st.session_state["appraisal_data"]
    st.markdown("---")
    st.subheader(f"📋 Appraisal: {data.get('paper_title', 'Untitled')}")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Study Design", data.get("study_design"))
    with c2:
        st.metric("Tool Selected", data.get("appraisal_tool_selected"))
    with c3:
        st.metric("Overall Verdict", data.get("overall_verdict"))

    st.info(f"**Synthesis:** {data.get('overall_synthesis')}")

    st.subheader("📊 Domain Breakdown")
    rows = [
        {
            "Domain": d.get("domain_name"),
            "Score/Rating": d.get("rating"),
            "Reference": d.get("location_reference"),
            "Verbatim Evidence": d.get("verbatim_evidence"),
            "Methodological Critique": d.get("critique"),
        }
        for d in data.get("domain_appraisals", [])
    ]
    st.dataframe(rows, use_container_width=True)

    doc_buf = build_docx_report(data)
    safe_title = re.sub(
        r"[^a-zA-Z0-9_-]", "_", data.get("paper_title", "Report")[:25]
    )
    st.download_button(
        label="📥 Download Completed Word Document (.docx)",
        data=doc_buf,
        file_name=(
            f"Appraisal_{data.get('appraisal_tool_selected')}_{safe_title}.docx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ),
        use_container_width=True,
    )
