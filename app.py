import difflib
import io
import json
import os
import re
import sqlite3
import time
from datetime import datetime
from typing import Dict, List, Optional, Tuple

from bs4 import BeautifulSoup
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from docx.shared import Pt, RGBColor
from google import genai
from google.genai import types
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import requests
import streamlit as st

DB_NAME = "appraisals.db"
APP_NAME = "Kitique Appraisal Tool"


def get_connection():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


SEED_TAXONOMY = {
    "Clinical Practice Guideline / Consensus Statement": {
        "tool": "AGREE II",
        "full_name": "Appraisal of Guidelines for Research & Evaluation II",
        "purpose": "Assesses methodological rigour, stakeholder involvement, and clinical transparency.",
        "domains": [
            {"name": "Domain 1: Scope and Purpose", "question": "Are overall objectives, health questions, and target population explicitly described?"},
            {"name": "Domain 2: Stakeholder Involvement", "question": "Does the group include relevant professionals and have user/target views been sought?"},
            {"name": "Domain 3: Rigour of Development", "question": "Were systematic search methods used, criteria explicit, and an update protocol documented?"},
            {"name": "Domain 4: Clarity of Presentation", "question": "Are recommendations specific, unambiguous, and easily identifiable?"},
            {"name": "Domain 5: Applicability", "question": "Does it describe facilitators/barriers, execution tools, and monitoring criteria?"},
            {"name": "Domain 6: Editorial Independence", "question": "Did funding bodies influence content, and were competing interests addressed?"},
        ],
        "scale": "1 (Strongly Disagree / Absent) to 7 (Strongly Agree / Exemplary). Standardized % >= 70% required for high quality.",
    },
    "Randomized Controlled Trial (RCT)": {
        "tool": "Cochrane RoB 2",
        "full_name": "Cochrane Risk of Bias Tool for Randomized Trials (Version 2)",
        "purpose": "Evaluates internal validity and potential systematic errors in trial design and conduct.",
        "domains": [
            {"name": "Domain 1: Randomization process", "question": "Was allocation random, concealed until assignment, and baseline differences balanced?"},
            {"name": "Domain 2: Deviations from intended interventions", "question": "Were participants/caregivers blinded, and was an ITT/modified-ITT analysis executed?"},
            {"name": "Domain 3: Missing outcome data", "question": "Were outcome data complete or is there evidence that attrition was non-informative?"},
            {"name": "Domain 4: Measurement of the outcome", "question": "Was the outcome assessor fully blinded and measurement instruments validated?"},
            {"name": "Domain 5: Selection of the reported result", "question": "Was the trial registered prospectively with an analysis plan finalized prior to unblinded analysis?"},
        ],
        "scale": "Low Risk | Some Concerns | High Risk. A single High Risk domain mandates an overall High Risk rating.",
    },
    "Non-Randomized Intervention / Quasi-Experimental": {
        "tool": "ROBINS-I",
        "full_name": "Risk Of Bias In Non-randomized Studies - of Interventions",
        "purpose": "Assesses risk of bias in comparative studies comparing intervention groups without true randomization.",
        "domains": [
            {"name": "Domain 1: Confounding", "question": "Did the study adjust for all critical baseline and time-varying confounders?"},
            {"name": "Domain 2: Selection of participants", "question": "Was selection into cohort free from immortal time bias and independent of outcomes?"},
            {"name": "Domain 3: Classification of interventions", "question": "Were interventions clearly defined prospectively without misclassification?"},
            {"name": "Domain 4: Deviations from intended interventions", "question": "Were there systematic deviations or co-interventions unbalanced between groups?"},
            {"name": "Domain 5: Missing data", "question": "Were attrition rates documented with sensitivity analyses or imputation?"},
            {"name": "Domain 6: Measurement of outcomes", "question": "Were outcome assessors blinded to exposure status with validated metrics?"},
            {"name": "Domain 7: Selection of reported result", "question": "Are published results consistent with pre-registered protocol without selective p-hacking?"},
        ],
        "scale": "Low Risk | Moderate Risk | Serious Risk | Critical Risk | No Information.",
    },
    "Systematic Review / Meta-Analysis": {
        "tool": "AMSTAR 2",
        "full_name": "A MeaSurement Tool to Assess systematic Reviews (Version 2)",
        "purpose": "Critically appraises systematic reviews evaluating randomized or non-randomized healthcare studies.",
        "domains": [
            {"name": "Domain 1: PICO & Protocol Pre-registration", "question": "Did review establish PICO and register protocol (e.g. PROSPERO) prior to screening?"},
            {"name": "Domain 2: Literature Search Strategy", "question": "Did authors search >=2 databases, provide full boolean search strings, and trial registries?"},
            {"name": "Domain 3: Duplicate Study Selection & Extraction", "question": "Did at least two reviewers independently screen titles, abstracts, and full texts?"},
            {"name": "Domain 4: Excluded Studies Justification", "question": "Did authors provide full bibliographic list of excluded studies with explicit rationale?"},
            {"name": "Domain 5: Primary Study RoB Assessment", "question": "Did the review apply established, validated risk of bias tools to included studies?"},
            {"name": "Domain 6: Meta-Analytic Synthesis Appropriateness", "question": "Were statistical pooling methods justified and statistical heterogeneity investigated?"},
            {"name": "Domain 7: Publication Bias Assessment", "question": "Did authors formally test for publication bias (funnel plots, Egger's test)?"},
            {"name": "Domain 8: Conflict of Interest & Funding", "question": "Did authors report their own funding/COIs and funding of primary included studies?"},
        ],
        "scale": "High Confidence | Moderate Confidence | Low Confidence | Critically Low Confidence.",
    },
    "Observational Study (Cohort or Case-Control)": {
        "tool": "Newcastle-Ottawa Scale (NOS) / CASP",
        "full_name": "Newcastle-Ottawa Quality Assessment Scale / CASP Cohort Tool",
        "purpose": "Evaluates selection, comparability, and exposure/outcome assessment in observational epidemiology.",
        "domains": [
            {"name": "Domain 1: Selection of Cohorts / Cases", "question": "Are cohorts representative of community cases with exposure ascertained objectively?"},
            {"name": "Domain 2: Comparability of Groups", "question": "Did multivariable statistical models adjust for primary confounders and key covariates?"},
            {"name": "Domain 3: Exposure and Outcome Ascertainment", "question": "Was outcome assessment blinded or linkage-derived, with follow-up retention >80%?"},
        ],
        "scale": "0 to 9 Stars: Good Quality (7-9 stars) | Fair Quality (4-6 stars) | Poor Quality (0-3 stars).",
    },
    "Diagnostic Accuracy Study": {
        "tool": "QUADAS-2",
        "full_name": "Quality Assessment of Diagnostic Accuracy Studies 2",
        "purpose": "Evaluates risk of bias and applicability concerns in primary diagnostic accuracy studies.",
        "domains": [
            {"name": "Domain 1: Patient Selection", "question": "Was a consecutive or random sample enrolled without inappropriate exclusions?"},
            {"name": "Domain 2: Index Test Execution & Interpretation", "question": "Were index test results interpreted blinded to reference standard with pre-specified cutoffs?"},
            {"name": "Domain 3: Reference Standard Validity", "question": "Is reference standard true gold-standard and interpreted blinded to index test?"},
            {"name": "Domain 4: Flow and Timing", "question": "Was time interval appropriate, and did all patients receive verified reference standard?"},
        ],
        "scale": "Low Risk | High Risk | Unclear Risk.",
    },
    "Narrative Review": {
        "tool": "SANRA",
        "full_name": "Scale for the Assessment of Narrative Review Articles",
        "purpose": "Evaluates quality and presentation standard of non-systematic scientific reviews.",
        "domains": [
            {"name": "Domain 1: Justification of article's importance", "question": "Is clinical/scientific rationale explicitly justified with background epidemiology?"},
            {"name": "Domain 2: Statement of concrete aims", "question": "Does manuscript state explicit questions or testable objective hypotheses?"},
            {"name": "Domain 3: Description of literature search", "question": "Are search terms, databases, and date restrictions transparently described?"},
            {"name": "Domain 4: Referencing of primary literature", "question": "Are key claims substantiated by original experimental studies rather than tertiary opinions?"},
            {"name": "Domain 5: Scientific reasoning & balance", "question": "Does review present contradictory findings and neutral debate without advocacy bias?"},
            {"name": "Domain 6: Presentation of relevant data", "question": "Are numerical endpoint values, trial sample sizes, and comparative effect estimates tabulated?"},
        ],
        "scale": "0 = Not Met | 1 = Partially Met | 2 = Met. Total score out of 12 (>=9 required for high quality).",
    },
    "Qualitative Study": {
        "tool": "CASP Qualitative / COREQ",
        "full_name": "Critical Appraisal Skills Programme (CASP) Qualitative Checklist",
        "purpose": "Assesses credibility, dependability, and transferability of qualitative health inquiries.",
        "domains": [
            {"name": "Domain 1: Research design appropriateness", "question": "Is qualitative inquiry justified over quantitative measurement for this research question?"},
            {"name": "Domain 2: Sampling and recruitment strategy", "question": "Was theoretical/purposive sampling described, and non-participants explained?"},
            {"name": "Domain 3: Data collection and reflexivity", "question": "Did investigators critically evaluate researcher positionality and potential interaction bias?"},
            {"name": "Domain 4: Ethical considerations addressed", "question": "Were ethics committee approval, informed consent, and anonymization explicitly reported?"},
            {"name": "Domain 5: Rigour and depth of data analysis", "question": "Was thematic saturation verified, and was duplicate independent coding conducted?"},
            {"name": "Domain 6: Clear statement of findings and value", "question": "Are emergent themes backed by direct participant quotes with contextualized clinical value?"},
        ],
        "scale": "Yes | Can't Tell | No.",
    },
}


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
        conn.execute("""
            CREATE TABLE IF NOT EXISTS custom_tools (
                design_name TEXT PRIMARY KEY,
                tool_data TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor = conn.cursor()
        cursor.execute("SELECT count(*) as cnt FROM custom_tools")
        if cursor.fetchone()["cnt"] == 0:
            for design_name, tool_obj in SEED_TAXONOMY.items():
                cursor.execute(
                    "INSERT OR IGNORE INTO custom_tools (design_name, tool_data) VALUES (?, ?)",
                    (design_name, json.dumps(tool_obj)),
                )
        conn.commit()


def get_all_tools() -> Dict[str, dict]:
    with get_connection() as conn:
        cursor = conn.cursor()
        cursor.execute("SELECT design_name, tool_data FROM custom_tools ORDER BY rowid ASC")
        return {row["design_name"]: json.loads(row["tool_data"]) for row in cursor.fetchall()}


def save_discovered_tool(design_name: str, tool_obj: dict):
    with get_connection() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO custom_tools (design_name, tool_data) VALUES (?, ?)",
            (design_name, json.dumps(tool_obj)),
        )
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
        cursor.execute("SELECT full_json FROM appraisals WHERE id = ?", (appraisal_id,))
        row = cursor.fetchone()
        return json.loads(row["full_json"]) if row else None


def delete_appraisal(appraisal_id: int):
    with get_connection() as conn:
        conn.execute("DELETE FROM appraisals WHERE id = ?", (appraisal_id,))
        conn.commit()


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        extracted = []
        for idx, page in enumerate(reader.pages):
            text = page.extract_text() or ""
            if text.strip():
                extracted.append(f"--- [Page {idx + 1}] ---\n{text.strip()}")
        full_text = "\n\n".join(extracted)
        return full_text if len(full_text.strip()) > 150 else ""
    except Exception:
        return ""


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
    return article.get_text(separator="\n", strip=True) if article else ""


def classify_study_design_heuristically(text: str) -> str:
    header_text = text[:10000].lower()

    patterns = [
        (r"clinical\s+practice\s+guideline|consensus\s+statement|delphi\s+consensus", "Clinical Practice Guideline / Consensus Statement"),
        (r"systematic\s+review|meta-analysis|prisma|cochrane\s+database", "Systematic Review / Meta-Analysis"),
        (r"randomi[zs]ed\s+controlled\s+trial|double-blind|placebo-controlled|\brct\b", "Randomized Controlled Trial (RCT)"),
        (r"diagnostic\s+accuracy|sensitivity\s+and\s+specificity|index\s+test", "Diagnostic Accuracy Study"),
        (r"retrospective|prospective|cohort\s+study|case-control|cross-sectional|observational|comparative\s+study|prevalence", "Observational Study (Cohort or Case-Control)"),
        (r"qualitative\s+study|thematic\s+analysis|semi-structured\s+interview|grounded\s+theory", "Qualitative Study"),
        (r"quasi-experimental|non-randomi[zs]ed|controlled\s+before-and-after", "Non-Randomized Intervention / Quasi-Experimental"),
        (r"narrative\s+review|scoping\s+review|literature\s+review", "Narrative Review"),
    ]
    for p, design in patterns:
        if re.search(p, header_text):
            return design

    return "🔍 Auto-Detect & Discover Tool (If not in list)"


def trim_manuscript_body(text: str) -> str:
    lower = text.lower()
    for marker in ["\nreferences\n", "\nbibliography\n", "\nliterature cited\n"]:
        pos = lower.rfind(marker)
        if pos != -1 and pos > (len(text) * 0.45):
            return text[:pos]
    return text


def clean_tokens(s: str) -> List[str]:
    s = s.lower()
    s = re.sub(r"[\u2018\u2019\u201c\u201d]", "'", s)
    s = re.sub(r"[\u2013\u2014-]", " ", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return s.split()


def fuzzy_quote_match(quote: str, source_text: str, threshold: float = 0.82) -> Tuple[bool, float]:
    quote_tokens = clean_tokens(quote)
    if not quote_tokens:
        return True, 1.0

    q_len = len(quote_tokens)
    q_str = " ".join(quote_tokens)
    source_clean = " ".join(clean_tokens(source_text))
    if q_str in source_clean:
        return True, 1.0

    source_tokens = clean_tokens(source_text)
    if not source_tokens:
        return False, 0.0

    if q_len < 5:
        overlap = len(set(quote_tokens) & set(source_tokens)) / float(len(set(quote_tokens)))
        return (overlap >= threshold), overlap

    best_ratio = 0.0
    step = max(1, int(q_len * 0.35))
    for i in range(0, len(source_tokens) - q_len + 1, step):
        window = " ".join(source_tokens[i : i + q_len])
        if abs(len(window) - len(q_str)) > len(q_str) * 0.4:
            continue
        ratio = difflib.SequenceMatcher(None, q_str, window).quick_ratio()
        if ratio > best_ratio:
            best_ratio = ratio
        if best_ratio >= threshold:
            return True, best_ratio

    return (best_ratio >= threshold), best_ratio


def run_comprehensive_checksum(
    data: dict, full_raw_text: str, fuzzy_threshold: float = 0.82
) -> Tuple[dict, float, List[str], List[str]]:
    grounding_warnings: List[str] = []
    flaw_warnings: List[str] = []

    if not full_raw_text or not full_raw_text.strip():
        return data, 100.0, ["Source text unavailable."], ["Cannot verify integrity without source text."]

    domains = data.get("domain_appraisals", [])
    if not domains:
        return data, 100.0, [], ["No domain appraisals populated."]

    verified_count = 0
    for d in domains:
        quote = d.get("verbatim_evidence", "").strip()
        if not quote or quote.lower() in ["not reported", "none", "n/a", "not specified", "unclear"]:
            d["checksum_status"] = "Verified (Not Applicable)"
            verified_count += 1
            continue

        is_match, ratio = fuzzy_quote_match(quote, full_raw_text, threshold=fuzzy_threshold)
        if is_match:
            d["checksum_status"] = f"Verified Grounded ({int(ratio * 100)}% Match)"
            verified_count += 1
        else:
            d["checksum_status"] = f"Unverified Grounding ({int(ratio * 100)}% Max Similarity)"
            grounding_warnings.append(f"Quote in '{d.get('domain_name')}' matched at {int(ratio * 100)}%.")

    grounded_score = round((verified_count / len(domains)) * 100, 1)
    tier = str(data.get("quality_tier", "")).upper()
    has_high_risk = any("high" in str(d.get("rating", "")).lower() or "critical" in str(d.get("rating", "")).lower() for d in domains)

    if "GREEN" in tier and has_high_risk:
        flaw_warnings.append("Methodological Contradiction: Green tier assigned despite high-risk domain.")

    data["groundedness_score"] = grounded_score
    data["groundedness_warnings"] = grounding_warnings
    data["checksum_flaws"] = flaw_warnings
    data["checksum_passed"] = len(flaw_warnings) == 0 and grounded_score >= 70.0

    return data, grounded_score, grounding_warnings, flaw_warnings


def format_harvard_filename(data: dict, extension: str) -> str:
    raw_authors = str(data.get("authors", "")).strip()
    year = str(data.get("year_published", "")).strip()
    tool = str(data.get("appraisal_tool_selected", "Tool")).strip()

    year_match = re.search(r"\b(19\d\d|20\d\d)\b", year)
    clean_year = year_match.group(1) if year_match else "UnknownYear"

    if not raw_authors or raw_authors.lower() in ["not specified", "unspecified", "unknown"]:
        author_part = "AuthorUnknown"
    else:
        has_multiple = any(marker in raw_authors.lower() for marker in ["et al", " and ", "&", ","])
        first_segment = re.split(r"[,;&]|\sand\s|\set\sal", raw_authors, flags=re.IGNORECASE)[0].strip()
        tokens = [t for t in re.split(r"\s+", first_segment) if re.search(r"[a-zA-Z]", t)]
        surname = tokens[-1] if (len(tokens) > 1 and len(tokens[0]) <= 2) else (tokens[0] if tokens else "Author")
        surname = re.sub(r"[^a-zA-Z0-9]", "", surname)
        author_part = f"{surname}_et_al" if has_multiple else surname

    clean_tool = re.sub(r"[^a-zA-Z0-9_-]", "_", tool)
    clean_author = re.sub(r"[^a-zA-Z0-9_-]", "_", author_part)
    clean_filename = f"{clean_author}_{clean_year}_{clean_tool}_KAT.{extension.lstrip('.')}"
    return re.sub(r"_+", "_", clean_filename)


def discover_appraisal_tool(client: genai.Client, text_sample: str) -> Tuple[str, dict]:
    discovery_instruction = """
    Examine research text and return JSON:
    {
      "design_name": "Study Architecture Name",
      "tool_taxonomy": {
        "tool": "Tool Acronym",
        "full_name": "Full Tool Name",
        "purpose": "Precise methodological purpose",
        "domains": [{"name": "Domain Title", "question": "Criterion question"}],
        "scale": "Scoring thresholds and rules"
      }
    }
    """
    response = client.models.generate_content(
        model="gemini-3.5-flash-lite",
        contents=[f"MANUSCRIPT EXCERPT:\n{text_sample[:8000]}\n\nIdentify study architecture and gold-standard appraisal tool."],
        config=types.GenerateContentConfig(
            system_instruction=discovery_instruction,
            response_mime_type="application/json",
            temperature=0.0,
        ),
    )
    res_json = json.loads(response.text)
    return res_json["design_name"], res_json["tool_taxonomy"]


def execute_manuscript_evaluation(
    client: genai.Client,
    text_content: Optional[str] = None,
    raw_pdf_bytes: Optional[bytes] = None,
    assigned_design: Optional[str] = None,
    active_tools_map: Optional[dict] = None,
    strictness_mode: str = "Strict",
) -> dict:
    active_tool_taxonomy = {assigned_design: active_tools_map[assigned_design]} if (assigned_design and active_tools_map and assigned_design in active_tools_map) else (active_tools_map or SEED_TAXONOMY)
    taxonomy_json = json.dumps(active_tool_taxonomy, indent=2)

    system_instruction = f"""
    You are an unyielding epidemiological auditor. Apply zero benefit of doubt.
    Appraisal Framework:
    {taxonomy_json}

    RULES:
    1. ABSENCE OF REPORTING IS PRESENCE OF BIAS.
    2. Extract exact verbatim evidence quotes for every single domain.
    3. Return quality_tier strictly as 'GREEN', 'YELLOW', or 'RED'.

    CRITICAL: Output ONLY a valid JSON object matching:
    {{
      "paper_title": "Identified full title",
      "authors": "Author list",
      "journal": "Journal Name",
      "year_published": "YYYY",
      "doi": "DOI or URL",
      "study_design": "{assigned_design or 'Identified design'}",
      "appraisal_tool_selected": "Tool name",
      "quality_tier": "GREEN / YELLOW / RED",
      "overall_verdict": "Low / Moderate / High Risk",
      "overall_synthesis": "Methodological critique",
      "tool_meta": {{"acronym": "Tool", "full_name": "Full Name", "purpose": "Purpose", "scoring_rules": "Rules"}},
      "domain_appraisals": [
        {{"domain_name": "Title", "evaluation_question": "Question", "rating": "Rating", "location_reference": "Section/Page", "verbatim_evidence": "Quote", "critique": "Critique"}}
      ]
    }}
    """

    clean_text = trim_manuscript_body(text_content) if text_content else ""
    contents = [f"PAPER CONTENT:\n{clean_text}\n\nExecute appraisal and return JSON."] if (clean_text and len(clean_text.strip()) > 150) else (
        [types.Part.from_bytes(data=raw_pdf_bytes, mime_type="application/pdf"), f"Appraise for {assigned_design}."] if raw_pdf_bytes else None
    )
    if not contents:
        raise ValueError("No valid manuscript text or PDF bytes supplied.")

    for model_name in ["gemini-3.5-flash-lite", "gemini-2.5-flash-lite", "gemini-3.8-flash", "gemini-2.0-flash"]:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=types.GenerateContentConfig(
                    system_instruction=system_instruction,
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            return json.loads(response.text)
        except Exception:
            time.sleep(1.0)
            continue

    raise RuntimeError("Appraisal pipeline exceeded retry threshold.")


def answer_interrogation_query(
    client: genai.Client,
    query: str,
    appraisal_json: dict,
    raw_manuscript_text: str,
    chat_history: List[dict]
) -> str:
    history_context = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in chat_history[-6:]])
    chat_instruction = f"""
    You are an objective methodological audit assistant.
    Answer user queries regarding the appraisal report and source manuscript.
    Cite specific lines, sections, and quotes. Do not mention external technology.

    APPRAISAL JSON:
    {json.dumps(appraisal_json, indent=1)}

    SOURCE TEXT:
    {raw_manuscript_text[:20000]}
    """
    response = client.models.generate_content(
        model="gemini-3.5-flash-lite",
        contents=[f"HISTORY:\n{history_context}\n\nUSER QUESTION: {query}"],
        config=types.GenerateContentConfig(system_instruction=chat_instruction, temperature=0.1)
    )
    return response.text


def build_docx_report(data: dict) -> io.BytesIO:
    doc = Document()
    title = doc.add_heading(level=0)
    run_t = title.add_run(f"{APP_NAME} — Methodological Evaluation")
    run_t.font.size = Pt(20)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(120, 187, 166)

    meta_table = doc.add_table(rows=7, cols=2)
    meta_table.style = "Table Grid"
    fields = [
        ("Evaluated Title:", data.get("paper_title", "N/A")),
        ("Authors:", data.get("authors", "Not Specified")),
        ("Journal & Year:", f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})"),
        ("DOI / Reference:", data.get("doi", "N/A")),
        ("Study Architecture:", data.get("study_design", "N/A")),
        ("Gold-Standard Tool Applied:", data.get("appraisal_tool_selected", "N/A")),
        ("Evidence Groundedness:", f"{data.get('groundedness_score', 100.0)}%"),
    ]
    for idx, (label, val) in enumerate(fields):
        row = meta_table.rows[idx]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EBF5F2"/>')
        row.cells[0]._tc.get_or_add_tcPr().append(shd)
        row.cells[0].paragraphs[0].add_run(label).bold = True
        row.cells[1].paragraphs[0].add_run(str(val))

    doc.add_paragraph("")
    doc.add_heading(f"1. Detailed Domain Appraisal ({data.get('appraisal_tool_selected', 'Tool')})", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, h in enumerate(["Domain & Criterion", "Score", "Evidence", "Critique"]):
        cell = table.rows[0].cells[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="78BBA6"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        r = cell.paragraphs[0].add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for item in data.get("domain_appraisals", []):
        row_cells = table.add_row().cells
        row_cells[0].paragraphs[0].add_run(f"{item.get('domain_name')}\n\n\"{item.get('evaluation_question')}\"")
        row_cells[1].paragraphs[0].add_run(str(item.get("rating")))
        row_cells[2].paragraphs[0].add_run(f"[{item.get('location_reference')}]\n\"{item.get('verbatim_evidence')}\"")
        row_cells[3].paragraphs[0].add_run(str(item.get("critique")))

    doc.add_paragraph("")
    doc.add_heading("2. Executive Synthesis", level=1)
    doc.add_paragraph(data.get("overall_synthesis", ""))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_pdf_report(data: dict) -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle("T", parent=styles["Heading1"], fontSize=18, leading=22, textColor=colors.HexColor("#4A8574"))
    body_style = ParagraphStyle("B", parent=styles["Normal"], fontSize=8.5, leading=11.5)
    bold_style = ParagraphStyle("BB", parent=body_style, fontName="Helvetica-Bold")
    header_style = ParagraphStyle("H", parent=body_style, fontName="Helvetica-Bold", textColor=colors.white)

    story = [
        Paragraph(f"{APP_NAME} — Methodological Evaluation", title_style),
        Spacer(1, 10),
    ]

    meta_data = [
        [Paragraph("Evaluated Title:", bold_style), Paragraph(data.get("paper_title", "N/A"), body_style)],
        [Paragraph("Authors:", bold_style), Paragraph(data.get("authors", "Not Specified"), body_style)],
        [Paragraph("Journal & Year:", bold_style), Paragraph(f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})", body_style)],
        [Paragraph("Tool Applied:", bold_style), Paragraph(data.get("appraisal_tool_selected", "N/A"), body_style)],
        [Paragraph("Verdict:", bold_style), Paragraph(data.get("overall_verdict", "N/A"), bold_style)],
    ]
    t_meta = Table(meta_data, colWidths=[1.8 * inch, 5.6 * inch])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EBF5F2')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 12))

    story.append(Paragraph("1. Domain Appraisal", styles["Heading2"]))
    domain_data = [[Paragraph("Domain & Criterion", header_style), Paragraph("Score", header_style), Paragraph("Evidence", header_style), Paragraph("Critique", header_style)]]
    for d in data.get("domain_appraisals", []):
        domain_data.append([
            Paragraph(f"<b>{d.get('domain_name')}</b><br/><i>{d.get('evaluation_question')}</i>", body_style),
            Paragraph(f"<b>{d.get('rating')}</b>", bold_style),
            Paragraph(f"[{d.get('location_reference')}]<br/>\"{d.get('verbatim_evidence')}\"", body_style),
            Paragraph(d.get("critique", ""), body_style),
        ])
    t_domains = Table(domain_data, colWidths=[2.2 * inch, 0.9 * inch, 2.3 * inch, 2.0 * inch], repeatRows=1)
    t_domains.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#78BBA6')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_domains)
    story.append(Spacer(1, 12))

    story.append(Paragraph("2. Synthesis", styles["Heading2"]))
    story.append(Paragraph(data.get("overall_synthesis", ""), body_style))

    doc.build(story)
    buf.seek(0)
    return buf


st.set_page_config(page_title=APP_NAME, page_icon="🌿", layout="wide")
init_db()

# Pastel Seafoam Green UI Palette
st.markdown(
    """
<style>
    /* Global Buttons */
    .stButton>button, .stDownloadButton>button {
        background-color: #78BBA6 !important;
        color: white !important;
        border: none !important;
        border-radius: 8px !important;
        font-weight: 600 !important;
        transition: all 0.2s ease-in-out !important;
    }
    .stButton>button:hover, .stDownloadButton>button:hover {
        background-color: #5FA38E !important;
        color: #FFFFFF !important;
        box-shadow: 0 4px 10px rgba(120, 187, 166, 0.35) !important;
    }
    .stButton>button:active, .stDownloadButton>button:active {
        background-color: #4A8574 !important;
    }
    /* Chat Input Button & Focus */
    div[data-testid="stChatInput"] button {
        background-color: #78BBA6 !important;
        color: white !important;
    }
    div[data-testid="stChatInput"] textarea:focus {
        border-color: #78BBA6 !important;
        box-shadow: 0 0 0 1px #78BBA6 !important;
    }
    /* Accent Cards & Badges */
    .metric-card {
        background-color: rgba(120, 187, 166, 0.08);
        border: 1px solid rgba(120, 187, 166, 0.3);
        border-radius: 8px;
        padding: 12px;
    }
    /* Radio & Checkboxes */
    div[role="radiogroup"] > label > div:first-child[aria-checked="true"] {
        background-color: #78BBA6 !important;
        border-color: #78BBA6 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

CORRECT_PASSCODE = "324598"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title(f"🔒 {APP_NAME}")
    passcode_input = st.text_input("Access Passcode", type="password")
    if st.button("Unlock Appraisal Workspace"):
        if passcode_input.strip() == CORRECT_PASSCODE:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Invalid passcode.")
    st.stop()

st.title(f"🌿 {APP_NAME}")
st.caption("Methodological Evidence Quality Synthesis & Diagnostic Review Engine")

api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", None))

if "history_limit" not in st.session_state:
    st.session_state["history_limit"] = 10
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []
if "raw_paper_text" not in st.session_state:
    st.session_state["raw_paper_text"] = ""

with st.sidebar:
    st.header("⚙️ Configuration")
    if not api_key:
        api_key = st.text_input("Service Access Key", type="password")
    else:
        st.success("Access Credential Active")

    strictness_choice = st.radio(
        "Adjudication Strictness:",
        options=["Strict (Zero Benefit of Doubt)", "Auditor / Forensic (Zero Tolerance)"],
        index=0,
    )

    if st.button("Log Out"):
        st.session_state["authenticated"] = False
        st.rerun()

    st.markdown("---")
    st.header("🗄️ Past Appraisals")
    records = get_all_appraisals()
    if not records:
        st.caption("No saved records.")
    else:
        for r in records[:st.session_state["history_limit"]]:
            short_title = r["paper_title"][:34] + "..." if len(r["paper_title"]) > 34 else r["paper_title"]
            st.markdown(f"**{short_title}**\n<small>{r['created_at'][:10]} · {r['tool_selected']}</small>", unsafe_allow_html=True)
            c_load, c_del = st.columns([3, 1])
            with c_load:
                if st.button("Load", key=f"v_{r['id']}", use_container_width=True):
                    st.session_state["appraisal_data"] = get_appraisal_by_id(r["id"])
                    st.session_state["chat_history"] = []
                    st.rerun()
            with c_del:
                if st.button("🗑️", key=f"d_{r['id']}", use_container_width=True):
                    delete_appraisal(r["id"])
                    st.rerun()
            st.markdown("<hr style='margin: 4px 0;'/>", unsafe_allow_html=True)

        if len(records) > 10:
            if st.session_state["history_limit"] == 10:
                if st.button("Show More (50)"):
                    st.session_state["history_limit"] = 50
                    st.rerun()
            else:
                if st.button("Show Less (10)"):
                    st.session_state["history_limit"] = 10
                    st.rerun()

tab_pdf, tab_url = st.tabs(["📤 Upload PDF", "🔗 Enter Link / DOI"])
paper_text = ""
raw_pdf_bytes = None
filename_tag = "Paper"

with tab_pdf:
    uploaded = st.file_uploader("Upload Manuscript PDF", type=["pdf"])
    if uploaded:
        raw_pdf_bytes = uploaded.read()
        filename_tag = uploaded.name
        with st.spinner("Extracting text..."):
            paper_text = extract_text_from_pdf_bytes(raw_pdf_bytes)
        if paper_text:
            st.session_state["raw_paper_text"] = paper_text
            st.success(f"Parsed {len(paper_text):,} characters successfully.")
        else:
            st.warning("Scanned document detected. System OCR active.")

with tab_url:
    url_input = st.text_input("Article Link or DOI")
    if st.button("Fetch Article"):
        if url_input.strip():
            with st.spinner("Retrieving text..."):
                try:
                    paper_text = extract_text_from_url(url_input.strip())
                    st.session_state["raw_paper_text"] = paper_text
                    filename_tag = "Web_Article"
                    st.success(f"Retrieved {len(paper_text):,} characters.")
                except Exception as e:
                    st.error(f"Error fetching URL: {e}")

active_tools = get_all_tools()
available_designs = list(active_tools.keys()) + ["🔍 Auto-Detect & Discover Tool (If not in list)"]
detected_design = classify_study_design_heuristically(paper_text) if paper_text else available_designs[-1]

st.markdown("---")
selected_design = st.selectbox(
    "Target Study Architecture & Framework:",
    options=available_designs,
    index=available_designs.index(detected_design) if detected_design in available_designs else (len(available_designs) - 1),
)

if st.button("⚡ Run Methodological Appraisal", type="primary", use_container_width=True):
    if not api_key:
        st.warning("Enter access key in sidebar.")
    elif not paper_text and not raw_pdf_bytes:
        st.warning("Upload PDF or enter DOI.")
    else:
        client = genai.Client(api_key=api_key)
        target_design = selected_design

        if target_design == "🔍 Auto-Detect & Discover Tool (If not in list)":
            with st.spinner("Discovering best appraisal tool..."):
                try:
                    discovered_name, discovered_tool = discover_appraisal_tool(client, paper_text or "Excerpt")
                    save_discovered_tool(discovered_name, discovered_tool)
                    target_design = discovered_name
                    active_tools = get_all_tools()
                    st.info(f"Adopted tool: **{discovered_tool.get('tool')}**")
                except Exception:
                    target_design = "Observational Study (Cohort or Case-Control)"

        with st.spinner("Executing rigorous appraisal..."):
            try:
                raw_data = execute_manuscript_evaluation(
                    client,
                    text_content=paper_text,
                    raw_pdf_bytes=raw_pdf_bytes if not paper_text else None,
                    assigned_design=target_design,
                    active_tools_map=active_tools,
                    strictness_mode=strictness_choice,
                )
                verified_data, score, _, flaw_warns = run_comprehensive_checksum(raw_data, paper_text)
                new_id = save_appraisal(verified_data, filename_tag)
                st.session_state["appraisal_data"] = verified_data
                st.session_state["chat_history"] = []
                st.success(f"Completed and saved (ID: #{new_id})")
                st.rerun()
            except Exception as e:
                st.error(f"Execution error: {e}")

if "appraisal_data" in st.session_state:
    data = st.session_state["appraisal_data"]
    st.markdown("---")
    st.subheader(f"📋 {data.get('paper_title', 'Untitled')}")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Architecture", data.get("study_design"))
    with c2:
        st.metric("Tool Applied", data.get("appraisal_tool_selected"))
    with c3:
        st.metric("Overall Verdict", data.get("overall_verdict"))
    with c4:
        st.metric("Groundedness", f"{data.get('groundedness_score', 100.0)}%")

    tier = str(data.get("quality_tier", "")).upper()
    g1, g2, g3 = st.columns(3)
    with g1:
        st.error("🚨 RED TIER (High Risk)" if tier == "RED" else "Red")
    with g2:
        st.warning("⚠️ YELLOW TIER (Moderate)" if tier == "YELLOW" else "Yellow")
    with g3:
        st.success("✅ GREEN TIER (Robust)" if tier == "GREEN" else "Green")

    st.subheader("📊 Domain Analysis")
    st.dataframe(
        [
            {
                "Domain": d.get("domain_name"),
                "Criterion": d.get("evaluation_question"),
                "Score": d.get("rating"),
                "Reference": d.get("location_reference"),
                "Quote": d.get("verbatim_evidence"),
                "Critique": d.get("critique"),
            }
            for d in data.get("domain_appraisals", [])
        ],
        use_container_width=True,
    )

    st.subheader("📝 Executive Summary")
    st.info(data.get("overall_synthesis"))

    c_doc, c_pdf = st.columns(2)
    with c_doc:
        st.download_button(
            "📥 Download Word (.docx)",
            data=build_docx_report(data),
            file_name=format_harvard_filename(data, "docx"),
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with c_pdf:
        st.download_button(
            "📥 Download PDF (.pdf)",
            data=build_pdf_report(data),
            file_name=format_harvard_filename(data, "pdf"),
            mime="application/pdf",
            use_container_width=True,
        )

    st.markdown("---")
    st.subheader("💬 Interrogation Assistant")
    for msg in st.session_state["chat_history"]:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    if query := st.chat_input("Ask about specific scores or cite line locations..."):
        if not api_key:
            st.error("Key required.")
        else:
            st.session_state["chat_history"].append({"role": "user", "content": query})
            with st.chat_message("user"):
                st.markdown(query)

            client = genai.Client(api_key=api_key)
            with st.chat_message("assistant"):
                with st.spinner("Checking manuscript text..."):
                    try:
                        reply = answer_interrogation_query(
                            client, query, data, st.session_state.get("raw_paper_text", ""), st.session_state["chat_history"]
                        )
                        st.markdown(reply)
                        st.session_state["chat_history"].append({"role": "assistant", "content": reply})
                    except Exception as err:
                        st.error(f"Query error: {err}")
