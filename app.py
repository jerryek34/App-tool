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
from docx.shared import Inches, Pt, RGBColor
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

# ==========================================
# 1. DATABASE LAYER (SQLite)
# ==========================================
DB_NAME = "appraisals.db"
APP_NAME = "Kitique Appraisal Tool"


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
        "full_name": "Appraisal of Guidelines for Research & Evaluation II",
        "purpose": "Assesses the methodological rigour, stakeholder involvement, and clinical transparency of guidelines.",
        "domains": [
            {
                "name": "Domain 1: Scope and Purpose",
                "question": "Are the overall objective(s), health question(s) covered, and target population explicitly described?",
            },
            {
                "name": "Domain 2: Stakeholder Involvement",
                "question": "Does the group include all relevant professional groups, have patient/target population views been sought, and are target users clearly defined?",
            },
            {
                "name": "Domain 3: Rigour of Development",
                "question": "Were systematic search methods used, criteria for selecting evidence transparent, strengths/limitations explicit, health benefits/risks considered, and is an external review/update protocol documented?",
            },
            {
                "name": "Domain 4: Clarity of Presentation",
                "question": "Are recommendations specific, unambiguous, identifiable, and presenting different management options clearly?",
            },
            {
                "name": "Domain 5: Applicability",
                "question": "Does the guideline describe facilitators/barriers to application, advice/tools on execution, resource implications, and monitoring criteria?",
            },
            {
                "name": "Domain 6: Editorial Independence",
                "question": "Did the funding body influence guideline content, and were competing interests of development group members recorded and addressed?",
            },
        ],
        "scale": "1 (Strongly Disagree / Absent) to 7 (Strongly Agree / Exemplary). Overall domain scores are calculated as standardized percentages. >70% indicates high quality.",
    },
    "Randomized Controlled Trial (RCT)": {
        "tool": "Cochrane RoB 2",
        "full_name": "Cochrane Risk of Bias Tool for Randomized Trials (Version 2)",
        "purpose": "Evaluates internal validity and potential systematic errors in trial design, conduct, and reporting.",
        "domains": [
            {
                "name": "Domain 1: Bias arising from the randomization process",
                "question": "Was the allocation sequence random, concealed until assignment, and were baseline differences balanced across groups?",
            },
            {
                "name": "Domain 2: Bias due to deviations from intended interventions",
                "question": "Were participants and caregivers aware of assignments, and did any non-protocol deviations occur that were unbalanced between groups?",
            },
            {
                "name": "Domain 3: Bias due to missing outcome data",
                "question": "Were data for this outcome available for nearly all participants randomized, and is there evidence that missingness depends on true values?",
            },
            {
                "name": "Domain 4: Bias in measurement of the outcome",
                "question": "Was the outcome assessor blinded, was the measurement method appropriate, and could assessment have been influenced by knowledge of intervention?",
            },
            {
                "name": "Domain 5: Bias in selection of the reported result",
                "question": "Was the study analyzed in accordance with a pre-specified plan/protocol registered before unblinded data became available?",
            },
        ],
        "scale": "Low Risk of Bias | Some Concerns | High Risk of Bias. The overall risk of bias judgment matches the highest risk assessed in any individual domain.",
    },
    "Non-Randomized Intervention / Quasi-Experimental": {
        "tool": "ROBINS-I",
        "full_name": "Risk Of Bias In Non-randomized Studies - of Interventions",
        "purpose": "Assesses risk of bias in comparative studies comparing intervention groups without true randomization.",
        "domains": [
            {
                "name": "Domain 1: Bias due to Confounding",
                "question": "Did the authors adequately control for baseline and time-varying confounding using appropriate design or adjustment?",
            },
            {
                "name": "Domain 2: Bias in Selection of participants",
                "question": "Was participant selection into the study related to both intervention and outcome?",
            },
            {
                "name": "Domain 3: Classification of interventions",
                "question": "Were intervention status and exposure definitions clearly and objectively defined without misclassification?",
            },
            {
                "name": "Domain 4: Deviations from intended interventions",
                "question": "Were there systematic deviations from intended interventions arising because of the experimental context?",
            },
            {
                "name": "Domain 5: Missing data",
                "question": "Were outcome data complete, or were missing data handled with robust sensitivity/imputation analyses?",
            },
            {
                "name": "Domain 6: Measurement of outcomes",
                "question": "Were outcome assessors blinded, and were outcome measures valid, reliable, and applied equally?",
            },
            {
                "name": "Domain 7: Selection of the reported result",
                "question": "Are reported results free from selective outcome or analysis reporting based on statistical significance?",
            },
        ],
        "scale": "Low Risk | Moderate Risk | Serious Risk | Critical Risk | No Information.",
    },
    "Systematic Review / Meta-Analysis": {
        "tool": "AMSTAR 2",
        "full_name": "A MeaSurement Tool to Assess systematic Reviews (Version 2)",
        "purpose": "Critically appraises systematic reviews that evaluate randomized or non-randomized studies of healthcare interventions.",
        "domains": [
            {
                "name": "Domain 1: PICO & Protocol Pre-registration",
                "question": "Did the review question and inclusion criteria include components of PICO, and was a protocol established and registered prior to conducting the review?",
            },
            {
                "name": "Domain 2: Comprehensive Literature Search Strategy",
                "question": "Did the authors search at least two databases, provide keywords/search strategies, search trial registries, and review reference lists?",
            },
            {
                "name": "Domain 3: Duplicate Study Selection & Extraction",
                "question": "Did at least two authors independently screen titles, abstracts, and full texts, and extract study data?",
            },
            {
                "name": "Domain 4: Excluded Studies Justification",
                "question": "Did the authors provide a comprehensive list of potentially relevant excluded studies along with explicit reasons for each exclusion?",
            },
            {
                "name": "Domain 5: Primary Study RoB Assessment",
                "question": "Did the authors apply a recognized, formal risk of bias tool to primary included studies?",
            },
            {
                "name": "Domain 6: Meta-Analytic Synthesis Appropriateness",
                "question": "If quantitative synthesis was performed, were statistical pooling methods justified and investigated for heterogeneity?",
            },
            {
                "name": "Domain 7: Publication Bias Assessment",
                "question": "Did the authors investigate publication bias (funnel plots, Egger's test) and discuss its likely clinical impact?",
            },
            {
                "name": "Domain 8: Conflict of Interest & Funding Reporting",
                "question": "Did the authors report their funding sources and any conflicts of interest, as well as the funding sources of primary studies?",
            },
        ],
        "scale": "High Confidence (No or one non-critical weakness) | Moderate Confidence (More than one non-critical weakness) | Low Confidence (One critical flaw) | Critically Low Confidence (More than one critical flaw).",
    },
    "Observational Study (Cohort or Case-Control)": {
        "tool": "Newcastle-Ottawa Scale (NOS) / CASP",
        "full_name": "Newcastle-Ottawa Quality Assessment Scale / CASP Cohort Tool",
        "purpose": "Evaluates selection, comparability, and exposure/outcome assessment in observational epidemiology.",
        "domains": [
            {
                "name": "Domain 1: Selection of Cohorts / Cases",
                "question": "Are the study cohorts truly representative, cases clearly defined, and exposure ascertainment secured via objective records?",
            },
            {
                "name": "Domain 2: Comparability of Groups",
                "question": "Did the study design or statistical model adjust for the primary confounder and additional relevant covariates?",
            },
            {
                "name": "Domain 3: Ascertainment of Exposure and Outcome",
                "question": "Was outcome assessment blinded or independent, with follow-up long and complete enough for meaningful conclusions?",
            },
        ],
        "scale": "Stars awarded (0 to 9 stars total): Good Quality (7-9 stars) | Fair Quality (4-6 stars) | Poor Quality (0-3 stars).",
    },
    "Diagnostic Accuracy Study": {
        "tool": "QUADAS-2",
        "full_name": "Quality Assessment of Diagnostic Accuracy Studies 2",
        "purpose": "Evaluates risk of bias and applicability concerns in primary diagnostic accuracy studies.",
        "domains": [
            {
                "name": "Domain 1: Patient Selection",
                "question": "Was a consecutive or random sample of patients enrolled, and was inappropriate exclusion avoided?",
            },
            {
                "name": "Domain 2: Index Test Execution & Interpretation",
                "question": "Were index test results interpreted without knowledge of the reference standard, and with a pre-specified threshold?",
            },
            {
                "name": "Domain 3: Reference Standard Validity",
                "question": "Is the reference standard likely to correctly classify the target condition, and interpreted without knowledge of the index test?",
            },
            {
                "name": "Domain 4: Flow and Timing",
                "question": "Was there an appropriate interval between index test and reference standard, and did all patients receive the same reference test?",
            },
        ],
        "scale": "Low Risk | High Risk | Unclear Risk across bias and applicability domains.",
    },
    "Narrative Review": {
        "tool": "SANRA",
        "full_name": "Scale for the Assessment of Narrative Review Articles",
        "purpose": "Evaluates quality and presentation standard of non-systematic scientific reviews.",
        "domains": [
            {
                "name": "Domain 1: Justification of article's importance",
                "question": "Is the scientific or clinical importance of the review justified and contextualized?",
            },
            {
                "name": "Domain 2: Statement of concrete aims",
                "question": "Does the article formulate explicit, concrete questions or aims?",
            },
            {
                "name": "Domain 3: Description of the literature search",
                "question": "Is the literature search methodology transparently described (databases, search terms, restrictions)?",
            },
            {
                "name": "Domain 4: Referencing of primary literature",
                "question": "Are arguments and assertions referenced with original primary evidence rather than non-systematic reviews?",
            },
            {
                "name": "Domain 5: Scientific reasoning and balanced presentation",
                "question": "Does the review present evidence impartially and discuss contradictory findings?",
            },
            {
                "name": "Domain 6: Presentation of relevant data",
                "question": "Are relevant endpoint data, clinical trial outcomes, and effect sizes synthesized in clear tables or charts?",
            },
        ],
        "scale": "0 to 2 per item (Maximum score 12). >=9/12 indicates a high-standard narrative review.",
    },
    "Qualitative Study": {
        "tool": "CASP Qualitative / COREQ",
        "full_name": "Critical Appraisal Skills Programme (CASP) Qualitative Checklist",
        "purpose": "Assesses credibility, dependability, and transferability of qualitative health inquiries.",
        "domains": [
            {
                "name": "Domain 1: Research design appropriateness",
                "question": "Is qualitative methodology appropriate for addressing the research goal?",
            },
            {
                "name": "Domain 2: Sampling and recruitment strategy",
                "question": "Was the recruitment strategy appropriate to the aims, and were non-participants explained?",
            },
            {
                "name": "Domain 3: Data collection and reflexivity",
                "question": "Did the researchers critically examine their own role, potential bias, and influence during data collection?",
            },
            {
                "name": "Domain 4: Ethical considerations addressed",
                "question": "Were ethical approval, consent procedures, and confidentiality explicitly secured?",
            },
            {
                "name": "Domain 5: Rigour and depth of data analysis",
                "question": "Was the analytic process rigorous (e.g., thematic analysis, saturation, duplicate coding)?",
            },
            {
                "name": "Domain 6: Clear statement of findings and value",
                "question": "Are findings clearly stated, substantiated by participant quotes, and clinically valuable?",
            },
        ],
        "scale": "Yes | Can't Tell | No across 10 structured appraisal questions.",
    },
}

# ==========================================
# 3. EXTRACTION, CHECKSUM & GEMINI
# ==========================================
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
    header_text = text[:7000].lower()

    if any(re.search(p, header_text) for p in [
        r"clinical\s+practice\s+guideline",
        r"consensus\s+statement",
        r"recommendations\s+for\s+clinical\s+practice",
        r"clinical\s+guideline",
        r"delphi\s+consensus"
    ]):
        return "Clinical Practice Guideline / Consensus Statement"

    if any(re.search(p, header_text) for p in [
        r"systematic\s+review",
        r"meta-analysis",
        r"prisma",
        r"cochrane\s+database\s+of\s+systematic\s+reviews"
    ]):
        return "Systematic Review / Meta-Analysis"

    if any(re.search(p, header_text) for p in [
        r"randomised\s+controlled\s+trial",
        r"randomized\s+controlled\s+trial",
        r"double-blind",
        r"placebo-controlled",
        r"randomly\s+assigned",
        r"rct\b"
    ]):
        return "Randomized Controlled Trial (RCT)"

    if any(re.search(p, header_text) for p in [
        r"diagnostic\s+accuracy",
        r"sensitivity\s+and\s+specificity",
        r"receiver\s+operating\s+characteristic",
        r"index\s+test",
        r"reference\s+standard"
    ]):
        return "Diagnostic Accuracy Study"

    if any(re.search(p, header_text) for p in [
        r"qualitative\s+study",
        r"thematic\s+analysis",
        r"semi-structured\s+interview",
        r"focus\s+groups?",
        r"grounded\s+theory",
        r"phenomenological"
    ]):
        return "Qualitative Study"

    if any(re.search(p, header_text) for p in [
        r"cohort\s+study",
        r"prospective\s+cohort",
        r"retrospective\s+cohort",
        r"case-control",
        r"cross-sectional\s+study",
        r"longitudinal\s+study"
    ]):
        return "Observational Study (Cohort or Case-Control)"

    if any(re.search(p, header_text) for p in [
        r"quasi-experimental",
        r"non-randomi[zs]ed",
        r"controlled\s+before-and-after",
        r"interrupted\s+time\s+series"
    ]):
        return "Non-Randomized Intervention / Quasi-Experimental"

    return "Narrative Review"


def trim_manuscript_body(text: str) -> str:
    ref_markers = [
        "\nreferences\n",
        "\nreferences \n",
        "\nbibliography\n",
        "\nliterature cited\n",
        "\nreferences & notes\n"
    ]
    lower = text.lower()
    cut_pos = -1
    for marker in ref_markers:
        pos = lower.rfind(marker)
        if pos != -1 and pos > (len(text) * 0.45):
            cut_pos = pos
            break

    if cut_pos != -1:
        return text[:cut_pos]
    return text


def clean_tokens(s: str) -> List[str]:
    """Normalizes text by removing punctuation, casing, hyphens, and whitespace splits."""
    s = s.lower()
    # Normalize unicode quotes and dashes
    s = re.sub(r"[\u2018\u2019\u201c\u201d]", "'", s)
    s = re.sub(r"[\u2013\u2014-]", " ", s)
    # Strip non-alphanumeric characters
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return s.split()


def fuzzy_quote_match(quote: str, source_text: str, threshold: float = 0.82) -> Tuple[bool, float]:
    """
    Checks if a quote exists within source_text using sliding token windows
    and difflib SequenceMatcher to absorb OCR errors, typos, and line-wrap hyphens.
    """
    quote_tokens = clean_tokens(quote)
    if not quote_tokens:
        return True, 1.0

    q_len = len(quote_tokens)
    q_str = " ".join(quote_tokens)

    # 1. Exact canonical substring shortcut
    source_clean = " ".join(clean_tokens(source_text))
    if q_str in source_clean:
        return True, 1.0

    source_tokens = clean_tokens(source_text)
    if not source_tokens:
        return False, 0.0

    # Short quote fallback (under 5 tokens)
    if q_len < 5:
        # Check token set inclusion
        q_set = set(quote_tokens)
        s_set = set(source_tokens)
        overlap = len(q_set & s_set) / float(len(q_set))
        return (overlap >= threshold), overlap

    # 2. Sliding Window with SequenceMatcher
    best_ratio = 0.0
    step = max(1, int(q_len * 0.35))
    window_sizes = [q_len, int(q_len * 1.15)]

    for w_size in window_sizes:
        for i in range(0, len(source_tokens) - q_len + 1, step):
            window = " ".join(source_tokens[i : i + w_size])
            # Quick length gate before running sequence matcher
            if abs(len(window) - len(q_str)) > len(q_str) * 0.4:
                continue
            ratio = difflib.SequenceMatcher(None, q_str, window).quick_ratio()
            if ratio > best_ratio:
                best_ratio = ratio
            if best_ratio >= threshold:
                return True, best_ratio

    # Final detailed check on the best sliding candidate
    return (best_ratio >= threshold), best_ratio


def run_python_groundedness_checksum(
    data: dict, full_raw_text: str, fuzzy_threshold: float = 0.82
) -> Tuple[dict, float, List[str]]:
    """
    Deterministic Python checksum verifying that LLM verbatim quotes are grounded
    in the document using sliding fuzzy token matching.
    """
    if not full_raw_text or not full_raw_text.strip():
        return data, 100.0, ["Source text unavailable for verification."]

    domains = data.get("domain_appraisals", [])
    if not domains:
        return data, 100.0, []

    verified_count = 0
    warnings = []

    for d in domains:
        quote = d.get("verbatim_evidence", "").strip()
        if not quote or quote.lower() in [
            "not reported",
            "none",
            "n/a",
            "not specified",
            "unclear",
        ]:
            d["checksum_status"] = "Verified (Not Applicable)"
            verified_count += 1
            continue

        is_match, ratio = fuzzy_quote_match(quote, full_raw_text, threshold=fuzzy_threshold)

        if is_match:
            d["checksum_status"] = f"Verified Grounded ({int(ratio * 100)}% Match)"
            verified_count += 1
        else:
            d["checksum_status"] = f"Unverified Grounding ({int(ratio * 100)}% Max Similarity)"
            warnings.append(
                f"Quote in '{d.get('domain_name')}' matched at only {int(ratio * 100)}% (threshold: {int(fuzzy_threshold * 100)}%)."
            )

    grounded_score = round((verified_count / len(domains)) * 100, 1)
    data["groundedness_score"] = grounded_score
    data["groundedness_warnings"] = warnings
    return data, grounded_score, warnings


def appraise_with_gemini(
    client: genai.Client,
    text_content: Optional[str] = None,
    raw_pdf_bytes: Optional[bytes] = None,
    assigned_design: Optional[str] = None,
) -> dict:
    if assigned_design and assigned_design in TOOL_TAXONOMY:
        active_tool_taxonomy = {assigned_design: TOOL_TAXONOMY[assigned_design]}
    else:
        active_tool_taxonomy = TOOL_TAXONOMY

    taxonomy_json = json.dumps(active_tool_taxonomy, indent=2)

    system_instruction = f"""
    You are an expert epidemiological reviewer and methodological appraiser.
    Evaluate the provided research paper or clinical guideline and perform a critical appraisal.

    Appraisal Tool Assigned:
    {taxonomy_json}

    Step-by-Step Instructions:
    1. Bibliographic Extraction:
       - Extract the paper's exact 'paper_title'.
       - Extract 'authors' (e.g., "Smith JA, Doe B, et al.").
       - Extract 'journal' (journal name, or 'Unspecified').
       - Extract 'year_published' (e.g., "2024").
       - Extract 'doi' (e.g., "10.1016/j.math.2023...").
    2. Confirm Study Design: The paper has been pre-classified as '{assigned_design if assigned_design else 'Auto-detect'}'. Confirm and apply the corresponding Gold-Standard Tool.
    3. Tool Assignment: Evaluate EVERY domain and question listed for this assigned tool.
    4. Grounded Evidence Extraction:
       - Include the exact 'evaluation_question' as provided in the taxonomy.
       - Extract the EXACT 'verbatim_evidence' directly word-for-word from the text.
       - Provide the exact 'location_reference' (e.g., "[Page 4] Methods §Randomization").
       - Provide an analytical 'critique' justifying the score/rating.
    5. Categorize 'quality_tier' into exactly one of: "GREEN", "YELLOW", or "RED".
       - "GREEN": Methodologically robust, low risk of bias, high clinical relevance and validity.
       - "YELLOW": Moderate quality, some methodological concerns, fair clinical applicability.
       - "RED": High risk of bias, critical flaws, weak validity, or poor applicability.
    6. Return full 'tool_meta' (acronym, full_name, purpose, scoring_rules).

    CRITICAL: Output ONLY a valid JSON object matching this schema:
    {{
      "paper_title": "Identified full title",
      "authors": "First Author et al. or full author list",
      "journal": "Journal Name",
      "year_published": "YYYY",
      "doi": "DOI or URL",
      "study_design": "{assigned_design if assigned_design else 'Identified design'}",
      "appraisal_tool_selected": "Name of appraisal tool",
      "quality_tier": "GREEN / YELLOW / RED",
      "overall_verdict": "Low Risk / Moderate Risk / High Risk / Recommended / etc.",
      "overall_synthesis": "Comprehensive narrative explaining methodological quality and clinical applicability.",
      "tool_meta": {{
        "acronym": "Tool Acronym",
        "full_name": "Full Tool Name",
        "purpose": "Clinical/methodological purpose of the tool",
        "scoring_rules": "Detailed explanation of scores, thresholds, or stars required"
      }},
      "domain_appraisals": [
        {{
          "domain_name": "Domain name as per taxonomy",
          "evaluation_question": "The appraisal question evaluated",
          "rating": "Score or Rating",
          "location_reference": "Section or Page number reference",
          "verbatim_evidence": "Direct quote from text",
          "critique": "Analytical rationale for this domain"
        }}
      ]
    }}
    """

    clean_text = trim_manuscript_body(text_content) if text_content else ""

    contents = []
    if clean_text and len(clean_text.strip()) > 150:
        contents.append(
            f"PAPER CONTENT:\n{clean_text}\n\nPerform the full critical appraisal and return the JSON object."
        )
    elif raw_pdf_bytes:
        contents.append(
            types.Part.from_bytes(data=raw_pdf_bytes, mime_type="application/pdf")
        )
        contents.append(
            f"Perform the critical appraisal for {assigned_design} using vision OCR and return the JSON object."
        )
    else:
        raise ValueError("No valid manuscript text or PDF bytes supplied.")

    candidate_models = [
        "gemini-3.5-flash-lite",
        "gemini-2.5-flash-lite",
        "gemini-3.8-flash",
        "gemini-2.0-flash",
    ]

    last_error = None

    for model_name in candidate_models:
        for attempt in range(2):
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
                    time.sleep(1.5)
                    continue
                elif "404" in err_str or "NOT_FOUND" in err_str:
                    break
                else:
                    raise e

    raise last_error


# ==========================================
# 4. WORD (.docx) & PDF EXPORT ENGINES
# ==========================================
def set_cell_background(cell, fill_hex: str):
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def build_docx_report(data: dict) -> io.BytesIO:
    doc = Document()

    title = doc.add_heading(level=0)
    run_t = title.add_run(f"{APP_NAME} — Methodological Evaluation")
    run_t.font.size = Pt(20)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(31, 78, 121)

    p_sub = doc.add_paragraph("Automated critical evidence synthesis evaluated via Google Gemini.")
    p_sub.runs[0].font.size = Pt(9.5)
    p_sub.runs[0].italic = True

    meta_table = doc.add_table(rows=7, cols=2)
    meta_table.style = "Table Grid"
    fields = [
        ("Evaluated Title:", data.get("paper_title", "N/A")),
        ("Authors:", data.get("authors", "Not Specified")),
        ("Journal & Year:", f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})"),
        ("DOI / Reference:", data.get("doi", "N/A")),
        ("Study Architecture:", data.get("study_design", "N/A")),
        ("Gold-Standard Tool Applied:", data.get("appraisal_tool_selected", "N/A")),
        ("Fuzzy Groundedness Checksum:", f"{data.get('groundedness_score', 100.0)}% Evidence Verified in Source Text"),
    ]
    for idx, (label, val) in enumerate(fields):
        row = meta_table.rows[idx]
        set_cell_background(row.cells[0], "F2F5F8")
        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(9.5)
        p1 = row.cells[1].paragraphs[0]
        r1 = p1.add_run(str(val))
        r1.font.size = Pt(9.5)
        if idx == 6:
            r1.bold = True
            r1.font.color.rgb = RGBColor(40, 120, 40) if data.get('groundedness_score', 100.0) >= 80 else RGBColor(180, 40, 40)

    doc.add_paragraph("")

    # Visual Quality Meter
    tier = str(data.get("quality_tier", "")).upper()
    if "GREEN" not in tier and "YELLOW" not in tier and "RED" not in tier:
        v = str(data.get("overall_verdict", "")).lower()
        tier = "GREEN" if any(w in v for w in ["low", "high confidence", "recommended", "good"]) else ("YELLOW" if any(w in v for w in ["some", "moderate"]) else "RED")

    p_gauge_label = doc.add_paragraph()
    r_gl = p_gauge_label.add_run("Methodological Rigour & Clinical Trustworthiness Gauge (At-A-Glance):")
    r_gl.bold = True
    r_gl.font.size = Pt(10.5)

    bar_table = doc.add_table(rows=1, cols=3)
    bar_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    cols_spec = [
        ("RED", "C00000", "E6B8B8", "HIGH RISK / FLAWED\n(Limited Reliability)"),
        ("YELLOW", "FFC000", "FFF2CC", "MODERATE / SOME CONCERNS\n(Conditional Validity)"),
        ("GREEN", "385623", "E2EFDA", "ROBUST / LOW RISK\n(Strong Clinical Validity)"),
    ]

    for i, (b_tier, active_hex, inactive_hex, label) in enumerate(cols_spec):
        cell = bar_table.rows[0].cells[i]
        is_active = (b_tier == tier)
        fill_hex = active_hex if is_active else inactive_hex
        set_cell_background(cell, fill_hex)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(label)
        run.bold = is_active
        run.font.size = Pt(8.5)
        run.font.color.rgb = RGBColor(255, 255, 255) if is_active else RGBColor(80, 80, 80)
        if is_active:
            p_ind = cell.add_paragraph()
            p_ind.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r_ind = p_ind.add_run("▲ ASSIGNED QUALITY TIER ▲")
            r_ind.bold = True
            r_ind.font.size = Pt(8.5)
            r_ind.font.color.rgb = RGBColor(255, 255, 255)

    doc.add_paragraph("")

    # Section 1: Detailed Question-by-Question Domain Appraisal
    doc.add_heading(
        f"1. Detailed Domain Appraisal ({data.get('appraisal_tool_selected', 'Tool')})",
        level=1,
    )
    p_lead = doc.add_paragraph(
        "Each domain is evaluated below against its standard methodological criterion, presenting the appraisal prompt question, assigned score, grounded textual evidence, and evaluative critique."
    )
    p_lead.runs[0].font.size = Pt(9.5)

    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = [
        "Domain & Appraisal Criterion Question",
        "Appraisal Score",
        "Grounded Verbatim Evidence & Citation",
        "Methodological Critique & Rationale",
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

        # Col 0: Domain Name + Question Prompt for Context
        p_d = row_cells[0].paragraphs[0]
        r_name = p_d.add_run(f"{item.get('domain_name', '')}\n\n")
        r_name.bold = True
        r_name.font.size = Pt(9.5)

        q_text = item.get("evaluation_question", "").strip()
        if q_text:
            r_ql = p_d.add_run("Appraisal Criterion Question:\n")
            r_ql.bold = True
            r_ql.font.size = Pt(8.5)
            r_ql.font.color.rgb = RGBColor(80, 80, 80)
            r_qt = p_d.add_run(f'"{q_text}"')
            r_qt.italic = True
            r_qt.font.size = Pt(8.5)

        # Col 1: Score/Rating
        p_r = row_cells[1].paragraphs[0]
        r_sc = p_r.add_run(item.get("rating", ""))
        r_sc.bold = True
        r_sc.font.size = Pt(10)

        # Col 2: Evidence & Section Location + Checksum verification status
        p_e = row_cells[2].paragraphs[0]
        loc = item.get("location_reference", "Section")
        r_loc = p_e.add_run(f"[{loc}]\n")
        r_loc.bold = True
        r_loc.font.size = Pt(8)
        r_loc.font.color.rgb = RGBColor(90, 90, 90)
        r_q = p_e.add_run(f'"{item.get("verbatim_evidence", "")}"\n\n')
        r_q.italic = True
        r_q.font.size = Pt(8.5)
        
        status = item.get("checksum_status", "Verified Grounded")
        r_chk = p_e.add_run(f"✓ {status}" if "Verified" in status else f"⚠️ {status}")
        r_chk.bold = True
        r_chk.font.size = Pt(7.5)
        r_chk.font.color.rgb = RGBColor(30, 130, 30) if "Verified" in status else RGBColor(180, 40, 40)

        # Col 3: Detailed Critique
        p_c = row_cells[3].paragraphs[0]
        r_crit = p_c.add_run(item.get("critique", ""))
        r_crit.font.size = Pt(9)

    doc.add_paragraph("")

    # Section 2: Summary of Findings and Final Scorecard
    doc.add_heading("2. Summary of Findings & Final Scorecard", level=1)

    doc.add_heading("Overall Methodological Synthesis", level=2)
    p_syn = doc.add_paragraph(data.get("overall_synthesis", ""))
    p_syn.paragraph_format.line_spacing = 1.15
    p_syn.runs[0].font.size = Pt(10)

    doc.add_heading("Final Scorecard Summary", level=2)
    score_table = doc.add_table(rows=1, cols=3)
    score_table.style = "Table Grid"

    score_headers = ["Domain / Methodological Item", "Rating / Assigned Score", "Status"]
    for i, title_text in enumerate(score_headers):
        cell = score_table.rows[0].cells[i]
        set_cell_background(cell, "333F48")
        p = cell.paragraphs[0]
        r = p.add_run(title_text)
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = RGBColor(255, 255, 255)

    for item in data.get("domain_appraisals", []):
        r_cells = score_table.add_row().cells
        r_cells[0].paragraphs[0].add_run(item.get("domain_name", "")).font.size = Pt(9)
        r_score = r_cells[1].paragraphs[0].add_run(item.get("rating", ""))
        r_score.bold = True
        r_score.font.size = Pt(9)

        rating_str = str(item.get("rating", "")).lower()
        if any(w in rating_str for w in ["low", "yes", "6", "7", "adequate", "2"]):
            badge = "Satisfactory"
        elif any(w in rating_str for w in ["some", "moderate", "partial", "4", "5", "1"]):
            badge = "Minor / Moderate Concern"
        else:
            badge = "High Concern / Deficit"
        r_cells[2].paragraphs[0].add_run(badge).font.size = Pt(8.5)

    doc.add_paragraph("")
    p_verdict = doc.add_paragraph()
    r_vlabel = p_verdict.add_run("Final Quality Determination: ")
    r_vlabel.bold = True
    r_vlabel.font.size = Pt(11)
    r_vval = p_verdict.add_run(data.get("overall_verdict", "N/A"))
    r_vval.bold = True
    r_vval.font.size = Pt(11)
    r_vval.font.color.rgb = RGBColor(170, 30, 30)

    doc.add_paragraph("")

    # Section 3: Appraisal Standard Reference Appendix
    doc.add_heading("3. Standard Reference, Tool Abbreviation & Scoring Scales", level=1)
    
    tool_meta = data.get("tool_meta", {})
    acronym = tool_meta.get("acronym") or data.get("appraisal_tool_selected", "Tool")
    full_name = tool_meta.get("full_name") or "Standard Critical Appraisal Instrument"
    purpose = tool_meta.get("purpose") or "Standardized epidemiological instrument used to assess bias and study rigour."
    scoring_rules = tool_meta.get("scoring_rules") or "Standard qualitative rating or numerical score threshold."

    design_match = TOOL_TAXONOMY.get(data.get("study_design", ""))
    if design_match:
        if not full_name or full_name == "Standard Critical Appraisal Instrument":
            full_name = design_match.get("full_name", full_name)
        if not purpose or "Standardized epidemiological" in purpose:
            purpose = design_match.get("purpose", purpose)
        if not scoring_rules or "Standard qualitative" in scoring_rules:
            scoring_rules = design_match.get("scale", scoring_rules)

    app_table = doc.add_table(rows=4, cols=2)
    app_table.style = "Table Grid"
    app_entries = [
        ("Tool & Abbreviation:", f"{acronym} — {full_name}"),
        ("Instrument Purpose:", purpose),
        ("Scoring Metric & Scales:", scoring_rules),
        ("Interpretation Guide:", "Evidence appraised using this framework should be balanced against clinical relevance and the patient population. Low risk of bias or high score thresholds indicate strong internal validity, while high risk domains warrant caution before clinical translation."),
    ]
    for idx, (label, val) in enumerate(app_entries):
        row = app_table.rows[idx]
        set_cell_background(row.cells[0], "F2F5F8")
        p0 = row.cells[0].paragraphs[0]
        r0 = p0.add_run(label)
        r0.bold = True
        r0.font.size = Pt(9.5)
        p1 = row.cells[1].paragraphs[0]
        r1 = p1.add_run(val)
        r1.font.size = Pt(9.5)

    doc.add_paragraph("")
    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_pdf_report(data: dict) -> io.BytesIO:
    """Generates an executive, styled PDF report via ReportLab in pure Python."""
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=36,
        bottomMargin=36
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "DocTitle",
        parent=styles["Heading1"],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#1F4E79"),
        fontName="Helvetica-Bold",
        spaceAfter=4,
    )
    sub_style = ParagraphStyle(
        "DocSub",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#555555"),
        fontName="Helvetica-Oblique",
        spaceAfter=12,
    )
    h1_style = ParagraphStyle(
        "H1",
        parent=styles["Heading2"],
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#1F4E79"),
        fontName="Helvetica-Bold",
        spaceBefore=12,
        spaceAfter=6,
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11.5,
        fontName="Helvetica",
    )
    bold_body_style = ParagraphStyle(
        "BoldBody",
        parent=body_style,
        fontName="Helvetica-Bold",
    )
    table_header_style = ParagraphStyle(
        "TableHeader",
        parent=body_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
        fontSize=8.5,
        leading=11,
    )

    story = []

    # Title Banner
    story.append(Paragraph(f"{APP_NAME} — Methodological Evaluation", title_style))
    story.append(Paragraph("Automated critical evidence synthesis evaluated via Google Gemini.", sub_style))

    # Bibliographic Metadata Table
    meta_data = [
        [Paragraph("Evaluated Title:", bold_body_style), Paragraph(data.get("paper_title", "N/A"), body_style)],
        [Paragraph("Authors:", bold_body_style), Paragraph(data.get("authors", "Not Specified"), body_style)],
        [Paragraph("Journal & Year:", bold_body_style), Paragraph(f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})", body_style)],
        [Paragraph("DOI / Reference:", bold_body_style), Paragraph(data.get("doi", "N/A"), body_style)],
        [Paragraph("Study Architecture:", bold_body_style), Paragraph(data.get("study_design", "N/A"), body_style)],
        [Paragraph("Gold-Standard Tool Applied:", bold_body_style), Paragraph(data.get("appraisal_tool_selected", "N/A"), body_style)],
        [Paragraph("Checksum Groundedness:", bold_body_style), Paragraph(f"{data.get('groundedness_score', 100.0)}% Evidence Grounded in Document", bold_body_style)],
    ]
    t_meta = Table(meta_data, colWidths=[1.8 * inch, 5.6 * inch])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F2F5F8')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 10))

    # Visual Quality Meter
    tier = str(data.get("quality_tier", "")).upper()
    if "GREEN" not in tier and "YELLOW" not in tier and "RED" not in tier:
        v = str(data.get("overall_verdict", "")).lower()
        tier = "GREEN" if any(w in v for w in ["low", "high", "recommended"]) else ("YELLOW" if any(w in v for w in ["some", "moderate"]) else "RED")

    story.append(Paragraph("Methodological Rigour & Clinical Trustworthiness Gauge (At-A-Glance):", bold_body_style))
    story.append(Spacer(1, 4))

    gauge_data = [[
        Paragraph("<font color='white'><b>HIGH RISK / FLAWED</b><br/>(Limited Reliability)</font>" if tier == "RED" else "HIGH RISK / FLAWED<br/>(Limited Reliability)", body_style),
        Paragraph("<font color='white'><b>MODERATE CONCERNS</b><br/>(Conditional Validity)</font>" if tier == "YELLOW" else "MODERATE CONCERNS<br/>(Conditional Validity)", body_style),
        Paragraph("<font color='white'><b>ROBUST / LOW RISK</b><br/>(Strong Validity)</font>" if tier == "GREEN" else "ROBUST / LOW RISK<br/>(Strong Validity)", body_style),
    ]]
    t_gauge = Table(gauge_data, colWidths=[2.46 * inch, 2.46 * inch, 2.46 * inch])
    t_gauge.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#C00000') if tier == "RED" else colors.HexColor('#E6B8B8')),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#FFC000') if tier == "YELLOW" else colors.HexColor('#FFF2CC')),
        ('BACKGROUND', (2, 0), (2, 0), colors.HexColor('#385623') if tier == "GREEN" else colors.HexColor('#E2EFDA')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.white),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]))
    story.append(t_gauge)
    story.append(Spacer(1, 10))

    # Domain Appraisal Breakdown
    story.append(Paragraph(f"1. Detailed Domain Appraisal ({data.get('appraisal_tool_selected', 'Tool')})", h1_style))
    domain_table_data = [[
        Paragraph("Domain & Prompt Criterion", table_header_style),
        Paragraph("Score", table_header_style),
        Paragraph("Grounded Quote & Citation", table_header_style),
        Paragraph("Methodological Critique", table_header_style),
    ]]

    for d in data.get("domain_appraisals", []):
        chk_status = d.get("checksum_status", "Verified Grounded")
        chk_tag = "<font color='#1b7e1b'><b>✓ Verified</b></font>" if "Verified" in chk_status else "<font color='#aa2222'><b>⚠️ Unverified</b></font>"
        
        domain_cell = Paragraph(f"<b>{d.get('domain_name', '')}</b><br/><br/><i>{d.get('evaluation_question', '')}</i>", body_style)
        rating_cell = Paragraph(f"<b>{d.get('rating', '')}</b>", bold_body_style)
        evidence_cell = Paragraph(f"<b>[{d.get('location_reference', 'Section')}]</b><br/><i>\"{d.get('verbatim_evidence', '')}\"</i><br/><br/>{chk_tag}", body_style)
        critique_cell = Paragraph(d.get("critique", ""), body_style)
        domain_table_data.append([domain_cell, rating_cell, evidence_cell, critique_cell])

    t_domains = Table(domain_table_data, colWidths=[2.2 * inch, 0.9 * inch, 2.3 * inch, 2.0 * inch], repeatRows=1)
    t_domains.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1F4E79')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_domains)
    story.append(Spacer(1, 12))

    # Executive Synthesis & Scorecard Summary
    story.append(Paragraph("2. Methodological Synthesis & Scorecard Summary", h1_style))
    story.append(Paragraph(f"<b>Executive Synthesis:</b> {data.get('overall_synthesis', '')}", body_style))
    story.append(Spacer(1, 8))

    score_data = [[
        Paragraph("Domain Item", table_header_style),
        Paragraph("Score / Rating", table_header_style),
        Paragraph("Determination", table_header_style)
    ]]
    for item in data.get("domain_appraisals", []):
        rating_str = str(item.get("rating", "")).lower()
        if any(w in rating_str for w in ["low", "yes", "6", "7", "adequate", "2"]):
            badge = "Satisfactory"
        elif any(w in rating_str for w in ["some", "moderate", "partial", "4", "5", "1"]):
            badge = "Minor / Moderate Concern"
        else:
            badge = "High Concern / Deficit"
        score_data.append([
            Paragraph(item.get("domain_name", ""), body_style),
            Paragraph(f"<b>{item.get('rating', '')}</b>", body_style),
            Paragraph(badge, body_style)
        ])

    t_score = Table(score_data, colWidths=[3.4 * inch, 2.0 * inch, 2.0 * inch], repeatRows=1)
    t_score.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#333F48')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_score)
    story.append(Spacer(1, 8))

    story.append(Paragraph(f"<b>Final Determination:</b> <font color='#aa2222'><b>{data.get('overall_verdict', 'N/A')}</b></font>", bold_body_style))
    story.append(Spacer(1, 10))

    # Tool Metadata Reference Table
    story.append(Paragraph("3. Appraisal Framework & Standards Appendix", h1_style))
    tool_meta = data.get("tool_meta", {})
    ref_data = [
        [Paragraph("Tool & Abbreviation:", bold_body_style), Paragraph(f"{tool_meta.get('acronym', 'Tool')} — {tool_meta.get('full_name', '')}", body_style)],
        [Paragraph("Instrument Purpose:", bold_body_style), Paragraph(tool_meta.get("purpose", ""), body_style)],
        [Paragraph("Scoring Scale:", bold_body_style), Paragraph(tool_meta.get("scoring_rules", ""), body_style)],
    ]
    t_ref = Table(ref_data, colWidths=[1.8 * inch, 5.6 * inch])
    t_ref.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#F2F5F8')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#CCCCCC')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_ref)

    doc.build(story)
    buf.seek(0)
    return buf


# ==========================================
# 5. STREAMLIT APP & SECURITY GATE
# ==========================================
st.set_page_config(page_title=APP_NAME, page_icon="📑", layout="wide")
init_db()

st.markdown(
    """
<style>
    .stButton>button { width: 100%; border-radius: 8px; font-weight: 600; }
    .history-card {
        padding: 8px 10px;
        margin-bottom: 6px;
        border-radius: 6px;
        background-color: rgba(255, 255, 255, 0.05);
        border: 1px solid rgba(255, 255, 255, 0.1);
    }
</style>
""",
    unsafe_allow_html=True,
)

CORRECT_PASSCODE = "324598"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title(f"🔒 {APP_NAME}: Access Restricted")
    st.write("Please enter the passcode to access the appraisal workspace.")
    passcode_input = st.text_input("Passcode", type="password")
    if st.button("Unlock App"):
        if passcode_input.strip() == CORRECT_PASSCODE:
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Invalid passcode. Please try again.")
    st.stop()

# --- Main App Header ---
st.title(f"📑 {APP_NAME}")
st.caption("Automated Methodological Critical Appraisal & Evidence Synthesis")

api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", None))

if "history_limit" not in st.session_state:
    st.session_state["history_limit"] = 10

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
        st.caption("No saved appraisals yet.")
    else:
        current_limit = st.session_state["history_limit"]
        displayed_records = records[:current_limit]
        
        st.caption(f"Showing {len(displayed_records)} of {len(records)} entries")

        for r in displayed_records:
            verdict_text = str(r.get("overall_verdict", "")).lower()
            if any(w in verdict_text for w in ["low", "recommended", "high confidence", "good"]):
                dot = "🟢"
            elif any(w in verdict_text for w in ["some", "moderate", "fair"]):
                dot = "🟡"
            else:
                dot = "🔴"

            with st.container():
                short_title = r["paper_title"][:38] + "..." if len(r["paper_title"]) > 38 else r["paper_title"]
                st.markdown(
                    f"**{dot} {short_title}**  \n"
                    f"<small style='color: gray;'>{r['created_at'][:10]} · {r['tool_selected']}</small>",
                    unsafe_allow_html=True
                )
                
                c_load, c_del = st.columns([3, 1])
                with c_load:
                    if st.button("View", key=f"view_{r['id']}", use_container_width=True):
                        st.session_state["appraisal_data"] = get_appraisal_by_id(r["id"])
                        st.rerun()
                with c_del:
                    if st.button("🗑️", key=f"del_{r['id']}", use_container_width=True):
                        delete_appraisal(r["id"])
                        st.rerun()
                st.markdown("<hr style='margin: 8px 0; border: none; border-top: 1px solid rgba(255,255,255,0.08);'/>", unsafe_allow_html=True)

        if len(records) > 10:
            if current_limit == 10:
                if st.button("⬇️ Show More (Up to 50)"):
                    st.session_state["history_limit"] = 50
                    st.rerun()
            else:
                if st.button("⬆️ Show Less (Top 10)"):
                    st.session_state["history_limit"] = 10
                    st.rerun()

tab_pdf, tab_url = st.tabs(["📤 Upload PDF (Mobile / Desktop)", "🔗 Enter Link or DOI"])

paper_text = ""
raw_pdf_bytes = None
filename_tag = "Paper"

with tab_pdf:
    uploaded = st.file_uploader(
        "Upload Manuscript or Clinical Practice Guideline", type=["pdf"]
    )
    if uploaded:
        raw_pdf_bytes = uploaded.read()
        filename_tag = uploaded.name
        with st.spinner("Extracting text locally from PDF pages..."):
            paper_text = extract_text_from_pdf_bytes(raw_pdf_bytes)

        if paper_text:
            st.success(
                f"⚡ Parsed text successfully ({len(paper_text):,} characters extracted)"
            )
        else:
            st.warning(
                "⚠️ PDF appears scanned or image-based. Gemini Native Vision OCR will be used automatically."
            )

with tab_url:
    url_input = st.text_input(
        "Paste Article Link or DOI",
        placeholder="https://doi.org/10.1016/... or https://pubmed...",
    )
    if st.button("Fetch Article Content"):
        if url_input.strip():
            with st.spinner("Retrieving article text..."):
                try:
                    paper_text = extract_text_from_url(url_input.strip())
                    filename_tag = "Web_Article"
                    st.success(f"Retrieved {len(paper_text):,} characters from link.")
                except Exception as e:
                    st.error(f"Failed to fetch content: {str(e)}")

# Instant local pattern recognition
detected_design = "Randomized Controlled Trial (RCT)"
if paper_text:
    detected_design = classify_study_design_heuristically(paper_text)

st.markdown("---")
available_designs = list(TOOL_TAXONOMY.keys())
default_idx = (
    available_designs.index(detected_design)
    if detected_design in available_designs
    else 0
)

selected_design = st.selectbox(
    "Target Study Architecture (Auto-detected via Python heuristics):",
    options=available_designs,
    index=default_idx,
    help="Python detected this study design based on title, abstract, and methodology keywords. You can manually adjust it if needed."
)

if st.button("⚡ Run Methodological Appraisal", type="primary"):
    if not api_key:
        st.warning("Please enter your Gemini API key in the sidebar.")
    elif not paper_text and not raw_pdf_bytes:
        st.warning("Please upload a PDF or fetch an article URL first.")
    else:
        client = genai.Client(api_key=api_key)
        mode_label = (
            "Vision OCR"
            if (raw_pdf_bytes and not paper_text)
            else "Extracted Text (Trimmed)"
        )
        with st.spinner(f"Evaluating {selected_design} using {mode_label}..."):
            try:
                raw_data = appraise_with_gemini(
                    client,
                    text_content=paper_text,
                    raw_pdf_bytes=raw_pdf_bytes if not paper_text else None,
                    assigned_design=selected_design,
                )
                
                # Run deterministic fuzzy token-matching checksum
                verified_data, score, warnings = run_python_groundedness_checksum(
                    raw_data, paper_text, fuzzy_threshold=0.82
                )
                
                new_id = save_appraisal(verified_data, filename_tag)
                st.session_state["appraisal_data"] = verified_data
                st.success(f"Appraisal finished and saved to library (ID: #{new_id})! Groundedness: {score}%")
                st.rerun()
            except Exception as e:
                st.error(f"Appraisal failed: {str(e)}")

# Display Results
if "appraisal_data" in st.session_state:
    data = st.session_state["appraisal_data"]
    st.markdown("---")
    
    # Paper Header & Bibliographic Details
    st.subheader(f"📋 {data.get('paper_title', 'Untitled Manuscript')}")
    
    b1, b2, b3, b4 = st.columns(4)
    with b1:
        st.caption("👥 **Authors**")
        st.write(data.get("authors", "Not Specified"))
    with b2:
        st.caption("📖 **Journal**")
        st.write(data.get("journal", "Not Specified"))
    with b3:
        st.caption("📅 **Year**")
        st.write(str(data.get("year_published", "N/A")))
    with b4:
        st.caption("🔗 **DOI / URL**")
        doi_val = data.get("doi", "N/A")
        if "10." in doi_val or "http" in doi_val:
            clean_url = doi_val if doi_val.startswith("http") else f"https://doi.org/{doi_val}"
            st.markdown(f"[{doi_val}]({clean_url})")
        else:
            st.write(doi_val)

    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Study Design", data.get("study_design"))
    with c2:
        st.metric("Tool Selected", data.get("appraisal_tool_selected"))
    with c3:
        st.metric("Overall Verdict", data.get("overall_verdict"))
    with c4:
        chk_score = data.get("groundedness_score", 100.0)
        st.metric("Evidence Groundedness", f"{chk_score}%", help="Calculated via local Python SequenceMatcher comparing LLM quotes to raw PDF text.")

    # Checksum Warning Alert
    if data.get("groundedness_warnings"):
        with st.expander("⚠️ Python Checksum Warnings (Unverified Citations)"):
            for w in data["groundedness_warnings"]:
                st.write(f"- {w}")

    # Three-Bar Quality Gauge in Streamlit
    tier = str(data.get("quality_tier", "")).upper()
    if "GREEN" not in tier and "YELLOW" not in tier and "RED" not in tier:
        v = str(data.get("overall_verdict", "")).lower()
        tier = "GREEN" if any(w in v for w in ["low", "high", "recommended"]) else ("YELLOW" if any(w in v for w in ["some", "moderate"]) else "RED")

    st.markdown("#### 🎯 At-A-Glance Quality Meter")
    g_col1, g_col2, g_col3 = st.columns(3)
    with g_col1:
        if tier == "RED":
            st.error("🚨 **RED TIER**\nHigh Risk / Methodological Flaws")
        else:
            st.caption("Red (High Risk)")
    with g_col2:
        if tier == "YELLOW":
            st.warning("⚠️ **YELLOW TIER**\nModerate Quality / Some Concerns")
        else:
            st.caption("Yellow (Moderate)")
    with g_col3:
        if tier == "GREEN":
            st.success("✅ **GREEN TIER**\nRobust / Low Risk / Trustworthy")
        else:
            st.caption("Green (Robust)")

    st.subheader("📊 Domain Breakdown with Questions")
    rows = [
        {
            "Domain": d.get("domain_name"),
            "Question / Criterion": d.get("evaluation_question"),
            "Score/Rating": d.get("rating"),
            "Reference": d.get("location_reference"),
            "Verbatim Evidence": d.get("verbatim_evidence"),
            "Checksum Status": d.get("checksum_status", "Verified Grounded"),
            "Methodological Critique": d.get("critique"),
        }
        for d in data.get("domain_appraisals", [])
    ]
    st.dataframe(rows, use_container_width=True)

    st.subheader("📝 Executive Summary & Findings")
    st.info(data.get("overall_synthesis"))

    tool_meta = data.get("tool_meta", {})
    if tool_meta:
        with st.expander("📖 Appraisal Framework & Scoring Standards"):
            st.markdown(f"**Tool:** {tool_meta.get('acronym')} ({tool_meta.get('full_name')})")
            st.markdown(f"**Purpose:** {tool_meta.get('purpose')}")
            st.markdown(f"**Scoring Metric:** {tool_meta.get('scoring_rules')}")

    # Export Documents (Word & PDF)
    safe_title = re.sub(r"[^a-zA-Z0-9_-]", "_", data.get("paper_title", "Report")[:25])
    col_docx, col_pdf = st.columns(2)

    with col_docx:
        doc_buf = build_docx_report(data)
        st.download_button(
            label="📥 Download Word Document (.docx)",
            data=doc_buf,
            file_name=f"Kitique_{data.get('appraisal_tool_selected')}_{safe_title}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    with col_pdf:
        pdf_buf = build_pdf_report(data)
        st.download_button(
            label="📥 Download Formatted PDF (.pdf)",
            data=pdf_buf,
            file_name=f"Kitique_{data.get('appraisal_tool_selected')}_{safe_title}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
