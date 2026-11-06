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

# Optional OCR fallback imports
try:
    from pdf2image import convert_from_bytes
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

DB_NAME = "appraisals.db"
APP_NAME = "Kitique Appraisal Tool"

FREE_ACTIVE_MODELS = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-3.8-flash",
]


def get_connection():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


SEED_TAXONOMY = {
    "Systematic Review / Meta-Analysis": {
        "tool": "AMSTAR 2",
        "full_name": "A MeaSurement Tool to Assess systematic Reviews (Version 2)",
        "purpose": "Critically appraises systematic reviews evaluating randomized or non-randomized healthcare studies.",
        "domains": [
            {"name": "Domain 1: PICO & Protocol Pre-registration", "question": "Did the review establish PICO and register protocol (e.g. PROSPERO) prior to screening?"},
            {"name": "Domain 2: Literature Search Strategy", "question": "Did authors search >=2 databases, provide full boolean search strings, and search trial registries?"},
            {"name": "Domain 3: Duplicate Study Selection & Extraction", "question": "Did at least two reviewers independently screen titles, abstracts, and full texts?"},
            {"name": "Domain 4: Excluded Studies Justification", "question": "Did authors provide full bibliographic list of excluded studies with explicit rationale?"},
            {"name": "Domain 5: Primary Study RoB Assessment", "question": "Did the review apply established, validated risk of bias tools to included studies?"},
            {"name": "Domain 6: Meta-Analytic Synthesis Appropriateness", "question": "Were statistical pooling methods justified and statistical heterogeneity investigated?"},
            {"name": "Domain 7: Publication Bias Assessment", "question": "Did authors formally test for publication bias (funnel plots, Egger's test)?"},
            {"name": "Domain 8: Conflict of Interest & Funding", "question": "Did authors report their own funding/COIs and funding of primary included studies?"},
        ],
        "scale": "High Confidence | Moderate Confidence | Low Confidence | Critically Low Confidence.",
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
    "Observational Study (Cohort)": {
        "tool": "Newcastle-Ottawa Scale (NOS) Cohort",
        "full_name": "Newcastle-Ottawa Quality Assessment Scale for Cohort Studies",
        "purpose": "Evaluates selection, comparability, and outcome assessment in longitudinal prospective and retrospective cohorts.",
        "domains": [
            {"name": "Selection: Representativeness & Exposure", "question": "Is the exposed cohort truly representative, non-exposed drawn from same community, and exposure ascertainment secured via records? (0-4 Stars)"},
            {"name": "Comparability: Confounder Control", "question": "Did the study control for the most important factor and additional relevant covariates via design/analysis? (0-2 Stars)"},
            {"name": "Outcome: Assessment & Follow-up", "question": "Was outcome assessed independently/blinded, follow-up long enough, and adequacy of cohort follow-up complete (>80%)? (0-3 Stars)"},
        ],
        "scale": "Stars Awarded (0-9 Total): Good Quality (7-9 Stars) | Fair Quality (4-6 Stars) | Poor Quality (0-3 Stars).",
    },
    "Observational Study (Case-Control)": {
        "tool": "Newcastle-Ottawa Scale (NOS) Case-Control",
        "full_name": "Newcastle-Ottawa Quality Assessment Scale for Case-Control Studies",
        "purpose": "Evaluates case/control definition, selection bias, and exposure ascertainment in retrospective case-control designs.",
        "domains": [
            {"name": "Selection: Case and Control Definition", "question": "Are cases adequately defined/validated, representativeness verified, and controls derived from same population with no disease history? (0-4 Stars)"},
            {"name": "Comparability: Matching and Adjustment", "question": "Are cases and controls matched or adjusted for primary and secondary prognostic factors? (0-2 Stars)"},
            {"name": "Exposure: Ascertainment & Non-Response", "question": "Was exposure ascertained via blinded interview/structured records and same non-response rate for both groups? (0-3 Stars)"},
        ],
        "scale": "Stars Awarded (0-9 Total): Good Quality (7-9 Stars) | Fair Quality (4-6 Stars) | Poor Quality (0-3 Stars).",
    },
    "Cross-Sectional Study": {
        "tool": "AXIS",
        "full_name": "Appraisal tool for Cross-Sectional Studies (AXIS)",
        "purpose": "Assesses study design, target population framing, non-response bias, and outcome measurement validity in observational cross-sectional surveys.",
        "domains": [
            {"name": "Domain 1: Aims, Target Population & Sampling", "question": "Were objectives clear, study population clearly defined, and sampling frame appropriate to represent target group?"},
            {"name": "Domain 2: Selection Bias & Response Rate", "question": "Were measures taken to address non-responders, and was response rate high enough to prevent non-response bias?"},
            {"name": "Domain 3: Measurement Validity & Reliability", "question": "Were risk factors and outcome variables measured using instruments shown to be valid and reliable?"},
            {"name": "Domain 4: Confounding & Statistical Reporting", "question": "Were methods sufficiently detailed, statistical tests appropriate, and potential confounding variables accounted for?"},
            {"name": "Domain 5: Ethical Standards & Internal Consistency", "question": "Was ethical approval documented, and are discussions/conclusions justified by the results presented?"},
        ],
        "scale": "Yes | No | Don't Know across 20 structured criteria. High Quality (>=15 Yes) | Moderate Quality (10-14 Yes) | Low Quality (<10 Yes).",
    },
    "Prediction Model Study": {
        "tool": "PROBAST",
        "full_name": "Prediction model Risk Of Bias ASsessment Tool",
        "purpose": "Assesses risk of bias and applicability of diagnostic and prognostic clinical prediction model studies.",
        "domains": [
            {"name": "Domain 1: Participants", "question": "Were data sources appropriate and participant inclusion/exclusion avoided inappropriate selection bias?"},
            {"name": "Domain 2: Predictors", "question": "Were predictors defined and assessed blindly and consistently without knowledge of outcomes?"},
            {"name": "Domain 3: Outcome", "question": "Was outcome pre-specified, determined objectively, and without knowledge of predictor status?"},
            {"name": "Domain 4: Analysis", "question": "Were participant-to-variable ratio adequate, continuous predictors not inappropriately dichotomized, missing data handled, and model performance properly evaluated?"},
        ],
        "scale": "Low Risk | High Risk | Unclear Risk. If at least one domain is High Risk, overall model risk of bias is High Risk.",
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


def clean_pdf_artifacts(raw_text: str) -> str:
    cleaned = re.sub(r"(\w+)-\s*\n\s*(\w+)", r"\1\2", raw_text)
    cleaned = re.sub(r"(?i)\bpage\s+\d+\s+(?:of|\/)\s+\d+\b", "", cleaned)
    cleaned = re.sub(r"[ \t]+", " ", cleaned)
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned)
    return cleaned.strip()


def ocr_pdf(pdf_bytes: bytes) -> str:
    if not OCR_AVAILABLE:
        return ""
    try:
        pages = convert_from_bytes(pdf_bytes)
        text_pages = []
        for idx, page in enumerate(pages[:15]):
            extracted = pytesseract.image_to_string(page)
            if extracted.strip():
                text_pages.append(f"--- [Page {idx + 1}] ---\n{extracted.strip()}")
        return "\n\n".join(text_pages)
    except Exception:
        return ""


def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    text = ""
    try:
        reader = PdfReader(io.BytesIO(pdf_bytes))
        extracted = []
        for idx, page in enumerate(reader.pages):
            page_text = page.extract_text() or ""
            if page_text.strip():
                extracted.append(f"--- [Page {idx + 1}] ---\n{page_text.strip()}")
        text = "\n\n".join(extracted)
    except Exception:
        text = ""

    if len(text.strip()) < 100 and OCR_AVAILABLE:
        ocr_text = ocr_pdf(pdf_bytes)
        if len(ocr_text.strip()) > len(text.strip()):
            text = ocr_text

    return clean_pdf_artifacts(text) if len(text.strip()) > 100 else ""


def extract_text_from_url(raw_input: str) -> str:
    target = raw_input.strip()
    if not target.startswith("http://") and not target.startswith("https://"):
        if target.startswith("10.") or "/10." in target:
            clean_doi = target[target.find("10.") :]
            target = f"https://doi.org/{clean_doi}"
        else:
            target = f"https://{target}"

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    resp = requests.get(target, headers=headers, timeout=15)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")
    for elem in soup(["script", "style", "nav", "footer", "header", "noscript"]):
        elem.decompose()
    article = soup.find("article") or soup.find("main") or soup.body
    return article.get_text(separator="\n", strip=True) if article else ""


# ==========================================
# 3. HARVARD CITATION & AUTHOR FORMATTER (MAX 3)
# ==========================================
def format_harvard_authors(raw_authors: str, max_authors: int = 3) -> str:
    if not raw_authors or str(raw_authors).strip().lower() in ["not specified", "unspecified", "unknown", "n/a", "none"]:
        return "Not Specified"

    raw_clean = re.sub(r"\b(and|et al\.?|&)\b", ",", str(raw_authors), flags=re.IGNORECASE)
    raw_segments = [seg.strip() for seg in raw_clean.split(",") if seg.strip()]

    parsed_authors = []
    for seg in raw_segments:
        tokens = [t.strip(" .") for t in seg.split() if re.search(r"[a-zA-Z]", t)]
        if not tokens:
            continue
        if len(tokens) == 1:
            parsed_authors.append(tokens[0].title())
        elif len(tokens[0]) <= 2 and len(tokens) > 1 and len(tokens[-1]) > 2:
            surname = tokens[-1].title()
            initials = "".join([f"{t[0].upper()}." for t in tokens[:-1]])
            parsed_authors.append(f"{surname}, {initials}")
        else:
            surname = tokens[-1].title()
            initials = "".join([f"{t[0].upper()}." for t in tokens[:-1]])
            parsed_authors.append(f"{surname}, {initials}")

    if not parsed_authors:
        return str(raw_authors).strip()

    if len(parsed_authors) > max_authors:
        return f"{', '.join(parsed_authors[:max_authors])} et al."
    elif len(parsed_authors) == 1:
        return parsed_authors[0]
    elif len(parsed_authors) == 2:
        return f"{parsed_authors[0]} and {parsed_authors[1]}"
    else:
        return f"{', '.join(parsed_authors[:-1])} and {parsed_authors[-1]}"


def format_harvard_filename(data: dict, extension: str) -> str:
    raw_authors = str(data.get("authors", "")).strip()
    year = str(data.get("year_published", "")).strip()
    tool = str(data.get("appraisal_tool_selected", "Tool")).strip()

    year_match = re.search(r"\b(19\d\d|20\d\d)\b", year)
    clean_year = year_match.group(1) if year_match else "UnknownYear"

    if not raw_authors or raw_authors.lower() in ["not specified", "unspecified", "unknown", "n/a"]:
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


# ==========================================
# 4. 4-SOURCE RANKED ENSEMBLE CLASSIFIER
# ==========================================
def extract_declared_design_ranked(text: str) -> Optional[Tuple[str, str, int]]:
    head = text[:6000].lower()

    ranked_definitions = [
        ("Systematic Review / Meta-Analysis", [
            r"\bsystematic\s+review\s+and\s+meta-analysis\b",
            r"\bmeta-analysis\s+and\s+systematic\s+review\b",
            r"\bnetwork\s+meta-analysis\b",
            r"\bsystematic\s+review\b",
            r"\bmeta-analysis\b",
        ], 100),
        ("Clinical Practice Guideline / Consensus Statement", [
            r"\bclinical\s+practice\s+guideline\b",
            r"\bconsensus\s+statement\b",
            r"\bdelphi\s+consensus\b",
            r"\bexpert\s+consensus\b",
        ], 95),
        ("Randomized Controlled Trial (RCT)", [
            r"\brandomi[zs]ed\s+controlled\s+trial\b",
            r"\brandomi[zs]ed\s+clinical\s+trial\b",
            r"\bdouble-blind(?:ed)?,\s+placebo-controlled\b",
        ], 90),
        ("Prediction Model Study", [
            r"\bclinical\s+prediction\s+model\b",
            r"\bprognostic\s+model\s+study\b",
            r"\brisk\s+prediction\s+model\b",
            r"\bprediction\s+rule\s+validation\b",
        ], 85),
        ("Diagnostic Accuracy Study", [
            r"\bdiagnostic\s+accuracy\s+study\b",
            r"\bdiagnostic\s+test\s+accuracy\b",
        ], 82),
        ("Observational Study (Cohort)", [
            r"\bretrospective\s+cohort\s+study\b",
            r"\bprospective\s+cohort\s+study\b",
            r"\bretrospective\s+comparative\s+study\b",
            r"\bcomparative\s+cohort\s+study\b",
            r"\bretrospective\s+cohort\b",
            r"\bprospective\s+cohort\b",
            r"\bcohort\s+study\b",
        ], 80),
        ("Observational Study (Case-Control)", [
            r"\bcase-control\s+study\b",
            r"\bcase\s+control\s+study\b",
            r"\bmatched\s+case-control\b",
        ], 78),
        ("Cross-Sectional Study", [
            r"\bcross-sectional\s+study\b",
            r"\bcross\s+sectional\s+study\b",
            r"\bcross-sectional\s+survey\b",
            r"\bprevalence\s+study\b",
        ], 75),
        ("Qualitative Study", [
            r"\bqualitative\s+study\b",
            r"\bthematic\s+analysis\s+study\b",
            r"\bphenomenological\s+study\b",
        ], 70),
        ("Narrative Review", [
            r"\bnarrative\s+review\b",
            r"\bscoping\s+review\b",
            r"\bliterature\s+review\b",
            r"\bstate-of-the-art\s+review\b",
        ], 65),
    ]

    discovered = []
    for design, patterns, rank in ranked_definitions:
        for p in patterns:
            match = re.search(p, head)
            if match:
                discovered.append((design, match.group(0), rank))
                break

    if not discovered:
        return None

    discovered.sort(key=lambda x: x[2], reverse=True)
    return discovered[0]


def extract_study_features(text: str) -> dict:
    head = text[:9000].lower()

    is_review = bool(re.search(r"\b(review|systematic\s+review|meta-analysis|literature\s+search|prisma|scoping\s+review)\b", head))
    is_guideline = bool(re.search(r"\b(clinical\s+practice\s+guideline|consensus\s+statement|delphi\s+consensus|expert\s+panel\s+recommendations)\b", head))
    is_rct = bool(re.search(r"\b(randomized\s+controlled\s+trial|randomised\s+controlled\s+trial|double-blind|placebo-controlled|\brct\b)\b", head))
    is_intervention = bool(re.search(r"\b(intervention|allocated\s+to|treated\s+with|pre-post|quasi-experimental|before-and-after)\b", head))
    is_diagnostic = bool(re.search(r"\b(diagnostic\s+accuracy|index\s+test|reference\s+standard|sensitivity\s+and\s+specificity|roc\s+curve|auc)\b", head))
    is_prediction = bool(re.search(r"\b(prediction\s+model|risk\s+score|nomogram|prognostic\s+model|validation\s+cohort|c-index|probast)\b", head))
    is_cross_sectional = bool(re.search(r"\b(cross-sectional|survey|prevalence\s+study|questionnaire\s+survey|axis\s+tool)\b", head))
    is_case_control = bool(re.search(r"\b(case-control|cases\s+and\s+controls|matched\s+controls)\b", head))
    is_cohort = bool(re.search(r"\b(cohort\s+study|prospective\s+cohort|retrospective\s+cohort|longitudinal|follow-up\s+period|retrospective\s+comparative)\b", head))
    is_qualitative = bool(re.search(r"\b(qualitative\s+study|thematic\s+analysis|semi-structured\s+interview|focus\s+groups|grounded\s+theory|phenomenolog)\b", head))

    is_secondary = is_review or is_guideline
    is_primary = not is_secondary or bool(re.search(r"\b(we\s+enrolled|patients\s+were\s+recruited|our\s+cohort|consecutive\s+patients)\b", head))

    return {
        "primary_research": is_primary and not is_guideline,
        "secondary_synthesis": is_secondary,
        "meta_analysis": bool(re.search(r"\b(meta-analysis|forest\s+plot|pooled\s+odds|i\^2|heterogeneity)\b", head)),
        "systematic_search": bool(re.search(r"\b(prisma|database\s+search|pubmed\s+and\s+embase|search\s+strategy)\b", head)),
        "guideline_document": is_guideline,
        "randomized": is_rct,
        "intervention_assigned": is_intervention,
        "diagnostic_accuracy": is_diagnostic,
        "prediction_model": is_prediction,
        "cross_sectional": is_cross_sectional,
        "case_control": is_case_control,
        "cohort": is_cohort,
        "qualitative": is_qualitative,
    }


def classify_from_features(features: dict) -> str:
    if features.get("guideline_document"):
        return "Clinical Practice Guideline / Consensus Statement"
    if features.get("secondary_synthesis") and not features.get("primary_research"):
        if features.get("systematic_search") or features.get("meta_analysis"):
            return "Systematic Review / Meta-Analysis"
        return "Narrative Review"
    if features.get("randomized"):
        return "Randomized Controlled Trial (RCT)"
    if features.get("diagnostic_accuracy"):
        return "Diagnostic Accuracy Study"
    if features.get("prediction_model"):
        return "Prediction Model Study"
    if features.get("cross_sectional"):
        return "Cross-Sectional Study"
    if features.get("case_control"):
        return "Observational Study (Case-Control)"
    if features.get("cohort"):
        return "Observational Study (Cohort)"
    if features.get("intervention_assigned"):
        return "Non-Randomized Intervention / Quasi-Experimental"
    if features.get("qualitative"):
        return "Qualitative Study"

    return "Observational Study (Cohort)"


def compute_weighted_keyword_frequencies(text: str) -> str:
    head_title = text[:600].lower()
    abstract_zone = text[600:3000].lower()
    body_zone = text[3000:12000].lower()

    scores = {
        "Observational Study (Cohort)": 0,
        "Observational Study (Case-Control)": 0,
        "Cross-Sectional Study": 0,
        "Systematic Review / Meta-Analysis": 0,
        "Randomized Controlled Trial (RCT)": 0,
        "Diagnostic Accuracy Study": 0,
    }

    terms = {
        "Observational Study (Cohort)": ["cohort", "follow-up", "retrospective", "prospective", "incidence", "risk ratio"],
        "Observational Study (Case-Control)": ["case-control", "case control", "odds ratio", "matched controls"],
        "Cross-Sectional Study": ["cross-sectional", "prevalence", "survey", "questionnaire"],
        "Systematic Review / Meta-Analysis": ["prisma", "database search", "search strategy", "systematic review", "meta-analysis"],
        "Randomized Controlled Trial (RCT)": ["randomized", "randomised", "placebo", "double-blind", "intention-to-treat"],
        "Diagnostic Accuracy Study": ["sensitivity", "specificity", "index test", "reference standard", "roc curve"],
    }

    for design, word_list in terms.items():
        for w in word_list:
            scores[design] += (head_title.count(w) * 5)
            scores[design] += (abstract_zone.count(w) * 3)
            scores[design] += (body_zone.count(w) * 1)

    return max(scores, key=scores.get)


def verify_architecture_with_first_page(
    client: Optional[genai.Client], first_page_text: str, candidate_tools: List[str]
) -> str:
    if not client:
        return ""
    tools_str = ", ".join(f"'{t}'" for t in candidate_tools if not t.startswith("🔍"))
    verify_prompt = f"""
    You are an expert epidemiological classifier.
    Examine ONLY the title, abstract, and study design descriptors from Page 1 of this manuscript:

    {first_page_text[:4500]}

    Determine whether this is a PRIMARY study (gathering original patient/cohort data) or a SECONDARY SYNTHESIS (review, meta-analysis, guideline).
    Select the EXACT matching study design from this permitted list:
    [{tools_str}]

    CRITICAL: Output ONLY a JSON object:
    {{
      "confirmed_design": "Selected or identified design name",
      "rationale": "One-sentence rationale citing title/abstract text"
    }}
    """
    for model_name in FREE_ACTIVE_MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[verify_prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            parsed = json.loads(response.text)
            return parsed.get("confirmed_design", "")
        except Exception:
            continue
    return ""


def run_four_source_ensemble(
    text: str, client: Optional[genai.Client], available_designs: List[str]
) -> dict:
    evidence_points = []

    declared_res = extract_declared_design_ranked(text)
    declared_pred = declared_res[0] if declared_res else None
    if declared_pred:
        evidence_points.append(f"Declared: Author explicitly states '{declared_res[1]}' (Priority Rank: {declared_res[2]})")

    features = extract_study_features(text)
    feature_pred = classify_from_features(features)
    evidence_points.append(f"Features: Primary={features.get('primary_research')}, Cohort={features.get('cohort')}, RCT={features.get('randomized')}")

    keyword_pred = compute_weighted_keyword_frequencies(text)
    evidence_points.append(f"Positional Keywords: Highest weighted score points to '{keyword_pred}'")

    llm_pred = verify_architecture_with_first_page(client, text, available_designs)
    if llm_pred:
        evidence_points.append(f"Page-1 Verifier: Independent analysis confirmed '{llm_pred}'")

    if declared_pred:
        if feature_pred == declared_pred:
            assigned_confidence = 0.98
            override_msg = "Author declaration confirmed by structural features"
        else:
            assigned_confidence = 0.82
            override_msg = "Author declaration used with caveat (Structural features diverged)"
        evidence_points.append(f"Adjudication: {override_msg}")

        trace = {
            "declared_design": declared_pred,
            "feature_prediction": feature_pred,
            "keyword_prediction": keyword_pred,
            "llm_prediction": llm_pred or "N/A",
            "consensus": declared_pred,
            "confidence": assigned_confidence,
        }

        return {
            "winner": declared_pred,
            "confidence": assigned_confidence,
            "trace": trace,
            "evidence": evidence_points,
            "override_applied": True,
        }

    votes = [
        (feature_pred, 0.50),
        (llm_pred, 0.35 if llm_pred else 0.0),
        (keyword_pred, 0.15),
    ]

    scores = {}
    total_active_weight = 0.0
    for pred, w in votes:
        if not pred:
            continue
        scores[pred] = scores.get(pred, 0.0) + w
        total_active_weight += w

    winner = max(scores, key=scores.get) if scores else "Observational Study (Cohort)"
    raw_confidence = scores[winner] / (total_active_weight if total_active_weight > 0 else 1.0)

    trace = {
        "declared_design": "None Detected",
        "feature_prediction": feature_pred,
        "keyword_prediction": keyword_pred,
        "llm_prediction": llm_pred or "N/A",
        "consensus": winner,
        "confidence": round(raw_confidence, 2),
    }

    return {
        "winner": winner,
        "confidence": round(raw_confidence, 2),
        "trace": trace,
        "evidence": evidence_points,
        "override_applied": False,
    }


def detect_classification_contradiction(manuscript_text: str, predicted_design: str) -> Optional[str]:
    header = manuscript_text[:4000].lower()

    if "review" in predicted_design.lower():
        if any(term in header for term in ["retrospective cohort", "prospective cohort", "comparative study", "consecutive patients", "patients enrolled"]):
            return "Manuscript explicitly declares an empirical cohort/comparative study, but assigned framework is Review (SANRA/AMSTAR). Appraisal blocked."

    if "cohort" in predicted_design.lower() or "rct" in predicted_design.lower():
        if "systematic review and meta-analysis" in header or "systematic review" in header:
            return "Manuscript explicitly declares a Systematic Review, but assigned framework is Empirical (NOS/RoB 2). Appraisal blocked."

    return None


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
    ratings_summary = str(data.get("ratings_summary", {})).upper()
    has_high_risk = any("high" in str(d.get("rating", "")).lower() or "critical" in str(d.get("rating", "")).lower() for d in domains)

    if "GOOD" in ratings_summary and has_high_risk:
        flaw_warnings.append("Methodological Contradiction: Good quality assigned despite high-risk domain.")

    data["groundedness_score"] = grounded_score
    data["groundedness_warnings"] = grounding_warnings
    data["checksum_flaws"] = flaw_warnings
    data["checksum_passed"] = len(flaw_warnings) == 0 and grounded_score >= 70.0

    return data, grounded_score, grounding_warnings, flaw_warnings


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
    for model_name in FREE_ACTIVE_MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[f"MANUSCRIPT EXCERPT:\n{text_sample[:8000]}\n\nIdentify study architecture and gold-standard appraisal tool."],
                config=types.GenerateContentConfig(
                    system_instruction=discovery_instruction,
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            res_json = json.loads(response.text)
            return res_json["design_name"], res_json["tool_taxonomy"]
        except Exception:
            continue
    raise RuntimeError("Failed to discover tool schema across all free models.")


def execute_manuscript_evaluation(
    client: genai.Client,
    text_content: Optional[str] = None,
    raw_pdf_bytes: Optional[bytes] = None,
    assigned_design: Optional[str] = None,
    active_tools_map: Optional[dict] = None,
    strictness_mode: str = "Strict",
    classification_trace: Optional[dict] = None,
) -> dict:
    active_tool_taxonomy = {assigned_design: active_tools_map[assigned_design]} if (assigned_design and active_tools_map and assigned_design in active_tools_map) else (active_tools_map or SEED_TAXONOMY)
    taxonomy_json = json.dumps(active_tool_taxonomy, indent=2)

    system_instruction = f"""
    You are an objective epidemiological auditor. Apply gold-standard criteria rigorously.
    Appraisal Framework Assigned:
    {taxonomy_json}

    CORE ADJUDICATION PRINCIPLE:
    Absence of reporting must be recorded as 'Unclear Risk' or 'Not Reported' unless the tool explicitly mandates documented evidence for compliance. Do not assume good practice took place without explicit text, but do not invent flaws where reporting is simply terse.

    SPECIALIZED NOS COHORT & CASE-CONTROL SCORING (IF APPLICABLE):
    If using Newcastle-Ottawa Scale (NOS), calculate exact star counts:
    - Selection: 0 to 4 Stars
    - Comparability: 0 to 2 Stars
    - Outcome/Exposure: 0 to 3 Stars
    - Total: 0 to 9 Stars
    - Quality: Good (7-9), Fair (4-6), Poor (0-3).

    AUTHORS CITATION FORMAT:
    Extract up to 3 primary authors in Harvard format (e.g. 'Smith, J., Doe, A. and Brown, C.' or 'Smith, J. et al.' if >3).

    METHODOLOGICAL SYNTHESIS FORMATTING RULE:
    The overall synthesis must NOT be a wall of academic jargon. Provide it as 4 to 6 clear, easy-to-read bullet points written in simple, accessible language:
    - Point 1 (Study Setup): Plain-language summary of what the authors did and who was tested.
    - Point 2 (Key Strengths): What was done well and why the results are trustworthy.
    - Point 3 (Key Methodological Vulnerabilities): Major risks of bias, missing details, confounding, or unblinded assessment explained simply.
    - Point 4 (Practical Limitations): What stops us from applying this to every patient/setting (e.g. sample size, short follow-up).
    - Point 5 (Bottom-Line Reliability): Final verdict on how reliable these findings are for clinical decision-making.

    STORY-LIKE NARRATIVE ARC SUMMARY (4 TO 10 BULLETS REQUIRED):
    Provide a compelling, chronological story-like summary of the entire paper and its post-grading appraisal in plain words:
    1. Act I (The Clinical Dilemma & Spark): What unresolved problem or uncertainty prompted this inquiry?
    2. Act II (The Patient Cohort & Investigative Journey): Who was studied, how were they followed/intervened, and under what conditions?
    3. Act III (The Empirical Revelations): What did the data actually show regarding the primary endpoints?
    4. Act IV (The Auditor's Methodological Reality Check): How did the study hold up under our appraisal tool? Highlight key risks of bias, confounding, blinding/attrition caveats, or notable procedural rigour.
    5. Act V (The Translational Epilogue): What does this mean for clinical practice today? Adopt, adapt, or treat with extreme caution?

    CRITICAL: Output ONLY a valid JSON object matching:
    {{
      "paper_title": "Identified full title",
      "authors": "Up to 3 authors in Harvard format",
      "journal": "Journal Name",
      "year_published": "YYYY",
      "doi": "DOI or URL",
      "study_design": "{assigned_design or 'Identified design'}",
      "appraisal_tool_selected": "Tool name",
      "ratings_summary": {{
        "methodology_quality": "GOOD / FAIR / POOR",
        "reporting_quality": "GOOD / FAIR / POOR",
        "certainty_score": "HIGH / MODERATE / LOW / VERY LOW"
      }},
      "nos_score_breakdown": {{
        "selection_stars": 0,
        "comparability_stars": 0,
        "outcome_stars": 0,
        "total_stars": 0
      }},
      "overall_verdict": "Low Risk / Moderate Risk / High Risk / Recommended / etc.",
      "overall_synthesis": [
        "Study Setup: [Plain language summary]",
        "Methodological Strengths: [What was done well in simple terms]",
        "Risks of Bias & Red Flags: [Plain explanation of flaws or missing data]",
        "Real-World Limitations: [Who this might not apply to]",
        "Auditor Verdict: [Bottom line in plain English]"
      ],
      "practical_takeaways": [
        "Primary clinical finding or translational headline",
        "Key methodological caveat / risk of bias limitation to consider before clinical adoption",
        "Actionable practice or implementation guideline"
      ],
      "narrative_arc_summary": [
        "Act I — The Clinical Dilemma: [Describe clinical context and question]",
        "Act II — The Cohort & Journey: [Describe methodology, participants, intervention]",
        "Act III — The Data Speaks: [Describe findings and statistical endpoints]",
        "Act IV — The Critical Reality Check: [Appraisal tool post-grading critique and flaws/strengths]",
        "Act V — The Translational Epilogue: [Final clinical adoption verdict]"
      ],
      "audit_log": {{
        "classification_rationale": "Keywords and structural features identifying this design",
        "decision_pathway": [
          {{"stage": "Domain Evaluation", "observation": "Key observation", "rule_applied": "Grading principle"}}
        ],
        "verdict_derivation": "Direct justification for final scores"
      }},
      "tool_meta": {{"acronym": "Tool", "full_name": "Full Name", "purpose": "Purpose", "scoring_rules": "Rules"}},
      "domain_appraisals": [
        {{"domain_name": "Title", "evaluation_question": "Question", "rating": "Rating", "location_reference": "Section/Page", "verbatim_evidence": "Quote", "critique": "Critique"}}
      ]
    }}
    """

    clean_text = trim_manuscript_body(text_content) if text_content else ""
    contents = []
    if clean_text and len(clean_text.strip()) > 150:
        contents = [f"PAPER CONTENT:\n{clean_text}\n\nExecute appraisal and return JSON."]
    elif raw_pdf_bytes:
        contents = [
            types.Part.from_bytes(data=raw_pdf_bytes, mime_type="application/pdf"),
            f"Execute the critical appraisal for {assigned_design} based on the document and return the JSON object."
        ]
    else:
        raise ValueError("No valid manuscript text or PDF bytes supplied.")

    last_exception = None

    for model_name in FREE_ACTIVE_MODELS:
        for attempt in range(2):
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
                raw_text = response.text.strip()
                if raw_text.startswith("```"):
                    raw_text = re.sub(r"^```(?:json)?\n", "", raw_text)
                    raw_text = re.sub(r"\n```$", "", raw_text)

                parsed = json.loads(raw_text)

                if "authors" in parsed:
                    parsed["authors"] = format_harvard_authors(parsed["authors"], max_authors=3)

                if classification_trace:
                    if "audit_log" not in parsed:
                        parsed["audit_log"] = {}
                    parsed["audit_log"]["classification_trace"] = classification_trace

                return parsed
            except Exception as e:
                last_exception = e
                err_str = str(e)
                if "429" in err_str or "503" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    time.sleep(2.0)
                    continue
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                time.sleep(1.0)

    raise RuntimeError(f"Appraisal pipeline failed on all free models. Details: {str(last_exception)}")


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
    Answer user queries regarding the appraisal report and source manuscript in simple, direct language.
    Avoid overly dense jargon. Cite specific lines, sections, and quotes clearly.

    APPRAISAL JSON:
    {json.dumps(appraisal_json, indent=1)}

    SOURCE TEXT:
    {raw_manuscript_text[:20000]}
    """
    for model_name in FREE_ACTIVE_MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[f"HISTORY:\n{history_context}\n\nUSER QUESTION: {query}"],
                config=types.GenerateContentConfig(system_instruction=chat_instruction, temperature=0.1)
            )
            return response.text
        except Exception:
            continue
    return "Failed to query audit assistant across all active models."


# ==========================================
# 5. DOCUMENT GENERATION (WORD & PDF)
# ==========================================
def build_docx_report(data: dict) -> io.BytesIO:
    doc = Document()
    title = doc.add_heading(level=0)
    run_t = title.add_run(f"{APP_NAME} — Methodological Evaluation")
    run_t.font.size = Pt(20)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(120, 187, 166)

    ratings = data.get("ratings_summary", {})
    harvard_authors = format_harvard_authors(data.get("authors", "Not Specified"), max_authors=3)

    meta_table = doc.add_table(rows=8, cols=2)
    meta_table.style = "Table Grid"
    fields = [
        ("Evaluated Title:", data.get("paper_title", "N/A")),
        ("Authors (Harvard):", harvard_authors),
        ("Journal & Year:", f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})"),
        ("DOI / Reference:", data.get("doi", "N/A")),
        ("Study Architecture:", data.get("study_design", "N/A")),
        ("Tool Applied:", data.get("appraisal_tool_selected", "N/A")),
        ("Methodology / Reporting / Certainty:", f"{ratings.get('methodology_quality', 'N/A')} / {ratings.get('reporting_quality', 'N/A')} / {ratings.get('certainty_score', 'N/A')}"),
        ("Evidence Groundedness:", f"{data.get('groundedness_score', 100.0)}%"),
    ]
    for idx, (label, val) in enumerate(fields):
        row = meta_table.rows[idx]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EBF5F2"/>')
        row.cells[0]._tc.get_or_add_tcPr().append(shd)
        row.cells[0].paragraphs[0].add_run(label).bold = True
        row.cells[1].paragraphs[0].add_run(str(val))

    doc.add_paragraph("")

    takeaways = data.get("practical_takeaways", [])
    if takeaways:
        doc.add_heading("Practical Takeaways from the Paper (At-A-Glance)", level=2)
        for idx, pt in enumerate(takeaways, 1):
            p_t = doc.add_paragraph(style='List Bullet')
            p_t.add_run(f"Takeaway #{idx}: ").bold = True
            p_t.add_run(pt)

    story_arc = data.get("narrative_arc_summary", [])
    if story_arc:
        doc.add_heading("Narrative Story Arc & Post-Grading Appraisal Journey", level=2)
        for idx, pt in enumerate(story_arc, 1):
            p_s = doc.add_paragraph(style='List Number')
            parts = pt.split(":", 1)
            if len(parts) == 2:
                p_s.add_run(parts[0] + ":").bold = True
                p_s.add_run(parts[1])
            else:
                p_s.add_run(pt)

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
    doc.add_heading("2. Methodological Synthesis & Discussion (Key Points)", level=1)
    synth = data.get("overall_synthesis", [])
    if isinstance(synth, list):
        for pt in synth:
            parts = pt.split(":", 1)
            p_synth = doc.add_paragraph(style='List Bullet')
            if len(parts) == 2:
                p_synth.add_run(parts[0] + ":").bold = True
                p_synth.add_run(parts[1])
            else:
                p_synth.add_run(pt)
    else:
        doc.add_paragraph(str(synth))

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_pdf_report(data: dict) -> io.BytesIO:
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
        "T",
        parent=styles["Heading1"],
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#4A8574"),
        fontName="Helvetica-Bold",
        spaceAfter=4,
    )
    h2_style = ParagraphStyle(
        "H2",
        parent=styles["Heading2"],
        fontSize=12,
        leading=16,
        textColor=colors.HexColor("#2C5247"),
        fontName="Helvetica-Bold",
        spaceBefore=10,
        spaceAfter=6,
    )
    body_style = ParagraphStyle(
        "B",
        parent=styles["Normal"],
        fontSize=8.5,
        leading=11.5,
        fontName="Helvetica",
    )
    bold_style = ParagraphStyle(
        "BB",
        parent=body_style,
        fontName="Helvetica-Bold",
    )
    header_style = ParagraphStyle(
        "H",
        parent=body_style,
        fontName="Helvetica-Bold",
        textColor=colors.white,
    )
    takeaway_style = ParagraphStyle(
        "TA",
        parent=body_style,
        fontSize=8.5,
        leading=12,
        textColor=colors.HexColor("#1D3B32"),
    )

    story = [
        Paragraph(f"{APP_NAME} — Methodological Evaluation", title_style),
        Paragraph("<font color='#666666'><i>Standardized critical appraisal & translational synthesis</i></font>", body_style),
        Spacer(1, 8),
    ]

    ratings = data.get("ratings_summary", {})
    methodology_tier = str(ratings.get("methodology_quality", "")).upper()

    if "GOOD" in methodology_tier:
        active_tier = "GREEN"
    elif "POOR" in methodology_tier:
        active_tier = "RED"
    elif "FAIR" in methodology_tier:
        active_tier = "YELLOW"
    else:
        v = str(data.get("overall_verdict", "")).lower()
        if any(w in v for w in ["low risk", "recommended", "high quality"]):
            active_tier = "GREEN"
        elif any(w in v for w in ["high risk", "serious", "critical"]):
            active_tier = "RED"
        else:
            active_tier = "YELLOW"

    harvard_authors = format_harvard_authors(data.get("authors", "Not Specified"), max_authors=3)

    meta_data = [
        [Paragraph("Evaluated Title:", bold_style), Paragraph(data.get("paper_title", "N/A"), body_style)],
        [Paragraph("Authors (Harvard):", bold_style), Paragraph(harvard_authors, body_style)],
        [Paragraph("Journal & Year:", bold_style), Paragraph(f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})", body_style)],
        [Paragraph("Study Architecture:", bold_style), Paragraph(data.get("study_design", "N/A"), body_style)],
        [Paragraph("Tool Applied:", bold_style), Paragraph(data.get("appraisal_tool_selected", "N/A"), body_style)],
        [Paragraph("Quality Triad:", bold_style), Paragraph(f"Methodology: <b>{ratings.get('methodology_quality', 'N/A')}</b> | Reporting: <b>{ratings.get('reporting_quality', 'N/A')}</b> | Certainty: <b>{ratings.get('certainty_score', 'N/A')}</b>", body_style)],
        [Paragraph("Overall Verdict:", bold_style), Paragraph(f"<font color='#9E2A2B'><b>{data.get('overall_verdict', 'N/A')}</b></font>", body_style)],
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
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>Methodological Trustworthiness Tier (At-A-Glance):</b>", bold_style))
    story.append(Spacer(1, 4))

    r_text = (
        "<font color='white'><b>✓ HIGH RISK / POOR</b><br/><b>[ ✓ SELECTED TIER ]</b><br/>(Critical Deficits / High Bias)</font>"
        if active_tier == "RED"
        else "<font color='#7A7A7A'>HIGH RISK / POOR<br/>(Critical Deficits / High Bias)</font>"
    )
    y_text = (
        "<font color='white'><b>✓ MODERATE / FAIR</b><br/><b>[ ✓ SELECTED TIER ]</b><br/>(Some Methodological Concerns)</font>"
        if active_tier == "YELLOW"
        else "<font color='#7A7A7A'>MODERATE / FAIR<br/>(Some Methodological Concerns)</font>"
    )
    g_text = (
        "<font color='white'><b>✓ ROBUST / GOOD</b><br/><b>[ ✓ SELECTED TIER ]</b><br/>(Low Risk / High Methodological Rigour)</font>"
        if active_tier == "GREEN"
        else "<font color='#7A7A7A'>ROBUST / GOOD<br/>(Low Risk / High Methodological Rigour)</font>"
    )

    gauge_data = [[
        Paragraph(r_text, body_style),
        Paragraph(y_text, body_style),
        Paragraph(g_text, body_style),
    ]]
    t_gauge = Table(gauge_data, colWidths=[2.46 * inch, 2.46 * inch, 2.46 * inch])

    gauge_styles = [
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#C0392B') if active_tier == "RED" else colors.HexColor('#FADBD8')),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#D4AC0D') if active_tier == "YELLOW" else colors.HexColor('#FCF3CF')),
        ('BACKGROUND', (2, 0), (2, 0), colors.HexColor('#229954') if active_tier == "GREEN" else colors.HexColor('#D5F5E3')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]

    active_col = 0 if active_tier == "RED" else (1 if active_tier == "YELLOW" else 2)
    border_color = (
        colors.HexColor('#78281F') if active_tier == "RED" else (
            colors.HexColor('#7D6608') if active_tier == "YELLOW" else colors.HexColor('#145A32')
        )
    )
    gauge_styles.append(('BOX', (active_col, 0), (active_col, 0), 2.5, border_color))
    t_gauge.setStyle(TableStyle(gauge_styles))
    story.append(t_gauge)
    story.append(Spacer(1, 10))

    takeaways = data.get("practical_takeaways", [])
    if takeaways:
        story.append(Paragraph("Practical Takeaways from the Paper (At-A-Glance):", h2_style))
        takeaway_rows = []
        for idx, item in enumerate(takeaways):
            takeaway_rows.append([
                Paragraph(f"<b>Key Takeaway #{idx + 1}:</b>", bold_style),
                Paragraph(item, takeaway_style)
            ])
        t_takeaways = Table(takeaway_rows, colWidths=[1.8 * inch, 5.6 * inch])
        t_takeaways.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F4F9F7')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#78BBA6')),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(t_takeaways)
        story.append(Spacer(1, 10))

    story_arc = data.get("narrative_arc_summary", [])
    if story_arc:
        story.append(Paragraph("Narrative Story Arc & Post-Grading Appraisal Journey:", h2_style))
        story_rows = []
        for idx, point in enumerate(story_arc):
            story_rows.append([
                Paragraph(f"<b>Step #{idx + 1}:</b>", bold_style),
                Paragraph(point, takeaway_style)
            ])
        t_story = Table(story_rows, colWidths=[1.8 * inch, 5.6 * inch])
        t_story.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F7FBF9')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#A3D9C9')),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E1F2ED')),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ]))
        story.append(t_story)
        story.append(Spacer(1, 10))

    story.append(Paragraph(f"1. Detailed Domain Appraisal ({data.get('appraisal_tool_selected', 'Tool')})", h2_style))
    domain_data = [[
        Paragraph("Domain & Criterion", header_style),
        Paragraph("Score", header_style),
        Paragraph("Evidence", header_style),
        Paragraph("Critique", header_style)
    ]]
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
    story.append(Spacer(1, 10))

    # --- 2. Methodological Synthesis (Formatted in Clear, Distinct Points) ---
    story.append(Paragraph("2. Methodological Synthesis & Discussion (Plain-English Points)", h2_style))
    synth = data.get("overall_synthesis", [])
    synth_rows = []
    if isinstance(synth, list):
        for idx, pt in enumerate(synth, 1):
            synth_rows.append([
                Paragraph(f"<b>Point {idx}:</b>", bold_style),
                Paragraph(pt, takeaway_style)
            ])
    else:
        synth_rows.append([
            Paragraph("<b>Summary:</b>", bold_style),
            Paragraph(str(synth), takeaway_style)
        ])

    t_synth = Table(synth_rows, colWidths=[1.8 * inch, 5.6 * inch])
    t_synth.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F4F9F7')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_synth)
    story.append(Spacer(1, 10))

    story.append(Paragraph("3. Appraisal Framework & Standards Appendix", h2_style))
    tool_meta = data.get("tool_meta", {})
    acronym = tool_meta.get("acronym") or data.get("appraisal_tool_selected", "Tool")
    full_name = tool_meta.get("full_name") or "Standard Critical Appraisal Instrument"
    purpose = tool_meta.get("purpose") or "Standardized epidemiological instrument used to assess study rigour and bias."
    scoring_rules = tool_meta.get("scoring_rules") or "Standard qualitative rating or numerical score threshold."

    design_match = SEED_TAXONOMY.get(data.get("study_design", ""))
    if design_match:
        if not full_name or full_name == "Standard Critical Appraisal Instrument":
            full_name = design_match.get("full_name", full_name)
        if not purpose or "Standardized epidemiological" in purpose:
            purpose = design_match.get("purpose", purpose)
        if not scoring_rules or "Standard qualitative" in scoring_rules:
            scoring_rules = design_match.get("scale", scoring_rules)

    appendix_rows = [
        [Paragraph("Tool Acronym / Short Name:", bold_style), Paragraph(acronym, body_style)],
        [Paragraph("Full Name of Instrument:", bold_style), Paragraph(f"<b>{full_name}</b>", body_style)],
        [Paragraph("Methodological Purpose:", bold_style), Paragraph(purpose, body_style)],
        [Paragraph("Scoring Scale & Thresholds:", bold_style), Paragraph(scoring_rules, body_style)],
        [Paragraph("Clinical Interpretation:", bold_style), Paragraph("Findings appraised using this framework should be balanced against the patient population. Low risk of bias indicates strong internal validity, while highlighted deficiencies warrant caution prior to clinical implementation.", body_style)],
    ]
    t_app = Table(appendix_rows, colWidths=[1.8 * inch, 5.6 * inch])
    t_app.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EBF5F2')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
    ]))
    story.append(t_app)

    doc.build(story)
    buf.seek(0)
    return buf


# ==========================================
# 6. STREAMLIT APP & UI INTERACTION
# ==========================================
st.set_page_config(page_title=APP_NAME, page_icon="🐱", layout="wide")
init_db()

st.markdown(
    """
<style>
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
    div[data-testid="stChatInput"] button {
        background-color: #78BBA6 !important;
        color: white !important;
    }
    div[data-testid="stChatInput"] textarea:focus {
        border-color: #78BBA6 !important;
        box-shadow: 0 0 0 1px #78BBA6 !important;
    }
    .metric-card {
        background-color: rgba(120, 187, 166, 0.08);
        border: 1px solid rgba(120, 187, 166, 0.3);
        border-radius: 8px;
        padding: 12px;
    }
    div[role="radiogroup"] > label > div:first-child[aria-checked="true"] {
        background-color: #78BBA6 !important;
        border-color: #78BBA6 !important;
    }
</style>
""",
    unsafe_allow_html=True,
)

CORRECT_PASSCODE = st.secrets.get("APP_PASSCODE", os.environ.get("APP_PASSCODE", "324598"))

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

st.title(f"🐱 {APP_NAME}")
st.caption("Epidemiological Critical Appraisal & Evidence Verification Engine")

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
        with st.spinner("Extracting text (pypdf + OCR fallback)..."):
            paper_text = extract_text_from_pdf_bytes(raw_pdf_bytes)
        if paper_text:
            st.session_state["raw_paper_text"] = paper_text
            st.success(f"Parsed {len(paper_text):,} characters successfully.")
        else:
            st.warning("Could not extract readable text. Image OCR fallback pipeline active.")

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

ensemble_winner = "Observational Study (Cohort)"
confidence_score = 1.0
contradiction_flag = None
ensemble_data = {}

if paper_text:
    client_preflight = genai.Client(api_key=api_key) if api_key else None
    ensemble_data = run_four_source_ensemble(paper_text, client_preflight, available_designs)
    ensemble_winner = ensemble_data["winner"]
    confidence_score = ensemble_data["confidence"]
    contradiction_flag = detect_classification_contradiction(paper_text, ensemble_winner)

st.markdown("---")

col_d1, col_d2 = st.columns([3, 1])
with col_d1:
    default_idx = (
        available_designs.index(ensemble_winner)
        if ensemble_winner in available_designs
        else (len(available_designs) - 1)
    )
    selected_design = st.selectbox(
        "Target Study Architecture & Framework:",
        options=available_designs,
        index=default_idx,
    )
with col_d2:
    conf_badge = f"{int(confidence_score * 100)}%"
    st.metric("Ensemble Confidence", conf_badge)

if ensemble_data.get("evidence"):
    override_badge = " [Supreme Signal Override Applied]" if ensemble_data.get("override_applied") else ""
    with st.expander(f"📋 Classification Evidence & Trace Rationale{override_badge}"):
        st.markdown(f"**Selected Design:** `{ensemble_winner}`  \n**Consensus Confidence:** `{int(confidence_score * 100)}%`")
        for pt in ensemble_data["evidence"]:
            st.markdown(f"- {pt}")
        if ensemble_data.get("trace"):
            st.json(ensemble_data["trace"])

if contradiction_flag:
    st.error(f"🛑 **Critical Architecture Conflict Detected:**\n\n{contradiction_flag}\n\n*Appraisal halted to prevent scoring an empirical study with a review framework (or vice versa). Please manually correct the dropdown selection above.*")
    st.stop()

if confidence_score < 0.80:
    st.info("ℹ️ **Confidence Gating (<80%):** Multiple designs detected across ensemble sources. Please confirm or adjust the dropdown selection before proceeding.")

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
                    target_design = "Observational Study (Cohort)"

        with st.spinner(f"Executing rigorous appraisal under {target_design}..."):
            try:
                raw_data = execute_manuscript_evaluation(
                    client,
                    text_content=paper_text,
                    raw_pdf_bytes=raw_pdf_bytes if not paper_text else None,
                    assigned_design=target_design,
                    active_tools_map=active_tools,
                    strictness_mode=strictness_choice,
                    classification_trace=ensemble_data.get("trace"),
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

    ratings = data.get("ratings_summary", {})
    harvard_display_authors = format_harvard_authors(data.get("authors", "Not Specified"), max_authors=3)
    st.markdown(f"**Authors (Harvard):** {harvard_display_authors}")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Architecture", data.get("study_design"))
    with c2:
        st.metric("Tool Applied", data.get("appraisal_tool_selected"))
    with c3:
        st.metric("Methodology Quality", ratings.get("methodology_quality", "N/A"))
    with c4:
        st.metric("Certainty Score", ratings.get("certainty_score", "N/A"))

    nos_stars = data.get("nos_score_breakdown")
    if nos_stars and "total_stars" in nos_stars and nos_stars["total_stars"] > 0:
        st.info(f"⭐ **Newcastle-Ottawa Scale Star Count:** {nos_stars.get('total_stars', 0)}/9 Stars (Selection: {nos_stars.get('selection_stars', 0)}★ | Comparability: {nos_stars.get('comparability_stars', 0)}★ | Outcome: {nos_stars.get('outcome_stars', 0)}★)")

    takeaways = data.get("practical_takeaways", [])
    if takeaways:
        st.markdown("### 💡 Practical Takeaways from the Paper (At-A-Glance)")
        for idx, item in enumerate(takeaways, 1):
            st.markdown(f"**{idx}.** {item}")

    story_arc = data.get("narrative_arc_summary", [])
    if story_arc:
        st.markdown("### 📖 Narrative Story Arc & Translational Journey")
        with st.container():
            for idx, pt in enumerate(story_arc, 1):
                parts = pt.split(":", 1)
                if len(parts) == 2:
                    st.markdown(f"**{idx}. {parts[0].strip()}:** {parts[1].strip()}")
                else:
                    st.markdown(f"**{idx}.** {pt}")

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

    st.subheader("📝 Methodological Synthesis & Discussion (Plain-English Points)")
    synth = data.get("overall_synthesis", [])
    if isinstance(synth, list):
        for idx, pt in enumerate(synth, 1):
            parts = pt.split(":", 1)
            if len(parts) == 2:
                st.markdown(f"**• {parts[0].strip()}:** {parts[1].strip()}")
            else:
                st.markdown(f"**•** {pt}")
    else:
        st.info(str(synth))

    audit = data.get("audit_log", {})
    if audit:
        with st.expander("🔍 Methodological Decision Log & Audit Trail"):
            st.markdown("#### Study Classification Rationale")
            st.write(audit.get("classification_rationale", "N/A"))
            if audit.get("classification_trace"):
                st.markdown("#### 4-Source Ensemble Trace")
                st.json(audit["classification_trace"])
            steps = audit.get("decision_pathway", [])
            if steps:
                st.markdown("#### Decision Pathway & Applied Rules")
                st.dataframe(steps, use_container_width=True)
            st.markdown("#### Verdict Derivation")
            st.write(audit.get("verdict_derivation", "N/A"))

    c_doc, c_pdf, c_log = st.columns(3)
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
    with c_log:
        audit_payload = json.dumps(
            {
                "paper_title": data.get("paper_title"),
                "authors_harvard": harvard_display_authors,
                "study_design": data.get("study_design"),
                "appraisal_tool_selected": data.get("appraisal_tool_selected"),
                "ratings_summary": data.get("ratings_summary", {}),
                "nos_score_breakdown": data.get("nos_score_breakdown", {}),
                "overall_verdict": data.get("overall_verdict"),
                "overall_synthesis": data.get("overall_synthesis", []),
                "practical_takeaways": data.get("practical_takeaways", []),
                "narrative_arc_summary": data.get("narrative_arc_summary", []),
                "audit_log": data.get("audit_log", {}),
                "checksum_flaws": data.get("checksum_flaws", []),
                "groundedness_score": data.get("groundedness_score"),
            },
            indent=2,
        )
        st.download_button(
            "📥 Download Audit Log (.json)",
            data=audit_payload,
            file_name=format_harvard_filename(data, "json"),
            mime="application/json",
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
