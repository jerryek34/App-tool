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
ASSISTANT_NAME = "Ki's Assistant"
MIN_CONFIDENCE_MARGIN = 15

FREE_ACTIVE_MODELS = [
    "gemini-2.0-flash",
    "gemini-2.0-flash-lite",
    "gemini-1.5-flash",
]


class ClassificationConsistencyError(Exception):
    """Raised ONLY on objective, impossible structural contradictions (Hard Fails)."""
    pass


def get_connection():
    conn = sqlite3.connect(DB_NAME, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


# ==========================================
# 1. LEGALITY MATRIX, SIGNATURES & TAXONOMY
# ==========================================
LEGAL_PAIRS = {
    "Clinical Practice Guideline / Consensus Statement": ["AGREE II"],
    "Systematic Review / Meta-Analysis": ["AMSTAR 2"],
    "Narrative Review": ["SANRA"],
    "Randomized Controlled Trial (RCT)": ["Cochrane RoB 2"],
    "Non-Randomized Intervention / Quasi-Experimental": ["ROBINS-I"],
    "Observational Study (Cohort)": ["Newcastle-Ottawa Scale (NOS) Cohort", "NOS"],
    "Observational Study (Case-Control)": ["Newcastle-Ottawa Scale (NOS) Case-Control", "NOS"],
    "Cross-Sectional Study": ["AXIS"],
    "Diagnostic Accuracy Study": ["QUADAS-2"],
    "Prediction Model Study": ["PROBAST"],
    "Qualitative Study": ["CASP Qualitative / COREQ", "CASP"],
}

FRAMEWORK_SIGNATURES = {
    "Clinical Practice Guideline / Consensus Statement": {
        "source_must_have_any": [r"\brecommendation\s*\d+", r"\bconsensus\b", r"\bguideline\b", r"\bgrade\s+of\s+recommendation\b"],
        "source_must_not_have": [r"\bpatients\s+were\s+randomized\s+to\b", r"\bwe\s+retrospectively\s+reviewed\s+the\s+medical\s+records\b"],
        "expected_domain_keywords": ["scope", "stakeholder", "rigour", "clarity", "applicability", "editorial"],
        "tool": "AGREE II"
    },
    "Systematic Review / Meta-Analysis": {
        "source_must_have_any": [r"\bsearch\s+strategy\b", r"\bprisma\b", r"\bdatabases?\b", r"\bmeta-analysis\b", r"\bsystematic\s+review\b"],
        "source_must_not_have": [r"\bgrade\s+of\s+recommendation\s+gpp\b", r"\bstrong\s+consensus\s+\(\d+%\s+agreement\)\b"],
        "expected_domain_keywords": ["pico", "search", "duplicate", "excluded", "rob", "meta-analytic", "publication bias", "conflict"],
        "tool": "AMSTAR 2"
    },
    "Narrative Review": {
        "source_must_have_any": [r"\breview\b", r"\bliterature\b", r"\boverview\b"],
        "source_must_not_have": [r"\bpatients\s+were\s+randomized\b"],
        "expected_domain_keywords": ["justification", "aims", "search", "referencing", "reasoning", "presentation"],
        "tool": "SANRA"
    },
    "Randomized Controlled Trial (RCT)": {
        "source_must_have_any": [r"\brandomi[zs]ed\b", r"\ballocat(?:ed|ion)\b", r"\btrial\b", r"\bplacebo\b"],
        "source_must_not_have": [r"\bsystematic\s+review\s+and\s+meta-analysis\b"],
        "expected_domain_keywords": ["randomization", "deviations", "missing", "measurement", "selection"],
        "tool": "Cochrane RoB 2"
    },
    "Non-Randomized Intervention / Quasi-Experimental": {
        "source_must_have_any": [r"\bintervention\b", r"\bpatients\b", r"\btreated\b", r"\bgroup\b"],
        "source_must_not_have": [r"\bprisma\s+flowchart\b"],
        "expected_domain_keywords": ["confounding", "selection", "classification", "deviations", "missing", "measurement", "reported"],
        "tool": "ROBINS-I"
    },
    "Observational Study (Cohort)": {
        "source_must_have_any": [r"\bcohort\b", r"\bpatients\b", r"\bfollow-up\b", r"\bretrospective\b", r"\bprospective\b", r"\bconsecutive\b"],
        "source_must_not_have": [r"\bgrade\s+of\s+recommendation\s+gpp\b"],
        "expected_domain_keywords": ["selection", "comparability", "outcome"],
        "tool": "Newcastle-Ottawa Scale (NOS) Cohort"
    },
    "Observational Study (Case-Control)": {
        "source_must_have_any": [r"\bcases\b", r"\bcontrols\b", r"\bodds\s+ratio\b", r"\bcase-control\b"],
        "source_must_not_have": [r"\bgrade\s+of\s+recommendation\s+gpp\b"],
        "expected_domain_keywords": ["selection", "comparability", "exposure"],
        "tool": "Newcastle-Ottawa Scale (NOS) Case-Control"
    },
    "Cross-Sectional Study": {
        "source_must_have_any": [r"\bcross-sectional\b", r"\bsurvey\b", r"\bprevalence\b", r"\bquestionnaire\b"],
        "source_must_not_have": [r"\brandomized\s+controlled\s+trial\b"],
        "expected_domain_keywords": ["aims", "selection", "measurement", "confounding", "ethical"],
        "tool": "AXIS"
    },
    "Prediction Model Study": {
        "source_must_have_any": [r"\bmodel\b", r"\bpredict(?:ion|ors)\b", r"\bvalidation\b", r"\bauc\b", r"\bc-index\b", r"\bnomogram\b"],
        "source_must_not_have": [r"\bconsensus\s+guideline\b"],
        "expected_domain_keywords": ["participants", "predictors", "outcome", "analysis"],
        "tool": "PROBAST"
    },
    "Diagnostic Accuracy Study": {
        "source_must_have_any": [r"\bsensitivity\b", r"\bspecificity\b", r"\bindex\s+test\b", r"\breference\s+standard\b", r"\bdiagnostic\b"],
        "source_must_not_have": [r"\bconsensus\s+guideline\b"],
        "expected_domain_keywords": ["patient selection", "index test", "reference standard", "flow"],
        "tool": "QUADAS-2"
    },
    "Qualitative Study": {
        "source_must_have_any": [r"\bqualitative\b", r"\binterview\b", r"\bthematic\b", r"\bthemes\b", r"\bfocus\s+group\b"],
        "source_must_not_have": [r"\brandomized\s+controlled\s+trial\b"],
        "expected_domain_keywords": ["appropriateness", "sampling", "reflexivity", "ethical", "analysis", "findings"],
        "tool": "CASP Qualitative / COREQ"
    },
}

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
        "scale": "Yes | Partial Yes | No across 16 specific items. Adjudicated into High, Moderate, Low, or Critically Low Confidence based on critical domain flaws.",
        "score_interpretation": (
            "• 'Yes': Complete compliance with methodological standard.\n"
            "• 'Partial Yes': Partial adherence with minor reporting omission.\n"
            "• 'No': Absent compliance or unaddressed critical risk of bias."
        ),
        "scoring_guide": (
            "• High Confidence: No or one non-critical weakness; provides an accurate, comprehensive summary of available studies.\n"
            "• Moderate Confidence: More than one non-critical weakness, but zero critical domain flaws; summary remains accurate.\n"
            "• Low Confidence: Exactly one critical domain flaw (with or without non-critical weaknesses); review findings may not be dependable.\n"
            "• Critically Low Confidence: More than one critical flaw; review should not be relied upon for clinical practice decisions."
        )
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
        "scale": "Item scale: 0 (No / Not Met), 1 (Partially Met), 2 (Yes / Completely Met). Total score out of 12 points.",
        "score_interpretation": (
            "• Score 0 (No / Not Met): The criterion is completely absent, unaddressed, or inadequately documented.\n"
            "• Score 1 (Partially Met): The requirement is partially addressed or vague, but key methodological details are omitted.\n"
            "• Score 2 (Yes / Completely Met): Criterion is fully satisfied with explicit transparency and documented rigor."
        ),
        "scoring_guide": (
            "• 6 items scored on a 3-point scale (0, 1, or 2), yielding a maximum total of 12 points.\n"
            "• High Quality: 9 to 12 points (indicates strong literature referencing, scientific balance, and explicit search reporting).\n"
            "• Moderate Quality: 6 to 8 points (adequate presentation but noticeable gaps in search transparency or data synthesis).\n"
            "• Low Quality: 0 to 5 points (insufficient literature substantiation or substantial risk of advocacy/narrative bias)."
        )
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
        "scale": "Items scored 1 (Strongly Disagree / Absent) to 7 (Strongly Agree / Exemplary). Standardized % >= 70% indicates high quality.",
        "score_interpretation": (
            "• Score 1 (Strongly Disagree): Relevant criteria are absent or poorly reported.\n"
            "• Score 2–6: Partial fulfillment with increasing levels of completeness and rigor.\n"
            "• Score 7 (Strongly Agree): Criterion is addressed with exceptional methodological transparency."
        ),
        "scoring_guide": (
            "• Scaled domain score = (Obtained Score - Minimum Possible) / (Maximum Possible - Minimum Possible) × 100%.\n"
            "• Strongly Recommended: Domain 3 (Rigour of Development) >= 70% and most remaining domains >= 70%.\n"
            "• Recommended with Modifications: Most domains score between 50% and 69%.\n"
            "• Not Recommended: Domain 3 (Rigour) < 50%, reflecting substandard development procedures."
        )
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
        "scale": "Domain ratings: Low Risk | Some Concerns | High Risk. A single High Risk domain mandates an overall High Risk rating.",
        "score_interpretation": (
            "• 'Low Risk': Robust trial design with low potential for systematic bias.\n"
            "• 'Some Concerns': Methodological question marks or omissions that may modestly compromise validity.\n"
            "• 'High Risk': Systematic errors or unblinded threats that introduce significant bias into findings."
        ),
        "scoring_guide": (
            "• Low Risk of Bias: Study is judged at low risk across all 5 key trial domains.\n"
            "• Some Concerns: Study has some concerns in at least one domain, but no domain is judged at high risk.\n"
            "• High Risk of Bias: Judged at high risk in at least one domain, or multiple domains with 'some concerns' that substantially reduce confidence."
        )
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
        "score_interpretation": (
            "• 'Low Risk': Comparable to a high-quality randomized trial with minimal bias.\n"
            "• 'Moderate Risk': Sound for a non-randomized study; some residual confounding is likely.\n"
            "• 'Serious Risk': Methodological flaws substantially compromise confidence in effect estimates.\n"
            "• 'Critical Risk': Problematic design flaws render the evidence clinically uninformative."
        ),
        "scoring_guide": (
            "• Low Risk: Comparable to a well-conducted randomized controlled trial.\n"
            "• Moderate Risk: Sound methodology for a non-randomized study; some residual confounding is likely.\n"
            "• Serious Risk: Important confounding, unbalanced co-interventions, or unblinded assessment present; interpret with caution.\n"
            "• Critical Risk: Flaws are so severe that the study provides no dependable evidence."
        )
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
        "score_interpretation": (
            "• ★ Awarded (Yes / Compliant): Criterion was explicitly met using validated, objective methods.\n"
            "• No ★ (No / Deficient): Item was omitted, self-reported without verification, or failed retention thresholds."
        ),
        "scoring_guide": (
            "• Selection (Max 4 ★): Evaluates exposed representativeness, non-exposed selection, exposure verification, and absence of outcome at baseline.\n"
            "• Comparability (Max 2 ★): 1 ★ for controlling primary confounder (e.g. age/sex), 1 ★ for controlling secondary critical confounders.\n"
            "• Outcome (Max 3 ★): 1 ★ for blinded/record outcome assessment, 1 ★ for sufficient duration, 1 ★ for follow-up retention > 80%.\n"
            "• Thresholds: 7–9 Stars = Good Quality; 4–6 Stars = Fair Quality; 0–3 Stars = Poor Quality."
        )
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
        "score_interpretation": (
            "• ★ Awarded (Yes / Compliant): Specific case/control definition and blinded ascertainment satisfied.\n"
            "• No ★ (No / Deficient): Unvalidated definitions, surrogate controls, or unblinded interviews."
        ),
        "scoring_guide": (
            "• Selection (Max 4 ★): Independent case validation, representativeness, control selection, and no history of endpoint.\n"
            "• Comparability (Max 2 ★): Matching or multivariable adjustment for primary and additional covariates.\n"
            "• Exposure (Max 3 ★): Blinded exposure ascertainment, identical interview method, and comparable non-response rates.\n"
            "• Thresholds: 7–9 Stars = Good Quality; 4–6 Stars = Fair Quality; 0–3 Stars = Poor Quality."
        )
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
        "scale": "Scored across 20 items: Yes | No | Don't Know. High Quality (>=15 Yes) | Moderate (10-14 Yes) | Low (<10 Yes).",
        "score_interpretation": (
            "• 'Yes': Clear, transparent fulfillment of epidemiological standard.\n"
            "• 'No': Item unfulfilled or evidence of significant bias.\n"
            "• 'Don't Know': Ambiguous or insufficient reporting to make a definitive judgment."
        ),
        "scoring_guide": (
            "• 20 individual cross-sectional criteria scored as Yes, No, or Don't Know.\n"
            "• High Quality (>=15 Yes): Low risk of sampling, non-response, or reporting bias.\n"
            "• Moderate Quality (10–14 Yes): Acceptable reporting, but non-response or confounding concerns exist.\n"
            "• Low Quality (<10 Yes): Substantial vulnerability to sampling bias or unvalidated survey instruments."
        )
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
        "score_interpretation": (
            "• 'Low Risk': Sound participant selection, blinded predictors, pre-specified outcomes, and adequate event-per-variable ratio.\n"
            "• 'High Risk': High risk of overfitting, optimistic performance, or misleading discrimination.\n"
            "• 'Unclear Risk': Key validation or statistical modeling details omitted."
        ),
        "scoring_guide": (
            "• Low Risk: All 4 domains (Participants, Predictors, Outcome, Analysis) judged low risk.\n"
            "• High Risk: At least one domain evaluated as high risk, indicating substantial risk of overfitting or optimistic performance."
        )
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
        "score_interpretation": (
            "• 'Low Risk': Strict blinded test sequence, valid reference standard, and appropriate diagnostic timing.\n"
            "• 'High Risk': Presence of verification bias, unblinded interpretation, or inappropriate spectrum exclusions.\n"
            "• 'Unclear Risk': Missing timing intervals or test execution details."
        ),
        "scoring_guide": (
            "• Low Risk: All 4 domains judged low risk; strong diagnostic validity.\n"
            "• High Risk: One or more domains show unblinded interpretation, verification bias, or spectrum bias.\n"
            "• Unclear Risk: Incomplete reporting of test execution sequence or blinding."
        )
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
        "score_interpretation": (
            "• 'Yes': Thoroughly demonstrated with qualitative reflexivity and thematic coding.\n"
            "• 'Can't Tell': Unclear reporting of saturation, interview depth, or positionality.\n"
            "• 'No': Substandard qualitative reporting or complete omission of researcher positionality."
        ),
        "scoring_guide": (
            "• Assesses 10 key qualitative validity questions.\n"
            "• High Quality: Research reflexivity, thematic saturation, and verbatim quotes fully demonstrated.\n"
            "• Moderate/Low Quality: Inadequate description of data analysis or omission of researcher positionality."
        )
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


# ==========================================
# 2. CROSSREF-STYLE HEURISTIC TITLE & SECTION PARSER
# ==========================================
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


@st.cache_data(show_spinner=False)
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


@st.cache_data(show_spinner=False)
def extract_text_from_url(raw_input: str) -> str:
    target = raw_input.strip()
    if not target.startswith("http://) and not target.startswith("https://):
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


def find_probable_title_and_header(text: str) -> Tuple[str, str]:
    """
    Crossref-inspired heuristic title detector.
    Evaluates the top 45 lines of the document, penalizing publisher boilerplate
    (Elsevier, ScienceDirect, DOI, ISSN, affiliations, emails) and rewarding
    title-length lines and study-architecture keywords.
    """
    lines = [line.strip() for line in text[:4500].split("\n") if line.strip()]
    top_lines = lines[:45]

    boilerplate_patterns = re.compile(
        r"(?i)(?:sciencedirect|elsevier|springer|wiley|contents\s+lists\s+available|"
        r"journal\s+homepage|http://|https://|doi\.org|10\.\d{4}/|issn\s+\d|"
        r"received\s+\d|accepted\s+\d|available\s+online|copyright|all\s+rights\s+reserved|"
        r"corresponding\s+author|e-mail\s+address|department\s+of|university\s+of|"
        r"hospital|institute\s+of|school\s+of\s+medicine|faculty\s+of|--- \[page)"
    )

    scored_candidates = []
    category_tags = []

    for idx, line in enumerate(top_lines):
        # Capture short category headers like "ESPEN Guideline" or "Original Article"
        if len(line) < 40 and re.search(r"(?i)\b(?:guideline|consensus|original\s+article|review\s+article|clinical\s+trial)\b", line):
            category_tags.append(line)

        if boilerplate_patterns.search(line):
            continue

        # Skip author lists (many commas and superscripts/initials)
        if line.count(",") >= 3:
            continue

        score = 0
        line_len = len(line)
        if 25 <= line_len <= 220:
            score += 35
        elif 15 <= line_len < 25:
            score += 15
        else:
            score -= 20

        # Position bonus (usually lines 2 to 18)
        if 1 <= idx <= 18:
            score += 20

        # Study design keyword bonus in title
        if re.search(r"(?i)\b(?:guideline|consensus|systematic\s+review|meta-analysis|randomized|trial|cohort|case-control|cross-sectional|prevalence|complications)\b", line):
            score += 30

        # Penalty for ending with a period (sentences vs titles)
        if line.endswith("."):
            score -= 10

        scored_candidates.append((score, idx, line))

    scored_candidates.sort(key=lambda x: x[0], reverse=True)
    best_title = scored_candidates[0][2] if scored_candidates else (top_lines[0] if top_lines else "")

    # Also grab the line immediately following best_title if it looks like a subtitle continuation
    if scored_candidates:
        best_idx = scored_candidates[0][1]
        if best_idx + 1 < len(top_lines):
            next_line = top_lines[best_idx + 1]
            if 10 < len(next_line) < 150 and not boilerplate_patterns.search(next_line) and next_line.count(",") < 2:
                best_title = f"{best_title} {next_line}"

    combined_title_zone = " | ".join(category_tags + [best_title]) + "\n" + "\n".join(top_lines[:15])
    return best_title, combined_title_zone


def parse_manuscript_sections(text: str) -> Dict[str, str]:
    """
    Section-aware parser using heuristic title detection and explicit regex boundary anchors.
    """
    detected_title, title_zone = find_probable_title_and_header(text)
    lower = text.lower()

    # Locate Abstract / Summary
    abs_match = re.search(r"\n(?:\s*)(?:abstract|summary|s\s+u\s+m\s+m\s+a\s+r\s+y|background\s+and\s+aims)\b", lower[:5000])
    abs_start = abs_match.start() if abs_match else min(600, len(text))

    # Locate Methods / Methodology
    m_match = re.search(r"\n(?:\d+\.?\s*)?(?:methods|methodology|materials\s+and\s+methods|patients\s+and\s+methods|study\s+design|search\s+strategy)\b", lower)
    r_match = re.search(r"\n(?:\d+\.?\s*)?(?:results|findings|basic\s+questions|recommendations)\b", lower)
    d_match = re.search(r"\n(?:\d+\.?\s*)?(?:discussion|commentary|conclusion)\b", lower)
    ref_match = re.search(r"\n(?:\d+\.?\s*)?(?:references|bibliography|literature\s+cited)\b", lower)

    m_start = m_match.start() if m_match else min(2200, len(text))
    r_start = r_match.start() if (r_match and r_match.start() > m_start) else min(m_start + 5500, len(text))
    d_start = d_match.start() if (d_match and d_match.start() > r_start) else min(r_start + 5500, len(text))
    ref_start = ref_match.start() if (ref_match and ref_match.start() > m_start) else len(text)

    abstract_zone = text[abs_start:max(abs_start + 2500, m_start)]

    return {
        "detected_title": detected_title,
        "title": title_zone,
        "abstract": abstract_zone[:3000],
        "methods": text[m_start:min(r_start, m_start + 8000)],
        "results": text[r_start:min(d_start, r_start + 8000)],
        "discussion": text[d_start:min(ref_start, d_start + 8000)],
        "body_no_refs": text[:ref_start],
    }


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
        surname = tokens[-1].title()
        initials = "".join([f"{t[0].upper()}." for t in tokens[:-1]])
        parsed_authors.append(f"{surname}, {initials}" if initials else surname)

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


def format_harvard_filename(data: dict, extension: str, tag: str = "KAT") -> str:
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
    clean_filename = f"{clean_author}_{clean_year}_{clean_tool}_{tag}.{extension.lstrip('.')}"
    return re.sub(r"_+", "_", clean_filename)


# ==========================================
# 3. HIERARCHICAL OVERRIDES & CONSISTENCY AUDITOR
# ==========================================
def detect_hierarchical_overrides(sections: Dict[str, str]) -> Optional[Tuple[str, List[str], List[str]]]:
    """
    Evidence Hierarchy Short-Circuit:
    Guidelines contain systematic reviews; systematic reviews do not contain graded consensus recommendations.
    If an unambiguous structural signature of a higher-order design is present, return immediately without voting.
    """
    title_l = sections["title"].lower()
    abs_l = sections["abstract"].lower()
    meth_l = sections["methods"].lower()
    body_head = sections["body_no_refs"][:12000].lower()

    # HIERARCHY TIER 1: Clinical Practice Guideline / Consensus Statement
    has_guideline_title = bool(re.search(r"\b(?:practical\s+guideline|clinical\s+practice\s+guideline|espen\s+guideline|consensus\s+statement|consensus\s+guidelines?|practice\s+guideline)\b", title_l))
    has_graded_recommendations = bool(re.search(r"(?:grade\s+of\s+recommendation|strong\s+consensus\s*\(\d+%\s*agreement\)|recommendation\s+\d+\s*\n|sign\s+methodology|delphi\s+rounds?)", meth_l + body_head))

    if has_guideline_title or has_graded_recommendations:
        decision_path = [
            "Tier 1 Hierarchical Override Triggered: Clinical Practice Guideline / Consensus Statement",
        ]
        if has_guideline_title:
            decision_path.append(f"Probable Title/Header contains explicit guideline declaration ('{sections['detected_title'][:80]}')")
        if has_graded_recommendations:
            decision_path.append("Manuscript body/methods contains formal recommendation grading & consensus agreement percentages")
        excluded = [
            "AMSTAR 2 excluded by hierarchy: Guidelines synthesize systematic reviews, but systematic reviews do not issue graded consensus recommendations",
            "Cochrane RoB 2 & NOS excluded: Document is a secondary clinical practice guideline, not a primary patient cohort"
        ]
        return "Clinical Practice Guideline / Consensus Statement", decision_path, excluded

    # HIERARCHY TIER 2: Systematic Review / Meta-Analysis
    has_sr_title = bool(re.search(r"\b(?:systematic\s+review|meta-analysis|network\s+meta-analysis)\b", title_l))
    has_prisma_methods = bool(re.search(r"\b(?:prisma|preferred\s+reporting\s+items\s+for\s+systematic|prospero|boolean|databases?\s+were\s+searched)\b", meth_l + abs_l))

    if has_sr_title and has_prisma_methods:
        return (
            "Systematic Review / Meta-Analysis",
            [
                "Tier 2 Hierarchical Override Triggered: Systematic Review / Meta-Analysis",
                "Title explicitly declares Systematic Review / Meta-Analysis AND Methods/Abstract confirms PRISMA/database search protocol"
            ],
            [
                "SANRA excluded: Formal systematic search protocol (PRISMA/PROSPERO) present",
                "Primary study tools (RoB 2 / NOS) excluded: Secondary evidence synthesis"
            ]
        )

    # HIERARCHY TIER 3: Randomized Controlled Trial (RCT)
    has_rct_title = bool(re.search(r"\b(?:randomi[zs]ed\s+controlled\s+trial|randomi[zs]ed\s+clinical\s+trial|double-blind.*trial|phase\s+iii\s+trial)\b", title_l))
    has_random_allocation = bool(re.search(r"\b(?:randomly\s+assigned|randomi[zs]ed\s+to\s+receive|allocation\s+concealment|computer-generated\s+random)\b", meth_l + abs_l))

    if has_rct_title and has_random_allocation:
        return (
            "Randomized Controlled Trial (RCT)",
            [
                "Tier 3 Hierarchical Override Triggered: Randomized Controlled Trial (RCT)",
                "Title declares randomized trial AND Methods/Abstract confirms active participant randomization"
            ],
            [
                "ROBINS-I & NOS excluded: Active randomized allocation confirmed",
                "AMSTAR 2 excluded: Primary interventional trial"
            ]
        )

    return None


def run_deterministic_section_classifier(sections: Dict[str, str]) -> Tuple[str, Dict[str, int], List[str], List[str], bool]:
    """
    Classifier A: First checks Hierarchical Short-Circuits, then falls back to Section-Aware Scoring.
    """
    override = detect_hierarchical_overrides(sections)
    if override:
        winner, decision_path, excluded = override
        scores = {k: (100 if k == winner else 0) for k in LEGAL_PAIRS.keys()}
        return winner, scores, decision_path, excluded, True

    title_l = sections["title"].lower()
    abs_l = sections["abstract"].lower()
    meth_l = sections["methods"].lower()

    scores = {k: 0 for k in LEGAL_PAIRS.keys()}
    decision_path = []
    excluded_alternatives = []

    # 1. Title-Zone Rules
    title_rules = [
        ("Clinical Practice Guideline / Consensus Statement", [r"\bguidelines?\b", r"\bconsensus\b"], 85),
        ("Systematic Review / Meta-Analysis", [r"\bsystematic\s+review\b", r"\bmeta-analysis\b"], 80),
        ("Randomized Controlled Trial (RCT)", [r"\brandomi[zs]ed\b", r"\bclinical\s+trial\b"], 75),
        ("Observational Study (Cohort)", [r"\b(?:retrospective|prospective)\s+cohort\b", r"\bcohort\s+study\b", r"\bcomplications\s+of\b"], 70),
        ("Observational Study (Case-Control)", [r"\bcase-control\b"], 75),
        ("Cross-Sectional Study", [r"\bcross-sectional\b", r"\bprevalence\b", r"\bsurvey\b"], 75),
        ("Diagnostic Accuracy Study", [r"\bdiagnostic\s+accuracy\b", r"\bsensitivity\s+and\s+specificity\b"], 75),
        ("Prediction Model Study", [r"\bprediction\s+model\b", r"\bprognostic\s+model\b", r"\brisk\s+score\b"], 75),
        ("Qualitative Study", [r"\bqualitative\b", r"\bthematic\s+analysis\b"], 75),
        ("Narrative Review", [r"\bnarrative\s+review\b", r"\bliterature\s+review\b", r"\bscoping\s+review\b"], 70),
    ]

    for design, patterns, pts in title_rules:
        for p in patterns:
            m = re.search(p, title_l)
            if m:
                scores[design] += pts
                decision_path.append(f"Title zone matched '{m.group(0)}' (+{pts} pts for {design})")
                break

    # 2. Methods & Abstract Empirical Signals
    if re.search(r"\b(?:retrospective\s+review|medical\s+records\s+were\s+reviewed|consecutive\s+patients|patients\s+who\s+underwent|institutional\s+review\s+board|irb\s+approval)\b", meth_l + abs_l):
        scores["Observational Study (Cohort)"] += 65
        scores["Narrative Review"] -= 45
        scores["Systematic Review / Meta-Analysis"] -= 45
        decision_path.append("Methods/Abstract confirms empirical patient cohort & chart review (+65 Cohort; -45 Review)")
        excluded_alternatives.append("AMSTAR 2 & SANRA excluded because study analyzes primary patient medical records")

    if re.search(r"\b(?:cases\s+and\s+controls|matched\s+controls)\b", meth_l + abs_l):
        scores["Observational Study (Case-Control)"] += 60
        decision_path.append("Methods/Abstract describes case-control matching (+60 Case-Control)")

    if re.search(r"\b(?:cross-sectional\s+survey|questionnaire\s+was\s+administered)\b", meth_l + abs_l):
        scores["Cross-Sectional Study"] += 60
        decision_path.append("Methods/Abstract describes cross-sectional survey (+60 Cross-Sectional)")

    if re.search(r"\b(?:review\s+of\s+the\s+literature|this\s+review\s+summarizes|narrative\s+overview)\b", abs_l + meth_l) and scores["Observational Study (Cohort)"] < 30:
        scores["Narrative Review"] += 55
        decision_path.append("Abstract/Methods describes non-systematic literature overview (+55 Narrative Review)")

    winner = max(scores, key=scores.get)
    if scores[winner] <= 0:
        winner = "Observational Study (Cohort)"
        scores[winner] = 40
        decision_path.append("Defaulted to Observational Study (Cohort) based on clinical structure")

    return winner, scores, decision_path, excluded_alternatives, False


def run_llm_classifier_b(client: Optional[genai.Client], sections: Dict[str, str], available_designs: List[str]) -> Tuple[str, str]:
    if not client:
        return "", "No API key active"
    valid_list = [d for d in available_designs if not d.startswith("🔍")]
    prompt = f"""
    You are an epidemiological study-design classifier.
    Evaluate the isolated sections below. Ignore background literature cited in the introduction; focus on what THIS document actually is.

    PROBABLE TITLE:
    {sections['detected_title']}

    TITLE ZONE:
    {sections['title'][:600]}

    ABSTRACT / SUMMARY:
    {sections['abstract'][:1500]}

    METHODS SECTION:
    {sections['methods'][:2000]}

    Select the EXACT design from: {json.dumps(valid_list)}
    Output JSON:
    {{
      "llm_design": "Exact design name",
      "llm_rationale": "1 sentence explanation"
    }}
    """
    for model_name in FREE_ACTIVE_MODELS:
        try:
            resp = client.models.generate_content(
                model=model_name,
                contents=[prompt],
                config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.0),
            )
            parsed = json.loads(resp.text)
            return parsed.get("llm_design", ""), parsed.get("llm_rationale", "")
        except (json.JSONDecodeError, KeyError):
            continue
        except Exception:
            continue
    return "", "LLM classifier unavailable"


def run_framework_consistency_auditor(
    text: str, client: Optional[genai.Client], available_designs: List[str]
) -> dict:
    """
    Executes Hierarchical Override Check, Section-Aware Classification, and Calibrated Disagreement Penalty.
    Does NOT trigger annoying hard stops on borderline Cohort vs Case-Control disagreements;
    instead deducts 0.15 confidence and only flags human review if confidence < 0.80 or margin < 15.
    """
    sections = parse_manuscript_sections(text)
    rule_winner, score_map, decision_path, excluded_alts, is_hierarchical = run_deterministic_section_classifier(sections)
    llm_winner, llm_rationale = run_llm_classifier_b(client, sections, available_designs)

    sorted_scores = sorted(score_map.items(), key=lambda x: x[1], reverse=True)
    first_design, first_score = sorted_scores[0]
    second_design, second_score = sorted_scores[1] if len(sorted_scores) > 1 else ("None", 0)
    margin = first_score - second_score

    base_confidence = 0.98 if is_hierarchical else min(0.95, round(0.75 + (margin / 250.0), 2))
    disagreement_detected = bool(llm_winner and llm_winner in score_map and llm_winner != rule_winner)

    if disagreement_detected and not is_hierarchical:
        base_confidence = round(max(0.50, base_confidence - 0.15), 2)
        decision_path.append(f"LLM Classifier B suggested '{llm_winner}' (-0.15 confidence adjustment applied)")
    elif disagreement_detected and is_hierarchical:
        decision_path.append(f"Hierarchical Override ('{rule_winner}') superseded LLM suggestion ('{llm_winner}')")

    requires_human_confirmation = (base_confidence < 0.80) or (margin < MIN_CONFIDENCE_MARGIN and not is_hierarchical)

    if is_hierarchical:
        status = "Verified by Hierarchical Override"
    elif requires_human_confirmation:
        status = f"Advisory: Low Margin or Classifier Split (Confidence: {int(base_confidence * 100)}%)"
    else:
        status = "Verified Consistent"

    expected_tool = SEED_TAXONOMY.get(rule_winner, {}).get("tool", "Assigned Tool")

    return {
        "design": rule_winner,
        "tool": expected_tool,
        "status": status,
        "hierarchical_override": is_hierarchical,
        "classifier_a_rule_engine": f"{rule_winner} ({first_score} pts; runner-up: {second_design} at {second_score} pts)",
        "classifier_b_llm": f"{llm_winner or 'Not Run'} ({llm_rationale})",
        "confidence_margin": margin,
        "disagreement_detected": disagreement_detected,
        "requires_human_confirmation": requires_human_confirmation,
        "decision_tree": decision_path,
        "excluded_alternatives": excluded_alts,
        "confidence": base_confidence,
        "sections_preview": {
            "detected_title": sections["detected_title"],
            "title_zone": sections["title"][:500],
            "abstract_zone": sections["abstract"][:500],
            "methods_zone": sections["methods"][:500],
        }
    }


def validate_framework_signature_and_legality(
    design: str,
    selected_tool: str,
    generated_domains: List[dict],
    manuscript_text: str
) -> Tuple[List[str], List[str]]:
    """
    Split into:
    1. HARD FAILS (ClassificationConsistencyError): Objective impossibilities
       - Illegal Design <-> Tool pairing (e.g. Guideline + AMSTAR 2)
       - Cross-contaminated tool domains (e.g. AMSTAR 2 outputting AGREE II 'Stakeholder Involvement')
       - Source manuscript impossible combinations (e.g. RCT with zero mention of random/allocation/trial)
    2. WARNINGS ONLY: Probabilistic keyword overlap or short methods sections.
    """
    passed_notes = []
    warnings = []
    text_lower = manuscript_text.lower() if manuscript_text else ""

    # --- HARD FAIL 1: Legality Matrix ---
    if design in LEGAL_PAIRS:
        allowed_tools = LEGAL_PAIRS[design]
        if not any(allowed.lower() in selected_tool.lower() or selected_tool.lower() in allowed.lower() for allowed in allowed_tools):
            raise ClassificationConsistencyError(
                f"Hard Fail — Illegal Design-Tool Pair: '{design}' cannot be evaluated with '{selected_tool}'. Allowed: {allowed_tools}"
            )
        passed_notes.append(f"Legality Matrix Verified: '{design}' <-> '{selected_tool}'")

    # --- HARD FAIL 2: Foreign Tool Domain Cross-Contamination ---
    domain_names_blob = " ".join([str(d.get("domain_name", "")).lower() for d in generated_domains])
    if "AMSTAR" in selected_tool.upper() and ("stakeholder involvement" in domain_names_blob or "editorial independence" in domain_names_blob):
        raise ClassificationConsistencyError("Hard Fail — Cross-Contamination: AMSTAR 2 output contains AGREE II guideline domains.")
    if "AGREE" in selected_tool.upper() and ("duplicate study selection" in domain_names_blob or "publication bias" in domain_names_blob):
        raise ClassificationConsistencyError("Hard Fail — Cross-Contamination: AGREE II output contains AMSTAR 2 review domains.")

    # --- HARD FAIL 3: Source-Manuscript Direct Verification ---
    sig = FRAMEWORK_SIGNATURES.get(design)
    if sig and text_lower:
        must_have_any = sig.get("source_must_have_any", [])
        if must_have_any and not any(re.search(p, text_lower) for p in must_have_any):
            raise ClassificationConsistencyError(
                f"Hard Fail — Source Verification Failed: Manuscript lacks any required structural markers for '{design}' ({must_have_any})."
            )
        passed_notes.append(f"Source-Text Verification Passed: Manuscript contains required '{design}' markers.")

        for forbidden_pat in sig.get("source_must_not_have", []):
            if re.search(forbidden_pat, text_lower[:8000]):
                warnings.append(f"Source Advisory: Detected phrase matching '{forbidden_pat}' in manuscript; verify architecture.")

    # --- SOFT WARNING: Expected Domain Keyword Overlap ---
    if sig and generated_domains:
        expected_kws = sig.get("expected_domain_keywords", [])
        matches = sum(1 for kw in expected_kws if kw in domain_names_blob)
        if matches < max(1, len(expected_kws) // 2):
            warnings.append(
                f"Low Domain Keyword Overlap ({matches}/{len(expected_kws)} canonical keywords matched for {selected_tool})."
            )
        else:
            passed_notes.append(f"Domain Keyword Overlap: {matches}/{len(expected_kws)} canonical terms matched.")

    return passed_notes, warnings


# ==========================================
# 4. DUAL-THRESHOLD GROUNDEDNESS & SCORE RECONCILER
# ==========================================
def clean_tokens(s: str) -> List[str]:
    s = s.lower()
    s = re.sub(r"[\u2018\u2019\u201c\u201d]", "'", s)
    s = re.sub(r"[\u2013\u2014-]", " ", s)
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return s.split()


def fuzzy_quote_match(quote: str, source_text: str, threshold: float = 0.85) -> Tuple[bool, float]:
    quote_tokens = clean_tokens(quote)
    if not quote_tokens:
        return False, 0.0
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
    data: dict, full_raw_text: str
) -> Tuple[dict, float, List[str], List[str]]:
    """
    Fixes the 100% groundedness bug when source text is missing (returns 0.0%).
    Splits threshold: Direct quotes require >= 0.95 similarity; Paraphrased evidence requires >= 0.85.
    """
    grounding_warnings: List[str] = []
    flaw_warnings: List[str] = []

    if not full_raw_text or not full_raw_text.strip():
        data["groundedness_score"] = 0.0
        data["groundedness_warnings"] = ["Source text unavailable for verification (0.0% Grounded)."]
        data["checksum_flaws"] = ["Cannot verify quote integrity without extracted source text."]
        data["checksum_passed"] = False
        return data, 0.0, data["groundedness_warnings"], data["checksum_flaws"]

    domains = data.get("domain_appraisals", [])
    if not domains:
        data["groundedness_score"] = 0.0
        return data, 0.0, [], ["No domain appraisals populated."]

    verified_count = 0
    for d in domains:
        raw_quote = d.get("verbatim_evidence", "").strip()
        if not raw_quote or raw_quote.lower() in ["not reported", "none", "n/a", "not specified", "unclear"]:
            d["checksum_status"] = "Verified (Absence Noted)"
            verified_count += 1
            continue

        is_direct_quote = raw_quote.startswith('"') or raw_quote.startswith("“") or raw_quote.startswith("'")
        req_threshold = 0.95 if is_direct_quote else 0.85

        is_match, ratio = fuzzy_quote_match(raw_quote, full_raw_text, threshold=req_threshold)
        if is_match:
            mode_tag = "Direct Quote" if is_direct_quote else "Evidence"
            d["checksum_status"] = f"Verified {mode_tag} ({int(ratio * 100)}%)"
            verified_count += 1
        elif ratio >= 0.85 and is_direct_quote:
            d["checksum_status"] = f"Paraphrase Match ({int(ratio * 100)}% < 95% Direct Quote Threshold)"
            verified_count += 0.5
            grounding_warnings.append(f"'{d.get('domain_name')}' quote matched at {int(ratio * 100)}% (paraphrased rather than verbatim).")
        else:
            d["checksum_status"] = f"Unverified ({int(ratio * 100)}% Max Match)"
            grounding_warnings.append(f"Evidence in '{d.get('domain_name')}' fell below {int(req_threshold * 100)}% threshold ({int(ratio * 100)}%).")

    grounded_score = round((verified_count / len(domains)) * 100, 1)
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
        "scale": "Scoring thresholds and rules",
        "score_interpretation": "Explicit definition of what each point, star, or score level represents.",
        "scoring_guide": "Step-by-step plain-English explanation."
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
        except (json.JSONDecodeError, KeyError):
            continue
        except Exception:
            continue
    raise RuntimeError("Failed to discover tool schema across active models.")


def reconcile_appraisal_scores(parsed: dict, assigned_design: Optional[str] = None) -> dict:
    if assigned_design and assigned_design in SEED_TAXONOMY:
        parsed["study_design"] = assigned_design
        parsed["appraisal_tool_selected"] = SEED_TAXONOMY[assigned_design]["tool"]

    tool_name = str(parsed.get("appraisal_tool_selected", "")).upper()
    domains = parsed.get("domain_appraisals", [])

    if not domains:
        return parsed

    # 1. SANRA Deterministic Reconciliation
    if "SANRA" in tool_name:
        total_pts = 0
        for d in domains:
            r_str = str(d.get("rating", "")).strip().lower()
            digit_match = re.search(r"\b([0-2])\b", r_str)
            if digit_match:
                pts = int(digit_match.group(1))
            elif "not met" in r_str or "no" in r_str or "absent" in r_str:
                pts = 0
            elif "partially" in r_str or "partial" in r_str:
                pts = 1
            else:
                pts = 2
            total_pts += pts

        if total_pts >= 9:
            verdict, meth_quality = "High Quality / Low Risk", "GOOD"
        elif total_pts >= 6:
            verdict, meth_quality = "Moderate Quality / Moderate Risk", "FAIR"
        else:
            verdict, meth_quality = "Low Quality / High Risk", "POOR"

        parsed["overall_verdict"] = verdict
        parsed.setdefault("ratings_summary", {})["methodology_quality"] = meth_quality

    # 2. AGREE II Deterministic Reconciliation (1-7 Scale)
    elif "AGREE" in tool_name:
        total_pts = 0
        count = len(domains)
        for d in domains:
            r_str = str(d.get("rating", "")).strip()
            m = re.search(r"\b([1-7])\b", r_str)
            pts = int(m.group(1)) if m else 5
            total_pts += pts
        min_poss = 1 * count
        max_poss = 7 * count
        scaled_pct = round(((total_pts - min_poss) / max(1, (max_poss - min_poss))) * 100, 1)
        if scaled_pct >= 70.0:
            verdict, meth_quality = f"Strongly Recommended ({scaled_pct}% Scaled Rigour)", "GOOD"
        elif scaled_pct >= 50.0:
            verdict, meth_quality = f"Recommended with Modifications ({scaled_pct}% Scaled Rigour)", "FAIR"
        else:
            verdict, meth_quality = f"Not Recommended ({scaled_pct}% Scaled Rigour)", "POOR"
        parsed["overall_verdict"] = verdict
        parsed["agree_scaled_percentage"] = scaled_pct
        parsed.setdefault("ratings_summary", {})["methodology_quality"] = meth_quality

    # 3. Newcastle-Ottawa Scale (NOS)
    elif "NOS" in tool_name or "NEWCASTLE" in tool_name:
        sel_stars, comp_stars, out_stars = 0, 0, 0
        for d in domains:
            d_name = str(d.get("domain_name", "")).lower()
            r_str = str(d.get("rating", ""))
            stars = r_str.count("★") or r_str.count("*")
            if stars == 0:
                m = re.search(r"\b([0-4])\s*(?:stars?|★)?\b", r_str.lower())
                stars = int(m.group(1)) if m else (1 if "yes" in r_str.lower() or "awarded" in r_str.lower() else 0)
            if "selection" in d_name:
                sel_stars += min(stars, 4)
            elif "comparability" in d_name:
                comp_stars += min(stars, 2)
            elif "outcome" in d_name or "exposure" in d_name:
                out_stars += min(stars, 3)

        total_stars = sel_stars + comp_stars + out_stars
        parsed["nos_score_breakdown"] = {
            "selection_stars": sel_stars,
            "comparability_stars": comp_stars,
            "outcome_stars": out_stars,
            "total_stars": total_stars
        }
        if total_stars >= 7:
            verdict, meth_quality = "Good Quality / Low Risk", "GOOD"
        elif total_stars >= 4:
            verdict, meth_quality = "Fair Quality / Moderate Risk", "FAIR"
        else:
            verdict, meth_quality = "Poor Quality / High Risk", "POOR"
        parsed["overall_verdict"] = verdict
        parsed.setdefault("ratings_summary", {})["methodology_quality"] = meth_quality

    # 4. Cochrane RoB 2 / ROBINS-I / AMSTAR 2 / QUADAS-2
    elif "ROB" in tool_name or "AMSTAR" in tool_name or "QUADAS" in tool_name:
        ratings = [str(d.get("rating", "")).lower() for d in domains]
        high_crit_count = sum(1 for r in ratings if ("high" in r or "critical" in r or "serious" in r or r == "no" or "not met" in r))
        some_concerns = sum(1 for r in ratings if ("some" in r or "moderate" in r or "unclear" in r or "partial" in r))

        if high_crit_count > 1 and "AMSTAR" in tool_name:
            parsed["overall_verdict"] = "Critically Low Confidence"
            parsed.setdefault("ratings_summary", {})["methodology_quality"] = "POOR"
        elif high_crit_count == 1 and "AMSTAR" in tool_name:
            parsed["overall_verdict"] = "Low Confidence"
            parsed.setdefault("ratings_summary", {})["methodology_quality"] = "POOR"
        elif high_crit_count >= 1:
            parsed["overall_verdict"] = "High Risk of Bias"
            parsed.setdefault("ratings_summary", {})["methodology_quality"] = "POOR"
        elif some_concerns >= 1:
            parsed["overall_verdict"] = "Moderate Confidence / Some Concerns"
            parsed.setdefault("ratings_summary", {})["methodology_quality"] = "FAIR"
        else:
            parsed["overall_verdict"] = "High Confidence / Low Risk"
            parsed.setdefault("ratings_summary", {})["methodology_quality"] = "GOOD"

    return parsed


def extract_explicit_score_from_data(data: dict) -> str:
    tool_name = str(data.get("appraisal_tool_selected", "")).upper()
    verdict = str(data.get("overall_verdict", "N/A"))

    nos_stars = data.get("nos_score_breakdown")
    if isinstance(nos_stars, dict) and nos_stars.get("total_stars") is not None and nos_stars.get("total_stars", 0) > 0:
        sel = nos_stars.get("selection_stars", 0)
        comp = nos_stars.get("comparability_stars", 0)
        out = nos_stars.get("outcome_stars", 0)
        tot = nos_stars.get("total_stars", 0)
        return f"{tot}/9 Stars (Selection: {sel}★ | Comparability: {comp}★ | Outcome: {out}★) — {verdict}"

    if "AGREE" in tool_name and data.get("agree_scaled_percentage") is not None:
        return f"{data.get('agree_scaled_percentage')}% Standardized AGREE II Score — {verdict}"

    arith = data.get("arithmetic_breakdown")
    if isinstance(arith, dict) and arith.get("final_tally"):
        return f"{arith.get('final_tally')} — {verdict}"

    domains = data.get("domain_appraisals", [])
    if domains:
        if "SANRA" in tool_name:
            total_points = 0
            for d in domains:
                r_str = str(d.get("rating", "")).strip().lower()
                match = re.search(r"\b([0-2])\b", r_str)
                if match:
                    total_points += int(match.group(1))
                elif "not met" in r_str or "no" in r_str:
                    total_points += 0
                elif "partial" in r_str:
                    total_points += 1
                else:
                    total_points += 2
            return f"{total_points}/12 Points — {verdict}"

        elif "AMSTAR" in tool_name or "ROB" in tool_name or "QUADAS" in tool_name:
            ratings = [str(d.get("rating", "")).lower() for d in domains]
            low_count = sum(1 for r in ratings if ("low" in r or "yes" in r or "met" in r) and "not" not in r and "partial" not in r)
            some_concerns = sum(1 for r in ratings if "some" in r or "moderate" in r or "partial" in r or "unclear" in r)
            high_count = sum(1 for r in ratings if "high" in r or "serious" in r or "critical" in r or "not met" in r or r == "no")
            return f"Compliant/Low: {low_count} | Partial/Concerns: {some_concerns} | Deficient/High: {high_count} (of {len(domains)} Domains) — {verdict}"

    return verdict


def resolve_tool_scoring_guide(data: dict) -> Tuple[str, str, str, str, str, str]:
    tool_meta = data.get("tool_meta") or {}
    acronym = tool_meta.get("acronym") or data.get("appraisal_tool_selected", "Tool")
    full_name = tool_meta.get("full_name") or "Standard Critical Appraisal Instrument"
    purpose = tool_meta.get("purpose") or "Standardized epidemiological instrument used to assess study rigour and bias."
    scoring_rules = tool_meta.get("scoring_rules") or "Standard qualitative rating or numerical score threshold."
    score_interpretation = tool_meta.get("score_interpretation") or ""
    scoring_guide = tool_meta.get("scoring_guide") or ""

    raw_design = str(data.get("study_design", "")).lower()
    for design_key, tool_info in SEED_TAXONOMY.items():
        if design_key.lower() in raw_design or raw_design in design_key.lower():
            full_name = tool_info.get("full_name", full_name)
            purpose = tool_info.get("purpose", purpose)
            scoring_rules = tool_info.get("scale", scoring_rules)
            score_interpretation = tool_info.get("score_interpretation", score_interpretation)
            scoring_guide = tool_info.get("scoring_guide", scoring_guide)
            acronym = tool_info.get("tool", acronym)
            break

    return acronym, full_name, purpose, scoring_rules, score_interpretation, scoring_guide


def execute_manuscript_evaluation(
    client: genai.Client,
    text_content: Optional[str] = None,
    raw_pdf_bytes: Optional[bytes] = None,
    assigned_design: Optional[str] = None,
    active_tools_map: Optional[dict] = None,
    strictness_mode: str = "Strict",
    self_audit_report: Optional[dict] = None,
) -> dict:
    active_tool_taxonomy = {assigned_design: active_tools_map[assigned_design]} if (assigned_design and active_tools_map and assigned_design in active_tools_map) else (active_tools_map or SEED_TAXONOMY)
    mandatory_tool_name = active_tools_map[assigned_design]["tool"] if (assigned_design and active_tools_map and assigned_design in active_tools_map) else "Assigned Tool"
    taxonomy_json = json.dumps(active_tool_taxonomy, indent=2)

    system_instruction = f"""
    You are an objective epidemiological auditor.
    MANDATORY ARCHITECTURE: "{assigned_design}"
    MANDATORY APPRAISAL TOOL: "{mandatory_tool_name}"

    You MUST evaluate the manuscript using ONLY the domains listed in this schema:
    {taxonomy_json}

    DO NOT substitute another tool or invent domains from another instrument.

    DOMAIN RATING RULES:
    - If tool is AGREE II: Score each domain from 1 to 7 (e.g., 'Score 6 / 7').
    - If tool is SANRA: Score each domain 0, 1, or 2 (e.g., 'Score 2 (Yes / Completely Met)', 'Score 1 (Partially Met)', 'Score 0 (No / Not Met)').
    - If tool is NOS: Award stars for each domain (e.g., '3/4 Stars (★★★)').
    - If tool is AMSTAR 2: Rate each domain 'Yes', 'Partial Yes', or 'No'.
    - If tool is Cochrane RoB 2 / ROBINS-I / QUADAS-2 / PROBAST: Rate 'Low Risk', 'Some Concerns', or 'High Risk'.

    CRITICAL: Output ONLY a valid JSON object matching:
    {{
      "paper_title": "Identified full title",
      "authors": "Up to 3 authors in Harvard format",
      "journal": "Journal Name",
      "year_published": "YYYY",
      "doi": "DOI or URL",
      "study_design": "{assigned_design}",
      "appraisal_tool_selected": "{mandatory_tool_name}",
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
      "overall_verdict": "Overall rating",
      "overall_synthesis": [
        "Study Setup: [Plain language summary]",
        "Methodological Strengths: [What was done well]",
        "Risks of Bias & Red Flags: [Plain explanation of flaws]",
        "Real-World Limitations: [Who this might not apply to]",
        "Auditor Verdict: [Bottom line in plain English]"
      ],
      "practical_takeaways": ["Takeaway 1", "Takeaway 2", "Takeaway 3"],
      "narrative_arc_summary": ["Act I: ...", "Act II: ...", "Act III: ...", "Act IV: ...", "Act V: ..."],
      "audit_log": {{
        "classification_rationale": "Rationale",
        "decision_pathway": [{{"stage": "Stage", "observation": "Obs", "rule_applied": "Rule"}}],
        "verdict_derivation": "Justification"
      }},
      "tool_meta": {{
        "acronym": "{mandatory_tool_name}",
        "full_name": "Full Name",
        "purpose": "Purpose",
        "scoring_rules": "Rules",
        "score_interpretation": "Explicit definition of each score point",
        "scoring_guide": "Grading explanation"
      }},
      "domain_appraisals": [
        {{
          "domain_name": "Exact Domain Name from Schema",
          "evaluation_question": "Plain-English question",
          "rating": "Score / Rating",
          "location_reference": "Section/Page",
          "verbatim_evidence": "Quote",
          "critique": "Explanation"
        }}
      ]
    }}
    """

    sections = parse_manuscript_sections(text_content) if text_content else {}
    clean_text = sections.get("body_no_refs", text_content) if text_content else ""
    contents = [f"PAPER CONTENT:\n{clean_text}\n\nExecute appraisal under {mandatory_tool_name} and return JSON."] if (clean_text and len(clean_text.strip()) > 150) else (
        [types.Part.from_bytes(data=raw_pdf_bytes, mime_type="application/pdf"), f"Appraise for {assigned_design} using {mandatory_tool_name}."] if raw_pdf_bytes else None
    )
    if not contents:
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

                # 1. Deterministic Score Reconciliation
                parsed = reconcile_appraisal_scores(parsed, assigned_design=assigned_design)

                # 2. Framework Consistency Auditor (Hard Fails + Soft Warnings + Source Verification)
                passed_notes, auditor_warnings = validate_framework_signature_and_legality(
                    assigned_design or parsed.get("study_design", ""),
                    parsed.get("appraisal_tool_selected", ""),
                    parsed.get("domain_appraisals", []),
                    clean_text
                )

                parsed.setdefault("audit_log", {})["consistency_auditor_notes"] = passed_notes
                parsed["audit_log"]["consistency_auditor_warnings"] = auditor_warnings
                if self_audit_report:
                    parsed["audit_log"]["classification_decision_tree"] = self_audit_report.get("decision_tree", [])
                    parsed["audit_log"]["self_audit_report"] = self_audit_report

                return parsed
            except ClassificationConsistencyError as ce:
                raise ce
            except (json.JSONDecodeError, KeyError, ValueError) as parse_err:
                last_exception = parse_err
                time.sleep(1.0)
                continue
            except Exception as e:
                last_exception = e
                err_str = str(e)
                if "429" in err_str or "503" in err_str or "RESOURCE_EXHAUSTED" in err_str:
                    time.sleep(2.0)
                    continue
                if "404" in err_str or "NOT_FOUND" in err_str:
                    break
                time.sleep(1.0)

    raise RuntimeError(f"Appraisal pipeline failed on active models. Details: {str(last_exception)}")


# ==========================================
# 5. FORENSIC DEEP-DIVE GENERATOR
# ==========================================
def generate_forensic_deep_dive(
    client: genai.Client,
    base_data: dict,
    full_text: str = ""
) -> dict:
    locked_tool = base_data.get("appraisal_tool_selected", "Assigned Tool")
    locked_design = base_data.get("study_design", "Assigned Design")
    base_domains = base_data.get("domain_appraisals", [])
    explicit_score = extract_explicit_score_from_data(base_data)
    audit_trail = base_data.get("audit_log", {})

    short_context = full_text[:8000] if full_text else ""

    deep_prompt = f"""
    You are a distinguished university professor of clinical epidemiology and biostatistics conducting a formal, publication-standard forensic appraisal report.

    IMMUTABLE FRAMEWORK CONSTRAINT:
    - YOU MUST AUDIT THIS PAPER UNDER THE EXACT ASSIGNED TOOL: "{locked_tool}" ({locked_design}).
    - DO NOT switch or reclassify this tool under any circumstances.
    - Verified Score & Verdict: {explicit_score}
    - Base Evaluated Domains:
    {json.dumps(base_domains, indent=1)}
    - Self-Audit & Decision Tree:
    {json.dumps(audit_trail, indent=1)}

    MANUSCRIPT CONTEXT EXCERPT:
    {short_context}

    AUDIT OBJECTIVES & REQUIREMENTS (PUBLISHING-GRADE RIGOUR):
    1. TOOL SELECTION RATIONALE & COMPARATIVE REJECTION ALGORITHM:
       - Explain with formal academic reasoning why "{locked_tool}" was methodologically mandated for this manuscript architecture ({locked_design}).
       - Document the explicit algorithmic filtering steps and rejection criteria for competing tools.
    2. EXPLICIT SCORE INTERPRETATION & ARITHMETIC DERIVATIONS:
       - Detail what each individual score level means for "{locked_tool}".
       - Show the exact arithmetic derivation leading to "{explicit_score}".
    3. ZERO RISK-OF-BIAS BOILERPLATE:
       - Strip all generic textbook filler. Jump directly to concrete empirical facts from the paper.
    4. AT LEAST 2 TO 3 SUBSTANTIVE CONTENT CRITICISMS:
       - Detail deep, specific methodological vulnerabilities.
    5. DETAILED DOMAIN AUTOPSY:
       - Walk through each domain evaluated, providing forensic empirical observations and statistical impact.

    CRITICAL: Output ONLY a valid JSON object matching:
    {{
      "tool_selection_rationale": {{
        "assigned_tool": "{locked_tool}",
        "architecture_classification": "{locked_design}",
        "selection_justification": "Detailed academic rationale explaining why this tool is gold-standard for this design.",
        "competing_tools_rejected": "Comparison against other instruments and algorithmic rules explaining why each was rejected."
      }},
      "scoring_system_breakdown": {{
        "individual_score_meanings": "Explicit breakdown of what each grade or point means.",
        "how_scores_are_awarded": "Step-by-step description of how domain points are aggregated under {locked_tool}.",
        "quality_thresholds": "What specific scores indicate High, Moderate, or Poor quality under this instrument."
      }},
      "arithmetic_breakdown": {{
        "formula_used": "Exact mathematical equation or grading rule",
        "step_by_step_calculation": "Step-by-step arithmetic derivation",
        "final_tally": "{explicit_score}"
      }},
      "substantive_criticisms": [
        {{
          "flaw_title": "Deficit 1",
          "manuscript_evidence": "Exact numbers, line citations, or empirical data",
          "epidemiological_impact": "Direct scientific impact on validity"
        }},
        {{
          "flaw_title": "Deficit 2",
          "manuscript_evidence": "Exact numbers, line citations, or empirical data",
          "epidemiological_impact": "Direct scientific impact on validity"
        }},
        {{
          "flaw_title": "Deficit 3",
          "manuscript_evidence": "Exact numbers, line citations, or empirical data",
          "epidemiological_impact": "Direct scientific impact on validity"
        }}
      ],
      "domain_deep_dives": [
        {{
          "domain_name": "Domain Title",
          "score_awarded": "Score",
          "direct_empirical_finding": "Direct observation from paper",
          "statistical_or_design_flaw": "Vulnerability identified"
        }}
      ],
      "translational_verdict_and_advisory": "Final academic synthesis on whether clinicians should adopt, adapt, or reject these findings."
    }}
    """

    for model_name in FREE_ACTIVE_MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[deep_prompt],
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                    temperature=0.0,
                ),
            )
            raw = response.text.strip()
            if raw.startswith("```"):
                raw = re.sub(r"^```(?:json)?\n", "", raw)
                raw = re.sub(r"\n```$", "", raw)
            return json.loads(raw)
        except (json.JSONDecodeError, KeyError):
            continue
        except Exception:
            continue

    raise RuntimeError("Deep-dive forensic audit failed on active models.")


# ==========================================
# 6. UNIFIED SEAFOAM GREEN REPORT GENERATORS
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
    explicit_score = extract_explicit_score_from_data(data)
    acronym, full_name, purpose, scoring_rules, score_interpretation, scoring_guide = resolve_tool_scoring_guide(data)

    meta_table = doc.add_table(rows=8, cols=2)
    meta_table.style = "Table Grid"
    fields = [
        ("Evaluated Title:", data.get("paper_title", "N/A")),
        ("Authors (Harvard):", harvard_authors),
        ("Journal & Year:", f"{data.get('journal', 'N/A')} ({data.get('year_published', 'N/A')})"),
        ("Study Architecture:", data.get("study_design", "N/A")),
        ("Tool Applied:", f"{full_name} ({acronym})"),
        ("Quality Score / Star Tally:", explicit_score),
        ("Methodology / Reporting / Certainty:", f"{ratings.get('methodology_quality', 'N/A')} / {ratings.get('reporting_quality', 'N/A')} / {ratings.get('certainty_score', 'N/A')}"),
        ("Evidence Groundedness:", f"{data.get('groundedness_score', 0.0)}%"),
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
    doc.add_heading(f"1. Detailed Domain Appraisal ({acronym})", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    for i, h in enumerate(["Domain & Question", "Score", "Evidence", "Critique"]):
        cell = table.rows[0].cells[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="78BBA6"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        r = cell.paragraphs[0].add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for item in data.get("domain_appraisals", []):
        row_cells = table.add_row().cells
        row_cells[0].paragraphs[0].add_run(f"{item.get('domain_name')}\n\nQuestion: \"{item.get('evaluation_question')}\"")
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

    doc.add_paragraph("")
    doc.add_heading("3. Scoring System & Grading Criteria Explanation", level=1)
    doc.add_paragraph(f"Instrument Applied: {full_name} ({acronym})")
    doc.add_paragraph(f"Scored Verdict / Result: {explicit_score}")
    
    if score_interpretation:
        doc.add_paragraph("Interpretation of Individual Scores:")
        for line in score_interpretation.split("\n"):
            if line.strip():
                doc.add_paragraph(line.strip(), style='List Bullet')
                
    doc.add_paragraph("How Overall Quality Tiers Are Determined:")
    for line in scoring_guide.split("\n"):
        if line.strip():
            doc.add_paragraph(line.strip(), style='List Bullet')

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def build_pdf_report(data: dict) -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        rightMargin=32,
        leftMargin=32,
        topMargin=26,
        bottomMargin=26
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "T", parent=styles["Heading1"], fontSize=15, leading=18, textColor=colors.HexColor("#2C5247"), fontName="Helvetica-Bold", spaceAfter=2
    )
    sec_heading = ParagraphStyle(
        "SH", parent=styles["Heading2"], fontSize=10, leading=13, textColor=colors.HexColor("#1D3B32"), fontName="Helvetica-Bold", spaceBefore=7, spaceAfter=4
    )
    body_style = ParagraphStyle("B", parent=styles["Normal"], fontSize=8, leading=10.5, fontName="Helvetica")
    bold_style = ParagraphStyle("BB", parent=body_style, fontName="Helvetica-Bold")
    header_style = ParagraphStyle("H", parent=body_style, fontName="Helvetica-Bold", textColor=colors.white)
    question_style = ParagraphStyle("QS", parent=body_style, fontSize=7.5, leading=9.5, textColor=colors.HexColor("#1E3F35"), fontName="Helvetica-Oblique")
    quote_style = ParagraphStyle("Q", parent=body_style, fontSize=7.5, leading=9.5, textColor=colors.HexColor("#555555"), fontName="Helvetica-Oblique")

    story = [
        Paragraph(f"{APP_NAME} — Clinical Evidence Appraisal", title_style),
        Spacer(1, 4),
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
        active_tier = "GREEN" if any(w in v for w in ["low risk", "strongly recommended", "good quality", "high quality"]) else ("RED" if any(w in v for w in ["high risk", "critical", "not recommended", "poor"]) else "YELLOW")

    harvard_authors = format_harvard_authors(data.get("authors", "Not Specified"), max_authors=3)
    explicit_score = extract_explicit_score_from_data(data)
    acronym, full_name, purpose, scoring_rules, score_interpretation, scoring_guide = resolve_tool_scoring_guide(data)

    meta_data = [
        [
            Paragraph(f"<b>Title:</b> {data.get('paper_title', 'N/A')}", body_style),
            Paragraph(f"<b>Authors:</b> {harvard_authors} ({data.get('year_published', 'N/A')})", body_style),
        ],
        [
            Paragraph(f"<b>Design:</b> {data.get('study_design', 'N/A')} | <b>Tool:</b> {acronym}", body_style),
            Paragraph(f"<b>Tool Score / Tally:</b> <font color='#1D3B32'><b>{explicit_score}</b></font>", body_style),
        ],
        [
            Paragraph(f"<b>Quality Triad:</b> Method: <b>{ratings.get('methodology_quality', 'N/A')}</b> | Report: <b>{ratings.get('reporting_quality', 'N/A')}</b> | Certainty: <b>{ratings.get('certainty_score', 'N/A')}</b>", body_style),
            Paragraph(f"<b>Consistency Audit:</b> Verified Legal Pair ({data.get('groundedness_score', 0.0)}% Grounded)", body_style),
        ]
    ]
    t_meta = Table(meta_data, colWidths=[4.2 * inch, 3.4 * inch])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#EBF5F2')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 6))

    r_txt = "<b>✓ HIGH RISK / POOR</b>" if active_tier == "RED" else "HIGH RISK / POOR"
    y_txt = "<b>✓ MODERATE / FAIR</b>" if active_tier == "YELLOW" else "MODERATE / FAIR"
    g_txt = "<b>✓ ROBUST / GOOD</b>" if active_tier == "GREEN" else "ROBUST / GOOD"

    gauge_data = [[
        Paragraph(f"<font color='{'white' if active_tier == 'RED' else '#7A7A7A'}'>{r_txt}</font>", body_style),
        Paragraph(f"<font color='{'white' if active_tier == 'YELLOW' else '#7A7A7A'}'>{y_txt}</font>", body_style),
        Paragraph(f"<font color='{'white' if active_tier == 'GREEN' else '#7A7A7A'}'>{g_txt}</font>", body_style),
    ]]
    t_gauge = Table(gauge_data, colWidths=[2.53 * inch, 2.53 * inch, 2.53 * inch])
    t_gauge.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, 0), colors.HexColor('#C0392B') if active_tier == "RED" else colors.HexColor('#FADBD8')),
        ('BACKGROUND', (1, 0), (1, 0), colors.HexColor('#D4AC0D') if active_tier == "YELLOW" else colors.HexColor('#FCF3CF')),
        ('BACKGROUND', (2, 0), (2, 0), colors.HexColor('#2E7D32') if active_tier == "GREEN" else colors.HexColor('#E8F5E9')),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#A3D9C9')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_gauge)
    story.append(Spacer(1, 6))

    takeaways = data.get("practical_takeaways", [])
    synth = data.get("overall_synthesis", [])
    col_left = [Paragraph("<b>💡 Key Practical Takeaways</b>", sec_heading)]
    for idx, item in enumerate(takeaways[:3], 1):
        col_left.append(Paragraph(f"• <b>#{idx}:</b> {item}", body_style))
        col_left.append(Spacer(1, 2))

    col_right = [Paragraph("<b>🔍 Methodological Synthesis (Key Points)</b>", sec_heading)]
    if isinstance(synth, list):
        for pt in synth[:4]:
            parts = pt.split(":", 1)
            text = f"• <b>{parts[0]}:</b> {parts[1]}" if len(parts) == 2 else f"• {pt}"
            col_right.append(Paragraph(text, body_style))
            col_right.append(Spacer(1, 2))
    else:
        col_right.append(Paragraph(str(synth), body_style))

    t_split = Table([[col_left, col_right]], colWidths=[3.75 * inch, 3.85 * inch])
    t_split.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F7FBF9')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_split)
    story.append(Spacer(1, 8))

    story.append(Paragraph(f"<b>Domain-by-Domain Evaluation & Core Questions ({acronym})</b>", sec_heading))
    domain_data = [[Paragraph("Domain & Specific Question", header_style), Paragraph("Score", header_style), Paragraph("Plain-English Assessment & Grounded Evidence", header_style)]]
    for d in data.get("domain_appraisals", []):
        r_str = str(d.get("rating", ""))
        question = d.get("evaluation_question", "")
        quote = d.get("verbatim_evidence", "")
        loc = d.get("location_reference", "")
        critique = d.get("critique", "")

        domain_cell = [Paragraph(f"<b>{d.get('domain_name')}</b>", bold_style), Spacer(1, 2), Paragraph(f"<b>Question:</b> \"{question}\"", question_style)]
        evidence_cell = [Paragraph(critique, body_style)]
        if quote and quote.lower() not in ["none", "not reported", "n/a"]:
            evidence_cell.append(Spacer(1, 2))
            evidence_cell.append(Paragraph(f"[{loc}] \"{quote}\"", quote_style))

        domain_data.append([domain_cell, Paragraph(f"<b>{r_str}</b>", bold_style), evidence_cell])

    t_domains = Table(domain_data, colWidths=[2.5 * inch, 1.0 * inch, 4.1 * inch], repeatRows=1)
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#78BBA6')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(domain_data)):
        if i % 2 == 0:
            t_style.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F4F9F7')))
    t_domains.setStyle(TableStyle(t_style))
    story.append(t_domains)
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>Appraisal Framework & Grading System Explanation</b>", sec_heading))
    
    scoring_blocks = []
    if score_interpretation:
        scoring_blocks.append(Paragraph("<b>Meaning of Individual Domain Scores:</b>", bold_style))
        for line in score_interpretation.strip().split("\n"):
            if line.strip():
                scoring_blocks.append(Paragraph(line.strip(), body_style))
                scoring_blocks.append(Spacer(1, 1))
        scoring_blocks.append(Spacer(1, 3))

    scoring_blocks.append(Paragraph("<b>Overall Quality Thresholds & Logic:</b>", bold_style))
    if scoring_guide:
        for line in scoring_guide.strip().split("\n"):
            if line.strip():
                scoring_blocks.append(Paragraph(line.strip(), body_style))
                scoring_blocks.append(Spacer(1, 1))
    else:
        scoring_blocks.append(Paragraph(scoring_rules, body_style))

    appendix_rows = [
        [Paragraph("Tool Applied:", bold_style), Paragraph(f"<b>{full_name} ({acronym})</b>", body_style)],
        [Paragraph("Clinical Purpose:", bold_style), Paragraph(purpose, body_style)],
        [Paragraph("Scored Result:", bold_style), Paragraph(f"<font color='#1D3B32'><b>{explicit_score}</b></font>", bold_style)],
        [Paragraph("Scoring Criteria & Interpretation:", bold_style), scoring_blocks],
        [Paragraph("Clinical Significance:", bold_style), Paragraph("Scores indicate methodological rigour. Lower scores or missing criteria highlight methodological vulnerabilities that necessitate caution before changing clinical practice.", body_style)],
    ]
    t_app = Table(appendix_rows, colWidths=[1.8 * inch, 5.8 * inch])
    t_app.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EBF5F2')),
        ('BACKGROUND', (1, 0), (1, -1), colors.HexColor('#FBFDFD')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_app)

    doc.build(story)
    buf.seek(0)
    return buf


def build_deep_dive_pdf_report(base_data: dict, deep_dive: dict) -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf,
        pagesize=letter,
        rightMargin=36,
        leftMargin=36,
        topMargin=32,
        bottomMargin=32
    )
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "DDT", parent=styles["Heading1"], fontSize=16, leading=20, textColor=colors.HexColor("#2C5247"), fontName="Helvetica-Bold", spaceAfter=3
    )
    sub_title = ParagraphStyle(
        "DDST", parent=styles["Normal"], fontSize=8.5, leading=11, textColor=colors.HexColor("#4A8574"), fontName="Helvetica-Oblique", spaceAfter=8
    )
    sec_heading = ParagraphStyle(
        "DDSH", parent=styles["Heading2"], fontSize=10.5, leading=14, textColor=colors.HexColor("#1D3B32"), fontName="Helvetica-Bold", spaceBefore=9, spaceAfter=4
    )
    body_style = ParagraphStyle("DDB", parent=styles["Normal"], fontSize=8, leading=11, fontName="Helvetica")
    bold_style = ParagraphStyle("DDBB", parent=body_style, fontName="Helvetica-Bold")
    header_style = ParagraphStyle("DDH", parent=body_style, fontName="Helvetica-Bold", textColor=colors.white)

    story = [
        Paragraph(f"{APP_NAME} — Publication-Grade Forensic Critical Appraisal", title_style),
        Paragraph("Comprehensive Methodological Autopsy, Explicit Arithmetic Derivations & Bias Analysis", sub_title),
        Spacer(1, 4),
    ]

    harvard_authors = format_harvard_authors(base_data.get("authors", "Not Specified"), max_authors=3)
    explicit_score = extract_explicit_score_from_data(base_data)
    arith = deep_dive.get("arithmetic_breakdown", {})
    calc_tally = arith.get("final_tally") or explicit_score
    acronym, full_name, purpose, scoring_rules, score_interpretation, scoring_guide = resolve_tool_scoring_guide(base_data)

    meta_data = [
        [Paragraph(f"<b>Manuscript:</b> {base_data.get('paper_title', 'N/A')}", body_style), Paragraph(f"<b>Authors:</b> {harvard_authors} ({base_data.get('year_published', 'N/A')})", body_style)],
        [Paragraph(f"<b>Architecture:</b> {base_data.get('study_design', 'N/A')} | <b>Tool:</b> {full_name} ({acronym})", body_style), Paragraph(f"<b>Verified Score:</b> <font color='#1D3B32'><b>{calc_tally}</b></font>", bold_style)],
    ]
    t_meta = Table(meta_data, colWidths=[4.1 * inch, 3.3 * inch])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#EBF5F2')),
        ('BOX', (0, 0), (-1, -1), 1.2, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>1. Methodological Tool Selection & Algorithmic Rejection Rationale</b>", sec_heading))
    rationale = deep_dive.get("tool_selection_rationale", {})
    t_rat_data = [
        [Paragraph("<b>Mandated Framework:</b>", bold_style), Paragraph(f"<b>{rationale.get('assigned_tool', acronym)}</b> ({rationale.get('architecture_classification', base_data.get('study_design'))})", body_style)],
        [Paragraph("<b>Selection Justification:</b>", bold_style), Paragraph(rationale.get("selection_justification", "N/A"), body_style)],
        [Paragraph("<b>Alternative Tools Rejected & Decision Rule:</b>", bold_style), Paragraph(rationale.get("competing_tools_rejected", "N/A"), body_style)],
    ]
    t_rat = Table(t_rat_data, colWidths=[1.8 * inch, 5.6 * inch])
    t_rat.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EBF5F2')),
        ('BACKGROUND', (1, 0), (1, -1), colors.HexColor('#FBFCFC')),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_rat)
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>2. Scoring Mechanics, Arithmetic Derivations & Quality Thresholds</b>", sec_heading))
    scoring_system = deep_dive.get("scoring_system_breakdown", {})
    individual_meanings = scoring_system.get("individual_score_meanings") or score_interpretation
    how_awarded = scoring_system.get("how_scores_are_awarded") or scoring_guide
    thresholds = scoring_system.get("quality_thresholds") or scoring_rules

    arith_rows = [
        [Paragraph("<b>Individual Score Meaning:</b>", bold_style), Paragraph(individual_meanings, body_style)],
        [Paragraph("<b>Aggregation & Logic:</b>", bold_style), Paragraph(how_awarded, body_style)],
        [Paragraph("<b>Quality Thresholds:</b>", bold_style), Paragraph(thresholds, body_style)],
        [Paragraph("<b>Mathematical Formula / Rule:</b>", bold_style), Paragraph(arith.get("formula_used", "N/A"), body_style)],
        [Paragraph("<b>Step-by-Step Derivation:</b>", bold_style), Paragraph(arith.get("step_by_step_calculation", "N/A"), body_style)],
        [Paragraph("<b>Final Calculated Score / Tally:</b>", bold_style), Paragraph(f"<font color='#1D3B32'><b>{calc_tally}</b></font>", bold_style)],
    ]
    t_arith = Table(arith_rows, colWidths=[1.8 * inch, 5.6 * inch])
    t_arith.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (0, -1), colors.HexColor('#EBF5F2')),
        ('BACKGROUND', (1, 0), (1, -1), colors.HexColor('#FDFEFE')),
        ('BOX', (0, 0), (-1, -1), 0.75, colors.HexColor('#78BBA6')),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    story.append(t_arith)
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>3. Substantive Content Criticisms & Methodological Vulnerabilities</b>", sec_heading))
    crit_rows = [[Paragraph("Identified Deficit & Mechanism", header_style), Paragraph("Manuscript Empirical Data / Line Citations", header_style), Paragraph("Scientific & Epidemiological Threat", header_style)]]
    for c in deep_dive.get("substantive_criticisms", []):
        crit_rows.append([
            Paragraph(f"<b>{c.get('flaw_title')}</b>", bold_style),
            Paragraph(c.get("manuscript_evidence", ""), body_style),
            Paragraph(c.get("epidemiological_impact", ""), body_style),
        ])
    t_crit = Table(crit_rows, colWidths=[2.2 * inch, 2.6 * inch, 2.6 * inch], repeatRows=1)
    t_crit.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4A8574')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    for i in range(1, len(crit_rows)):
        if i % 2 == 0:
            t_crit.setStyle(TableStyle([('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F4F9F7'))]))
    story.append(t_crit)
    story.append(Spacer(1, 8))

    story.append(Paragraph("<b>4. Domain-by-Domain Forensic Audit (Boilerplate-Free)</b>", sec_heading))
    d_rows = [[Paragraph("Domain Evaluated", header_style), Paragraph("Score", header_style), Paragraph("Direct Empirical Finding in Manuscript", header_style), Paragraph("Identified Design / Statistical Vulnerability", header_style)]]
    for d in deep_dive.get("domain_deep_dives", []):
        d_rows.append([
            Paragraph(f"<b>{d.get('domain_name')}</b>", bold_style),
            Paragraph(d.get("score_awarded", ""), bold_style),
            Paragraph(d.get("direct_empirical_finding", ""), body_style),
            Paragraph(d.get("statistical_or_design_flaw", ""), body_style),
        ])
    t_d = Table(d_rows, colWidths=[1.8 * inch, 0.9 * inch, 2.4 * inch, 2.3 * inch], repeatRows=1)
    t_d.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2C5247')),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#D1E7DF')),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ]))
    for i in range(1, len(d_rows)):
        if i % 2 == 0:
            t_d.setStyle(TableStyle([('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F4F9F7'))]))
    story.append(t_d)
    story.append(Spacer(1, 8))

    adv = deep_dive.get("translational_verdict_and_advisory", "")
    if adv:
        story.append(Paragraph("<b>5. Translational Clinical Advisory & Adoption Verdict</b>", sec_heading))
        t_adv = Table([[Paragraph(f"<b>Clinical Adoption Recommendation:</b> {adv}", body_style)]], colWidths=[7.4 * inch])
        t_adv.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#EBF5F2')),
            ('BOX', (0, 0), (-1, -1), 1.2, colors.HexColor('#78BBA6')),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
        ]))
        story.append(t_adv)

    doc.build(story)
    buf.seek(0)
    return buf


def build_deep_dive_docx_report(base_data: dict, deep_dive: dict) -> io.BytesIO:
    doc = Document()
    title = doc.add_heading(level=0)
    run_t = title.add_run(f"{APP_NAME} — Publication-Grade Forensic Critical Appraisal")
    run_t.font.size = Pt(18)
    run_t.font.bold = True
    run_t.font.color.rgb = RGBColor(44, 82, 71)

    harvard_authors = format_harvard_authors(base_data.get("authors", "Not Specified"), max_authors=3)
    explicit_score = extract_explicit_score_from_data(base_data)
    arith = deep_dive.get("arithmetic_breakdown", {})
    calc_tally = arith.get("final_tally") or explicit_score
    acronym, full_name, purpose, scoring_rules, score_interpretation, scoring_guide = resolve_tool_scoring_guide(base_data)

    p_meta = doc.add_paragraph()
    p_meta.add_run("Evaluated Manuscript: ").bold = True
    p_meta.add_run(f"{base_data.get('paper_title')} ({harvard_authors}, {base_data.get('year_published')})\n")
    p_meta.add_run(f"Architecture: {base_data.get('study_design')} | Gold-Standard Tool: {full_name} ({acronym})\n")
    p_meta.add_run(f"Verified Tool Score / Tally: {calc_tally}").bold = True

    doc.add_heading("1. Methodological Tool Selection & Comparative Justification", level=1)
    rat = deep_dive.get("tool_selection_rationale", {})
    doc.add_paragraph(f"Mandated Tool: {rat.get('assigned_tool', acronym)} ({rat.get('architecture_classification', base_data.get('study_design'))})")
    doc.add_paragraph(f"Selection Justification:\n{rat.get('selection_justification')}")
    doc.add_paragraph(f"Alternative Frameworks Rejected & Algorithmic Rules:\n{rat.get('competing_tools_rejected')}")

    doc.add_heading("2. Scoring Mechanics, Mathematical Derivations & Score Interpretation", level=1)
    scoring_system = deep_dive.get("scoring_system_breakdown", {})
    doc.add_paragraph(f"Meaning of Individual Domain Scores:\n{scoring_system.get('individual_score_meanings', score_interpretation)}")
    doc.add_paragraph(f"Aggregation & Logic:\n{scoring_system.get('how_scores_are_awarded', scoring_guide)}")
    doc.add_paragraph(f"Quality Thresholds:\n{scoring_system.get('quality_thresholds', scoring_rules)}")
    doc.add_paragraph(f"Formula / Decision Gate: {arith.get('formula_used', 'N/A')}")
    doc.add_paragraph(f"Step-by-Step Arithmetic Calculation:\n{arith.get('step_by_step_calculation', 'N/A')}")
    p_fin = doc.add_paragraph()
    p_fin.add_run(f"Final Calculated Score / Tally: {calc_tally}").bold = True

    doc.add_heading("3. Substantive Content Criticisms & Flaw Analysis", level=1)
    for idx, c in enumerate(deep_dive.get("substantive_criticisms", []), 1):
        p_c = doc.add_paragraph(style='List Bullet')
        p_c.add_run(f"{idx}. {c.get('flaw_title')}: ").bold = True
        p_c.add_run(f"{c.get('manuscript_evidence')}\n")
        p_c.add_run(f"Epidemiological Threat: {c.get('epidemiological_impact')}")

    doc.add_heading("4. Domain-by-Domain Forensic Audit (Boilerplate-Free)", level=1)
    table = doc.add_table(rows=1, cols=4)
    table.style = "Table Grid"
    headers = ["Domain", "Score", "Direct Manuscript Finding", "Methodological Vulnerability"]
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="78BBA6"/>')
        cell._tc.get_or_add_tcPr().append(shd)
        r = cell.paragraphs[0].add_run(h)
        r.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)

    for d in deep_dive.get("domain_deep_dives", []):
        row = table.add_row().cells
        row[0].paragraphs[0].add_run(d.get("domain_name", "")).bold = True
        row[1].paragraphs[0].add_run(d.get("score_awarded", ""))
        row[2].paragraphs[0].add_run(d.get("direct_empirical_finding", ""))
        row[3].paragraphs[0].add_run(d.get("statistical_or_design_flaw", ""))

    adv = deep_dive.get("translational_verdict_and_advisory", "")
    if adv:
        doc.add_heading("5. Translational Clinical Advisory & Adoption Verdict", level=1)
        doc.add_paragraph(adv)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


def answer_interrogation_query(
    client: genai.Client,
    query: str,
    appraisal_json: dict,
    raw_manuscript_text: str,
    chat_history: List[dict]
) -> str:
    history_context = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in chat_history[-6:]])
    chat_instruction = f"""
    You are {ASSISTANT_NAME}, an objective methodological audit and interrogation assistant for the Kitique Appraisal Tool.
    Answer user queries regarding the appraisal report and source manuscript in simple, direct, academically grounded language.
    Cite specific lines, sections, and quotes clearly.

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
    return f"Failed to query {ASSISTANT_NAME} across active models."


# ==========================================
# 7. STREAMLIT APP & UI INTERACTION
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
else:
    st.title(f"🐱 {APP_NAME}")
    st.caption("Epidemiological Critical Appraisal & Hierarchical Consistency Auditor")

    api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", None))

    if "history_limit" not in st.session_state:
        st.session_state["history_limit"] = 10
    if "chat_history" not in st.session_state:
        st.session_state["chat_history"] = []
    if "raw_paper_text" not in st.session_state:
        st.session_state["raw_paper_text"] = ""
    if "deep_dive_data" not in st.session_state:
        st.session_state["deep_dive_data"] = None

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
                        st.session_state["deep_dive_data"] = None
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
            with st.spinner("Extracting text & isolating structural sections..."):
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

    self_audit_report = {}
    ensemble_winner = "Observational Study (Cohort)"
    confidence_score = 1.0

    if paper_text:
        client_preflight = genai.Client(api_key=api_key) if api_key else None
        self_audit_report = run_framework_consistency_auditor(paper_text, client_preflight, available_designs)
        ensemble_winner = self_audit_report["design"]
        confidence_score = self_audit_report["confidence"]

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
        st.metric("Consistency Confidence", conf_badge)

    can_proceed = True

    if self_audit_report:
        with st.expander(f"🛡️ Framework Consistency Auditor — {self_audit_report.get('status', 'Verified')}"):
            st.markdown(f"**Assigned Design:** `{self_audit_report.get('design')}` | **Mandated Tool:** `{self_audit_report.get('tool')}`")
            st.markdown(f"- **Classifier A (Hierarchical / Section Engine):** {self_audit_report.get('classifier_a_rule_engine')}")
            st.markdown(f"- **Classifier B (Constrained LLM):** {self_audit_report.get('classifier_b_llm')}")
            st.markdown("**Classification Decision Tree:**")
            for step in self_audit_report.get("decision_tree", []):
                st.markdown(f"  - ✅ {step}")
            st.markdown("**Excluded Alternatives:**")
            for ex in self_audit_report.get("excluded_alternatives", []):
                st.markdown(f"  - ❌ {ex}")

        with st.expander("🔍 Classification & Section Parser Debug View"):
            prev = self_audit_report.get("sections_preview", {})
            st.markdown("**Heuristic Probable Title Detected:**")
            st.code(prev.get("detected_title", "N/A"))
            st.markdown("**Isolated Title Zone (Top 500 chars):**")
            st.code(prev.get("title_zone", "N/A"))
            st.markdown("**Isolated Abstract Zone (Top 500 chars):**")
            st.code(prev.get("abstract_zone", "N/A"))
            st.markdown("**Isolated Methods Zone (Top 500 chars):**")
            st.code(prev.get("methods_zone", "N/A"))

        if self_audit_report.get("requires_human_confirmation"):
            st.warning(
                f"⚠️ **Low Confidence Margin or Classifier Split ({int(confidence_score * 100)}%):**\n\n"
                f"- **Rule Engine:** `{self_audit_report.get('design')}`\n"
                f"- **LLM Classifier:** `{self_audit_report.get('classifier_b_llm')}`\n\n"
                f"Please verify the dropdown above and check the box below to unlock execution."
            )
            can_proceed = st.checkbox("I have verified the selected Target Study Architecture above.", value=False)

    if st.button("⚡ Run Methodological Appraisal", type="primary", use_container_width=True, disabled=not can_proceed):
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

            with st.spinner(f"Executing appraisal & verifying framework consistency under {target_design}..."):
                try:
                    raw_data = execute_manuscript_evaluation(
                        client,
                        text_content=paper_text,
                        raw_pdf_bytes=raw_pdf_bytes if not paper_text else None,
                        assigned_design=target_design,
                        active_tools_map=active_tools,
                        strictness_mode=strictness_choice,
                        self_audit_report=self_audit_report,
                    )
                    verified_data, score, _, flaw_warns = run_comprehensive_checksum(raw_data, paper_text)
                    new_id = save_appraisal(verified_data, filename_tag)
                    st.session_state["appraisal_data"] = verified_data
                    st.session_state["deep_dive_data"] = None
                    st.session_state["chat_history"] = []
                    st.success(f"Verified by Framework Consistency Auditor and saved (ID: #{new_id})")
                    st.rerun()
                except ClassificationConsistencyError as ce:
                    st.error(f"🛑 **Framework Consistency Auditor Blocked Report Generation (Hard Fail):**\n\n{ce}")
                except Exception as e:
                    st.error(f"Execution error: {e}")

    if "appraisal_data" in st.session_state:
        data = st.session_state["appraisal_data"]
        st.markdown("---")
        st.subheader(f"📋 {data.get('paper_title', 'Untitled')}")

        ratings = data.get("ratings_summary", {})
        harvard_display_authors = format_harvard_authors(data.get("authors", "Not Specified"), max_authors=3)
        explicit_score = extract_explicit_score_from_data(data)

        st.markdown(f"**Authors (Harvard):** {harvard_display_authors} | **Tool Score / Tally:** `{explicit_score}`")

        auditor_warns = data.get("audit_log", {}).get("consistency_auditor_warnings", [])
        for w in auditor_warns:
            st.warning(f"⚠️ **Auditor Advisory:** {w}")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Architecture", data.get("study_design"))
        with c2:
            st.metric("Tool Applied", data.get("appraisal_tool_selected"))
        with c3:
            st.metric("Methodology Quality", ratings.get("methodology_quality", "N/A"))
        with c4:
            st.metric("Groundedness", f"{data.get('groundedness_score', 0.0)}%")

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
                    "Core Question": d.get("evaluation_question"),
                    "Score": d.get("rating"),
                    "Reference": d.get("location_reference"),
                    "Quote": d.get("verbatim_evidence"),
                    "Grounding Status": d.get("checksum_status", "Verified"),
                    "Assessment Critique": d.get("critique"),
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

        c_doc, c_pdf, c_log = st.columns(3)
        with c_doc:
            st.download_button(
                "📥 Download Executive Word (.docx)",
                data=build_docx_report(data),
                file_name=format_harvard_filename(data, "docx"),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
            )
        with c_pdf:
            st.download_button(
                "📥 Download Executive PDF (.pdf)",
                data=build_pdf_report(data),
                file_name=format_harvard_filename(data, "pdf"),
                mime="application/pdf",
                use_container_width=True,
            )
        with c_log:
            audit_payload = json.dumps(data, indent=2)
            st.download_button(
                "📥 Download Self-Audit Log (.json)",
                data=audit_payload,
                file_name=format_harvard_filename(data, "json"),
                mime="application/json",
                use_container_width=True,
            )

        # --- ADVANCED PUBLICATION-GRADE FORENSIC DEEP-DIVE ---
        st.markdown("---")
        st.subheader("🔬 Publication-Grade Forensic Deep-Dive")
        st.caption("Generates a full peer-review autopsy: tool selection comparative rationale, explicit arithmetic derivations, zero generic boilerplate, and 2–3 substantive content critiques (Seafoam/Green Themed).")

        if st.button("Generate Forensic Deep-Dive Report", type="secondary"):
            if not api_key:
                st.error("Access key required.")
            else:
                client = genai.Client(api_key=api_key)
                with st.spinner("Executing rigorous mathematical and methodological deep-dive..."):
                    try:
                        deep_dive_res = generate_forensic_deep_dive(client, data, st.session_state.get("raw_paper_text", ""))
                        st.session_state["deep_dive_data"] = deep_dive_res
                        st.success("Publication-grade forensic deep-dive generated.")
                        st.rerun()
                    except Exception as err:
                        st.error(f"Error generating deep-dive: {err}")

        if st.session_state.get("deep_dive_data"):
            dd = st.session_state["deep_dive_data"]
            with st.container():
                st.markdown("#### ⚖️ Methodological Tool Selection Rationale & Rejection Rules")
                rat = dd.get("tool_selection_rationale", {})
                st.markdown(f"**Mandated Instrument:** `{rat.get('assigned_tool', 'N/A')}`")
                st.write(rat.get("selection_justification", "N/A"))
                with st.expander("Comparative Framework Analysis & Rejection Decision Algorithm"):
                    st.write(rat.get("competing_tools_rejected", "N/A"))

                st.markdown("#### 🧮 Individual Score Interpretations & Arithmetic Derivations")
                scoring_system = dd.get("scoring_system_breakdown", {})
                st.write(f"**Meaning of Score Levels:**\n{scoring_system.get('individual_score_meanings', 'N/A')}")
                st.write(f"**Aggregation Logic:**\n{scoring_system.get('how_scores_are_awarded', 'N/A')}")
                st.write(f"**Quality Thresholds:**\n{scoring_system.get('quality_thresholds', 'N/A')}")

                arith = dd.get("arithmetic_breakdown", {})
                st.info(f"**Mathematical Rule / Equation:** {arith.get('formula_used', 'N/A')}\n\n**Step-by-Step Calculation:**\n{arith.get('step_by_step_calculation', 'N/A')}\n\n**Final Result:** {arith.get('final_tally', 'N/A')}")

                st.markdown("#### ⚠️ Substantive Content Criticisms")
                for c in dd.get("substantive_criticisms", []):
                    st.error(f"**{c.get('flaw_title')}:** {c.get('manuscript_evidence')}\n\n*Epidemiological Threat:* {c.get('epidemiological_impact')}")

                adv = dd.get("translational_verdict_and_advisory", "")
                if adv:
                    st.markdown("#### 🎯 Translational Clinical Advisory")
                    st.success(f"**Clinical Practice Recommendation:** {adv}")

                c_dd_doc, c_dd_pdf = st.columns(2)
                with c_dd_doc:
                    st.download_button(
                        "📥 Download Forensic Deep-Dive (.docx)",
                        data=build_deep_dive_docx_report(data, dd),
                        file_name=format_harvard_filename(data, "docx", tag="Forensic_Audit"),
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                    )
                with c_dd_pdf:
                    st.download_button(
                        "📥 Download Forensic Deep-Dive (.pdf)",
                        data=build_deep_dive_pdf_report(data, dd),
                        file_name=format_harvard_filename(data, "pdf", tag="Forensic_Audit"),
                        mime="application/pdf",
                        use_container_width=True,
                    )

        st.markdown("---")
        st.subheader(f"💬 {ASSISTANT_NAME}")
        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])

        if query := st.chat_input("Ask about specific scores, point rules, or line citations..."):
            if not api_key:
                st.error("Key required.")
            else:
                st.session_state["chat_history"].append({"role": "user", "content": query})
                with st.chat_message("user"):
                    st.markdown(query)

                client = genai.Client(api_key=api_key)
                with st.chat_message("assistant"):
                    with st.spinner(f"Consulting {ASSISTANT_NAME}..."):
                        try:
                            reply = answer_interrogation_query(
                                client, query, data, st.session_state.get("raw_paper_text", ""), st.session_state["chat_history"]
                            )
                            st.markdown(reply)
                            st.session_state["chat_history"].append({"role": "assistant", "content": reply})
                        except Exception as err:
                            st.error(f"Query error: {err}") any good?
