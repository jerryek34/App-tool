"""
Regression Validation Corpus for Kitique Appraisal Tool
Run via: python validate_classifier.py
"""
import os
from app import (
    ClassificationConsistencyError,
    LEGAL_PAIRS,
    SEED_TAXONOMY,
    parse_manuscript_sections,
    run_deterministic_section_classifier,
    validate_framework_signature_and_legality,
)

GOLD_STANDARD_CORPUS = [
    {
        "name": "ESPEN Practical Guideline",
        "raw_text": """
        Clinical Nutrition 40 (2021) 4745-4761
        ESPEN Guideline
        ESPEN practical guideline: Clinical nutrition in surgery
        Arved Weimann, Marco Braga
        SUMMARY
        Early oral feeding is the preferred mode. In a recent meta-analysis of 29 studies including 7179 patients...
        2. Methodology
        The present practical guideline consists of 37 recommendations.
        Recommendation 1: Preoperative fasting from midnight is unnecessary.
        Grade of recommendation A - strong consensus (97% agreement)
        """,
        "expected_design": "Clinical Practice Guideline / Consensus Statement",
        "expected_tool": "AGREE II",
    },
    {
        "name": "Cates 2021 Charcot Reconstruction (Empirical Cohort)",
        "raw_text": """
        Journal of Foot & Ankle Surgery 60 (2021)
        Original Research
        Complications of Charcot Reconstruction in Patients With Peripheral Arterial Disease
        Nicole K. Cates, DPM, Wynter Chiodo, DPM
        Abstract
        A retrospective single-center review of 284 diabetic patients with Charcot neuroarthropathy who underwent osseous reconstruction.
        Patients and Methods
        After institutional review board approval, medical records were reviewed for consecutive patients who underwent surgical reconstruction.
        Results
        Within a cohort of 284 patients, the rate of PAD was 20.8%.
        """,
        "expected_design": "Observational Study (Cohort)",
        "expected_tool": "Newcastle-Ottawa Scale (NOS) Cohort",
    },
    {
        "name": "PRISMA Systematic Review & Meta-Analysis",
        "raw_text": """
        Annals of Surgery 2022
        Early Oral Feeding After Upper Gastrointestinal Surgery: A Systematic Review and Meta-Analysis
        Abstract
        Background: We conducted a systematic review and meta-analysis of randomized controlled trials.
        Methods
        This review was registered on PROSPERO and followed PRISMA guidelines. PubMed, Embase, and Cochrane databases were searched using boolean strings.
        Results
        15 studies were included in the forest plot pooling.
        """,
        "expected_design": "Systematic Review / Meta-Analysis",
        "expected_tool": "AMSTAR 2",
    },
    {
        "name": "Double-Blind Randomized Controlled Trial (PROCY Trial)",
        "raw_text": """
        Annals of Surgery
        Preoperative Oral Carbohydrate Load Versus Placebo in Major Elective Abdominal Surgery: A Randomized Controlled Trial
        Abstract
        Objective: To evaluate preoperative carbohydrate loading.
        Methods
        In this multicenter, double-blind, placebo-controlled trial, 662 patients were randomly assigned (1:1 allocation concealment) to receive 800 mL carbohydrate drink or placebo.
        Results
        Intention-to-treat analysis showed lower insulin resistance.
        """,
        "expected_design": "Randomized Controlled Trial (RCT)",
        "expected_tool": "Cochrane RoB 2",
    },
    {
        "name": "Cross-Sectional Prevalence Survey",
        "raw_text": """
        BMC Public Health
        Prevalence of Burnout Among Orthopaedic Residents: A National Cross-Sectional Study
        Abstract
        We conducted a cross-sectional survey to estimate prevalence.
        Methods
        A validated questionnaire was administered to a representative sample of 450 participants. Non-response bias was evaluated.
        """,
        "expected_design": "Cross-Sectional Study",
        "expected_tool": "AXIS",
    },
]

def run_regression_suite():
    print("=" * 72)
    print("KITIQUE APPRAISAL TOOL — CLASSIFICATION REGRESSION SUITE")
    print("=" * 72)
    passed = 0

    for idx, item in enumerate(GOLD_STANDARD_CORPUS, 1):
        sections = parse_manuscript_sections(item["raw_text"])
        winner, scores, decision_path, excluded, is_hierarchical = run_deterministic_section_classifier(sections)
        assigned_tool = SEED_TAXONOMY[winner]["tool"]

        design_ok = winner == item["expected_design"]
        tool_ok = assigned_tool == item["expected_tool"]

        canonical_domains = [{"domain_name": d["name"]} for d in SEED_TAXONOMY[winner]["domains"]]
        auditor_ok = True
        try:
            validate_framework_signature_and_legality(winner, assigned_tool, canonical_domains, item["raw_text"])
        except ClassificationConsistencyError as err:
            auditor_ok = False
            print(f"  [!] Auditor Error: {err}")

        if design_ok and tool_ok and auditor_ok:
            passed += 1
            status = "✅ PASS"
        else:
            status = "❌ FAIL"

        print(f"{idx}. {status} | {item['name']}")
        print(f"   Detected Title : {sections['detected_title']}")
        print(f"   Predicted      : {winner} -> {assigned_tool} (Hierarchical: {is_hierarchical})")
        if not (design_ok and tool_ok):
            print(f"   Expected       : {item['expected_design']} -> {item['expected_tool']}")
        print("-" * 72)

    print("Testing Negative Safety Gate (ESPEN Guideline forced into AMSTAR 2)...")
    try:
        validate_framework_signature_and_legality(
            "Clinical Practice Guideline / Consensus Statement",
            "AMSTAR 2",
            [],
            GOLD_STANDARD_CORPUS[0]["raw_text"]
        )
        print("❌ FAIL: Negative safety gate did NOT block illegal Guideline <-> AMSTAR 2 pair!")
    except ClassificationConsistencyError as ce:
        print(f"✅ PASS: Negative safety gate blocked illegal pair.")

    print("=" * 72)
    print(f"FINAL RESULT: {passed}/{len(GOLD_STANDARD_CORPUS)} Gold-Standard Corpus Papers Passed.")
    print("=" * 72)

if __name__ == "__main__":
    run_regression_suite()
