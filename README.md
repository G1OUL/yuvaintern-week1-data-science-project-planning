# Customer Churn Prediction and Retention Strategy Using Python
### Week 1 — Data Science Project Planning and Strategy Design

**Internship Organization:** YuvaIntern  
**Internship Role:** Virtual Data Science Explorer Intern  
**Student Name:** Hariom Upadhyay  
**Submission Date:** 10 October 2026  
**Assignment Designation:** Week 1 — Data Science Project Planning and Strategy Design  
**Planned Project Effort:** 33 Hours (Strict Allocation Target Met)  
**Report Document:** [report/Hariom_Upadhyay_YuvaIntern_Week1_Project_Planning_Report.docx](report/Hariom_Upadhyay_YuvaIntern_Week1_Project_Planning_Report.docx)  

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Status](https://img.shields.io/badge/Status-Submission--Ready-success.svg)]()
[![Internship](https://img.shields.io/badge/YuvaIntern-Week%201%20Milestone-orange.svg)]()  

---

## 1. Executive Summary & Project Overview

Customer retention represents the primary financial engine of subscription-based business models (SaaS, cloud platforms, and telecommunications). Acquiring new customers typically costs 5x to 7x more than retaining existing accounts. Consequently, unmitigated customer attrition (churn) directly diminishes Monthly Recurring Revenue (MRR), compresses Customer Lifetime Value (CLV), and creates an inefficient growth treadmill.

This repository hosts the **Week 1 Technical Milestone Deliverable** for the **YuvaIntern Virtual Data Science Explorer Internship**. It establishes an end-to-end strategic blueprint and methodology design for predicting customer churn and structuring proactive, cost-effective retention interventions using Python.

> **ACADEMIC & EVALUATION INTEGRITY DISCLOSURE:**  
> This project is strictly a **proposed planning and methodology design assignment**. No real customer data has been analyzed, and no machine learning models have been trained, deployed, or claimed as completed. All schemas, feature engineering recipes, validation protocols, and model architectures serve as an implementation blueprint for future project phases.

---

## 2. Business Problem Statement & Objectives

### Problem Statement
> *"To develop an end-to-end predictive data science strategy using Python to identify subscription accounts at high risk of voluntary attrition within a forward-looking 30-day forecast horizon, utilizing historical multi-source customer records. The resulting system outputs calibrated churn probabilities and interpretable feature attributions, enabling Customer Success teams to deploy targeted, cost-effective retention interventions before irreversible contract termination occurs."*

### Key Strategic Objectives
1. **Business Problem Formalization:** Map contractual renewal cycles, define voluntary vs. involuntary churn, and align on the 30-day forward prediction horizon across stakeholders.
2. **Multi-Source Data Strategy:** Catalog 6 hypothetical enterprise sources (CRM Demographics, Subscriptions, Billing/Payments, Product Telemetry, Support Tickets, and Churn Logs).
3. **Data Quality & Schema Governance:** Establish schema validation, missingness categorization (MCAR/MAR/MNAR), duplicate detection, and range audit rules.
4. **Hypothesis-Driven EDA:** Plan visual and statistical analyses across contract duration, tenure, usage decay, and billing friction.
5. **Leakage-Free Feature Engineering:** Derive dynamic features (usage velocity, seat utilization, payment friction) while strictly enforcing temporal cutoffs ($t_0$).
6. **Model Benchmarking Architecture:** Benchmark a majority-class baseline and heuristic rules against regularized Logistic Regression (interpretable) and Random Forest (non-linear ensemble).
7. **Multi-Metric Imbalanced Evaluation:** Emphasize Precision-Recall AUC (PR-AUC), cost-sensitive confusion matrices, probability calibration, and subgroup demographic parity.
8. **Actionable Retention Playbooks:** Map model risk deciles to stratified operational retention tiers within Customer Success staffing bandwidth.
9. **Reproducibility & Governance:** Establish modular Python pipelines, pinned dependency environments, and ethical data minimization protocols.

---

## 3. Project Scope: Boundaries & Constraints

| In-Scope (Planning Deliverables) | Out-of-Scope (Excluded from Week 1) |
| :--- | :--- |
| • Mathematical problem formulation & target definition ($Y \in \{0, 1\}$) | • Real customer database querying or live data extraction |
| • Schema mapping across 6 hypothetical enterprise data sources | • Production software deployment (APIs, Kubernetes, microservices) |
| • Leakage-free preprocessing & imputation pipeline design | • Training machine learning models or claiming empirical accuracy |
| • Comprehensive EDA visualization mapping & hypothesis framing | • Automated CRM webhook integration without human oversight |
| • Baseline, Logistic Regression, and Random Forest architectures | • Live customer retention campaigns or monetary discounting |
| • Cost-sensitive evaluation framework & threshold tuning design | • Fabricating experimental results or test metrics |
| • 33-hour project timeline & 10-risk mitigation matrix | • Processing unanonymized Personally Identifiable Information (PII) |

---

## 4. Proposed Python Data Science Methodology

The future implementation follows a disciplined 8-stage lifecycle:

```text
[1. Business Understanding] ──► [2. Data Collection] ──► [3. Data Quality & Cleaning]
                                                                  │
                                                                  ▼
[6. ML Modeling] ◄── [5. Feature Engineering] ◄── [4. Exploratory Data Analysis]
       │
       ▼
[7. Model Evaluation] ──► [8. Communication & Monitoring]
       │                              ▲
       └────── Iteration Loops ───────┘
```

### Core Methodological Highlights
- **Unit of Analysis:** Individual customer account (`account_id`) evaluated at monthly observation cutoff date $t_0$.
- **Temporal Windows:** 12-month rolling historical feature window $[t_0 - 365, t_0)$ and 30-day forward prediction horizon $[t_0, t_0 + 30)$.
- **Strict Leakage Prevention:** All scalers, encoders, and imputers encapsulated inside `sklearn.pipeline.Pipeline`, fitted strictly on training data splits.
- **Model Candidates:**
  - *Baseline:* `DummyClassifier(strategy='most_frequent')` + rule-based operational heuristic.
  - *Candidate 1 (Linear):* `LogisticRegression` with ElasticNet/L2 regularization for transparent Odds Ratios.
  - *Candidate 2 (Ensemble):* `RandomForestClassifier` for non-linear interactions, class weighting, and feature importances.
- **Evaluation under Class Imbalance:** Primary metric is **PR-AUC (Precision-Recall AUC)**; accuracy is rejected as a primary KPI due to class imbalance vulnerability. Cost-sensitive threshold optimization weighs False Negatives (lost CLV) against False Positives (outreach cost).

---

## 5. Technical Stack & Tools

- **Programming Language:** Python 3.10+ / 3.11
- **Interactive Prototyping:** JupyterLab / Jupyter Notebooks
- **Data Ingestion & Manipulation:** `pandas` (>=2.0), `NumPy` (>=1.24)
- **Machine Learning & Preprocessing:** `scikit-learn` (>=1.3)
- **Statistical Testing:** `SciPy` (>=1.10), `statsmodels`
- **Visualization Engines:** `Matplotlib` (>=3.7), `Seaborn` (>=0.12)
- **Report Automation:** `python-docx`
- **Version Control & Collaboration:** Git, GitHub

---

## 6. Project Workstream Timeline (33 Hours Total)

The work plan is structured into eight distinct workstreams totaling **exactly 33 hours**, fulfilling the 30–35 hour allocation requirement:

| Workstream | Activity Name | Allocated Effort | Key Deliverable Output |
| :---: | :--- | :---: | :--- |
| **WS 1** | Problem framing and success criteria | **3 Hours** | Signed Business Specification & KPI Charter |
| **WS 2** | Data-source and privacy planning | **3 Hours** | Data Source Catalog & Privacy Compliance Blueprint |
| **WS 3** | Data quality and preparation strategy | **5 Hours** | Automated Validation & Imputation Pipeline Code |
| **WS 4** | Exploratory analysis and feature planning | **4 Hours** | EDA Diagnostic Guide & Feature Recipe Store |
| **WS 5** | Baseline and model strategy | **5 Hours** | Model Benchmarking Blueprint & Hyperparameter Spec |
| **WS 6** | Evaluation and responsible-use planning | **4 Hours** | PR-AUC Evaluation Matrix & Subgroup Fairness Rubric |
| **WS 7** | Deliverables, risks, and documentation | **5 Hours** | 10-Risk Matrix, Technical README, & Report Draft |
| **WS 8** | Final review and presentation | **4 Hours** | Verified Final Report Package & Defense Slides |
| **TOTAL** | **Full Project Planning Milestone** | **33 Hours** | **Submission-Ready Deliverable Package** |

---

## 7. Technical Diagrams Description

All diagrams are available in high resolution (300 DPI) in the `diagrams/` directory and embedded in the report:

1. **Diagram 1: Data Science Lifecycle Flowchart** (`diagrams/diagram1_lifecycle_flowchart.png`)  
   Illustrates the 8 connected lifecycle stages from Business Understanding to Communication & Monitoring, highlighting iterative feedback loops for model refinement, feature iteration, data quality audits, and business KPI realignment.
2. **Diagram 2: Conceptual Project Architecture & Data Flow** (`diagrams/diagram2_conceptual_architecture.png`)  
   Details the flow from Approved Hypothetical Data Sources through Ingestion & Staging, Validation Pipelines, Feature Store, Model Benchmarking, Evaluation, to Business Recommendations, reinforced by Governance/Privacy and Integrity/Provenance layers.
3. **Diagram 3: 33-Hour Project Timeline Workstream Chart** (`diagram3_project_timeline.png`)  
   A horizontal bar chart visualizing the effort distribution across all 8 project activities, showing the 33-hour allocation.

---

## 8. Expected Outcomes, Risks, and Responsible AI

### Anticipated Future Deliverables
- Automated Data Quality & Schema Audit Report
- Exploratory Data Analysis & Diagnostic Visualization Suite
- Reusable, Leakage-Free Preprocessing Pipeline
- Benchmarked Baseline, Logistic Regression, and Random Forest Models
- Model Evaluation & Subgroup Demographic Parity Report
- Actionable Retention Playbook & Controlled A/B Pilot Protocol

### Key Risks & Mitigation Summary
- **Data Leakage:** Mitigated via strict temporal cutoffs ($t_0$) and pipeline encapsulation.
- **Class Imbalance:** Mitigated by prioritizing PR-AUC, tuning decision thresholds, and using balanced class weights.
- **Customer Privacy (GDPR/CCPA):** Mitigated via data minimization, PII exclusion, and surrogate key tokenization.
- **Concept Drift:** Mitigated by quarterly Population Stability Index (PSI) monitoring and scheduled model retraining.
- **Association vs. Causation:** Mitigated by framing scores as risk indicators, enforcing human-in-the-loop CSM decision-making, and conducting controlled A/B pilots.

---

## 9. Repository Structure

```text
yuvaintern-week1-data-science-project-planning/
├── README.md                                 # Comprehensive project documentation & guide
├── requirements.txt                          # Pinned Python dependencies
├── .gitignore                                # Git ignore file for Python and local artifacts
│
├── report/                                   # Finalized Word report deliverable
│   └── Hariom_Upadhyay_YuvaIntern_Week1_Project_Planning_Report.docx
│
├── diagrams/                                 # 300 DPI technical diagrams
│   ├── diagram1_lifecycle_flowchart.png     # Data Science Lifecycle with feedback loops
│   ├── diagram2_conceptual_architecture.png # Conceptual architecture & data flow
│   └── diagram3_project_timeline.png        # 33-hour workstream allocation chart
│
└── scripts/                                  # Reproducible Python automation scripts
    ├── docx_helpers.py                       # Word document styling, tables, callouts & headers
    ├── generate_diagrams.py                  # Script to generate all 3 technical diagrams
    ├── build_full_report.py                  # Word report generation pipeline
    └── validate_submission.py                # Automated 10-point deliverable validation script
```

---

## 10. Execution & Reproduction Guide

### Prerequisites
Clone the repository and install required Python dependencies:
```bash
git clone https://github.com/G1OUL/yuvaintern-week1-data-science-project-planning.git
cd yuvaintern-week1-data-science-project-planning
pip install -r requirements.txt
```

### 1. Regenerate Technical Diagrams
```bash
python scripts/generate_diagrams.py
```
*Generates and saves `diagram1_lifecycle_flowchart.png`, `diagram2_conceptual_architecture.png`, and `diagram3_project_timeline.png` in `diagrams/`.*

### 2. Generate Word Report Deliverable
```bash
python scripts/build_full_report.py
```
*Builds and formats the full Word report deliverable in `report/`.*

### 3. Run Automated Validation Suite
```bash
python scripts/validate_submission.py
```
*Executes the automated 10-point check suite verifying document integrity, metadata presence, 33-hour timeline, embedded diagrams, and zero credential leakage.*

---

## 11. Final Report Deliverable

The official finalized report is located at:  
👉 **[report/Hariom_Upadhyay_YuvaIntern_Week1_Project_Planning_Report.docx](report/Hariom_Upadhyay_YuvaIntern_Week1_Project_Planning_Report.docx)**

---
*Authored by Hariom Upadhyay | Virtual Data Science Explorer Intern | YuvaIntern | October 2026*
