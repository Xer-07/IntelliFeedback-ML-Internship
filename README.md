# IntelliFeedback: End-to-End Sentiment & Customer Feedback Intelligence Engine

[![Python](https://img.shields.io/badge/Python-3.10%2B-blue.svg)](https://www.python.org/)
[![Machine Learning](https://img.shields.io/badge/Machine%20Learning-Scikit--Learn%20%7C%20PyTorch-orange.svg)](https://scikit-learn.org/)
[![NLP](https://img.shields.io/badge/NLP-TF--IDF%20%7C%20DistilBERT-green.svg)](https://huggingface.co/)
[![Internship](https://img.shields.io/badge/YuvaIntern-ML%20Virtual%20Internship-purple.svg)](https://yuvaintern.com/)
[![License](https://img.shields.io/badge/License-MIT-lightgrey.svg)](LICENSE)

---

## 📌 Project Overview
**IntelliFeedback** is an enterprise-grade Natural Language Processing (NLP) and Machine Learning solution engineered to automatically ingest, analyze, classify, and extract operational intelligence from high-volume customer feedback, reviews, and support interactions.

By classifying reviews into multiclass sentiment categories (**Negative**, **Neutral**, **Positive**) with confidence scores and root-cause indicators, IntelliFeedback enables product and customer success teams to remediate customer dissatisfaction signals before churn occurs.

Developed as part of the **YuvaIntern 5-Week Machine Learning Virtual Internship**, this repository hosts the complete codebase, data pipelines, experimental notebooks, model weights, and deliverables.

---

## 🗓️ 5-Week Internship Roadmap

- [x] **Week 1: Project Planning & Strategy Development**
  - Problem definition, business impact, and hypothesis formulation.
  - Empirical literature review and architectural design.
  - 5-week milestone roadmap and risk management plan.
  - Formal Project Charter documentation (`Week1_Project_Planning_and_Strategy_Report.docx`).
- [x] **Week 2: Algorithm Exploration & Preliminary Model Design**
  - Theoretical & empirical review of MNB, Logistic Regression, LinearSVC, Tree Ensembles, and DistilBERT.
  - Comprehensive 7-dimension algorithm comparison matrix.
  - Dual-tier system architecture: sub-linear TF-IDF + Calibrated LinearSVC (<15ms) with DistilBERT contextual fallback.
  - Baseline model pipeline implementation (`src/model_pipeline.py`) & interactive demo (`run_week2_benchmarks.py`).
  - Formal Week 2 Report documentation (`Week2_Algorithm_Exploration_and_Model_Design_Report.docx`).
- [x] **Week 3: Experimental Design & Evaluation Metric Formulation**
  - Metric formulation addressing asymmetric churn cost: Macro F1 ($\ge 0.88$), Negative Recall ($\ge 0.90$), Balanced Accuracy ($\ge 0.85$).
  - Leakage-free experimental protocol: Golden Blind Test Set (15%) + Stratified 5-Fold Cross-Validation.
  - Systematic 4-tier qualitative error analysis taxonomy (Sarcasm, Concessive Polarity, Negation, Neutral Ambiguity).
  - Experimental evaluator module (`src/experimental_framework.py`) & runner script (`run_week3_experiments.py`).
  - Formal Week 3 Report documentation (`Week3_Experimental_Design_and_Metric_Formulation_Report.docx`).
- [x] **Week 4: Implementation Strategy & Code Architecture Outline** *(Current)*
  - SOLID-compliant modular software architecture: `data_loader`, `preprocessor`, `trainer`, `inference_engine`.
  - Scikit-Learn custom `TextCleaner(BaseEstimator, TransformerMixin)` eliminating training-serving skew.
  - Automated unit testing suite (`tests/test_pipeline.py`) with 100% passing test assertion coverage.
  - Model pipeline persistence (`models/production_sentiment_pipeline.joblib`) and runner demo (`run_week4_pipeline.py`).
  - Formal Week 4 Report documentation (`Week4_Implementation_Strategy_and_Code_Architecture_Report.docx`).
- [ ] **Week 5: Packaging, Interactive Web Application & Deployment**
  - Interactive web interface development using **Streamlit**.
  - Real-time single review scoring and bulk CSV inference mode.
  - Final project report, demo video walkthrough, and completion submission.

---

## 🏛️ System Architecture

```
Raw Customer Reviews (CSV / API Streams)
                  │
                  ▼
┌────────────────────────────────────────┐
│     Text Preprocessing & Cleaning      │
│  - Lowercasing, HTML/URL stripping     │
│  - Negation-preserving stopword filter │
│  - WordNet Lemmatization               │
└──────────────────┬─────────────────────┘
                   │
                   ▼
┌────────────────────────────────────────┐
│         Feature Extraction Track       │
│  ┌──────────────────┬────────────────┐ │
│  │ TF-IDF (1-3 ngrams) │ DistilBERT    │ │
│  │ (High Throughput)│ (Contextual DL)│ │
│  └──────────────────┴────────────────┘ │
└──────────────────┬─────────────────────┘
                   │
                   ▼
┌────────────────────────────────────────┐
│        Model Inference Engine          │
│  - LinearSVC / Regularized Logistic    │
│  - Calibrated Probability Scoring      │
└──────────────────┬─────────────────────┘
                   │
                   ▼
┌────────────────────────────────────────┐
│     Interactive UI & Export Layer      │
│  - Streamlit Dashboard                 │
│  - Sentiment Tags & Actionable Alerts  │
└────────────────────────────────────────┘
```

---

## 📁 Repository Structure

```
ML-Internship/
├── .gitignore               # Git exclusions (data files, virtualenvs, cache)
├── README.md                # Project overview and sprint documentation
├── requirements.txt         # Python library dependencies
├── generate_week1_doc.py    # Script to build formatted Week 1 Word document
├── data/
│   └── README.md            # Dataset schema and ingestion instructions
├── docs/
│   └── Week1_Project_Planning_and_Strategy_Report.docx # Week 1 Official Report
├── notebooks/
│   └── README.md            # Experimental Jupyter Notebooks roadmap
└── src/
    ├── __init__.py          # Source package initialization
    ├── config.py            # Global constants and pipeline parameters
    └── utils.py             # Reusable preprocessing and metric helpers
```

---

## ⚙️ Installation & Quickstart

### 1. Clone the Repository
```bash
git clone https://github.com/<your-username>/IntelliFeedback-ML-Internship.git
cd IntelliFeedback-ML-Internship
```

### 2. Create and Activate Virtual Environment
```bash
python -m venv venv
# On Windows:
.\venv\Scripts\activate
# On Linux/macOS:
source venv/bin/activate
```

### 3. Install Required Dependencies
```bash
pip install -r requirements.txt
```

### 4. Build or Re-generate Week 1 Deliverable Report
```bash
python generate_week1_doc.py
```
The output file will be generated in `docs/Week1_Project_Planning_and_Strategy_Report.docx`.

---

## 🎯 Evaluation Framework & Target KPIs

| KPI Metric | Target Threshold | Business Justification |
| :--- | :--- | :--- |
| **Macro F1-Score** | $\ge 0.88$ | Balances predictive quality equally across imbalanced classes. |
| **Negative Class Recall** | $\ge 0.90$ | Minimizes missed customer dissatisfaction alerts and churn risks. |
| **Balanced Accuracy** | $\ge 0.85$ | Unbiased indicator across minority and majority sentiments. |
| **Inference Latency** | $< 20\text{ ms}$ | Enables sub-second real-time scoring in enterprise customer support queues. |

---

## 👤 Author & Acknowledgements
- **Intern:** Virtual Machine Learning Intern
- **Program:** YuvaIntern 5-Week Machine Learning Virtual Internship
- **Track:** Technical / Applied Machine Learning
- **Mentor / Evaluator:** YuvaIntern Assessment Team
