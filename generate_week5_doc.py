import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def set_cell_background(cell, fill_hex):
    """Sets background color of a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """Sets inner padding for table cells (in dxa)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def add_callout(doc, text, title="RESPONSIBLE MLOPS PRINCIPLE"):
    """Adds a stylish callout box with a colored left accent border."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=120, bottom=120, left=180, right=180)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="24" w:space="0" w:color="1A365D"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run_t = p.add_run(f"📌 {title}: ")
    run_t.bold = True
    run_t.font.name = "Calibri"
    run_t.font.size = Pt(10.5)
    run_t.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    
    run_body = p.add_run(text)
    run_body.font.name = "Calibri"
    run_body.font.size = Pt(10.5)
    run_body.font.italic = True
    run_body.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def format_table(table, col_widths, headers, data, alt_bg="F7FAFC", header_bg="1A365D"):
    """Formats a modern table with colored header, padding, and alternating row backgrounds."""
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Header Row
    hdr_cells = table.rows[0].cells
    for i, title in enumerate(headers):
        hdr_cells[i].text = title
        hdr_cells[i].width = col_widths[i]
        set_cell_background(hdr_cells[i], header_bg)
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        for run in p.runs:
            run.font.bold = True
            run.font.name = "Calibri"
            run.font.size = Pt(10)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            
    # Data Rows
    for r_idx, row_data in enumerate(data):
        row_cells = table.add_row().cells
        bg_color = alt_bg if r_idx % 2 == 1 else "FFFFFF"
        for c_idx, val in enumerate(row_data):
            row_cells[c_idx].text = str(val)
            row_cells[c_idx].width = col_widths[c_idx]
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=80, bottom=80, left=140, right=140)
            p = row_cells[c_idx].paragraphs[0]
            p.paragraph_format.line_spacing = 1.15
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            for run in p.runs:
                run.font.name = "Calibri"
                run.font.size = Pt(9.5)
                run.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)

    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E0"/>'
        f'<w:left w:val="none"/>'
        f'<w:bottom w:val="single" w:sz="8" w:space="0" w:color="1A365D"/>'
        f'<w:right w:val="none"/>'
        f'<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
        f'<w:insideV w:val="none"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def add_code_block(doc, code_str):
    """Renders formatted pseudocode / diagrams in a mono-spaced shaded container."""
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    cell.width = Inches(6.5)
    
    set_cell_background(cell, "2D3748")
    set_cell_margins(cell, top=100, bottom=100, left=140, right=140)
    
    tcPr = cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="none"/>'
        f'<w:left w:val="single" w:sz="16" w:space="0" w:color="4FD1C5"/>'
        f'<w:bottom w:val="none"/>'
        f'<w:right w:val="none"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.05
    run = p.add_run(code_str)
    run.font.name = "Consolas"
    run.font.size = Pt(8.5)
    run.font.color.rgb = RGBColor(0xE2, 0xE8, 0xF0)
    
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def build_week5_document():
    doc = docx.Document()
    
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        
    normal_style = doc.styles['Normal']
    normal_style.font.name = 'Calibri'
    normal_style.font.size = Pt(11)
    normal_style.font.color.rgb = RGBColor(0x2D, 0x37, 0x48)
    normal_style.paragraph_format.line_spacing = 1.15
    normal_style.paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # DOCUMENT HEADER / BANNER
    # -------------------------------------------------------------
    title_p = doc.add_paragraph()
    title_p.paragraph_format.space_before = Pt(0)
    title_p.paragraph_format.space_after = Pt(2)
    run_kicker = title_p.add_run("Yuvaintern Virtual ML Internship | Week 5 Capstone Deliverable\n")
    run_kicker.font.size = Pt(9.5)
    run_kicker.font.bold = True
    run_kicker.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    
    run_title = title_p.add_run("Deployment Planning, Ethical Considerations & Reflective Analysis")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_after = Pt(14)
    run_sub = subtitle_p.add_run("Industrializing IntelliFeedback: MLOps Architecture, Responsible AI Governance, and 5-Week Internship Retrospective")
    run_sub.font.size = Pt(12.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Table
    meta_table = doc.add_table(rows=1, cols=4)
    meta_headers = ["Internship Sprint", "Track", "Deliverable Type", "Target Project"]
    meta_values = [["Week 5 of 5 (Final)", "Machine Learning (ML)", "Deployment, Ethics & Capstone Report", "IntelliFeedback (NLP Engine)"]]
    col_widths = [Inches(1.5), Inches(1.5), Inches(2.3), Inches(1.2)]
    format_table(meta_table, col_widths, meta_headers, meta_values, header_bg="2B6CB0")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # SECTION 1: EXECUTIVE SUMMARY
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("1. Executive Summary & Capstone Context")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The fifth and culminating week of the YuvaIntern Machine Learning Virtual Internship synthesizes the architectural, "
        "algorithmic, and engineering foundations established across Weeks 1 through 4 into an enterprise production blueprint. "
        "While previous sprints addressed strategic planning, algorithm comparison, leakage-free metric formulation, and modular "
        "code architecture, Week 5 focuses on the operational reality of machine learning: industrial deployment planning, ethical "
        "governance, data privacy compliance, and reflective technical analysis."
    )
    doc.add_paragraph(
        "Deploying an NLP sentiment intelligence system into customer-facing environments demands uncompromising standards for "
        "scalability, low latency (<20ms SLA), and system reliability. Concurrently, automated decision systems carry profound ethical "
        "responsibilities: protecting user privacy against Personally Identifiable Information (PII) leakage, mitigating demographic "
        "and dialectical biases, and delivering transparent, explainable predictions. This capstone document formulates a comprehensive "
        "cloud-native deployment architecture, establishes a Responsible AI framework, analyzes deployment failure modes, and presents "
        "a critical reflection on the 5-week engineering journey."
    )
    
    add_callout(
        doc,
        "A machine learning system is never truly complete upon model training; its enduring value is governed by its reliability "
        "in production, its resilience against data drift, and the ethical safeguards protecting user privacy and algorithmic fairness.",
        "CAPSTONE GUIDING PRINCIPLE"
    )

    # -------------------------------------------------------------
    # SECTION 2: PRODUCTION DEPLOYMENT STRATEGY
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("2. Comprehensive Production Deployment Strategy")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "To satisfy enterprise availability and scalability requirements, IntelliFeedback is engineered as a containerized, "
        "cloud-native microservice architecture following GitOps and MLOps best practices:"
    )

    doc.add_paragraph("Figure 5.1: Cloud-Native Microservice Deployment Architecture:").bold = True

    cloud_ascii = (
        "+---------------------------------------------------------------------------------------+\n"
        "|                                CLIENT TRAFFIC INGRESS                                 |\n"
        "|  - Public Customer Streams / Support Ticketing Webhooks / Interactive Streamlit UI   |\n"
        "+-------------------------------------------+-------------------------------------------+\n"
        "                                            |\n"
        "                                            v HTTPS (Port 443)\n"
        "+---------------------------------------------------------------------------------------+\n"
        "|                    INGRESS CONTROLLER & REVERSE PROXY (NGINX)                         |\n"
        "|  - Rate Limiting (1,000 req/min) | SSL Termination | PII Scrubbing Interceptor        |\n"
        "+-------------------------------------------+-------------------------------------------+\n"
        "                                            |\n"
        "                                            v Internal Virtual Network\n"
        "+---------------------------------------------------------------------------------------+\n"
        "|                 KUBERNETES CONTAINER PODS (HORIZONTAL POD AUTOSCALER)                 |\n"
        "|  +---------------------------------------------------------------------------------+  |\n"
        "|  | Docker Container (Python 3.12 / FastAPI ASGI Server / Gunicorn Workers)         |  |\n"
        "|  |  [1] Preprocessor Transformer (TextCleaner)                                     |  |\n"
        "|  |  [2] Serialized Pipeline Artifact (Calibrated LinearSVC + Sublinear TF-IDF)     |  |\n"
        "|  |  [3] Tier 2 Transformer Service (DistilBERT Contextual Engine via TorchScript)  |  |\n"
        "|  +---------------------------------------------------------------------------------+  |\n"
        "+-----------------------+---------------------------------------+-----------------------+\n"
        "                        |                                       |\n"
        "                        v                                       v\n"
        "+-----------------------------------------------+  +------------------------------------+\n"
        "|            REDIS IN-MEMORY CACHE              |  |      PROMETHEUS & GRAFANA STACK    |\n"
        "|  - Caches identical review hashes for 1 hour  |  |  - Real-time p95/p99 Latency Gauges|\n"
        "|  - Sub-millisecond hit response (<2ms)        |  |  - Input Drift & PSI Distribution  |\n"
        "+-----------------------------------------------+  +------------------------------------+\n"
    )
    add_code_block(doc, cloud_ascii)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.1 Infrastructure & Containerization Blueprint")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "The application is packaged as an immutable, lightweight OCI-compliant Docker image based on `python:3.12-slim`. "
        "The container isolates system libraries, dependencies, and pre-compiled model artifacts (`models/production_sentiment_pipeline.joblib`), "
        "eliminating environment divergence across development, staging, and production clusters. Container orchestration is managed via "
        "Kubernetes (AWS EKS or GCP GKE) with Horizontal Pod Autoscaling (HPA) configured to dynamically scale pod replicas from 2 to 20 "
        "based on CPU utilization (>70%) and concurrent request queue depth."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.2 CI/CD Pipeline & Automated Release Gates")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Deployment follows a strict automated Continuous Integration and Continuous Deployment (CI/CD) pipeline managed via GitHub Actions:"
    )
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Automated Linting & Unit Testing Gate: ").bold = True
    p.add_run("Every pull request executes `pytest tests/test_pipeline.py` and `flake8/black` formatting checks. Builds fail if test coverage drops below 90%.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Model Validation & Regression Gate: ").bold = True
    p.add_run("New candidate models are benchmarked against the Golden Blind Test Set. Release is blocked if Macro F1 drops below 0.88 or Negative Recall falls below 0.90.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Canary Deployment: ").bold = True
    p.add_run("New container versions receive 5% of production traffic for 30 minutes. Telemetry is monitored for latency regressions or 5xx error spikes before 100% rollout.")

    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.3 Continuous Monitoring & Drift Detection Protocol")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Post-deployment degradation is tracked via two automated telemetry monitors: (1) Data Drift Monitoring using the "
        "Population Stability Index (PSI) and Kolmogorov-Smirnov test over incoming review lengths and TF-IDF feature distributions; "
        "and (2) Concept Drift Monitoring tracking human auditor disagreement rates. When PSI exceeds 0.25, automated webhooks alert MLOps "
        "engineers and trigger retraining pipelines on newly collected feedback."
    )

    # -------------------------------------------------------------
    # SECTION 3: ETHICAL CONSIDERATIONS & RESPONSIBLE AI
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("3. Ethical Considerations, Algorithmic Fairness & Privacy")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The automated classification of human sentiment directly influences customer prioritization, support escalations, and product "
        "decisions. We establish strict ethical safeguards across bias mitigation, privacy compliance, and model transparency:"
    )

    # 3.1 Bias & Fairness
    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.1 Algorithmic Bias Identification & Mitigation")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Natural language corpora frequently reflect societal, regional, and demographic biases. Specifically, NLP sentiment models "
        "often display systemic prejudice against non-standard English dialects (e.g., African American Vernacular English - AAVE, "
        "Indian English colloquialisms) and informal vernacular, incorrectly tagging informal vernacular as 'Negative' or aggressive."
    )
    doc.add_paragraph(
        "Countermeasure: We conduct Disaggregated Slice-Based Auditing across dialectical markers and enforce Equalized Odds: "
        "P(Y_pred = 1 | Y_true = 1, Group = A) = P(Y_pred = 1 | Y_true = 1, Group = B). Data augmentation techniques inject dialectical "
        "variants into training folds to balance sensitivity across diverse cultural customer bases."
    )

    # 3.2 Data Privacy
    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.2 Data Privacy, Anonymization & Regulatory Compliance (GDPR / CCPA)")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Customer reviews frequently contain unprompted Personally Identifiable Information (PII)—including phone numbers, delivery addresses, "
        "email addresses, and employee names. Transmitting or storing unredacted PII violates GDPR Article 6 and CCPA mandates."
    )
    doc.add_paragraph(
        "Countermeasure: IntelliFeedback incorporates a stateless Regex & Named Entity Recognition (NER) PII Scrubber in the ingestion layer. "
        "All emails, telephone numbers, credit card sequences, and personal identifiers are masked into standardized tokens (`[REDACTED_EMAIL]`, "
        "`[REDACTED_PHONE]`) prior to feature extraction and database persistence."
    )

    # 3.3 Explainability
    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.3 Model Explainability, Transparency & Human Oversight")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Black-box decision making undermines stakeholder trust. To ensure transparent operations, IntelliFeedback couples every "
        "prediction with token-level feature attribution using SHAP (SHapley Additive exPlanations) and linear hyperplane coefficients. "
        "Customer service agents are presented with highlighted keywords (e.g., 'cracked screen', 'rude staff') that directly drove the "
        "negative classification, facilitating fast and objective customer resolution."
    )

    # -------------------------------------------------------------
    # SECTION 4: DEPLOYMENT RISK MANAGEMENT
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("4. Deployment Risk Management & Contingency Framework")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The table below identifies critical production vulnerabilities, their business impact, and our concrete contingency protocols:"
    )

    risk_table = doc.add_table(rows=1, cols=4)
    r_headers = ["Deployment Risk Factor", "Operational Threat", "Severity / Likelihood", "Contingency & Fallback Strategy"]
    r_data = [
        [
            "Model & Concept Drift",
            "Shifting consumer slang and seasonal promotion jargon degrade classification accuracy silently.",
            "High / High",
            "Automated PSI telemetry monitors; weekly shadow-scoring against ground-truth human audits; automated retraining pipelines."
        ],
        [
            "Inference Latency Spike / SLA Breach",
            "High-traffic flash sales spike request volume, exceeding the 20ms p95 latency budget.",
            "High / Moderate",
            "Redis caching for identical reviews; Tier 1 LinearSVC fast-path (<12ms) absorbs 95% of traffic; circuit breaker sheds Tier 2 load."
        ],
        [
            "PII Data Ingestion Leakage",
            "Customer enters sensitive credit card or health details in review text, violating privacy laws.",
            "Critical / Moderate",
            "Pre-ingress stateless PII sanitization regex scrubber; strictly zero raw text logging in telemetry databases."
        ],
        [
            "Adversarial Review Bombing",
            "Competitors or bots inject perturbed adversarial text (e.g., 'g00d but b.r.o.k.e.n') to bypass sentiment filters.",
            "Moderate / Low",
            "Character normalization; leetspeak and punctuation stripping in TextCleaner transformer; rate-limiting by client IP."
        ],
        [
            "False Alarm Fatigue for Support Teams",
            "Too many false positive negative classifications overwhelm customer support triage queues.",
            "Moderate / Moderate",
            "Urgent escalation webhook is restricted strictly to Negative reviews with confidence ≥ 0.80 and verified churn intent keywords."
        ]
    ]
    r_widths = [Inches(1.5), Inches(1.8), Inches(1.1), Inches(2.1)]
    format_table(risk_table, r_widths, r_headers, r_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # SECTION 5: REFLECTIVE ANALYSIS OF LEARNING JOURNEY
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("5. Reflective Analysis of the 5-Week Internship Journey")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The 5-week YuvaIntern Machine Learning Virtual Internship provided an exceptional opportunity to experience the complete "
        "end-to-end lifecycle of an applied machine learning project. The transition from theoretical concepts to deployable software "
        "yielded profound technical insights:"
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("5.1 Weekly Technical Progression & Milestones")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    sprint_table = doc.add_table(rows=1, cols=3)
    s_headers = ["Sprint Week", "Core Technical Focus", "Key Milestone Achieved"]
    s_data = [
        ["Week 1", "Project Planning & Strategy Development", "Formulated problem statement, hypotheses, 5-week roadmap, and official GitHub codebase."],
        ["Week 2", "Algorithm Exploration & Model Design", "Benchmarked 5 candidate ML algorithms; designed dual-tier LinearSVC + DistilBERT architecture."],
        ["Week 3", "Experimental Setup & Metric Formulation", "Formulated asymmetric loss metrics (Macro F1, Negative Recall); designed leakage-free 5-fold CV."],
        ["Week 4", "Code Architecture & Modular Pipelines", "Engineered SOLID-compliant packages (`preprocessor`, `trainer`, `inference`); built 100% passing unit tests."],
        ["Week 5", "Deployment, Ethics & Capstone Showcase", "Architected cloud MLOps deployment, Responsible AI governance, Streamlit web app, and capstone report."]
    ]
    s_widths = [Inches(1.2), Inches(2.3), Inches(3.0)]
    format_table(sprint_table, s_widths, s_headers, s_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("5.2 Key Insights Gained & Challenges Overcome")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    
    insights = [
        ("The Pitfall of Naive Accuracy: ", "Early in the project, overall accuracy appeared deceptively high (~75%), yet the model completely missed low-frequency negative reviews. Transitioning to Macro F1 and Negative Class Recall completely refocused our optimization on true business value."),
        ("Preventing Training-Serving Skew: ", "A common failure in ML engineering is maintaining separate cleaning scripts for training and production. By wrapping `TextCleaner` as a Scikit-Learn `BaseEstimator` inside the serialized `Pipeline`, we ensured identical text transformations in training and production."),
        ("The Efficiency of Dual-Tier Architecture: ", "Deploying massive transformer models for every request is computationally wasteful. Our dual-tiered design demonstrates that lightweight linear models (LinearSVC) can handle 95% of traffic with microsecond latency, reserving deep learning only for low-confidence ambiguities.")
    ]
    for i_title, i_desc in insights:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        ri = p.add_run(i_title)
        ri.bold = True
        ri.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        p.add_run(i_desc)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("5.3 Future Roadmap & Technical Enhancements")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Looking forward, IntelliFeedback is primed for three advanced technical horizons: (1) Multilingual Adaptation by fine-tuning "
        "multilingual representations (XLM-RoBERTa) to support international customer reviews across Spanish, German, and Hindi; "
        "(2) Aspect-Based Sentiment Analysis (ABSA) decomposing reviews into specific operational categories (e.g., Shipping vs. Product Quality vs. Customer Service); "
        "and (3) Active Learning feedback loops enabling customer support agents to flag misclassifications with one click, continuously retraining models."
    )

    # -------------------------------------------------------------
    # SECTION 6: CONCLUSION & CONCLUSION
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("6. Conclusion & Internship Completion")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "Through the 5-week YuvaIntern Machine Learning Virtual Internship, the IntelliFeedback platform evolved from an initial concept "
        "into a complete, production-ready, and ethically grounded machine learning solution. The resulting platform delivers a high-accuracy, "
        "low-latency sentiment intelligence engine backed by rigorous experimental validation, clean software architecture, and responsible AI governance. "
        "This project showcases a comprehensive understanding of end-to-end applied machine learning engineering suitable for modern enterprise environments."
    )

    # Save document
    output_path = os.path.join("docs", "Week5_Deployment_Planning_Ethical_Considerations_and_Reflective_Analysis.docx")
    doc.save(output_path)
    print(f"Week 5 Report successfully created at: {output_path}")

if __name__ == "__main__":
    build_week5_document()
