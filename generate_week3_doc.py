import os
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
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

def add_callout(doc, text, title="CORE EXPERIMENTAL PRINCIPLE"):
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

def build_week3_document():
    doc = docx.Document()
    
    # 1-inch standard margins
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
    run_kicker = title_p.add_run("Yuvaintern Virtual ML Internship | Week 3 Technical Deliverable\n")
    run_kicker.font.size = Pt(9.5)
    run_kicker.font.bold = True
    run_kicker.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    
    run_title = title_p.add_run("Experimental Design & Evaluation Metric Formulation")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_after = Pt(14)
    run_sub = subtitle_p.add_run("A Methodological Validation Framework: Stratified Protocols, Multi-Class Error Analysis, and Metric Alignment for IntelliFeedback")
    run_sub.font.size = Pt(12.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Table
    meta_table = doc.add_table(rows=1, cols=4)
    meta_headers = ["Internship Sprint", "Track", "Deliverable Type", "Target Project"]
    meta_values = [["Week 3 of 5", "Machine Learning (ML)", "Experimental Protocol & Metric Design", "IntelliFeedback (NLP Engine)"]]
    col_widths = [Inches(1.5), Inches(1.5), Inches(2.3), Inches(1.2)]
    format_table(meta_table, col_widths, meta_headers, meta_values, header_bg="2B6CB0")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # SECTION 1: EXECUTIVE OVERVIEW
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("1. Executive Summary & Experimental Objectives")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "A machine learning architecture is only as sound as the empirical framework deployed to measure its performance. "
        "Following the algorithm exploration and preliminary design conducted in Week 2, the objective of Week 3 is to "
        "establish a statistically rigorous, reproducible, and objective experimental methodology for validating the "
        "IntelliFeedback sentiment intelligence platform."
    )
    doc.add_paragraph(
        "In enterprise text classification, naïve evaluation strategies—such as evaluating models using standard classification "
        "accuracy or unstratified random train-test splits—consistently introduce severe optimistic bias, obscure failures on "
        "minority classes, and permit subtle data leakage. This document formulates an end-to-end experimental protocol: "
        "defining multi-dimensional evaluation metrics tailored for asymmetric business loss, structuring a leakage-free "
        "stratified cross-validation pipeline, defining systematic error analysis taxonomies, and outlining contingency plans "
        "for anticipated experimental risks."
    )
    
    add_callout(
        doc,
        "In customer feedback classification, the cost of a False Negative on severe dissatisfaction (missed churn signal) "
        "is 5x to 10x higher than a False Positive. Consequently, our experimental design elevates Class 0 (Negative) Recall "
        "and Macro F1-score as paramount optimization objectives over aggregate accuracy.",
        "ASYMMETRIC LOSS PRINCIPLE"
    )

    # -------------------------------------------------------------
    # SECTION 2: FORMULATION OF EVALUATION METRICS
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("2. Comprehensive Formulation of Evaluation Metrics")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "To capture every dimension of model behavior—from discrimination ability and probability calibration to operational "
        "latency—we establish the following evaluation suite:"
    )

    # 2.1 Multi-class Accuracy vs Balanced Accuracy
    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.1 Overall Accuracy vs. Balanced Accuracy")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Overall Accuracy measures the fraction of all correct predictions: Accuracy = (TP + TN) / (TP + TN + FP + FN). "
        "However, in consumer review corpora, positive reviews characteristically dominate (~65-75% of total volume), while neutral (~10-15%) "
        "and negative (~15-20%) reviews form a distinct minority. A degenerate classifier predicting 'Positive' indiscriminately could achieve "
        "75% accuracy while failing 100% of negative churn warnings."
    )
    doc.add_paragraph(
        "To guard against this majority-class delusion, we incorporate Balanced Accuracy, defined as the unweighted macro-average "
        "of recall obtained on each individual class: Balanced Accuracy = 1/K ∑_k (TP_k / (TP_k + FN_k)). It guarantees that performance "
        "on rare negative feedback is weighted identically to common positive feedback."
    )

    # 2.2 Precision, Recall, and F1 Formulation
    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.2 Precision, Recall, and Macro-Averaged F1-Score")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "For multiclass sentiment (k ∈ {Negative, Neutral, Positive}):"
    )
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Precision (Positive Predictive Value): ").bold = True
    p.add_run("Precision_k = TP_k / (TP_k + FP_k). Measures the purity of alarms. High precision ensures support agents do not waste time investigating false alarms.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Recall (Sensitivity): ").bold = True
    p.add_run("Recall_k = TP_k / (TP_k + FN_k). Measures coverage. For the Negative class, Recall is critical to ensure no defective product reviews bypass retention queues.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Macro F1-Score: ").bold = True
    p.add_run("Macro F1 = 1/3 (F1_neg + F1_neu + F1_pos), where F1_k = 2 · (Prec_k · Rec_k) / (Prec_k + Rec_k). Unlike Weighted F1 (which weights scores by class frequency, biasing results towards the majority), Macro F1 treats all three sentiment classes with equal importance.")

    # 2.3 Probabilistic & Operational Metrics
    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.3 Probabilistic Calibration & Operational Service-Level Metrics")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Because our architecture features a confidence-based routing layer (Tier 1 fast-path vs. Tier 2 DistilBERT escalation), "
        "the raw probability outputs must reflect true empirical likelihoods (well-calibrated probabilities):"
    )
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Multiclass ROC-AUC (One-vs-Rest): ").bold = True
    p.add_run("Evaluates the degree of class separability across all possible decision thresholds, computing the macro-average of pairwise OvR ROC areas.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Log-Loss (Cross-Entropy): ").bold = True
    p.add_run("LogLoss = - 1/N ∑_i ∑_k y_{i,k} log(p_{i,k}). Heavily penalizes overconfident incorrect classifications.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Inference Latency (p95): ").bold = True
    p.add_run("The 95th percentile execution time per single review prediction in milliseconds, ensuring sub-second response times under concurrent load.")

    # Summary Table
    h2 = doc.add_heading(level=2)
    r = h2.add_run("2.4 Consolidated Evaluation Metric Matrix")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)

    metric_table = doc.add_table(rows=1, cols=5)
    m_headers = ["Metric Name", "Mathematical Expression", "Optimization Goal", "Strategic Business Rationale", "Target KPI"]
    m_data = [
        ["Macro F1-Score", "1/K ∑ [2·P_k·R_k / (P_k + R_k)]", "Maximize", "Equal evaluation weight across all classes; guards against class collapse.", "≥ 0.88"],
        ["Negative Recall", "TP_neg / (TP_neg + FN_neg)", "Maximize", "Zero tolerance for missed churn & severe customer complaint signals.", "≥ 0.90"],
        ["Balanced Accuracy", "1/K ∑ Recall_k", "Maximize", "Unbiased accuracy measure independent of dataset class imbalance.", "≥ 0.85"],
        ["Multiclass ROC-AUC", "Macro Avg of OvR AUC", "Maximize", "Measures discriminatory power independent of chosen threshold cutoffs.", "≥ 0.92"],
        ["Log-Loss", "- 1/N ∑ ∑ y_{ik} log(p_{ik})", "Minimize", "Ensures calibrated confidence values for reliable Tier 2 escalation routing.", "≤ 0.35"],
        ["p95 Inference Latency", "95th percentile(t_predict)", "Minimize", "Maintains SLA guarantees for high-volume customer streaming APIs.", "< 20 ms (Tier 1)"]
    ]
    m_widths = [Inches(1.3), Inches(1.6), Inches(0.9), Inches(1.9), Inches(0.8)]
    format_table(metric_table, m_widths, m_headers, m_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # SECTION 3: EXPERIMENTAL DESIGN & PROTOCOL
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("3. End-to-End Experimental Architecture & Protocol")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "To generate unimpeachable scientific validity, the experiment is partitioned into strict stages with isolated data boundaries:"
    )

    doc.add_paragraph("Figure 3.1: Strict Experimental Workflow & Data Isolation Architecture:").bold = True

    ascii_flow = (
        "+----------------------------------------------------------------------------------------+\n"
        "|                        TOTAL CORPUS: 20,000+ CUSTOMER REVIEWS                          |\n"
        "+-------------------------------------------+--------------------------------------------+\n"
        "                                            |\n"
        "             +------------------------------+-----------------------------+\n"
        "             | (85% Development Split)                                    | (15% Blind Test)\n"
        "             v                                                            v\n"
        "+---------------------------------------+               +--------------------------------+\n"
        "|     DEVELOPMENT DATASET (N=17,000)    |               |    GOLDEN TEST SET (N=3,000)   |\n"
        "|  - Stratified by Target & Domain      |               |  - Completely Isolated         |\n"
        "+-------------------+-------------------+               |  - Zero Feature Fitting        |\n"
        "                    |                                   |  - Touched ONLY at final signoff|\n"
        "                    v                                   +----------------+---------------+\n"
        "+---------------------------------------+                                |\n"
        "|   5-FOLD STRATIFIED CROSS VALIDATION  |                                |\n"
        "|  +---------------------------------+  |                                |\n"
        "|  | Fold 1: 80% Train | 20% Val     |  |                                |\n"
        "|  | Fold 2: 80% Train | 20% Val     |  |                                |\n"
        "|  | Fold 3: 80% Train | 20% Val     |  |                                |\n"
        "|  | Fold 4: 80% Train | 20% Val     |  |                                |\n"
        "|  | Fold 5: 80% Train | 20% Val     |  |                                |\n"
        "|  +---------------------------------+  |                                |\n"
        "|  * Note: Vectorizer fitted ONLY on     |                                |\n"
        "|    internal fold training splits      |                                |\n"
        "+-------------------+-------------------+                                |\n"
        "                    |                                                    |\n"
        "                    v                                                    |\n"
        "+---------------------------------------+                                |\n"
        "|   HYPERPARAMETER GRID SEARCH & TUNING |                                |\n"
        "|   (C, sublinear_tf, n-gram ranges)    |                                |\n"
        "+-------------------+-------------------+                                |\n"
        "                    | Best Estimator Retrained on Full Dev Set           |\n"
        "                    +----------------------------------------------------+ \n"
        "                                            |\n"
        "                                            v\n"
        "+----------------------------------------------------------------------------------------+\n"
        "|                            FINAL BENCHMARK EVALUATION STAGE                            |\n"
        "|   - Blind Hold-Out Scoring (Macro F1, Confusion Matrix, Negative Recall)               |\n"
        "|   - Bootstrap Confidence Intervals (95% BCa, 1,000 iterations)                         |\n"
        "|   - Statistical Significance (Pairwise McNemar's Test vs. Baselines)                   |\n"
        "+----------------------------------------------------------------------------------------+"
    )
    add_code_block(doc, ascii_flow)

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.1 Data Partitioning & Stratification Protocol")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Data leakage is the most pervasive flaw in applied NLP experiments. It frequently occurs when feature transformations "
        "(such as computing IDF vocabulary frequencies or global target statistics) are performed on the complete corpus before splitting. "
        "Our protocol enforces strict isolation: the Golden Blind Test Set (15%, N ≈ 3,000) is locked away prior to any feature extraction, "
        "imputation, or vocabulary generation. All TF-IDF parameters and vocabulary dictionaries are learned exclusively inside the training folds."
    )

    h2 = doc.add_heading(level=2)
    r = h2.add_run("3.2 Stratified 5-Fold Cross-Validation Framework")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Rather than relying on a single train-validation split (which suffers from high sampling variance), we employ Stratified 5-Fold "
        "Cross-Validation. In each fold, the class frequency ratio (Negative : Neutral : Positive) is strictly preserved across the 80% "
        "training and 20% validation splits. Performance metrics are reported as (Mean ± Standard Deviation) across the 5 validation folds."
    )

    # -------------------------------------------------------------
    # SECTION 4: ERROR ANALYSIS & FAILURE TAXONOMY
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("4. Error Analysis Procedures & Failure Mode Taxonomy")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "High aggregate metrics often conceal dangerous localized systematic failures. To guarantee that IntelliFeedback is robust "
        "in production, our experimental design mandates a 3-tier Error Analysis Procedure:"
    )

    # 4.1 Confusion Matrix
    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.1 Normalized Confusion Matrix & Off-Diagonal Cost Analysis")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "Raw prediction counts are converted into row-normalized transition probabilities: C_{i,j} = P(Predicted = j | True = i). "
        "We specifically audit two severe off-diagonal failure cells:"
    )
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Critical Error Cell C(Negative, Positive): ").bold = True
    p.add_run("A truly negative review misclassified as positive. This is the costliest operational error, representing catastrophic failure to detect an angry customer.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Boundary Error Cell C(Neutral, Negative) & C(Neutral, Positive): ").bold = True
    p.add_run("Examines the ambiguity boundary around 3-star reviews where sentiment polarity is mixed or tepid.")

    # 4.2 Qualitative Failure Taxonomy
    h2 = doc.add_heading(level=2)
    r = h2.add_run("4.2 Qualitative Failure Taxonomy & Root-Cause Categories")
    r.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    doc.add_paragraph(
        "All misclassified instances from the hold-out validation folds are automatically routed into an error analysis inspector "
        "and categorized into four linguistic failure modes:"
    )

    err_table = doc.add_table(rows=1, cols=4)
    e_headers = ["Failure Category", "Linguistic Pattern / Example", "Primary Model Root Cause", "Engineering Remediation"]
    e_data = [
        [
            "Category A:\nSarcasm & Irony",
            "\"Oh wonderful, the phone died in 10 minutes. What a great invention!\"",
            "N-grams see positive lexical tokens ('wonderful', 'great') and ignore ironic context.",
            "Route reviews where Tier 1 confidence is borderline (<0.70) to Tier 2 DistilBERT contextual attention."
        ],
        [
            "Category B:\nMixed / Concessive Polarity",
            "\"The screen is gorgeous and sound is amazing, but the battery exploded.\"",
            "TF-IDF feature weights for multiple positive clauses overshadow the single critical negative clause.",
            "Weight negative keyword tokens higher; incorporate aspect-level sentence segmentation."
        ],
        [
            "Category C:\nComplex Negation Scope",
            "\"I can not say that I ever found a single feature that worked well.\"",
            "Long-distance dependency between negation 'not' and target phrase exceeds bi-gram/tri-gram window.",
            "Implement sublinear TF-IDF with tri-grams (1, 3); utilize dependency parsing / transformer attention."
        ],
        [
            "Category D:\nSubjective Neutral Ambiguity",
            "\"It is an ordinary cable. Arrived in a box. It plugs into the wall.\"",
            "Factual/objective reviews lack polarized sentiment vocabulary, drifting into random classification.",
            "Calibrate decision thresholds; assign Neutral as the default fallback for zero-polarity vectors."
        ]
    ]
    e_widths = [Inches(1.5), Inches(2.0), Inches(1.5), Inches(1.5)]
    format_table(err_table, e_widths, e_headers, e_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # SECTION 5: STATISTICAL HYPOTHESIS TESTING
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("5. Statistical Significance & Hypothesis Testing Protocol")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "To verify that performance improvements between baseline models (e.g., MNB vs. Calibrated LinearSVC vs. DistilBERT) "
        "are statistically significant rather than artifacts of stochastic data variation, we enforce formal hypothesis testing:"
    )
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Pairwise McNemar's Test: ").bold = True
    p.add_run("A non-parametric test applied to the contingency table of paired predictions on the test set. Evaluates whether the marginal frequencies of discordant classifications (Model A correct, Model B wrong vs. vice versa) differ significantly with p < 0.01.")
    p = doc.add_paragraph(style='List Bullet')
    p.add_run("Bootstrap Confidence Intervals (95% BCa): ").bold = True
    p.add_run("We compute 1,000 bootstrap resamples of the blind hold-out test set to estimate non-parametric 95% confidence intervals for Macro F1 and Negative Recall, ensuring reported gains are statistically robust.")

    # -------------------------------------------------------------
    # SECTION 6: RISK ASSESSMENT & MITIGATION PLAN
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("6. Experimental Risks, Limitations, and Mitigation Strategies")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The following risk matrix identifies foreseeable threats to experimental validity, their root causes, and our planned countermeasures:"
    )

    risk_table = doc.add_table(rows=1, cols=4)
    r_headers = ["Identified Risk", "Threat to Experimental Validity", "Likelihood / Severity", "Planned Technical Countermeasure"]
    r_data = [
        [
            "Information Leakage via Preprocessing",
            "Fitting TF-IDF vectorizers or tokenizers on entire dataset before splitting produces unrealistically high validation scores.",
            "High / Critical",
            "Encapsulate all vectorization and model steps inside Scikit-Learn `Pipeline` objects, guaranteeing vectorizers fit strictly on fold training splits."
        ],
        [
            "Class Imbalance Distortion",
            "High proportion of positive reviews causes models to ignore negative classes while maintaining deceptive 80%+ accuracy.",
            "High / High",
            "Set `class_weight='balanced'` in estimators, prioritize Balanced Accuracy and Macro F1, and monitor class-specific confusion matrices."
        ],
        [
            "Hyperparameter Optimization Overfitting",
            "Aggressive grid search on a single validation fold overfits hyperparameter choices to that specific sample.",
            "Medium / Moderate",
            "Perform GridSearchCV strictly under 5-Fold Stratified Cross-Validation with fixed random seeds (`random_state=42`)."
        ],
        [
            "Crowd-Sourced Label Noise",
            "Customer star ratings do not always match written text (e.g., 5-star rating with complaint text like 'terrible').",
            "Medium / Moderate",
            "Implement high-loss sample auditing: filter out reviews where rating contradicts text sentiment using heuristic sentiment rules."
        ],
        [
            "Hardware Constraints during Deep Learning",
            "Fine-tuning DistilBERT on local CPU incurs timeout or memory exhaustion during 5-fold evaluation.",
            "High / Low",
            "Pre-compute embeddings where feasible; restrict batch size to 16 with gradient accumulation; execute heavy fine-tuning on Google Colab GPU tier."
        ]
    ]
    r_widths = [Inches(1.5), Inches(1.8), Inches(1.1), Inches(2.1)]
    format_table(risk_table, r_widths, r_headers, r_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # SECTION 7: SUMMARY & WEEK 4 TRANSITION PLAN
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("7. Summary & Transition Plan for Week 4")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "Week 3 establishes an airtight experimental design and metric formulation for IntelliFeedback. By rejecting naive accuracy "
        "in favor of Macro F1, Negative Class Recall, and Balanced Accuracy, the experimental setup aligns directly with enterprise customer "
        "retention objectives. The strict data partition protocol, Stratified 5-Fold Cross-Validation, and automated error analysis "
        "ensure that model performance can be measured with scientific integrity."
    )
    doc.add_paragraph(
        "Armed with this evaluation framework, the project transitions into Week 4 (Model Tuning, Deep Learning & Optimization). "
        "Week 4 will execute the grid search optimization over the full 20,000+ review dataset, fine-tune the DistilBERT transformer, "
        "perform error analysis on edge cases, and produce final persisted model artifacts (.joblib / .pt) ready for deployment."
    )

    # Save document
    output_path = os.path.join("docs", "Week3_Experimental_Design_and_Metric_Formulation_Report.docx")
    doc.save(output_path)
    print(f"Week 3 Report successfully created at: {output_path}")

if __name__ == "__main__":
    build_week3_document()
