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

def add_callout(doc, text, title="SOFTWARE ENGINEERING PHILOSOPHY"):
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

def build_week4_document():
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
    run_kicker = title_p.add_run("Yuvaintern Virtual ML Internship | Week 4 Technical Deliverable\n")
    run_kicker.font.size = Pt(9.5)
    run_kicker.font.bold = True
    run_kicker.font.color.rgb = RGBColor(0x2B, 0x6C, 0xB0)
    
    run_title = title_p.add_run("Implementation Strategy & Code Architecture Outline")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    
    subtitle_p = doc.add_paragraph()
    subtitle_p.paragraph_format.space_after = Pt(14)
    run_sub = subtitle_p.add_run("Modular Software Design, UML Component Schematics, Pseudocode Pipelines, and Production Best Practices for IntelliFeedback")
    run_sub.font.size = Pt(12.5)
    run_sub.font.italic = True
    run_sub.font.color.rgb = RGBColor(0x4A, 0x55, 0x68)

    # Metadata Table
    meta_table = doc.add_table(rows=1, cols=4)
    meta_headers = ["Internship Sprint", "Track", "Deliverable Type", "Target Project"]
    meta_values = [["Week 4 of 5", "Machine Learning (ML)", "Code Architecture & Implementation Strategy", "IntelliFeedback (NLP Engine)"]]
    col_widths = [Inches(1.5), Inches(1.5), Inches(2.3), Inches(1.2)]
    format_table(meta_table, col_widths, meta_headers, meta_values, header_bg="2B6CB0")
    
    doc.add_paragraph().paragraph_format.space_after = Pt(8)

    # -------------------------------------------------------------
    # SECTION 1: EXECUTIVE OVERVIEW
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("1. Executive Summary & Architectural Vision")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "Transitioning a machine learning concept from exploratory Jupyter notebooks into an enterprise-ready production "
        "system requires uncompromising software engineering discipline. In Week 4 of the YuvaIntern Machine Learning Virtual "
        "Internship, we establish a comprehensive implementation strategy and modular code architecture for the IntelliFeedback "
        "sentiment intelligence engine."
    )
    doc.add_paragraph(
        "Adhering to SOLID object-oriented design principles, clean architecture patterns, and modern Python best practices, "
        "this document outlines an end-to-end blueprint designed for high maintainability, fault isolation, automated testability, "
        "and sub-20ms inference latency. We provide complete UML class and component interaction diagrams, articulate the precise "
        "responsibilities of decoupled functional packages, present production-grade pseudocode from data ingestion to model serving, "
        "and formulate a proactive risk mitigation matrix addressing concurrency, dependency drift, and artifact serialization."
    )
    
    add_callout(
        doc,
        "Production ML code must decouple feature transformations, model training, and inference serving into stateless, "
        "independently testable modules conforming to Scikit-Learn's BaseEstimator contract. This guarantees zero training-serving "
        "skew and enables seamless swapping of algorithmic backends without breaking client APIs.",
        "CORE ARCHITECTURAL PRINCIPLE"
    )

    # -------------------------------------------------------------
    # SECTION 2: STEP-BY-STEP IMPLEMENTATION STRATEGY
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("2. Step-by-Step Implementation Strategy")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The implementation lifecycle is decomposed into five progressive, verifiable phases designed to minimize technical debt "
        "and ensure seamless integration across modules:"
    )

    phases = [
        ("Phase 1: Environment & Dependency Isolation: ",
         "Establish strict runtime reproducibility using Python virtual environments, pinning package constraints in `requirements.txt`. "
         "Configure configuration management in `src/config.py` using typed dataclasses and environment variable overrides."),
        
        ("Phase 2: Ingestion & Schema Validation Layer: ",
         "Construct `src/data_loader.py` to ingest raw customer review corpora (CSV/JSON/API). Enforce schema validation (presence of `text`, "
         "`rating`, `sentiment`), strip corrupt UTF-8 sequences, drop null records, and generate deterministic train/validation/test partitions."),
        
        ("Phase 3: Decoupled Preprocessing & Feature Engineering: ",
         "Implement custom transformers (`TextCleaner`, `SublinearTfidfExtractor`) subclassing Scikit-Learn's `BaseEstimator` and `TransformerMixin`. "
         "This guarantees that text cleaning, negation retention, and vocabulary transformations are seamlessly serialized into unified pipeline objects."),
        
        ("Phase 4: Training Orchestration, Calibration & Cross-Validation: ",
         "Develop `src/trainer.py` to manage model training under Stratified 5-Fold Cross-Validation. Incorporate Probability Calibration "
         "(`CalibratedClassifierCV`) via Platt scaling / isotonic regression, evaluate multi-class metrics (Macro F1, Negative Recall), and "
         "persist winning model artifacts (`.joblib`) with SHA-256 metadata manifests."),
        
        ("Phase 5: Thread-Safe Production Inference Engine: ",
         "Build `src/inference_engine.py` as an encapsulated singleton inference service. Expose clean predict methods accepting single review "
         "strings or batched payloads, executing dual-tier routing (Tier 1 fast path vs. Tier 2 DistilBERT contextual fallback), and returning "
         "structured JSON dictionaries with actionable churn alerts.")
    ]
    for p_title, p_desc in phases:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rp = p.add_run(p_title)
        rp.bold = True
        rp.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        p.add_run(p_desc)

    # -------------------------------------------------------------
    # SECTION 3: SYSTEM ARCHITECTURE & UML COMPONENT DIAGRAMS
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("3. System Architecture & UML Component Diagrams")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The system follows a Layered Modular Architecture ensuring strict separation of concerns across data handling, "
        "feature transformation, model inference, and service presentation layers."
    )

    doc.add_paragraph("Figure 4.1: UML Class Structure & Component Hierarchy:").bold = True

    uml_diagram = (
        "+---------------------------------------------------------------------------------------+\n"
        "|                                     src.config                                        |\n"
        "|  + PipelineConfig (dataclass)                                                         |\n"
        "|    - max_features: int = 10000 | ngram_range: tuple = (1, 3) | random_state: int = 42     |\n"
        "|    - confidence_threshold: float = 0.70 | labels_map: dict                           |\n"
        "+-------------------------------------------+-------------------------------------------+\n"
        "                                            |\n"
        "                                            v\n"
        "+---------------------------------------------------------------------------------------+\n"
        "|                                  src.data_loader                                      |\n"
        "|  + DataLoader                                                                         |\n"
        "|    + load_raw_dataset(filepath: str) -> pd.DataFrame                                  |\n"
        "|    + validate_schema(df: pd.DataFrame) -> bool                                        |\n"
        "|    + split_dataset(df: pd.DataFrame, test_size=0.15) -> Tuple[X_tr, X_val, y_tr, y_val] |\n"
        "+-------------------------------------------+-------------------------------------------+\n"
        "                                            |\n"
        "                                            v\n"
        "+---------------------------------------------------------------------------------------+\n"
        "|                                 src.preprocessor                                      |\n"
        "|  + TextCleaner(BaseEstimator, TransformerMixin)                                       |\n"
        "|    + fit(X, y=None) -> self                                                           |\n"
        "|    + transform(X) -> List[str]  # URL removal, HTML stripping, negation preservation  |\n"
        "+-------------------------------------------+-------------------------------------------+\n"
        "                                            |\n"
        "                                            v\n"
        "+---------------------------------------------------------------------------------------+\n"
        "|                                    src.trainer                                        |\n"
        "|  + ModelTrainer                                                                       |\n"
        "|    - config: PipelineConfig                                                           |\n"
        "|    - pipeline: sklearn.pipeline.Pipeline                                              |\n"
        "|    + build_pipeline() -> Pipeline                                                     |\n"
        "|    + train_and_cross_validate(X_train, y_train) -> Dict[str, float]                   |\n"
        "|    + save_model(output_path: str) -> None                                             |\n"
        "+-------------------------------------------+-------------------------------------------+\n"
        "                                            |\n"
        "                                            v\n"
        "+---------------------------------------------------------------------------------------+\n"
        "|                               src.inference_engine                                    |\n"
        "|  + InferenceEngine                                                                    |\n"
        "|    - model_artifact: Any                                                              |\n"
        "|    - config: PipelineConfig                                                           |\n"
        "|    + load_artifact(artifact_path: str) -> None                                        |\n"
        "|    + predict_single(text: str) -> Dict[str, Any]                                      |\n"
        "|    + predict_batch(texts: List[str]) -> List[Dict[str, Any]]                          |\n"
        "+---------------------------------------------------------------------------------------+"
    )
    add_code_block(doc, uml_diagram)

    doc.add_paragraph("Figure 4.2: Inference Request Sequence Flow:").bold = True

    seq_flow = (
        "Client (Web UI / API)        InferenceEngine          TextCleaner          LinearSVC Classifier\n"
        "         |                         |                      |                          |\n"
        "         |--- 1. POST /predict --->|                      |                          |\n"
        "         |    (raw review text)    |--- 2. transform() -->|                          |\n"
        "         |                         |<-- 3. cleaned text --|                          |\n"
        "         |                         |                                                 |\n"
        "         |                         |--- 4. predict_proba() ------------------------->|\n"
        "         |                         |<-- 5. class probabilities [P_neg, P_neu, P_pos]-|\n"
        "         |                         |                                                 |\n"
        "         |                         |--- [Confidence >= 0.70?]                        |\n"
        "         |                         |    YES: Resolve Tier 1 Fast Path (<15ms)        |\n"
        "         |                         |    NO:  Trigger Tier 2 DistilBERT Attention     |\n"
        "         |<-- 6. JSON Payload -----|                                                 |\n"
        "         |    (Sentiment, Conf,    |                                                 |\n"
        "         |     Churn Alert Flag)   |                                                 |\n"
        "         v                         v                                                 v\n"
    )
    add_code_block(doc, seq_flow)

    # -------------------------------------------------------------
    # SECTION 4: MODULAR CODEBASE BREAKDOWN
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("4. Comprehensive Modular Codebase Breakdown")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "Each module in the `src/` directory possesses explicit boundaries, well-defined interface contracts, and isolated "
        "responsibilities to prevent circular dependencies and enable targeted unit testing:"
    )

    mod_table = doc.add_table(rows=1, cols=4)
    mod_headers = ["Module Path", "Primary Class / Functions", "Architectural Role", "Key Contracts & Dependencies"]
    mod_data = [
        [
            "src/config.py",
            "PipelineConfig, LABELS_MAP, PATHS",
            "Central source of truth for hyperparameters, directory paths, random seeds, and class constants.",
            "No dependencies. Imported by all downstream modules."
        ],
        [
            "src/utils.py",
            "clean_text(), format_metrics()",
            "Stateless text sanitization utilities, regular expressions, and reporting helpers.",
            "Standard library (re, string). Independent."
        ],
        [
            "src/data_loader.py",
            "DataLoader.load_raw_dataset(), split_dataset()",
            "Handles dataset ingestion, UTF-8 normalization, null dropping, and stratified train/val/test splitting.",
            "Pandas, Scikit-Learn (train_test_split)."
        ],
        [
            "src/preprocessor.py",
            "TextCleaner(BaseEstimator, TransformerMixin)",
            "Scikit-Learn compatible transformer implementing custom token cleaning and negation preservation.",
            "Scikit-Learn base classes. Compatible with Pipeline()."
        ],
        [
            "src/trainer.py",
            "ModelTrainer.fit(), evaluate_cv(), save_artifact()",
            "Constructs Scikit-Learn pipelines (TF-IDF + Calibrated LinearSVC), executes Stratified K-Fold CV, and serializes models.",
            "Scikit-Learn, Joblib, src.config, src.preprocessor."
        ],
        [
            "src/inference_engine.py",
            "InferenceEngine.predict_single(), predict_batch()",
            "Thread-safe production inference service executing confidence checks, sentiment mapping, and churn triage alerts.",
            "Joblib, Numpy, src.config."
        ]
    ]
    mod_widths = [Inches(1.3), Inches(1.8), Inches(2.0), Inches(1.4)]
    format_table(mod_table, mod_widths, mod_headers, mod_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # SECTION 5: PRODUCTION PSEUDOCODE
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("5. Production Pseudocode & Implementation Blueprint")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The following pseudocode details the core implementation across ingestion, preprocessing, training, and inference serving:"
    )

    core_pseudocode = (
        "# =====================================================================\n"
        "# 1. MODULAR PREPROCESSOR TRANSFORMER\n"
        "# =====================================================================\n"
        "FROM sklearn.base IMPORT BaseEstimator, TransformerMixin\n"
        "IMPORT re\n"
        "\n"
        "CLASS TextCleaner(BaseEstimator, TransformerMixin):\n"
        "    DEF fit(self, X, y=None):\n"
        "        RETURN self\n"
        "\n"
        "    DEF transform(self, X):\n"
        "        cleaned_docs = []\n"
        "        FOR doc IN X:\n"
        "            text = re.sub(r'<.*?>', ' ', str(doc))          # Strip HTML\n"
        "            text = re.sub(r'https?://\\S+', ' ', text)       # Strip URLs\n"
        "            text = re.sub(r'\\s+', ' ', text).strip().lower() # Normalize whitespace\n"
        "            cleaned_docs.append(text)\n"
        "        RETURN cleaned_docs\n"
        "\n"
        "# =====================================================================\n"
        "# 2. TRAINING ORCHESTRATOR MODULE\n"
        "# =====================================================================\n"
        "CLASS ModelTrainer:\n"
        "    DEF __init__(self, config: PipelineConfig):\n"
        "        self.config = config\n"
        "        self.pipeline = Pipeline([\n"
        "            ('cleaner', TextCleaner()),\n"
        "            ('tfidf', TfidfVectorizer(\n"
        "                ngram_range=config.ngram_range,\n"
        "                sublinear_tf=True,\n"
        "                max_features=config.max_features,\n"
        "                stop_words='english'\n"
        "            )),\n"
        "            ('clf', CalibratedClassifierCV(\n"
        "                estimator=LinearSVC(C=1.0, class_weight='balanced', random_state=config.random_state),\n"
        "                method='sigmoid',\n"
        "                cv=3\n"
        "            ))\n"
        "        ])\n"
        "\n"
        "    DEF train_and_evaluate(self, X_train, y_train, cv_folds=5):\n"
        "        cv = StratifiedKFold(n_splits=cv_folds, shuffle=True, random_state=self.config.random_state)\n"
        "        scores = cross_validate(self.pipeline, X_train, y_train, cv=cv, scoring=['f1_macro', 'recall_macro'])\n"
        "        self.pipeline.fit(X_train, y_train)\n"
        "        joblib.dump(self.pipeline, self.config.model_artifact_path)\n"
        "        RETURN {'mean_f1_macro': np.mean(scores['test_f1_macro'])}\n"
        "\n"
        "# =====================================================================\n"
        "# 3. THREAD-SAFE PRODUCTION INFERENCE SERVICE\n"
        "# =====================================================================\n"
        "CLASS InferenceEngine:\n"
        "    DEF __init__(self, artifact_path: str, config: PipelineConfig):\n"
        "        self.model = joblib.load(artifact_path)\n"
        "        self.config = config\n"
        "\n"
        "    DEF predict_single(self, text: str) -> Dict[str, Any]:\n"
        "        probabilities = self.model.predict_proba([text])[0]\n"
        "        pred_class = int(np.argmax(probabilities))\n"
        "        confidence = float(np.max(probabilities))\n"
        "\n"
        "        # Tier 2 Contextual Fallback Trigger\n"
        "        IF confidence < self.config.confidence_threshold:\n"
        "            final_sentiment = self.distilbert_predict(text)\n"
        "            routing = 'TIER2_DISTILBERT_FALLBACK'\n"
        "        ELSE:\n"
        "            final_sentiment = self.config.labels_map[pred_class]\n"
        "            routing = 'TIER1_FAST_PATH_RESOLVED'\n"
        "\n"
        "        RETURN {\n"
        "            'sentiment': final_sentiment,\n"
        "            'confidence': round(confidence, 4),\n"
        "            'routing_status': routing,\n"
        "            'churn_risk_alert': (final_sentiment == 'Negative' AND confidence >= 0.80)\n"
        "        }\n"
    )
    add_code_block(doc, core_pseudocode)

    # -------------------------------------------------------------
    # SECTION 6: BEST PRACTICES FOR MAINTAINABILITY
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("6. Python Engineering Best Practices & Maintainability")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "To prevent technical debt, ensure software maintainability, and guarantee enterprise reliability, the codebase incorporates "
        "the following industry standards:"
    )

    practices = [
        ("PEP 8 & Static Typing: ", "Every function signature includes strict Python type hints (`typing.List`, `typing.Dict`, `typing.Optional`) and descriptive docstrings adhering to the Google/NumPy standard."),
        ("Unit Testing Suite with PyTest: ", "Automated tests located in `tests/test_pipeline.py` validate transformer contracts, test serialization fidelity, and assert that edge cases (empty strings, all caps, pure punctuation) do not throw unhandled exceptions."),
        ("Stateless Design & Concurrency: ", "The `InferenceEngine` executes pure function evaluations over read-only loaded pipeline objects, enabling thread-safe parallel request handling under ASGI/WSGI servers without memory locks."),
        ("Artifact Versioning & Integrity Hashing: ", "Serialized model checkpoints (`.joblib`) are accompanied by a JSON manifest recording training timestamps, Git commit SHAs, Scikit-Learn versions, and SHA-256 artifact hashes.")
    ]
    for p_title, p_desc in practices:
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(3)
        rp = p.add_run(p_title)
        rp.bold = True
        rp.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
        p.add_run(p_desc)

    # -------------------------------------------------------------
    # SECTION 7: RISK ASSESSMENT & MITIGATION
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("7. Implementation Risks, Complexity Analysis & Mitigations")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "The following risk assessment highlights potential implementation bottlenecks and outlines proactive engineering countermeasures:"
    )

    risk_table = doc.add_table(rows=1, cols=4)
    r_headers = ["Implementation Risk", "Technical Impact", "Severity / Likelihood", "Proactive Engineering Mitigation"]
    r_data = [
        [
            "Serialization / Pickling Drift",
            "Model artifacts serialized in one environment fail to load in production due to minor Scikit-Learn version differences.",
            "High / High",
            "Enforce strict dependency pinning in `requirements.txt`. Save metadata manifests containing environment specs with automated startup checks."
        ],
        [
            "Memory Exhaustion on Sparse Matrices",
            "Constructing 10,000-dimensional TF-IDF matrices on large corpora (100k+ reviews) causes Out-Of-Memory (OOM) crashes.",
            "High / Moderate",
            "Utilize scipy.sparse CSR matrices natively supported by Scikit-Learn. Prune rare vocabulary (min_df=5) and set sublinear_tf=True."
        ],
        [
            "Training-Serving Skew",
            "Text preprocessing in production differs subtly from preprocessing during training, degrading live model accuracy.",
            "High / High",
            "Package text cleaning and tokenization inside the serialized Scikit-Learn `Pipeline` artifact itself via `TextCleaner(BaseEstimator)`."
        ],
        [
            "Concurrency Bottlenecks under Load",
            "Python Global Interpreter Lock (GIL) blocks CPU-bound inference when multiple web requests arrive simultaneously.",
            "Moderate / Moderate",
            "Deploy inference workers using multi-process web servers (Gunicorn/Uvicorn with multiple workers) rather than multi-threaded workers."
        ]
    ]
    r_widths = [Inches(1.5), Inches(1.8), Inches(1.1), Inches(2.1)]
    format_table(risk_table, r_widths, r_headers, r_data)

    doc.add_paragraph().paragraph_format.space_after = Pt(6)

    # -------------------------------------------------------------
    # SECTION 8: SUMMARY & TRANSITION TO WEEK 5
    # -------------------------------------------------------------
    h1 = doc.add_heading(level=1)
    h1.paragraph_format.space_before = Pt(14)
    h1.paragraph_format.space_after = Pt(4)
    r = h1.add_run("8. Summary & Transition to Week 5 (Deployment & Showcase)")
    r.font.color.rgb = RGBColor(0x1A, 0x36, 0x5D)
    r.font.bold = True

    doc.add_paragraph(
        "Week 4 formalizes an enterprise-grade code architecture and implementation strategy for IntelliFeedback. "
        "By decomposing the system into modular components (`data_loader`, `preprocessor`, `trainer`, `inference_engine`), "
        "enforcing strict transformer contracts, and instituting automated unit testing, the codebase achieves production readiness."
    )
    doc.add_paragraph(
        "This robust architectural foundation leads directly into Week 5 (Packaging, Deployment & Final Showcase). In Week 5, "
        "we will wrap the `InferenceEngine` in an interactive Streamlit web dashboard, provide real-time and CSV batch prediction "
        "modes, finalize complete end-to-end documentation, and deliver the final internship capstone showcase."
    )

    # Save document
    output_path = os.path.join("docs", "Week4_Implementation_Strategy_and_Code_Architecture_Report.docx")
    doc.save(output_path)
    print(f"Week 4 Report successfully created at: {output_path}")

if __name__ == "__main__":
    build_week4_document()
