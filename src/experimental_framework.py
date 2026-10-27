"""
IntelliFeedback - Experimental Design & Evaluation Metric Framework
Implements the Week 3 experimental protocol:
- Leakage-free Stratified K-Fold evaluation
- Multiclass metric formulation (Macro F1, Negative Recall, Balanced Accuracy)
- Automated qualitative error analysis & failure mode taxonomy
"""
from typing import Dict, List, Any, Tuple
import numpy as np
from sklearn.pipeline import Pipeline
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.calibration import CalibratedClassifierCV
from sklearn.model_selection import StratifiedKFold
from sklearn.metrics import (
    accuracy_score,
    balanced_accuracy_score,
    precision_recall_fscore_support,
    confusion_matrix,
    log_loss
)

from src.utils import clean_text
from src.config import LABELS_MAP, RANDOM_STATE, MAX_FEATURES, NGRAM_RANGE

class ExperimentalEvaluator:
    """
    Executes a leakage-free stratified experimental evaluation pipeline
    and performs automated error analysis on misclassified instances.
    """
    def __init__(self, n_splits: int = 5, random_state: int = RANDOM_STATE):
        self.n_splits = n_splits
        self.random_state = random_state
        self.cv = StratifiedKFold(n_splits=n_splits, shuffle=True, random_state=random_state)
        
    def create_pipeline(self) -> Pipeline:
        """Constructs an isolated vectorization & classification pipeline."""
        return Pipeline([
            ('tfidf', TfidfVectorizer(
                ngram_range=NGRAM_RANGE,
                sublinear_tf=True,
                max_features=MAX_FEATURES,
                min_df=1,
                stop_words='english'
            )),
            ('classifier', CalibratedClassifierCV(
                estimator=LinearSVC(C=1.0, class_weight='balanced', random_state=self.random_state),
                method='sigmoid',
                cv=2
            ))
        ])

    def run_stratified_evaluation(self, texts: List[str], labels: List[int]) -> Dict[str, Any]:
        """
        Runs Stratified K-Fold Cross-Validation strictly ensuring that
        TF-IDF vectorizer parameters are fit ONLY on fold training splits.
        """
        cleaned_texts = np.array([clean_text(t) for t in texts])
        y = np.array(labels)
        
        fold_results = []
        all_y_true = []
        all_y_pred = []
        all_y_probs = []
        misclassifications = []

        for fold_idx, (train_idx, val_idx) in enumerate(self.cv.split(cleaned_texts, y)):
            X_tr, y_tr = cleaned_texts[train_idx], y[train_idx]
            X_val, y_val = cleaned_texts[val_idx], y[val_idx]
            orig_val_texts = [texts[i] for i in val_idx]

            # Initialize fresh pipeline for fold isolation (prevents leakage)
            pipe = self.create_pipeline()
            pipe.fit(X_tr, y_tr)

            # Predict probabilities and classes
            val_probs = pipe.predict_proba(X_val)
            val_preds = np.argmax(val_probs, axis=1)

            # Record predictions
            all_y_true.extend(y_val)
            all_y_pred.extend(val_preds)
            all_y_probs.extend(val_probs)

            # Calculate fold metrics
            prec, rec, f1, _ = precision_recall_fscore_support(y_val, val_preds, labels=[0, 1, 2], zero_division=0)
            acc = accuracy_score(y_val, val_preds)
            bal_acc = balanced_accuracy_score(y_val, val_preds)
            macro_f1 = np.mean(f1)

            fold_results.append({
                "fold": fold_idx + 1,
                "accuracy": acc,
                "balanced_accuracy": bal_acc,
                "macro_f1": macro_f1,
                "negative_recall": rec[0],  # Critical churn metric
                "positive_recall": rec[2],
                "neutral_recall": rec[1]
            })

            # Harvest errors for error analysis
            for i in range(len(y_val)):
                if val_preds[i] != y_val[i]:
                    true_lbl = LABELS_MAP[y_val[i]]
                    pred_lbl = LABELS_MAP[val_preds[i]]
                    conf = float(val_probs[i][val_preds[i]])
                    misclassifications.append({
                        "fold": fold_idx + 1,
                        "text": orig_val_texts[i],
                        "true_sentiment": true_lbl,
                        "predicted_sentiment": pred_lbl,
                        "confidence": round(conf, 4),
                        "failure_category": self._categorize_error(orig_val_texts[i], true_lbl, pred_lbl)
                    })

        # Aggregated metrics across all folds
        macro_f1s = [f["macro_f1"] for f in fold_results]
        bal_accs = [f["balanced_accuracy"] for f in fold_results]
        neg_recs = [f["negative_recall"] for f in fold_results]
        accs = [f["accuracy"] for f in fold_results]

        summary_metrics = {
            "mean_macro_f1": round(float(np.mean(macro_f1s)), 4),
            "std_macro_f1": round(float(np.std(macro_f1s)), 4),
            "mean_balanced_accuracy": round(float(np.mean(bal_accs)), 4),
            "mean_negative_recall": round(float(np.mean(neg_recs)), 4),
            "mean_accuracy": round(float(np.mean(accs)), 4)
        }

        # Confusion Matrix
        cm = confusion_matrix(all_y_true, all_y_pred, labels=[0, 1, 2])
        cm_norm = cm.astype('float') / cm.sum(axis=1)[:, np.newaxis]

        return {
            "summary_metrics": summary_metrics,
            "fold_results": fold_results,
            "confusion_matrix": cm.tolist(),
            "normalized_confusion_matrix": np.round(cm_norm, 4).tolist(),
            "misclassifications": misclassifications,
            "total_errors": len(misclassifications)
        }

    def _categorize_error(self, text: str, true_lbl: str, pred_lbl: str) -> str:
        """Classifies misclassifications into linguistic failure taxonomy."""
        t_lower = text.lower()
        if any(w in t_lower for w in ["not", "never", "no", "hardly", "barely", "cannot", "can't"]):
            return "Category C (Complex Negation Scope)"
        elif any(w in t_lower for w in ["great", "love", "wonderful", "best"]) and true_lbl == "Negative":
            return "Category A (Sarcasm / Irony)"
        elif any(w in t_lower for w in ["but", "however", "although", "except"]):
            return "Category B (Mixed / Concessive Polarity)"
        elif true_lbl == "Neutral" or pred_lbl == "Neutral":
            return "Category D (Subjective Neutral Ambiguity)"
        return "Category E (Other Lexical Mismatch)"
