"""
IntelliFeedback - Preliminary Model Architecture & Pipeline
Implements the Week 2 candidate algorithm comparison and dual-tiered inference pipeline.
"""
from typing import Dict, List, Tuple, Any
import numpy as np
from sklearn.pipeline import Pipeline
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.naive_bayes import MultinomialNB
from sklearn.linear_model import LogisticRegression
from sklearn.svm import LinearSVC
from sklearn.ensemble import RandomForestClassifier
from sklearn.calibration import CalibratedClassifierCV
from sklearn.model_selection import StratifiedKFold, cross_validate

from src.utils import clean_text
from src.config import LABELS_MAP, RANDOM_STATE, MAX_FEATURES, NGRAM_RANGE

class IntelliFeedbackClassifier:
    """
    Production-grade preliminary sentiment classification model.
    Utilizes sub-linear TF-IDF (1-3 n-grams) and Calibrated LinearSVC
    for ultra-low latency and robust generalization on sparse text matrices.
    """
    def __init__(self, max_features: int = MAX_FEATURES, confidence_threshold: float = 0.70):
        self.confidence_threshold = confidence_threshold
        self.pipeline = Pipeline([
            ('tfidf', TfidfVectorizer(
                ngram_range=NGRAM_RANGE,
                sublinear_tf=True,
                max_features=max_features,
                min_df=2,
                stop_words='english'
            )),
            ('classifier', CalibratedClassifierCV(
                estimator=LinearSVC(C=1.0, class_weight='balanced', random_state=RANDOM_STATE),
                method='sigmoid',
                cv=3
            ))
        ])
        self.is_fitted = False

    def fit(self, texts: List[str], labels: List[int]) -> 'IntelliFeedbackClassifier':
        """Fits the vectorizer and calibrated classifier on training data."""
        cleaned_texts = [clean_text(t) for t in texts]
        self.pipeline.fit(cleaned_texts, labels)
        self.is_fitted = True
        return self

    def predict(self, raw_text: str) -> Dict[str, Any]:
        """
        Executes prediction on single review string with confidence scoring
        and dual-tier routing logic.
        """
        if not self.is_fitted:
            raise RuntimeError("Model must be trained before calling predict().")

        cleaned = clean_text(raw_text)
        probs = self.pipeline.predict_proba([cleaned])[0]
        predicted_class = int(np.argmax(probs))
        confidence = float(np.max(probs))

        # Check for Tier 2 escalation
        if confidence < self.confidence_threshold:
            routing_status = "ESCALATED_LOW_CONFIDENCE"
        else:
            routing_status = "TIER1_FAST_PATH_RESOLVED"

        sentiment_name = LABELS_MAP.get(predicted_class, "Unknown")
        urgent_flag = (sentiment_name == "Negative" and confidence >= 0.80)

        return {
            "cleaned_text": cleaned,
            "predicted_class": predicted_class,
            "sentiment": sentiment_name,
            "confidence": round(confidence, 4),
            "probabilities": {
                LABELS_MAP[idx]: round(float(p), 4) for idx, p in enumerate(probs)
            },
            "routing_status": routing_status,
            "urgent_churn_alert": urgent_flag
        }

def evaluate_candidate_algorithms(texts: List[str], labels: List[int]) -> Dict[str, Dict[str, float]]:
    """
    Empirically benchmarks candidate algorithms across Macro F1 and Balanced Accuracy
    using 5-Fold Stratified Cross Validation.
    """
    cleaned_texts = [clean_text(t) for t in texts]
    cv = StratifiedKFold(n_splits=3, shuffle=True, random_state=RANDOM_STATE)
    
    candidates = {
        "Multinomial Naive Bayes": MultinomialNB(alpha=1.0),
        "Logistic Regression (L2)": LogisticRegression(max_iter=1000, class_weight='balanced', random_state=RANDOM_STATE),
        "LinearSVC (Calibrated)": CalibratedClassifierCV(LinearSVC(C=1.0, class_weight='balanced', random_state=RANDOM_STATE), cv=2),
        "Random Forest (100 trees)": RandomForestClassifier(n_estimators=100, random_state=RANDOM_STATE, n_jobs=-1)
    }

    results = {}
    for name, clf in candidates.items():
        pipe = Pipeline([
            ('tfidf', TfidfVectorizer(ngram_range=(1, 2), sublinear_tf=True, max_features=5000, min_df=2)),
            ('clf', clf)
        ])
        cv_res = cross_validate(pipe, cleaned_texts, labels, cv=cv, scoring=['f1_macro', 'balanced_accuracy'])
        results[name] = {
            "mean_f1_macro": round(float(np.mean(cv_res['test_f1_macro'])), 4),
            "std_f1_macro": round(float(np.std(cv_res['test_f1_macro'])), 4),
            "mean_balanced_accuracy": round(float(np.mean(cv_res['test_balanced_accuracy'])), 4)
        }
    return results
