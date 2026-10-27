"""
IntelliFeedback - Model Training & Serialization Orchestrator
"""
from typing import Dict, List, Any
import os
import joblib
import numpy as np
from sklearn.pipeline import Pipeline
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.svm import LinearSVC
from sklearn.calibration import CalibratedClassifierCV
from sklearn.model_selection import StratifiedKFold, cross_validate

from src.preprocessor import TextCleaner
from src.config import RANDOM_STATE, MAX_FEATURES, NGRAM_RANGE

class ModelTrainer:
    """
    Orchestrates end-to-end pipeline assembly, cross-validated training,
    and model artifact serialization.
    """
    def __init__(self, max_features: int = MAX_FEATURES, random_state: int = RANDOM_STATE):
        self.max_features = max_features
        self.random_state = random_state
        self.pipeline = self.build_pipeline()
        self.is_trained = False

    def build_pipeline(self) -> Pipeline:
        """Assembles the Scikit-Learn Pipeline combining TextCleaner, TF-IDF, and Calibrated LinearSVC."""
        return Pipeline([
            ('cleaner', TextCleaner()),
            ('tfidf', TfidfVectorizer(
                ngram_range=NGRAM_RANGE,
                sublinear_tf=True,
                max_features=self.max_features,
                min_df=1,
                stop_words='english'
            )),
            ('classifier', CalibratedClassifierCV(
                estimator=LinearSVC(C=1.0, class_weight='balanced', random_state=self.random_state),
                method='sigmoid',
                cv=2
            ))
        ])

    def train_and_evaluate(self, X_train: List[str], y_train: List[int], n_splits: int = 5) -> Dict[str, float]:
        """Runs Stratified K-Fold CV, fits the final pipeline, and returns performance summary."""
        cv = StratifiedKFold(n_splits=n_splits, shuffle=True, random_state=self.random_state)
        scores = cross_validate(
            self.pipeline, X_train, y_train, cv=cv, scoring=['f1_macro', 'balanced_accuracy']
        )
        
        # Fit on entire training corpus
        self.pipeline.fit(X_train, y_train)
        self.is_trained = True

        return {
            "mean_macro_f1": round(float(np.mean(scores['test_f1_macro'])), 4),
            "std_macro_f1": round(float(np.std(scores['test_f1_macro'])), 4),
            "mean_balanced_accuracy": round(float(np.mean(scores['test_balanced_accuracy'])), 4)
        }

    def save_artifact(self, output_path: str) -> None:
        """Persists trained pipeline to disk via Joblib."""
        if not self.is_trained:
            raise RuntimeError("Cannot save an unfitted model pipeline.")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        joblib.dump(self.pipeline, output_path)
        print(f"[INFO] Production model artifact persisted at: {output_path}")
