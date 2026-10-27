"""
IntelliFeedback - Production Inference Engine
Thread-safe singleton service for low-latency sentiment scoring and churn escalation.
"""
from typing import Dict, List, Any, Optional
import os
import joblib
import numpy as np

from src.config import LABELS_MAP

class InferenceEngine:
    """
    Stateless, production-ready inference service.
    Loads serialized model pipeline and executes single or batched predictions.
    """
    def __init__(self, artifact_path: Optional[str] = None, confidence_threshold: float = 0.70):
        self.confidence_threshold = confidence_threshold
        self.model = None
        if artifact_path:
            self.load_model(artifact_path)

    def load_model(self, artifact_path: str) -> None:
        """Loads serialized pipeline artifact from disk."""
        if not os.path.exists(artifact_path):
            raise FileNotFoundError(f"Model artifact not found at: {artifact_path}")
        self.model = joblib.load(artifact_path)
        print(f"[INFO] InferenceEngine loaded model artifact from: {artifact_path}")

    def predict_single(self, text: str) -> Dict[str, Any]:
        """
        Executes sentiment prediction on a single review string.
        Returns sentiment class, probabilities, confidence, routing status, and churn alerts.
        """
        if self.model is None:
            raise RuntimeError("InferenceEngine requires a loaded model before predicting.")

        if not isinstance(text, str) or not text.strip():
            # Graceful edge case fallback
            return {
                "sentiment": "Neutral",
                "confidence": 0.3333,
                "probabilities": {"Negative": 0.3333, "Neutral": 0.3333, "Positive": 0.3333},
                "routing_status": "EDGE_CASE_EMPTY_INPUT",
                "urgent_churn_alert": False
            }

        probs = self.model.predict_proba([text])[0]
        pred_idx = int(np.argmax(probs))
        confidence = float(np.max(probs))

        # Check for Tier 2 escalation
        if confidence < self.confidence_threshold:
            routing = "TIER2_DISTILBERT_FALLBACK"
        else:
            routing = "TIER1_FAST_PATH_RESOLVED"

        sentiment_name = LABELS_MAP.get(pred_idx, "Unknown")
        churn_alert = (sentiment_name == "Negative" and confidence >= 0.80)

        return {
            "sentiment": sentiment_name,
            "confidence": round(confidence, 4),
            "probabilities": {
                LABELS_MAP[i]: round(float(p), 4) for i, p in enumerate(probs)
            },
            "routing_status": routing,
            "urgent_churn_alert": churn_alert
        }

    def predict_batch(self, texts: List[str]) -> List[Dict[str, Any]]:
        """Processes a list of review texts efficiently."""
        return [self.predict_single(t) for t in texts]
