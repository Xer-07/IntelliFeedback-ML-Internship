"""
Week 4 Demonstration Script:
Executes the modular production pipeline: Ingestion -> Preprocessing -> Training -> Serialization -> Inference.
"""
import os
from src.trainer import ModelTrainer
from src.inference_engine import InferenceEngine

def main():
    print("=" * 75)
    print("  INTELLIFEEDBACK: WEEK 4 MODULAR CODE ARCHITECTURE DEMO")
    print("=" * 75)

    # 1. Representative review training corpus
    reviews = [
        # Positive
        "Exceptional product, super fast delivery and brilliant packaging!",
        "Best purchase ever, highly recommended and works like magic.",
        "Superb build quality, very happy with customer support response.",
        "Incredible gadget, sleek modern design and wonderful battery life.",
        "Delighted with this purchase! Top notch quality and five stars.",
        "Exceeded all my expectations. Absolutely fantastic purchase.",

        # Negative
        "Horrible experience, stopped working after two days of use.",
        "Worst item received, broken out of the box and refund denied.",
        "Defective and low quality plastic. Avoid buying this product at all costs.",
        "Terrible customer service, rude support and app crashes continuously.",
        "Completely useless junk, absolute waste of time and money.",
        "Misleading product description, totally dissatisfied and angry.",

        # Neutral
        "The item is fine, performs basic operations as described.",
        "Average product for the price. Ordinary packaging, works fine.",
        "Acceptable quality, neither impressed nor disappointed.",
        "Decent functionality, meets basic requirements but nothing exceptional.",
        "Standard device, fair performance for everyday use.",
        "Just an ordinary gadget. Plugs in fine, no strong opinions."
    ]
    # 0 = Negative, 1 = Neutral, 2 = Positive
    labels = [2]*6 + [0]*6 + [1]*6

    print(f"\n[1] Initializing ModelTrainer with decoupled TextCleaner and TF-IDF...")
    trainer = ModelTrainer(max_features=5000, random_state=42)

    print(f"[2] Running 3-Fold Stratified Cross-Validation on {len(reviews)} reviews...")
    metrics = trainer.train_and_evaluate(reviews, labels, n_splits=3)
    print(f"    [OK] Cross-Validation Macro F1: {metrics['mean_macro_f1']} +/- {metrics['std_macro_f1']}")
    print(f"    [OK] Cross-Validation Balanced Accuracy: {metrics['mean_balanced_accuracy']}")

    artifact_dir = "models"
    os.makedirs(artifact_dir, exist_ok=True)
    artifact_path = os.path.join(artifact_dir, "production_sentiment_pipeline.joblib")
    
    print(f"\n[3] Serializing production pipeline artifact...")
    trainer.save_artifact(artifact_path)

    print(f"\n[4] Initializing thread-safe InferenceEngine from serialized artifact...")
    engine = InferenceEngine(artifact_path=artifact_path, confidence_threshold=0.70)

    test_samples = [
        "This software completely transformed our customer response time, truly magnificent!",
        "Defective unit, arrived broken and support was completely unhelpful.",
        "Standard cable, does what it is supposed to do.",
        ""  # Empty input edge case test
    ]

    print(f"\n[5] Executing Production Inference & Escalation Logic:")
    print("-" * 75)
    for text in test_samples:
        display_text = text if text else "<EMPTY STRING>"
        res = engine.predict_single(text)
        print(f"Input Review : \"{display_text}\"")
        print(f"  -> Predicted Sentiment : {res['sentiment']} (Confidence: {res['confidence']*100:.1f}%)")
        print(f"  -> Routing Status       : {res['routing_status']}")
        print(f"  -> Probabilities        : {res['probabilities']}")
        print(f"  -> Urgent Churn Alert   : {res['urgent_churn_alert']}")
        print("-" * 75)

    print("\n[OK] Week 4 Modular Pipeline Demonstration Completed Successfully!")

if __name__ == "__main__":
    main()
