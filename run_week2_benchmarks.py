"""
Week 2 Demonstration Script:
Benchmarks candidate algorithms and executes the preliminary model pipeline.
"""
from src.model_pipeline import IntelliFeedbackClassifier, evaluate_candidate_algorithms

def main():
    print("=" * 70)
    print("  INTELLIFEEDBACK: WEEK 2 PRELIMINARY MODEL BENCHMARK & DEMO")
    print("=" * 70)

    # Balanced, representative review dataset across Positive, Neutral, Negative
    sample_texts = [
        # Positive (Class 2)
        "Absolutely fantastic product! Exceeded my expectations, fast shipping.",
        "Best purchase I have made all year, works like a charm and great build quality.",
        "Outstanding customer support! They resolved my problem in ten minutes.",
        "I love this item so much! Will definitely buy again and recommend to everyone.",
        "Super high quality material, very pleased with the finish and packaging.",
        "Incredible performance and easy to use interface. Highly recommended!",
        "Flawless execution, five stars for sure. Worth every single penny.",
        "Great experience from ordering to delivery. Super reliable quality.",

        # Negative (Class 0)
        "Terrible experience, arrived damaged and customer service refused to refund.",
        "Worst product ever, broke on the first day of use. Complete waste of money.",
        "Extremely disappointed with the poor build quality. Do not buy this!",
        "Defective item received. Shipping was delayed by three weeks. Awful service.",
        "Horrible customer care, nobody answered my calls and the app crashed instantly.",
        "Cheap plastic feel and stopped working after two hours. Terrible item.",
        "I regret buying this so much. Faulty product and zero support from seller.",
        "Misleading advertisement! The device overheats and is completely useless.",

        # Neutral (Class 1)
        "The product is okay, does what it says but nothing extraordinary.",
        "Average quality for the price. Delivery took normal time, works fine.",
        "It is acceptable, neither great nor bad. Neutral opinion on this gadget.",
        "Decent functionality, however packaging was quite plain. Neutral impression.",
        "Fair product. Meets basic needs but there are better alternatives on the market.",
        "Standard item, neither impressed nor disappointed. Works as expected.",
        "Moderate performance. It fulfills its basic purpose, average quality.",
        "Usable device, nothing to complain about but nothing special either."
    ]
    # 0 = Negative, 1 = Neutral, 2 = Positive
    sample_labels = [2]*8 + [0]*8 + [1]*8

    print(f"\n[1] Evaluating candidate algorithms across {len(sample_texts)} balanced reviews...")
    benchmarks = evaluate_candidate_algorithms(sample_texts, sample_labels)
    
    print("\nCandidate Algorithm Comparison (3-Fold Stratified CV):")
    print("-" * 70)
    print(f"{'Algorithm':<30} | {'Mean Macro F1':<15} | {'Balanced Accuracy':<15}")
    print("-" * 70)
    for algo, metrics in benchmarks.items():
        print(f"{algo:<30} | {metrics['mean_f1_macro']:<15} | {metrics['mean_balanced_accuracy']:<15}")
    print("-" * 70)

    print("\n[2] Training Preliminary Production Classifier (Calibrated LinearSVC + Sublinear TF-IDF)...")
    clf = IntelliFeedbackClassifier(confidence_threshold=0.70)
    clf.fit(sample_texts, sample_labels)
    print("    [OK] Model successfully fitted and calibrated.")

    test_queries = [
        "This is an exceptional gadget, truly revolutionized my daily workflow!",
        "Completely useless junk, broke within an hour and total waste of cash.",
        "It works reasonably well, nothing special but it gets the job done.",
        "Not bad, but I would not say it is the greatest either."
    ]

    print("\n[3] Testing Inference Pipeline & Decision Routing:")
    print("-" * 70)
    for q in test_queries:
        res = clf.predict(q)
        print(f"Input Review : \"{q}\"")
        print(f"  -> Predicted Sentiment: {res['sentiment']} (Confidence: {res['confidence'] * 100:.1f}%)")
        print(f"  -> Class Probabilities: {res['probabilities']}")
        print(f"  -> Decision Routing   : {res['routing_status']}")
        print(f"  -> Urgent Churn Alert : {res['urgent_churn_alert']}")
        print("-" * 70)

    print("\n[OK] Week 2 Preliminary Model Pipeline Validation Complete!")

if __name__ == "__main__":
    main()
