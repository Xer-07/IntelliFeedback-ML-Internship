"""
Week 3 Demonstration Script:
Executes the experimental evaluation protocol and automated error analysis.
"""
from src.experimental_framework import ExperimentalEvaluator
from src.config import LABELS_MAP

def main():
    print("=" * 75)
    print("  INTELLIFEEDBACK: WEEK 3 EXPERIMENTAL PROTOCOL & METRIC EVALUATION")
    print("=" * 75)

    # Expanded representative customer review corpus
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
        "Fast response time, seamless setup, and brilliant user experience.",
        "Remarkable craftsmanship and sturdy construction. Love it!",

        # Negative (Class 0)
        "Terrible experience, arrived damaged and customer service refused to refund.",
        "Worst product ever, broke on the first day of use. Complete waste of money.",
        "Extremely disappointed with the poor build quality. Do not buy this!",
        "Defective item received. Shipping was delayed by three weeks. Awful service.",
        "Horrible customer care, nobody answered my calls and the app crashed instantly.",
        "Cheap plastic feel and stopped working after two hours. Terrible item.",
        "I regret buying this so much. Faulty product and zero support from seller.",
        "Misleading advertisement! The device overheats and is completely useless.",
        "Oh wonderful, broke right out of the box. Super happy with this junk.", # Sarcasm
        "I can not say that I ever found a single feature that worked properly.", # Negation

        # Neutral (Class 1)
        "The product is okay, does what it says but nothing extraordinary.",
        "Average quality for the price. Delivery took normal time, works fine.",
        "It is acceptable, neither great nor bad. Neutral opinion on this gadget.",
        "Decent functionality, however packaging was quite plain. Neutral impression.",
        "Fair product. Meets basic needs but there are better alternatives on the market.",
        "Standard item, neither impressed nor disappointed. Works as expected.",
        "Moderate performance. It fulfills its basic purpose, average quality.",
        "Usable device, nothing to complain about but nothing special either.",
        "Ordinary build, reasonable price point, performs adequately.",
        "Just a generic replacement part. Plugs in fine, no strong feelings."
    ]
    # 0 = Negative, 1 = Neutral, 2 = Positive
    sample_labels = [2]*10 + [0]*10 + [1]*10

    print(f"\n[1] Running Leakage-Free Stratified 5-Fold Cross-Validation (N={len(sample_texts)})...")
    evaluator = ExperimentalEvaluator(n_splits=5, random_state=42)
    results = evaluator.run_stratified_evaluation(sample_texts, sample_labels)

    metrics = results["summary_metrics"]
    print("\n[2] Summary Evaluation Metrics across 5 Folds:")
    print("-" * 75)
    print(f"  * Macro F1-Score         : {metrics['mean_macro_f1']} +/- {metrics['std_macro_f1']}")
    print(f"  * Negative Class Recall  : {metrics['mean_negative_recall']} (Primary Churn Prevention KPI)")
    print(f"  * Balanced Accuracy      : {metrics['mean_balanced_accuracy']}")
    print(f"  * Overall Accuracy       : {metrics['mean_accuracy']}")
    print("-" * 75)

    print("\n[3] Normalized Confusion Matrix (Rows: True, Cols: Predicted):")
    print(f"       {'Predicted Neg':<15} {'Predicted Neu':<15} {'Predicted Pos':<15}")
    cm_norm = results["normalized_confusion_matrix"]
    labels = ["True Negative", "True Neutral ", "True Positive"]
    for i, row in enumerate(cm_norm):
        print(f"{labels[i]} : {row[0]:<15.4f} {row[1]:<15.4f} {row[2]:<15.4f}")

    print(f"\n[4] Qualitative Error Analysis Breakdown (Total Misclassifications: {results['total_errors']}):")
    print("-" * 75)
    if results["misclassifications"]:
        for idx, err in enumerate(results["misclassifications"][:5], 1):
            print(f"Error #{idx} [Fold {err['fold']}]:")
            print(f"  Review Text     : \"{err['text']}\"")
            print(f"  True vs Pred    : {err['true_sentiment']} --> Predicted: {err['predicted_sentiment']} (Conf: {err['confidence']*100:.1f}%)")
            print(f"  Failure Category: {err['failure_category']}")
            print("-" * 75)
    else:
        print("  No misclassifications recorded on validation folds.")

    print("\n[OK] Week 3 Experimental Protocol & Metric Evaluation Completed Successfully!")

if __name__ == "__main__":
    main()
