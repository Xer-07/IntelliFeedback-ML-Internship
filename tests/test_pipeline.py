import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import unittest
from src.preprocessor import TextCleaner
from src.trainer import ModelTrainer
from src.inference_engine import InferenceEngine

class TestIntelliFeedbackPipeline(unittest.TestCase):
    
    def setUp(self):
        self.cleaner = TextCleaner()
        self.sample_train_texts = [
            "Outstanding device, works flawlessly!",
            "Terrible broken product, complete waste of money.",
            "Ordinary item, acceptable performance."
        ] * 6
        self.sample_labels = [2, 0, 1] * 6

    def test_text_cleaner_html_and_urls(self):
        raw = "<p>Great product! Visit https://example.com for details.</p>"
        cleaned = self.cleaner.transform([raw])[0]
        self.assertNotIn("<p>", cleaned)
        self.assertNotIn("https://", cleaned)
        self.assertIn("great product", cleaned)

    def test_text_cleaner_empty_input(self):
        cleaned = self.cleaner.transform(["", None, "   "])
        self.assertEqual(len(cleaned), 3)
        self.assertEqual(cleaned[0], "")
        self.assertEqual(cleaned[1], "")
        self.assertEqual(cleaned[2], "")

    def test_model_training_and_inference(self):
        trainer = ModelTrainer(max_features=100)
        metrics = trainer.train_and_evaluate(self.sample_train_texts, self.sample_labels, n_splits=2)
        self.assertIn("mean_macro_f1", metrics)
        
        # Test serialization
        artifact_path = "models/test_model.joblib"
        trainer.save_artifact(artifact_path)
        
        # Test inference engine loading & prediction
        engine = InferenceEngine(artifact_path=artifact_path)
        pred = engine.predict_single("Super happy with this item, fantastic quality!")
        self.assertIn("sentiment", pred)
        self.assertIn("confidence", pred)
        self.assertIn("urgent_churn_alert", pred)

    def test_inference_empty_string_handling(self):
        trainer = ModelTrainer(max_features=100)
        trainer.train_and_evaluate(self.sample_train_texts, self.sample_labels, n_splits=2)
        artifact_path = "models/test_model.joblib"
        trainer.save_artifact(artifact_path)

        engine = InferenceEngine(artifact_path=artifact_path)
        res = engine.predict_single("")
        self.assertEqual(res["sentiment"], "Neutral")
        self.assertEqual(res["routing_status"], "EDGE_CASE_EMPTY_INPUT")

if __name__ == "__main__":
    unittest.main()
