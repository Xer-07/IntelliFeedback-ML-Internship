"""
IntelliFeedback - Project Configuration
"""
from pathlib import Path

# Paths
BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
MODELS_DIR = BASE_DIR / "models"
DOCS_DIR = BASE_DIR / "docs"

# Model & Preprocessing Parameters
RANDOM_STATE = 42
TEST_SIZE = 0.15
VAL_SIZE = 0.15
MAX_FEATURES = 10000
NGRAM_RANGE = (1, 3)

# Target Classes
LABELS_MAP = {
    0: "Negative",
    1: "Neutral",
    2: "Positive"
}
