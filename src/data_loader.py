"""
IntelliFeedback - Data Ingestion & Splitting Module
"""
from typing import Tuple, List, Optional
import os
import pandas as pd
from sklearn.model_selection import train_test_split

from src.config import RANDOM_STATE, TEST_SIZE

class DataLoader:
    """Handles raw review dataset ingestion, schema validation, and partitioning."""
    
    REQUIRED_COLUMNS = ["review_text", "sentiment"]

    @staticmethod
    def validate_schema(df: pd.DataFrame) -> bool:
        """Asserts required columns exist and removes invalid null rows."""
        for col in DataLoader.REQUIRED_COLUMNS:
            if col not in df.columns:
                raise ValueError(f"Missing mandatory column: '{col}' in dataset.")
        return True

    @staticmethod
    def load_dataset(filepath: str) -> pd.DataFrame:
        """Loads CSV/JSON file, validates schema, and returns sanitized DataFrame."""
        if not os.path.exists(filepath):
            raise FileNotFoundError(f"Dataset not found at: {filepath}")
        
        df = pd.read_csv(filepath)
        DataLoader.validate_schema(df)
        df = df.dropna(subset=DataLoader.REQUIRED_COLUMNS).copy()
        df["sentiment"] = df["sentiment"].astype(int)
        return df

    @staticmethod
    def split_data(
        df: pd.DataFrame, test_size: float = TEST_SIZE, random_state: int = RANDOM_STATE
    ) -> Tuple[List[str], List[str], List[int], List[int]]:
        """Performs stratified train/test partitioning."""
        X = df["review_text"].tolist()
        y = df["sentiment"].tolist()
        
        X_train, X_test, y_train, y_test = train_test_split(
            X, y, test_size=test_size, stratify=y, random_state=random_state
        )
        return X_train, X_test, y_train, y_test
