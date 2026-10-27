"""
Utility functions for text preprocessing, metrics calculation, and logging.
"""
import re
import string

def clean_text(text: str) -> str:
    """
    Performs standard text cleaning while preserving sentiment cues.
    - Lowers casing
    - Strips URLs and HTML tags
    - Cleans whitespace
    """
    if not isinstance(text, str):
        return ""
    
    # Remove HTML tags
    text = re.sub(r'<.*?>', ' ', text)
    # Remove URLs
    text = re.sub(r'https?://\S+|www\.\S+', ' ', text)
    # Normalize whitespaces
    text = re.sub(r'\s+', ' ', text).strip()
    
    return text.lower()

def format_percentage(val: float) -> str:
    """Format floating point numbers as clean percentages."""
    return f"{val * 100:.2f}%"
