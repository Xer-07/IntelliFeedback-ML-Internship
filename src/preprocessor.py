"""
IntelliFeedback - Text Preprocessing Module
Implements Scikit-Learn compatible transformers for text sanitization.
"""
from typing import List, Union, Any
import re
from sklearn.base import BaseEstimator, TransformerMixin

class TextCleaner(BaseEstimator, TransformerMixin):
    """
    Stateless text sanitization transformer conforming to Scikit-Learn's
    BaseEstimator and TransformerMixin contract.
    """
    def __init__(self, strip_html: bool = True, strip_urls: bool = True, lower_case: bool = True):
        self.strip_html = strip_html
        self.strip_urls = strip_urls
        self.lower_case = lower_case

    def fit(self, X, y=None):
        """Fit is a no-op as the transformer is stateless."""
        return self

    def transform(self, X: Union[List[str], Any]) -> List[str]:
        """Sanitizes text strings while preserving negation tokens."""
        cleaned_docs = []
        for doc in X:
            if not isinstance(doc, str):
                cleaned_docs.append("")
                continue
            
            text = doc
            if self.strip_html:
                text = re.sub(r'<.*?>', ' ', text)
            if self.strip_urls:
                text = re.sub(r'https?://\S+|www\.\S+', ' ', text)
            
            text = re.sub(r'\s+', ' ', text).strip()
            if self.lower_case:
                text = text.lower()
                
            cleaned_docs.append(text)
        return cleaned_docs
