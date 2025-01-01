"""
Colab2Word - A package for generating Word documents from Google Colab notebooks
"""

from typing import List, Dict, Any, Optional, Union, Literal
from .document_generator import DocumentGenerator, DocumentSection, DocumentTheme, VisualizationSettings

__version__ = "0.1.0"
__author__ = "NativoCloud"

__all__ = ['DocumentGenerator', 'DocumentSection', 'DocumentTheme', 'VisualizationSettings']