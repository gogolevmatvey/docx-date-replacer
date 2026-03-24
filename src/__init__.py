"""
Docx Date Replacer - Замена дат в документах .docx.
"""

from .config import *
from .date_replacer import DateReplacer
from .docx_processor import DocxProcessor

__all__ = [
    'DateReplacer',
    'DocxProcessor',
]
