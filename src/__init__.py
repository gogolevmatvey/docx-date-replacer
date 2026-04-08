"""
Docx Date Replacer - Замена дат в документах .docx.
"""

from .config import Config, DEFAULT_CONFIG
from .date_replacer import DateReplacer
from .docx_processor import DocxProcessor

__all__ = [
    'Config',
    'DEFAULT_CONFIG',
    'DateReplacer',
    'DocxProcessor',
]
