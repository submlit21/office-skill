"""
Office Skill - Core Python package for office document processing.

This package provides a Python interface to cli-anything-libreoffice,
enabling creation, editing, modification, analysis, and annotation
of Word, Excel, and PowerPoint documents.
"""

__version__ = "0.1.0"
__author__ = "Office Skill Team"

from .cli_wrapper import LibreOfficeCLI
from .docx_handler import DocxHandler
from .pptx_handler import PptxHandler
from .xlsx_handler import XlsxHandler

__all__ = [
    "LibreOfficeCLI",
    "DocxHandler",
    "XlsxHandler",
    "PptxHandler",
    # Note: TemplateManager not imported here to avoid circular imports
    # Import directly: from office_main.core.template_handler import TemplateManager
]
