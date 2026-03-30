"""
Template manager package - Modular template management.

This package contains modular components for template management:
- storage: Filesystem storage and hierarchical directory management
- analyzer: Document structure analysis and markdown conversion
- search: Search and filtering functionality
- generator: Document generation from templates
"""

# Note: Importing modules here can cause circular imports.
# Instead, import directly from submodules:
# from template_manager.analyzer import TemplateAnalyzer
# from template_manager.generator import TemplateGenerator
# etc.

__all__ = [
    "TemplateStorage",
    "TemplateAnalyzer",
    "TemplateSearch",
    "TemplateGenerator",
]
