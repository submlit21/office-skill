"""
Template handler - High-level template management for office-skill.

This module provides a unified interface to template functionality,
delegating to specialized modules for storage, analysis, search, and generation.
"""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

# Import template manager components directly (avoid circular import through template_manager.__init__)
from template_manager.analyzer import TemplateAnalyzer
from template_manager.generator import TemplateGenerator
from template_manager.search import TemplateSearch
from template_manager.storage import TemplateStorage

from .cli_wrapper import LibreOfficeCLI


class TemplateManager:
    """High-level manager for document templates.

    This is the main entry point that coordinates all template functionality
    by delegating to specialized modular components:
    - TemplateStorage: Filesystem storage and hierarchical organization
    - TemplateAnalyzer: Document analysis and markdown conversion
    - TemplateSearch: Searching and filtering
    - TemplateGenerator: Document generation from templates
    """

    def __init__(self, templates_root: Optional[Union[str, Path]] = None):
        """
        Initialize template manager.

        Args:
            templates_root: Root directory for templates (default: project templates/)
        """
        if templates_root is None:
            # Try to find templates directory relative to project root
            # Current file: src/office_main/core/template_handler.py
            # → .parent 3 times gives project root (if in development layout)
            project_root = Path(__file__).parent.parent.parent.parent
            candidate = project_root / "templates"

            # Check if this looks like the development project root
            is_dev_layout = (project_root / "pyproject.toml").exists() and (
                candidate.exists() or (project_root / "src").exists()
            )

            if is_dev_layout:
                templates_root = candidate
            else:
                # Use user data directory for installed package
                templates_root = Path.home() / ".office_skill" / "templates"

            # Ensure directory exists
            templates_root.mkdir(parents=True, exist_ok=True)
        elif isinstance(templates_root, str):
            templates_root = Path(templates_root).absolute()

        # Initialize components
        self.cli = LibreOfficeCLI()
        self.storage = TemplateStorage(templates_root)
        self.analyzer = TemplateAnalyzer(self.cli)
        self.search = TemplateSearch(self.storage)
        self.generator = TemplateGenerator(self.storage, self.cli)

    def add_template(
        self,
        source_path: Union[str, Path],
        name: str,
        description: Optional[str] = None,
        tags: Optional[List[str]] = None,
        overwrite: bool = False,
    ) -> Dict[str, Any]:
        """
        Add a new template to the template library.

        Args:
            source_path: Path to source document
            name: Template name in format domain.type.purpose.variant.version
            description: Optional template description
            tags: Optional list of tags
            overwrite: Whether to overwrite existing template

        Returns:
            Dictionary with template metadata
        """
        source_path = Path(source_path).absolute()

        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Validate template name
        if not self.storage._validate_template_name(name):
            raise ValueError(
                f"Invalid template name format: {name}. "
                f"Expected: domain.type.purpose.variant.version"
            )

        # Check if template already exists
        if self.storage.exists(name) and not overwrite:
            raise FileExistsError(
                f"Template '{name}' already exists. Use overwrite=True to replace."
            )

        # Analyze document structure
        analysis = self.analyzer.analyze_document_structure(source_path)

        # Convert to markdown
        markdown_content = self.analyzer.convert_to_markdown(source_path)

        # Create metadata with enhanced SchemeB format
        created_time = datetime.now().isoformat()
        metadata = {
            "schema_version": "2.0",
            "name": name,
            "original_name": source_path.name,
            "original_path": str(source_path),
            "description": description or f"Template: {name}",
            "tags": tags or [],
            "created": created_time,
            "modified": created_time,
            "format": source_path.suffix.lower().lstrip("."),
            "analysis": analysis,
            "components": self.storage._parse_template_name(name),
            "variables": [],  # Placeholder for template variables
            "author": "",  # Optional author information
            "validation_rules": {},  # Validation rules for template usage
            "template_version": self.storage._parse_template_name(name)["version"],
        }

        # Save to storage
        return self.storage.save_template(name, metadata, markdown_content, source_path)

    def get_template(self, name: str) -> Dict[str, Any]:
        """
        Get template by name.

        Args:
            name: Template name

        Returns:
            Template metadata and content
        """
        return self.storage.load_template(name)

    def list_templates(
        self,
        domain: Optional[str] = None,
        type_filter: Optional[str] = None,
        purpose: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """
        List available templates with optional filtering.

        Args:
            domain: Filter by domain
            type_filter: Filter by type
            purpose: Filter by purpose

        Returns:
            List of template metadata
        """
        return self.search.list_filtered(domain, type_filter, purpose)

    def search_templates(self, query: str) -> List[Dict[str, Any]]:
        """
        Search templates by name, description, or tags.

        Args:
            query: Search query string

        Returns:
            List of matching templates
        """
        return self.search.search(query)

    def delete_template(self, name: str, force: bool = False) -> bool:
        """
        Delete a template.

        Args:
            name: Template name
            force: Force deletion without confirmation

        Returns:
            True if deleted, False otherwise
        """
        if not self.storage.exists(name):
            raise FileNotFoundError(f"Template not found: {name}")

        if not force:
            # For API usage, we just delete without confirmation
            pass

        return self.storage.delete_template(name)

    def generate_from_template(
        self,
        template_name: str,
        output_path: Union[str, Path],
        variables: Optional[Dict[str, str]] = None,
    ) -> Dict[str, Any]:
        """
        Generate a new document from a template.

        Args:
            template_name: Name of the template to use
            output_path: Path for the generated document
            variables: Optional variables to replace in template

        Returns:
            Dictionary with generation results
        """
        return self.generator.generate(template_name, output_path, variables)

    def convert_to_markdown(self, source_path: Union[str, Path]) -> str:
        """
        Convert office document to markdown format.

        Args:
            source_path: Path to source document (docx, xlsx, pptx)

        Returns:
            Markdown content as string
        """
        source_path = Path(source_path)
        return self.analyzer.convert_to_markdown(source_path)

    def analyze_document_structure(self, source_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Analyze document structure for template metadata.

        Args:
            source_path: Path to source document

        Returns:
            Dictionary with document analysis
        """
        source_path = Path(source_path)
        return self.analyzer.analyze_document_structure(source_path)
