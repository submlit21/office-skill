"""
Template handler for office-skill.

This module provides functionality to manage document templates,
including conversion to markdown format and hierarchical storage.
"""

import os
import re
import json
import tempfile
import shutil
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple, Union
from datetime import datetime

from .cli_wrapper import LibreOfficeCLI


class TemplateManager:
    """Manager for document templates with hierarchical storage."""

    def __init__(self, templates_root: Optional[str] = None):
        """
        Initialize template manager.

        Args:
            templates_root: Root directory for templates (default: project templates/)
        """
        if templates_root:
            self.templates_root = Path(templates_root).absolute()
        else:
            # Default to project templates directory
            project_root = Path(__file__).parent.parent.parent
            self.templates_root = project_root / "templates"

        # Ensure templates directory exists
        self.templates_root.mkdir(parents=True, exist_ok=True)

        self.cli = LibreOfficeCLI()

    def _parse_template_name(self, name: str) -> Dict[str, str]:
        """
        Parse template name in format: domain.type.purpose.variant.version

        Args:
            name: Template name string

        Returns:
            Dictionary with parsed components
        """
        parts = name.split(".")
        if len(parts) != 5:
            raise ValueError(
                f"Template name must be in format 'domain.type.purpose.variant.version', "
                f"got '{name}' with {len(parts)} parts"
            )

        return {
            "domain": parts[0],
            "type": parts[1],
            "purpose": parts[2],
            "variant": parts[3],
            "version": parts[4],
        }

    def _validate_template_name(self, name: str) -> bool:
        """
        Validate template name format.

        Args:
            name: Template name to validate

        Returns:
            True if valid, False otherwise
        """
        try:
            self._parse_template_name(name)
            return True
        except ValueError:
            return False

    def _get_template_path(self, name: str) -> Path:
        """
        Get filesystem path for a template based on its hierarchical name.

        Args:
            name: Template name in format domain.type.purpose.variant.version

        Returns:
            Path object for the template directory
        """
        if not self._validate_template_name(name):
            raise ValueError(f"Invalid template name format: {name}")

        # Create hierarchical directory structure
        parts = name.split(".")
        template_dir = self.templates_root

        for part in parts:
            template_dir = template_dir / part
            template_dir.mkdir(parents=True, exist_ok=True)

        return template_dir

    def convert_to_markdown(self, source_path: Union[str, Path]) -> str:
        """
        Convert office document to markdown format.

        Args:
            source_path: Path to source document (docx, xlsx, pptx)

        Returns:
            Markdown content as string
        """
        source_path = Path(source_path).absolute()

        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Check file extension
        ext = source_path.suffix.lower()
        if ext not in [".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"]:
            raise ValueError(f"Unsupported file format: {ext}")

        # Create temporary output file
        with tempfile.NamedTemporaryFile(suffix=".md", delete=False) as tmp:
            output_path = tmp.name

        try:
            # Use LibreOffice to convert to markdown
            # Note: This assumes cli-anything-libreoffice supports markdown export
            result = self.cli._run_command(
                ["export", "--input", str(source_path), "--output", output_path, "--format", "md"]
            )

            # Read the converted markdown
            with open(output_path, "r", encoding="utf-8") as f:
                markdown_content = f.read()

            return markdown_content

        except Exception as e:
            raise RuntimeError(f"Failed to convert {source_path} to markdown: {e}")
        finally:
            # Clean up temporary file
            if os.path.exists(output_path):
                os.unlink(output_path)

    def analyze_document_structure(self, source_path: Union[str, Path]) -> Dict[str, Any]:
        """
        Analyze document structure for template metadata.

        Args:
            source_path: Path to source document

        Returns:
            Dictionary with document analysis
        """
        source_path = Path(source_path).absolute()

        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        ext = source_path.suffix.lower()

        # Analyze based on document type
        if ext in [".docx", ".doc"]:
            return self._analyze_word_document(source_path)
        elif ext in [".xlsx", ".xls"]:
            return self._analyze_excel_document(source_path)
        elif ext in [".pptx", ".ppt"]:
            return self._analyze_powerpoint_document(source_path)
        else:
            raise ValueError(f"Unsupported file format: {ext}")

    def _analyze_word_document(self, path: Path) -> Dict[str, Any]:
        """Analyze Word document structure."""
        try:
            result = self.cli.writer("analyze", document=str(path))
            return {
                "type": "word",
                "pages": result.get("page_count", 0),
                "paragraphs": result.get("paragraph_count", 0),
                "tables": result.get("table_count", 0),
                "images": result.get("image_count", 0),
                "sections": result.get("section_count", 0),
                "styles": result.get("styles", []),
                "metadata": result.get("metadata", {}),
            }
        except Exception as e:
            return {
                "type": "word",
                "error": str(e),
                "pages": 0,
                "paragraphs": 0,
                "tables": 0,
                "images": 0,
            }

    def _analyze_excel_document(self, path: Path) -> Dict[str, Any]:
        """Analyze Excel document structure."""
        try:
            result = self.cli.calc("analyze", spreadsheet=str(path))
            return {
                "type": "excel",
                "sheets": result.get("sheet_count", 0),
                "cells": result.get("cell_count", 0),
                "formulas": result.get("formula_count", 0),
                "charts": result.get("chart_count", 0),
                "named_ranges": result.get("named_range_count", 0),
                "metadata": result.get("metadata", {}),
            }
        except Exception as e:
            return {"type": "excel", "error": str(e), "sheets": 0, "cells": 0, "formulas": 0}

    def _analyze_powerpoint_document(self, path: Path) -> Dict[str, Any]:
        """Analyze PowerPoint document structure."""
        try:
            result = self.cli.impress("analyze", presentation=str(path))
            return {
                "type": "powerpoint",
                "slides": result.get("slide_count", 0),
                "masters": result.get("master_count", 0),
                "layouts": result.get("layout_count", 0),
                "shapes": result.get("shape_count", 0),
                "images": result.get("image_count", 0),
                "notes": result.get("has_notes", False),
                "metadata": result.get("metadata", {}),
            }
        except Exception as e:
            return {"type": "powerpoint", "error": str(e), "slides": 0, "masters": 0, "layouts": 0}

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
        if not self._validate_template_name(name):
            raise ValueError(
                f"Invalid template name format: {name}. "
                f"Expected: domain.type.purpose.variant.version"
            )

        # Get template directory path
        template_dir = self._get_template_path(name)

        # Check if template already exists
        template_files = list(template_dir.glob("*"))
        if template_files and not overwrite:
            raise FileExistsError(
                f"Template '{name}' already exists at {template_dir}. "
                f"Use overwrite=True to replace."
            )

        # Analyze document structure
        analysis = self.analyze_document_structure(source_path)

        # Convert to markdown
        markdown_content = self.convert_to_markdown(source_path)

        # Create metadata
        metadata = {
            "name": name,
            "original_name": source_path.name,
            "original_path": str(source_path),
            "description": description or f"Template: {name}",
            "tags": tags or [],
            "created": datetime.now().isoformat(),
            "format": source_path.suffix.lower().lstrip("."),
            "analysis": analysis,
            "components": self._parse_template_name(name),
        }

        # Save template files
        template_dir.mkdir(parents=True, exist_ok=True)

        # Save markdown content
        markdown_path = template_dir / "template.md"
        with open(markdown_path, "w", encoding="utf-8") as f:
            f.write(markdown_content)

        # Save metadata
        metadata_path = template_dir / "metadata.json"
        with open(metadata_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)

        # Copy original file (optional)
        original_copy_path = template_dir / f"original{source_path.suffix}"
        shutil.copy2(source_path, original_copy_path)

        return metadata

    def get_template(self, name: str) -> Dict[str, Any]:
        """
        Get template by name.

        Args:
            name: Template name

        Returns:
            Template metadata and content
        """
        template_dir = self._get_template_path(name)

        if not template_dir.exists():
            raise FileNotFoundError(f"Template not found: {name}")

        # Load metadata
        metadata_path = template_dir / "metadata.json"
        if not metadata_path.exists():
            raise FileNotFoundError(f"Metadata not found for template: {name}")

        with open(metadata_path, "r", encoding="utf-8") as f:
            metadata = json.load(f)

        # Load markdown content
        markdown_path = template_dir / "template.md"
        if markdown_path.exists():
            with open(markdown_path, "r", encoding="utf-8") as f:
                markdown_content = f.read()
            metadata["markdown_content"] = markdown_content
        else:
            metadata["markdown_content"] = None

        # Check for original file
        original_files = list(template_dir.glob("original.*"))
        if original_files:
            metadata["original_file"] = str(original_files[0])

        metadata["template_path"] = str(template_dir)

        return metadata

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
        templates = []

        # Walk through template directory structure
        for root, dirs, files in os.walk(self.templates_root):
            # Check if this is a leaf directory (contains metadata.json)
            if "metadata.json" in files:
                # Extract template name from path
                rel_path = Path(root).relative_to(self.templates_root)
                template_name = ".".join(rel_path.parts)

                # Apply filters
                if domain and not template_name.startswith(f"{domain}."):
                    continue
                if type_filter and not template_name.startswith(f"{domain or '*'}.{type_filter}."):
                    # Need to parse to check type
                    parts = template_name.split(".")
                    if len(parts) >= 2 and parts[1] != type_filter:
                        continue
                if purpose and not template_name.startswith(
                    f"{domain or '*'}.{type_filter or '*'}."
                ):
                    # Need to parse to check purpose
                    parts = template_name.split(".")
                    if len(parts) >= 3 and parts[2] != purpose:
                        continue

                try:
                    template_data = self.get_template(template_name)
                    templates.append(template_data)
                except Exception as e:
                    # Skip templates with errors
                    print(f"Warning: Failed to load template {template_name}: {e}")

        return templates

    def search_templates(self, query: str) -> List[Dict[str, Any]]:
        """
        Search templates by name, description, or tags.

        Args:
            query: Search query string

        Returns:
            List of matching templates
        """
        all_templates = self.list_templates()
        query_lower = query.lower()

        results = []

        for template in all_templates:
            # Search in name
            if query_lower in template["name"].lower():
                results.append(template)
                continue

            # Search in description
            if template.get("description") and query_lower in template["description"].lower():
                results.append(template)
                continue

            # Search in tags
            if any(query_lower in tag.lower() for tag in template.get("tags", [])):
                results.append(template)
                continue

        return results

    def delete_template(self, name: str, force: bool = False) -> bool:
        """
        Delete a template.

        Args:
            name: Template name
            force: Force deletion without confirmation

        Returns:
            True if deleted, False otherwise
        """
        template_dir = self._get_template_path(name)

        if not template_dir.exists():
            raise FileNotFoundError(f"Template not found: {name}")

        if not force:
            # In a real implementation, you might ask for confirmation
            # For now, we'll just delete
            pass

        # Delete the template directory
        shutil.rmtree(template_dir)

        # Also clean up empty parent directories
        parent = template_dir.parent
        while parent != self.templates_root and not any(parent.iterdir()):
            parent.rmdir()
            parent = parent.parent

        return True

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
        template_data = self.get_template(template_name)

        if "original_file" not in template_data:
            raise ValueError(f"Template {template_name} has no original file for generation")

        original_path = Path(template_data["original_file"])
        output_path = Path(output_path).absolute()

        # Copy template to output location
        shutil.copy2(original_path, output_path)

        # Apply variable substitutions if provided
        if variables and template_data.get("markdown_content"):
            # This is a simplified version - in reality you'd need to
            # parse the document and replace placeholders
            markdown_content = template_data["markdown_content"]

            for key, value in variables.items():
                placeholder = f"{{{{{key}}}}}"
                markdown_content = markdown_content.replace(placeholder, value)

            # Note: Actually updating the document content would require
            # more complex processing with LibreOffice

            return {
                "template": template_name,
                "output_path": str(output_path),
                "variables_applied": list(variables.keys()),
                "note": "Document copied from template. Variable substitution requires manual processing.",
            }

        return {"template": template_name, "output_path": str(output_path), "variables_applied": []}
