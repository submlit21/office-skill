"""
Template storage - Handles filesystem paths and directory operations.

This module provides storage functionality for template metadata
and content files, managing the hierarchical directory structure.
"""

import json
import os
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional

import yaml


class TemplateStorage:
    """Manages template filesystem storage and hierarchical structure."""

    def __init__(self, templates_root: Path):
        """Initialize storage with root directory."""
        self.templates_root = templates_root
        self.templates_root.mkdir(parents=True, exist_ok=True)

    def _parse_template_name(self, name: str) -> Dict[str, str]:
        """Parse template name in format: domain.type.purpose.variant.version."""
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
        """Validate template name format."""
        try:
            self._parse_template_name(name)
            return True
        except ValueError:
            return False

    def _get_template_path(self, name: str) -> Path:
        """Get filesystem path for template based on its hierarchical name."""
        if not self._validate_template_name(name):
            raise ValueError(f"Invalid template name format: {name}")

        # Create hierarchical directory structure
        parts = name.split(".")
        template_dir = self.templates_root

        for part in parts:
            template_dir = template_dir / part
            template_dir.mkdir(parents=True, exist_ok=True)

        return template_dir

    def exists(self, name: str) -> bool:
        """Check if template exists."""
        template_dir = self._get_template_path(name)
        metadata_path = template_dir / "metadata.json"
        return metadata_path.exists()

    def save_template(
        self,
        name: str,
        metadata: Dict[str, Any],
        markdown_content: str,
        source_path: Path,
    ) -> Dict[str, Any]:
        """Save template files to storage."""
        template_dir = self._get_template_path(name)
        template_dir.mkdir(parents=True, exist_ok=True)

        # Ensure metadata has all SchemeB format fields
        metadata = self._enhance_metadata(metadata)

        # Save metadata (always JSON - consistent format for all template types)
        metadata_path = template_dir / "metadata.json"
        with open(metadata_path, "w", encoding="utf-8") as f:
            json.dump(metadata, f, indent=2, ensure_ascii=False)

        # Save content based on document format
        # Word documents (docx/doc) use markdown for better readability
        # Excel and PowerPoint use YAML for structured content
        ext = metadata.get("format", "docx")
        if ext in ["docx", "doc"]:
            # Word documents use markdown
            content_path = template_dir / "template.md"
            with open(content_path, "w", encoding="utf-8") as f:
                f.write(markdown_content)
        else:
            # Excel and PowerPoint use YAML for structured content
            try:
                content_path = template_dir / "template.yaml"
                yaml_data = {
                    "markdown_content": markdown_content,
                    "source_format": ext,
                }
                with open(content_path, "w", encoding="utf-8") as f:
                    yaml.dump(yaml_data, f, default_flow_style=False, allow_unicode=True)
            except ImportError:
                # Fallback to markdown if yaml not available
                content_path = template_dir / "template.md"
                with open(content_path, "w", encoding="utf-8") as f:
                    f.write(markdown_content)

        # Copy original file (always saved)
        original_copy_path = template_dir / f"original{source_path.suffix}"
        shutil.copy2(source_path, original_copy_path)

        metadata["template_path"] = str(template_dir)
        return metadata

    def _enhance_metadata(self, metadata: Dict[str, Any]) -> Dict[str, Any]:
        """Enhance metadata with SchemeB format fields if missing."""
        # Determine schema version
        if "schema_version" not in metadata:
            metadata["schema_version"] = "1.0"

        # Ensure required fields exist
        if "modified" not in metadata:
            metadata["modified"] = metadata.get("created", "")

        if "variables" not in metadata:
            metadata["variables"] = []

        if "author" not in metadata:
            metadata["author"] = ""

        if "validation_rules" not in metadata:
            metadata["validation_rules"] = {}

        if "template_version" not in metadata:
            # Extract from components if available, otherwise from name
            components = metadata.get("components", {})
            if isinstance(components, dict) and "version" in components:
                metadata["template_version"] = components["version"]
            else:
                # Try to parse from name
                try:
                    name = metadata.get("name", "")
                    if name and "." in name:
                        parts = name.split(".")
                        if len(parts) >= 5:
                            metadata["template_version"] = parts[-1]
                        else:
                            metadata["template_version"] = "v1"
                    else:
                        metadata["template_version"] = "v1"
                except Exception:
                    metadata["template_version"] = "v1"

        return metadata

    def load_template(self, name: str) -> Dict[str, Any]:
        """Load template metadata and content from storage."""
        template_dir = self._get_template_path(name)

        if not template_dir.exists():
            raise FileNotFoundError(f"Template not found: {name}")

        # Load metadata
        metadata_path = template_dir / "metadata.json"
        if not metadata_path.exists():
            raise FileNotFoundError(f"Metadata not found for template: {name}")

        with open(metadata_path, "r", encoding="utf-8") as f:
            metadata = json.load(f)

        # Enhance metadata with SchemeB format if needed
        metadata = self._enhance_metadata(metadata)

        # Load content based on format
        format_type = metadata.get("format", "docx")
        if format_type in ["docx", "doc"]:
            content_path = template_dir / "template.md"
            if content_path.exists():
                with open(content_path, "r", encoding="utf-8") as f:
                    markdown_content = f.read()
                metadata["markdown_content"] = markdown_content
            else:
                metadata["markdown_content"] = None
        else:
            # Excel and PowerPoint use YAML
            yaml_path = template_dir / "template.yaml"
            if yaml_path.exists():
                try:
                    with open(yaml_path, "r", encoding="utf-8") as f:
                        yaml_data = yaml.safe_load(f)
                    metadata["markdown_content"] = yaml_data.get("markdown_content", "")
                    metadata["yaml_data"] = yaml_data
                except ImportError:
                    # Fallback: check markdown if yaml not available
                    md_path = template_dir / "template.md"
                    if md_path.exists():
                        with open(md_path, "r", encoding="utf-8") as f:
                            metadata["markdown_content"] = f.read()
                    else:
                        metadata["markdown_content"] = None
            else:
                # Fallback to markdown
                md_path = template_dir / "template.md"
                if md_path.exists():
                    with open(md_path, "r", encoding="utf-8") as f:
                        metadata["markdown_content"] = f.read()
                else:
                    metadata["markdown_content"] = None

        # Check for original file
        original_files = list(template_dir.glob("original.*"))
        if original_files:
            metadata["original_file"] = str(original_files[0])

        metadata["template_path"] = str(template_dir)
        return metadata

    def delete_template(self, name: str) -> bool:
        """Delete a template from storage."""
        template_dir = self._get_template_path(name)

        if not template_dir.exists():
            raise FileNotFoundError(f"Template not found: {name}")

        # Delete the template directory
        shutil.rmtree(template_dir)

        # Also clean up empty parent directories
        parent = template_dir.parent
        while parent != self.templates_root and not any(parent.iterdir()):
            parent.rmdir()
            parent = parent.parent

        return True

    def list_templates(
        self,
        domain: Optional[str] = None,
        type_filter: Optional[str] = None,
        purpose: Optional[str] = None,
    ) -> List[str]:
        """List available template names with optional filtering."""
        template_names = []

        # Walk through template directory structure
        for root, _dirs, files in os.walk(self.templates_root):
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

                template_names.append(template_name)

        return template_names

    def get_parsed_components(self, name: str) -> Dict[str, str]:
        """Get parsed template name components."""
        return self._parse_template_name(name)
