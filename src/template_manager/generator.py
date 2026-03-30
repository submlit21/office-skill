"""
Template generation - Generate documents from templates.

This module provides functionality to generate new documents from
existing templates with optional variable substitution.

For Word documents, it leverages cli-anything-libreoffice to actually
update content placeholders in the document.
"""

import shutil
from pathlib import Path
from typing import Any, Dict, Optional, Union

import jinja2
from jinja2 import Template

from office_main.core.cli_wrapper import LibreOfficeCLI

from .storage import TemplateStorage


class TemplateGenerator:
    """Generates new documents from templates."""

    def __init__(self, storage: TemplateStorage, cli: LibreOfficeCLI):
        """Initialize generator with storage and CLI instance from manager."""
        self.storage = storage
        self.cli = cli

    def _render_with_jinja2(self, template_str: str, variables: Dict[str, str]) -> str:
        """Render a template string using Jinja2."""
        try:
            template = Template(template_str)
            return template.render(variables)
        except jinja2.exceptions.TemplateError as e:
            raise ValueError(f"Template rendering error: {e}") from e

    def _render_variables(self, variables: Dict[str, str]) -> Dict[str, str]:
        """Render variable values as Jinja2 templates (allowing variable interpolation)."""
        rendered = {}
        for key, value in variables.items():
            try:
                # Treat each value as a Jinja2 template with access to all variables
                template = Template(value)
                rendered[key] = template.render(variables)
            except jinja2.exceptions.TemplateError:
                # If rendering fails, use the original value
                rendered[key] = value
        return rendered

    def generate(
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
        template_data = self.storage.load_template(template_name)

        if "original_file" not in template_data:
            raise ValueError(f"Template {template_name} has no original file for generation")

        original_path = Path(template_data["original_file"])
        output_path = Path(output_path).absolute()

        # Copy template to output location
        shutil.copy2(original_path, output_path)

        variables_applied = []

        # Render variable values using Jinja2 (allows variable interpolation)
        rendered_variables = self._render_variables(variables) if variables else {}

        if (
            rendered_variables
            and len(rendered_variables) > 0
            and template_data.get("format") in ["docx", "doc"]
        ):
            # For Word documents, try to actually replace placeholders in document
            # We use libreoffice-cli to find and replace {{placeholders}}
            try:
                # Start session with the document - this creates the temporary
                # project json that references the document, with --project already set
                _ = self.cli.start_session(str(output_path))

                # We need to open the document in the current session
                args = ["document", "open", str(output_path)]
                self.cli._run_command(args)

                # Look for text content containing placeholders
                content = self.cli.writer("list")
                _replacements_done = False

                if content.get("items"):
                    for idx, item in enumerate(content["items"]):
                        if "text" in item:
                            text = item.get("text", "")
                            original_text = text
                            for key, value in rendered_variables.items():
                                placeholder = f"{{{{{key}}}}}"
                                if placeholder in text:
                                    text = text.replace(placeholder, value)
                                    if key not in variables_applied:
                                        variables_applied.append(key)
                            if text != original_text:
                                # Content changed, update it in document
                                self.cli.writer("set-text", index=idx, text=text)
                                _replacements_done = True

                # Save the modified document
                args = ["document", "save", str(output_path)]
                self.cli._run_command(args)
                self.cli.end_session()

                if variables_applied:
                    note = f"Document generated. {len(variables_applied)} variable(s) successfully substituted in original document."
                else:
                    note = "Document generated. No placeholders found matching provided variables."
            except Exception as e:
                # If replacement fails, fall back to original behavior
                note = f"Document copied from template. Failed to update document: {str(e)}. Variables only applied to cached markdown."
        elif variables and template_data.get("markdown_content"):
            # For other formats or if substitution fails, use Jinja2 templating on cached markdown
            markdown_content = template_data["markdown_content"]

            # First, detect which variables are used in the template (simple placeholder detection)
            for key in rendered_variables.keys():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in markdown_content:
                    variables_applied.append(key)

            # Render using Jinja2
            try:
                markdown_content = self._render_with_jinja2(markdown_content, rendered_variables)
                note = "Document copied from template. Variable substitution applied to cached markdown using Jinja2."
            except ValueError:
                # If Jinja2 rendering fails, fall back to simple replacement for backward compatibility
                for key, value in rendered_variables.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in markdown_content:
                        markdown_content = markdown_content.replace(placeholder, value)
                note = "Document copied from template. Variable substitution only applied to cached markdown (simple replacement)."
        else:
            note = "Document copied from template. No variables provided."

        return {
            "template": template_name,
            "output_path": str(output_path),
            "variables_applied": variables_applied,
            "note": note,
        }
