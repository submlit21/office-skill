"""
Template search - Search and filter templates in storage.

This module provides search functionality across the template library.
"""

from typing import Any, Dict, List, Optional

from .storage import TemplateStorage


class TemplateSearch:
    """Provides search and filtering capabilities for templates."""

    def __init__(self, storage: TemplateStorage):
        """Initialize search with storage instance."""
        self.storage = storage

    def search(self, query: str) -> List[Dict[str, Any]]:
        """Search templates by name, description, or tags."""
        all_templates = self.list_all()
        query_lower = query.lower()

        results = []

        for template_name in all_templates:
            try:
                template_data = self.storage.load_template(template_name)
                # Search in name
                if query_lower in template_data["name"].lower():
                    results.append(template_data)
                    continue

                # Search in description
                if (
                    template_data.get("description")
                    and query_lower in template_data["description"].lower()
                ):
                    results.append(template_data)
                    continue

                # Search in tags
                if any(query_lower in tag.lower() for tag in template_data.get("tags", [])):
                    results.append(template_data)
                    continue

            except Exception:
                # Skip templates with errors
                pass

        return results

    def list_filtered(
        self,
        domain: Optional[str] = None,
        type_filter: Optional[str] = None,
        purpose: Optional[str] = None,
    ) -> List[Dict[str, Any]]:
        """List templates with filtering."""
        template_names = self.storage.list_templates(domain, type_filter, purpose)
        results = []

        for template_name in template_names:
            try:
                template_data = self.storage.load_template(template_name)
                results.append(template_data)
            except Exception as e:
                # Skip templates with errors
                print(f"Warning: Failed to load template {template_name}: {e}")

        return results

    def list_all(self) -> List[str]:
        """List all template names without filtering."""
        return self.storage.list_templates()
