"""
Abstract base class for document handlers.

Provides shared initialization, export, cleanup, and context manager logic
for all Office document format handlers (DOCX/XLSX/PPTX).

All format-specific methods remain in the concrete subclasses.
"""

import os
from abc import ABC, abstractmethod
from typing import Any

from .cli_wrapper import LibreOfficeCLI


class BaseDocumentHandler(ABC):
    """Abstract base class for document format handlers.

    Subclasses must:
    - Call ``super().__init__(project_path=project_path)``
    - Set their own path attribute (e.g. ``self.document_path``)
    - Call ``self._start_session(str(self.path_attr))``
    - Provide ``analyze_structure()`` and ``create_from_template()``
    """

    def __init__(self, project_path: str | None = None):
        """Initialize shared handler state.

        Args:
            project_path: Optional path to .lo-cli.json project file.
                         If None, a temporary session file will be created
                         when ``_start_session()`` is called.
        """
        self.project_path = project_path
        self._temp_session = False
        self.cli: LibreOfficeCLI = None  # type: ignore[assignment]

    def _start_session(self, document_path: str) -> None:
        """Start a temporary LibreOffice session if no project_path was provided.

        Args:
            document_path: Path to the document to open in the session.
        """
        if self.project_path is None:
            temp_cli = LibreOfficeCLI()
            self.project_path = temp_cli.start_session(document_path)
            # Do NOT end session here - keep file for CLI operations
            self._temp_session = True

        self.cli = LibreOfficeCLI(project_path=self.project_path)

    def export(self, output_path: str, format: str = "pdf") -> dict[str, Any]:
        """Export the document to another format.

        Args:
            output_path: Path for the exported file
            format: Export format (pdf, docx, odt, ods, odp, etc.)

        Returns:
            Command result
        """
        return self.cli.export("render", positional=[output_path], preset=format)

    def close(self):
        """Close the handler and clean up temporary session files."""
        if self._temp_session and self.project_path:
            try:
                if os.path.exists(self.project_path):
                    os.unlink(self.project_path)
            except Exception:
                pass  # Ignore cleanup errors
        self.cli.end_session()

    @abstractmethod
    def analyze_structure(self) -> dict[str, Any]:
        """Analyze document structure.

        Returns:
            Document analysis including format-specific metadata.
        """

    @abstractmethod
    def create_from_template(self, template_path: str | None = None) -> "BaseDocumentHandler":
        """Create a new document from template.

        Args:
            template_path: Path to template document (optional)

        Returns:
            New handler instance

        Raises:
            NotImplementedError: If not implemented by the subclass
        """

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
