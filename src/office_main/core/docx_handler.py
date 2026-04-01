"""
DOCX (Word) document handler.

Provides high-level operations for Word document manipulation.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional

from .cli_wrapper import LibreOfficeCLI


class DocxHandler:
    """Handler for Word document operations."""

    def __init__(self, document_path: str, project_path: Optional[str] = None):
        """
        Initialize handler for a Word document.

        Args:
            document_path: Path to the .docx document
            project_path: Optional path to .lo-cli.json project file.
                         If None, a temporary session file will be created.
        """
        self.document_path = Path(document_path).absolute()
        self.project_path = project_path
        self._temp_session = False

        if self.project_path is None:
            # Create a temporary session file
            temp_cli = LibreOfficeCLI()
            self.project_path = temp_cli.start_session(str(self.document_path))
            # Do NOT end session here - keep file for CLI operations
            self._temp_session = True

        self.cli = LibreOfficeCLI(project_path=self.project_path)

    def add_paragraph(self, text: str, index: Optional[int] = None) -> Dict[str, Any]:
        """
        Add a paragraph to the document.

        Args:
            text: The text content
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        kwargs = {"text": text}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.writer("add-paragraph", positional=None, **kwargs)

    def add_heading(self, text: str, level: int = 1, index: Optional[int] = None) -> Dict[str, Any]:
        """
        Add a heading to the document.

        Args:
            text: The heading text
            level: Heading level (1-9)
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        kwargs = {"text": text, "level": str(level)}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.writer("add-heading", positional=None, **kwargs)

    def add_table(self, rows: int, cols: int, index: Optional[int] = None) -> Dict[str, Any]:
        """
        Add a table to the document.

        Args:
            rows: Number of rows
            cols: Number of columns
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        kwargs = {"rows": str(rows), "cols": str(cols)}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.writer("add-table", positional=None, **kwargs)

    def set_table_cell(self, table_index: int, row: int, col: int, value: str) -> Dict[str, Any]:
        """
        Set the value of a table cell.

        Args:
            table_index: Index of the table
            row: Row number (0-based)
            col: Column number (0-based)
            value: Cell value

        Returns:
            Command result
        """
        return self.cli.writer(
            "set-table-cell",
            positional=[str(table_index), str(row), str(col), value],
        )

    def get_table_cell(self, table_index: int, row: int, col: int) -> Dict[str, Any]:
        """
        Get the value of a table cell.

        Args:
            table_index: Index of the table
            row: Row number (0-based)
            col: Column number (0-based)

        Returns:
            Command result with cell value
        """
        return self.cli.writer(
            "get-table-cell",
            positional=[str(table_index), str(row), str(col)],
        )

    def add_list(self, items: List[str], index: Optional[int] = None) -> Dict[str, Any]:
        """
        Add a list to the document.

        Args:
            items: List items
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        # Note: cli-anything-libreoffice might have different API for lists
        # This is a placeholder implementation
        kwargs = {"items": ",".join(items)} if items else {}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.writer("add-list", positional=None, **kwargs)

    def add_page_break(self, index: Optional[int] = None) -> Dict[str, Any]:
        """
        Add a page break.

        Args:
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        kwargs = {}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.writer("add-page-break", positional=None, **kwargs)

    def set_text(self, index: int, text: str) -> Dict[str, Any]:
        """
        Set the text of a content item.

        Args:
            index: Index of the content item
            text: New text content

        Returns:
            Command result
        """
        return self.cli.writer("set-text", positional=[str(index), text])

    def remove(self, index: int) -> Dict[str, Any]:
        """
        Remove a content item by index.

        Args:
            index: Index of the content item to remove

        Returns:
            Command result
        """
        return self.cli.writer("remove", positional=[str(index)])

    def list_content(self) -> Dict[str, Any]:
        """
        List all content items in the document.

        Returns:
            List of content items with their indices and types
        """
        return self.cli.writer("list")

    def export(self, output_path: str, format: str = "pdf") -> Dict[str, Any]:
        """
        Export the document to another format.

        Args:
            output_path: Path for the exported file
            format: Export format (pdf, docx, odt, etc.)

        Returns:
            Command result
        """
        return self.cli.export("render", positional=[output_path], preset=format)

    def create_from_template(self, template_path: Optional[str] = None) -> "DocxHandler":
        """
        Create a new document from template.

        Note: This method is not yet implemented. For document creation,
        use the LibreOfficeCLI.document() method directly.

        Args:
            template_path: Path to template document (optional)

        Returns:
            New DocxHandler instance

        Raises:
            NotImplementedError: This method is not implemented
        """
        raise NotImplementedError(
            "create_from_template is not implemented. "
            "Use LibreOfficeCLI.document('new', ...) for document creation."
        )

    def analyze_structure(self) -> Dict[str, Any]:
        """
        Analyze document structure.

        Returns:
            Document analysis including sections, headings, tables, etc.
        """
        content = self.list_content()
        # Add additional analysis logic here
        if isinstance(content, list):
            items = content
        else:
            items = content.get("items", [])
        return {
            "document": str(self.document_path),
            "content_summary": content,
            "has_tables": any("table" in str(item).lower() for item in items),
            "has_headings": any("heading" in str(item).lower() for item in items),
        }

    def close(self):
        """Close the handler and clean up temporary session files."""
        if self._temp_session and self.project_path:
            try:
                import os

                if os.path.exists(self.project_path):
                    os.unlink(self.project_path)
            except Exception:
                pass  # Ignore cleanup errors
        self.cli.end_session()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
