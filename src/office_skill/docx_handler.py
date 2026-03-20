"""
DOCX (Word) document handler.

Provides high-level operations for Word document manipulation.
"""

from typing import Dict, List, Any, Optional
from pathlib import Path
from .cli_wrapper import LibreOfficeCLI


class DocxHandler:
    """Handler for Word document operations."""

    def __init__(self, document_path: str, project_path: Optional[str] = None):
        """
        Initialize handler for a Word document.

        Args:
            document_path: Path to the .docx document
            project_path: Optional path to .lo-cli.json project file
        """
        self.document_path = Path(document_path).absolute()
        self.project_path = project_path
        self.cli = LibreOfficeCLI(project_path=project_path)

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
            kwargs["index"] = index
        return self.cli.writer("add-paragraph", **kwargs)

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
        kwargs = {"text": text, "level": level}
        if index is not None:
            kwargs["index"] = index
        return self.cli.writer("add-heading", **kwargs)

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
        kwargs = {"rows": rows, "cols": cols}
        if index is not None:
            kwargs["index"] = index
        return self.cli.writer("add-table", **kwargs)

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
        return self.cli.writer("set-table-cell",
                              table_index=table_index,
                              row=row,
                              col=col,
                              value=value)

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
        return self.cli.writer("get-table-cell",
                              table_index=table_index,
                              row=row,
                              col=col)

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
            kwargs["index"] = index
        return self.cli.writer("add-list", **kwargs)

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
            kwargs["index"] = index
        return self.cli.writer("add-page-break", **kwargs)

    def set_text(self, index: int, text: str) -> Dict[str, Any]:
        """
        Set the text of a content item.

        Args:
            index: Index of the content item
            text: New text content

        Returns:
            Command result
        """
        return self.cli.writer("set-text", index=index, text=text)

    def remove(self, index: int) -> Dict[str, Any]:
        """
        Remove a content item by index.

        Args:
            index: Index of the content item to remove

        Returns:
            Command result
        """
        return self.cli.writer("remove", index=index)

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
        return self.cli.export("render",
                              input=str(self.document_path),
                              output=output_path,
                              preset=format)

    def create_from_template(self, template_path: Optional[str] = None) -> 'DocxHandler':
        """
        Create a new document from template.

        Args:
            template_path: Path to template document (optional)

        Returns:
            New DocxHandler instance
        """
        # This would need actual implementation
        # For now, just return self
        return self

    def analyze_structure(self) -> Dict[str, Any]:
        """
        Analyze document structure.

        Returns:
            Document analysis including sections, headings, tables, etc.
        """
        content = self.list_content()
        # Add additional analysis logic here
        return {
            "document": str(self.document_path),
            "content_summary": content,
            "has_tables": any("table" in str(item).lower() for item in content.get("items", [])),
            "has_headings": any("heading" in str(item).lower() for item in content.get("items", [])),
        }