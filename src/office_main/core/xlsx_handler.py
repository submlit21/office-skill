"""
XLSX (Excel) spreadsheet handler.

Provides high-level operations for Excel spreadsheet manipulation.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional

from .cli_wrapper import LibreOfficeCLI


class XlsxHandler:
    """Handler for Excel spreadsheet operations."""

    def __init__(self, spreadsheet_path: str, project_path: Optional[str] = None):
        """
        Initialize handler for an Excel spreadsheet.

        Args:
            spreadsheet_path: Path to the .xlsx spreadsheet
            project_path: Optional path to .lo-cli.json project file.
                         If None, a temporary session file will be created.
        """
        self.spreadsheet_path = Path(spreadsheet_path).absolute()
        self.project_path = project_path
        self._temp_session = False

        if self.project_path is None:
            # Create a temporary session file
            temp_cli = LibreOfficeCLI()
            self.project_path = temp_cli.start_session(str(self.spreadsheet_path))
            # Do NOT end session here - keep file for CLI operations
            self._temp_session = True

        self.cli = LibreOfficeCLI(project_path=self.project_path)

    def get_cell(self, sheet: str, cell_ref: str) -> Dict[str, Any]:
        """
        Get a cell value.

        Args:
            sheet: Sheet name or index
            cell_ref: Cell reference (e.g., "A1", "B5")

        Returns:
            Cell value and metadata
        """
        return self.cli.calc("get-cell", positional=[cell_ref], sheet=sheet)

    def set_cell(
        self, sheet: str, cell_ref: str, value: str, formula: bool = False
    ) -> Dict[str, Any]:
        """
        Set a cell value.

        Args:
            sheet: Sheet name or index
            cell_ref: Cell reference (e.g., "A1", "B5")
            value: Cell value (text, number, or formula)
            formula: Whether the value is a formula (starts with "=")

        Returns:
            Command result
        """
        if formula and not value.startswith("="):
            value = f"={value}"
        kwargs = {"sheet": sheet}
        if formula:
            kwargs["formula"] = ""
        return self.cli.calc("set-cell", positional=[cell_ref, value], **kwargs)

    def set_cell_format(self, sheet: str, cell_ref: str, format_spec: str) -> Dict[str, Any]:
        """
        Set formatting for a cell.

        Args:
            sheet: Sheet name or index
            cell_ref: Cell reference
            format_spec: Format specification (e.g., "$#,##0.00", "0.0%")

        Returns:
            Command result
        """
        return self.cli.calc(
            "set-cell-format", positional=[cell_ref], sheet=sheet, format=format_spec
        )

    def add_sheet(self, name: str, index: Optional[int] = None) -> Dict[str, Any]:
        """
        Add a new sheet.

        Args:
            name: Name for the new sheet
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        kwargs: Dict[str, Any] = {"name": name}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.calc("add-sheet", positional=None, **kwargs)

    def remove_sheet(self, sheet: str) -> Dict[str, Any]:
        """
        Remove a sheet.

        Args:
            sheet: Sheet name or index

        Returns:
            Command result
        """
        return self.cli.calc("remove-sheet", positional=[sheet])

    def rename_sheet(self, old_name: str, new_name: str) -> Dict[str, Any]:
        """
        Rename a sheet.

        Args:
            old_name: Current sheet name
            new_name: New sheet name

        Returns:
            Command result
        """
        return self.cli.calc("rename-sheet", positional=[old_name, new_name])

    def list_sheets(self) -> Dict[str, Any]:
        """
        List all sheets in the workbook.

        Returns:
            List of sheets with names and indices
        """
        return self.cli.calc("list-sheets")

    def merge_cells(self, sheet: str, range_ref: str) -> Dict[str, Any]:
        """
        Merge a range of cells.

        Args:
            sheet: Sheet name or index
            range_ref: Cell range (e.g., "A1:B5")

        Returns:
            Command result
        """
        return self.cli.calc("merge", positional=[range_ref], sheet=sheet)

    def unmerge_cells(self, sheet: str, cell_ref: str) -> Dict[str, Any]:
        """
        Unmerge cells.

        Args:
            sheet: Sheet name or index
            cell_ref: Reference to the top-left cell of the merged range

        Returns:
            Command result
        """
        return self.cli.calc("unmerge", positional=[cell_ref], sheet=sheet)

    def export(self, output_path: str, format: str = "pdf") -> Dict[str, Any]:
        """
        Export the spreadsheet to another format.

        Args:
            output_path: Path for the exported file
            format: Export format (pdf, xlsx, ods, csv, etc.)

        Returns:
            Command result
        """
        return self.cli.export("render", positional=[output_path], preset=format)

    def create_from_template(self, template_path: Optional[str] = None) -> "XlsxHandler":
        """
        Create a new spreadsheet from template.

        Note: This method is not yet implemented. For spreadsheet creation,
        use the LibreOfficeCLI.document() method directly.

        Args:
            template_path: Path to template spreadsheet (optional)

        Returns:
            New XlsxHandler instance

        Raises:
            NotImplementedError: This method is not implemented
        """
        raise NotImplementedError(
            "create_from_template is not implemented. "
            "Use LibreOfficeCLI.document('new', ...) for spreadsheet creation."
        )

    def analyze_structure(self) -> Dict[str, Any]:
        """
        Analyze spreadsheet structure.

        Returns:
            Spreadsheet analysis including sheets, formulas, etc.
        """
        sheets = self.list_sheets()
        sheet_names = [s.get("name", "") for s in sheets.get("sheets", [])]

        # Check for formulas in first sheet if available
        formula_count = 0
        if sheet_names:
            # Sample a few cells to check for formulas
            # This is a simplified analysis
            pass

        return {
            "spreadsheet": str(self.spreadsheet_path),
            "sheet_count": len(sheet_names),
            "sheet_names": sheet_names,
            "estimated_formula_count": formula_count,
        }

    def validate_formulas(self) -> Dict[str, Any]:
        """
        Validate formulas in the spreadsheet.

        Note: This would require integration with recalc.py from reference project.

        Returns:
            Validation results
        """
        # This is a placeholder - actual implementation would use recalc.py
        return {
            "status": "validation_not_implemented",
            "message": "Formula validation requires recalc.py integration",
            "recommendation": "Run python XLSX/recalc.py path/to/spreadsheet.xlsx",
        }

    def apply_financial_formatting(
        self,
        sheet: str,
        currency_cells: Optional[List[str]] = None,
        percent_cells: Optional[List[str]] = None,
    ) -> Dict[str, Any]:
        """
        Apply financial formatting standards.

        Args:
            sheet: Sheet name or index
            currency_cells: List of cell references for currency formatting
            percent_cells: List of cell references for percentage formatting

        Returns:
            Command results
        """
        results = []

        # Apply currency formatting
        if currency_cells:
            for cell in currency_cells:
                result = self.set_cell_format(sheet, cell, "$#,##0")
                results.append({"cell": cell, "format": "currency", "result": result})

        # Apply percentage formatting
        if percent_cells:
            for cell in percent_cells:
                result = self.set_cell_format(sheet, cell, "0.0%")
                results.append({"cell": cell, "format": "percentage", "result": result})

        return {"applied_formats": results}

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
