"""
Basic tests for office-skill package.

These tests verify the core functionality of the package.
Note: Some tests require cli-anything-libreoffice to be installed.
"""

import pytest
import os
import tempfile
from pathlib import Path

# Check if cli-anything-libreoffice is available
try:
    import subprocess
    subprocess.run(["cli-anything-libreoffice", "--help"],
                   capture_output=True, check=True)
    CLI_AVAILABLE = True
except (subprocess.CalledProcessError, FileNotFoundError):
    CLI_AVAILABLE = False

# Skip all tests if CLI not available
pytestmark = pytest.mark.skipif(
    not CLI_AVAILABLE,
    reason="cli-anything-libreoffice not installed or not in PATH"
)


class TestCLIWrapper:
    """Tests for the CLI wrapper."""

    def test_cli_initialization(self):
        """Test that CLI wrapper can be initialized."""
        from office_skill import LibreOfficeCLI
        cli = LibreOfficeCLI()
        assert cli is not None
        assert cli.json_output is True

    def test_cli_with_project_path(self):
        """Test CLI wrapper with project path."""
        from office_skill import LibreOfficeCLI
        cli = LibreOfficeCLI(project_path="/tmp/test.json", json_output=False)
        assert cli.project_path == "/tmp/test.json"
        assert cli.json_output is False

    def test_get_export_presets(self):
        """Test getting export presets."""
        from office_skill import LibreOfficeCLI
        cli = LibreOfficeCLI()
        result = cli.export("presets")
        assert isinstance(result, dict)
        # Should have presets key
        assert "presets" in result

    def test_writer_commands(self):
        """Test writer command listing."""
        from office_skill import LibreOfficeCLI
        cli = LibreOfficeCLI()
        result = cli.writer("list")
        assert isinstance(result, dict)
        # Should have commands key
        assert "commands" in result


class TestDocxHandler:
    """Tests for DOCX handler."""

    def test_handler_initialization(self):
        """Test that handler can be initialized."""
        from office_skill import DocxHandler
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            handler = DocxHandler(tmp.name)
            assert handler is not None
            assert Path(handler.document_path).name.endswith(".docx")

    def test_analyze_structure(self):
        """Test document structure analysis."""
        from office_skill import DocxHandler
        with tempfile.NamedTemporaryFile(suffix=".docx") as tmp:
            handler = DocxHandler(tmp.name)
            analysis = handler.analyze_structure()
            assert isinstance(analysis, dict)
            assert "document" in analysis
            assert "content_summary" in analysis


class TestXlsxHandler:
    """Tests for XLSX handler."""

    def test_handler_initialization(self):
        """Test that handler can be initialized."""
        from office_skill import XlsxHandler
        with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
            handler = XlsxHandler(tmp.name)
            assert handler is not None
            assert Path(handler.spreadsheet_path).name.endswith(".xlsx")

    def test_list_sheets(self):
        """Test sheet listing."""
        from office_skill import XlsxHandler
        with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
            handler = XlsxHandler(tmp.name)
            sheets = handler.list_sheets()
            assert isinstance(sheets, dict)
            assert "sheets" in sheets


class TestPptxHandler:
    """Tests for PPTX handler."""

    def test_handler_initialization(self):
        """Test that handler can be initialized."""
        from office_skill import PptxHandler
        with tempfile.NamedTemporaryFile(suffix=".pptx") as tmp:
            handler = PptxHandler(tmp.name)
            assert handler is not None
            assert Path(handler.presentation_path).name.endswith(".pptx")

    def test_list_slides(self):
        """Test slide listing."""
        from office_skill import PptxHandler
        with tempfile.NamedTemporaryFile(suffix=".pptx") as tmp:
            handler = PptxHandler(tmp.name)
            slides = handler.list_slides()
            assert isinstance(slides, dict)
            assert "slides" in slides


class TestPackageImport:
    """Tests for package imports and structure."""

    def test_package_imports(self):
        """Test that all main modules can be imported."""
        from office_skill import (
            LibreOfficeCLI,
            DocxHandler,
            XlsxHandler,
            PptxHandler
        )
        assert LibreOfficeCLI is not None
        assert DocxHandler is not None
        assert XlsxHandler is not None
        assert PptxHandler is not None

    def test_version(self):
        """Test package version."""
        from office_skill import __version__
        assert isinstance(__version__, str)
        assert __version__ == "0.1.0"


@pytest.mark.integration
class TestIntegration:
    """Integration tests that create actual documents."""

    def test_create_simple_docx(self):
        """Create a simple DOCX document with content."""
        from office_skill import DocxHandler
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            doc_path = tmp.name

        try:
            doc = DocxHandler(doc_path)
            # Add some content
            result = doc.add_heading("Test Heading", level=1)
            assert isinstance(result, dict)

            # Analyze structure
            analysis = doc.analyze_structure()
            assert analysis["has_headings"] is True

        finally:
            if os.path.exists(doc_path):
                os.unlink(doc_path)

    def test_create_simple_xlsx(self):
        """Create a simple XLSX spreadsheet."""
        from office_skill import XlsxHandler
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            xls_path = tmp.name

        try:
            xls = XlsxHandler(xls_path)
            # Add a sheet
            result = xls.add_sheet(name="TestSheet")
            assert isinstance(result, dict)

            # List sheets
            sheets = xls.list_sheets()
            assert len(sheets.get("sheets", [])) > 0

        finally:
            if os.path.exists(xls_path):
                os.unlink(xls_path)


if __name__ == "__main__":
    # Run tests directly for debugging
    pytest.main([__file__, "-v"])