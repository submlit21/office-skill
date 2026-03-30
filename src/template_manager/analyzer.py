"""
Template analysis - Analyzes document structure for template metadata.

This module provides functionality to analyze document structure
and extract metadata for template cataloging.
"""

import subprocess
import tempfile
from pathlib import Path
from typing import Any, Dict

from office_main.core.cli_wrapper import LibreOfficeCLI


class TemplateAnalyzer:
    """Analyzes document structure for template metadata."""

    def __init__(self, cli: LibreOfficeCLI):
        """Initialize analyzer with CLI instance."""
        self.cli = cli

    def convert_to_markdown(self, source_path: Path) -> str:
        """Convert office document to markdown format."""
        source_path = source_path.absolute()

        if not source_path.exists():
            raise FileNotFoundError(f"Source file not found: {source_path}")

        # Check file extension
        ext = source_path.suffix.lower()
        if ext not in [".docx", ".xlsx", ".pptx", ".doc", ".xls", ".ppt"]:
            raise ValueError(f"Unsupported file format: {ext}")

        # Create temporary output file
        with tempfile.NamedTemporaryFile(suffix=".md", delete=False) as tmp:
            output_path = tmp.name

        actual_output = None
        try:
            # First try: Use cli-anything-libreoffice export to text, then convert to markdown
            try:
                # Create a temporary session
                temp_cli = LibreOfficeCLI()
                session_file = temp_cli.start_session(str(source_path))
                cli_with_session = LibreOfficeCLI(project_path=session_file)

                # Export as text (document is already loaded via project file)
                text_output = Path(output_path).with_suffix(".txt")
                try:
                    cli_with_session.export(
                        "render", positional=[str(text_output)], preset="text", overwrite=True
                    )

                    # Read text output and convert to simple markdown
                    if text_output.exists():
                        with open(text_output, "r", encoding="utf-8") as f:
                            text_content = f.read()

                        # Check if content is meaningful (not empty or just whitespace)
                        if text_content.strip() and len(text_content.strip()) > 10:
                            # Convert to basic markdown
                            # Add title, preserve paragraphs
                            lines = text_content.split("\n")
                            markdown_lines = []
                            for line in lines:
                                if line.strip():
                                    markdown_lines.append(line)
                                else:
                                    markdown_lines.append("")

                            markdown_content = f"# {source_path.name}\n\n" + "\n".join(
                                markdown_lines
                            )

                            # Clean up
                            if text_output.exists():
                                text_output.unlink()
                            temp_cli.end_session()

                            return markdown_content
                        else:
                            # Content too short, fall back to other methods
                            if text_output.exists():
                                text_output.unlink()
                except Exception:
                    # Export failed, try fallback methods
                    pass

                # Clean up
                temp_cli.end_session()
            except Exception:
                # CLI-based conversion failed, fall back to original methods
                pass

            # Fallback to original pandoc/libreoffice conversion
            converted = False

            # Try to find libreoffice executable
            libreoffice_paths = [
                "libreoffice",  # Standard system path
                "/usr/bin/libreoffice",
                "/usr/local/bin/libreoffice",
                # Common conda/env paths
                "/opt/homebrew/bin/libreoffice",  # macOS Homebrew
                "/usr/local/software/miniconda3/envs/office-skill/bin/libreoffice",
                # Windows paths (if needed)
                # "C:\\Program Files\\LibreOffice\\program\\soffice.exe",
            ]

            libreoffice_found = None
            for path in libreoffice_paths:
                try:
                    subprocess.run([path, "--version"], capture_output=True, check=True)
                    libreoffice_found = path
                    break
                except (FileNotFoundError, subprocess.CalledProcessError):
                    continue

            # Try pandoc first - better markdown conversion from docx
            try:
                subprocess.run(
                    ["pandoc", "-s", str(source_path), "-o", output_path],
                    capture_output=True,
                    text=True,
                    check=True,
                )
                converted = True
            except (FileNotFoundError, subprocess.CalledProcessError):
                # pandoc not available, try libreoffice if found
                if libreoffice_found:
                    try:
                        subprocess.run(
                            [
                                libreoffice_found,
                                "--headless",
                                "--convert-to",
                                "md",
                                "--outdir",
                                str(Path(output_path).parent),
                                str(source_path),
                            ],
                            capture_output=True,
                            text=True,
                            check=True,
                        )
                        # LibreOffice outputs to output_dir/input_name.md, not the temp name we gave
                        actual_output = Path(output_path).parent / (source_path.stem + ".md")
                        if actual_output.exists():
                            output_path = actual_output
                        converted = True
                    except (FileNotFoundError, subprocess.CalledProcessError):
                        # libreoffice conversion failed
                        converted = False
                else:
                    # No libreoffice found
                    converted = False

            if converted:
                # Read the converted markdown
                with open(output_path, "r", encoding="utf-8") as f:
                    markdown_content = f.read()
            else:
                # Conversion failed
                markdown_content = (
                    f"# {source_path.name}\n\nConversion failed, no markdown preview available."
                )

            return markdown_content

        except Exception:
            # If conversion fails, fall back to simple text extraction
            markdown_content = (
                f"# {source_path.name}\n\nConversion failed, no markdown preview available."
            )
            return markdown_content

    def analyze_document_structure(self, source_path: Path) -> Dict[str, Any]:
        """Analyze document structure for template metadata."""
        source_path = source_path.absolute()

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
            # Create a temporary session to analyze the document
            temp_cli = LibreOfficeCLI()
            session_file = temp_cli.start_session(str(path))
            cli_with_session = LibreOfficeCLI(project_path=session_file)

            # Get document info
            info_result = cli_with_session.document("info")

            # List content items
            list_result = cli_with_session.writer("list")

            # Handle response format (could be list or dict)
            if isinstance(list_result, list):
                items = list_result
            elif isinstance(list_result, dict):
                items = list_result.get("items", [])
            else:
                items = []

            paragraphs = headings = tables = images = 0
            for item in items:
                if isinstance(item, dict):
                    item_type = item.get("type", "")
                else:
                    item_type = str(item)

                if item_type == "paragraph":
                    paragraphs += 1
                elif item_type == "heading":
                    headings += 1
                elif item_type == "table":
                    tables += 1
                elif item_type == "image":
                    images += 1

            # Clean up
            temp_cli.end_session()

            return {
                "type": "word",
                "pages": info_result.get("page_count", 0),
                "paragraphs": paragraphs,
                "headings": headings,
                "tables": tables,
                "images": images,
                "sections": info_result.get("section_count", 0),
                "styles": info_result.get("styles", []),
                "metadata": info_result.get("metadata", {}),
                "content_count": len(items),
            }
        except Exception as e:
            return {
                "type": "word",
                "error": str(e),
                "pages": 0,
                "paragraphs": 0,
                "tables": 0,
                "images": 0,
                "headings": 0,
                "sections": 0,
                "styles": [],
                "metadata": {},
                "content_count": 0,
            }

    def _analyze_excel_document(self, path: Path) -> Dict[str, Any]:
        """Analyze Excel document structure."""
        try:
            # Create a temporary session to analyze the document
            temp_cli = LibreOfficeCLI()
            session_file = temp_cli.start_session(str(path))
            cli_with_session = LibreOfficeCLI(project_path=session_file)

            # Get document info
            info_result = cli_with_session.document("info")

            # List sheets
            sheets_result = cli_with_session.calc("list-sheets")

            # Handle response format (could be list or dict)
            if isinstance(sheets_result, list):
                sheets = sheets_result
            elif isinstance(sheets_result, dict):
                sheets = sheets_result.get("sheets", [])
            else:
                sheets = []

            sheet_count = len(sheets)
            # Estimate cells: assume 1000 cells per sheet (rough estimate)
            estimated_cells = sheet_count * 1000

            # Clean up
            temp_cli.end_session()

            return {
                "type": "excel",
                "sheets": sheet_count,
                "cells": estimated_cells,
                "formulas": info_result.get("formula_count", 0),
                "charts": info_result.get("chart_count", 0),
                "named_ranges": info_result.get("named_range_count", 0),
                "metadata": info_result.get("metadata", {}),
                "sheet_names": [
                    s.get("name", "") if isinstance(s, dict) else str(s) for s in sheets
                ],
            }
        except Exception as e:
            return {
                "type": "excel",
                "error": str(e),
                "sheets": 0,
                "cells": 0,
                "formulas": 0,
                "charts": 0,
                "named_ranges": 0,
                "metadata": {},
                "sheet_names": [],
            }

    def _analyze_powerpoint_document(self, path: Path) -> Dict[str, Any]:
        """Analyze PowerPoint document structure."""
        try:
            # Create a temporary session to analyze the document
            temp_cli = LibreOfficeCLI()
            session_file = temp_cli.start_session(str(path))
            cli_with_session = LibreOfficeCLI(project_path=session_file)

            # Get document info
            info_result = cli_with_session.document("info")

            # List slides
            slides_result = cli_with_session.impress("list-slides")

            # Handle response format (could be list or dict)
            if isinstance(slides_result, list):
                slides = slides_result
            elif isinstance(slides_result, dict):
                slides = slides_result.get("slides", [])
            else:
                slides = []

            slide_count = len(slides)

            # Count layouts and check for notes
            layouts = {}
            has_notes = False
            for slide in slides:
                if isinstance(slide, dict):
                    layout = slide.get("layout", "Unknown")
                    if slide.get("has_notes", False):
                        has_notes = True
                else:
                    layout = "Unknown"
                layouts[layout] = layouts.get(layout, 0) + 1

            # Clean up
            temp_cli.end_session()

            return {
                "type": "powerpoint",
                "slides": slide_count,
                "masters": info_result.get("master_count", 0),
                "layouts": len(layouts),
                "layout_distribution": layouts,
                "shapes": info_result.get("shape_count", 0),
                "images": info_result.get("image_count", 0),
                "notes": has_notes,
                "metadata": info_result.get("metadata", {}),
            }
        except Exception as e:
            return {
                "type": "powerpoint",
                "error": str(e),
                "slides": 0,
                "masters": 0,
                "layouts": 0,
                "shapes": 0,
                "images": 0,
                "notes": False,
                "metadata": {},
                "layout_distribution": {},
            }
