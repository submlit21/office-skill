---
name: office
description: "Comprehensive office document processing tool built upon cli-anything-libreoffice, supporting creation, editing, modification, analysis, annotation, and execution of tasks for Word, Excel, and PowerPoint documents."
license: MIT
---

# Office document creation, editing, and analysis

## Overview

This skill provides AI-assisted document processing workflows for Microsoft Office documents using the `cli-anything-libreoffice` backend. It supports:

- **DOCX**: Word document creation, editing, analysis, and tracked changes
- **XLSX**: Excel spreadsheet manipulation with formula validation and recalculation
- **PPTX**: PowerPoint presentation creation and modification
- **Integrated CLI**: Unified command-line interface for all office operations

## Quick Start

### Installation Requirements
```bash
# Install cli-anything-libreoffice and dependencies
pip install cli-anything-libreoffice

# System dependencies
sudo apt install libreoffice  # Ubuntu/Debian
# brew install libreoffice    # macOS
```

### Basic Usage
```python
from office_skill import DocxHandler, XlsxHandler, PptxHandler

# Word document
doc = DocxHandler("document.docx")
doc.add_heading("Document Title", level=1)
doc.add_paragraph("This is a paragraph.")

# Excel spreadsheet
xls = XlsxHandler("spreadsheet.xlsx")
xls.set_cell("Sheet1", "A1", "Revenue", formula=False)
xls.set_cell("Sheet1", "B1", "=SUM(B2:B10)", formula=True)

# PowerPoint presentation
ppt = PptxHandler("presentation.pptx")
ppt.add_slide(layout="Title and Content")
ppt.set_content(0, title="Welcome", content="First slide content")
```

## Workflow Decision Tree

### Document Creation
- **New Word document**: Use `DocxHandler.create_from_template()` or direct editing
- **New Excel spreadsheet**: Use `XlsxHandler.create_from_template()` with financial formatting standards
- **New PowerPoint presentation**: Use `PptxHandler.create_slide_deck()` for structured creation

### Document Editing
- **Simple edits**: Use handler methods directly (add_paragraph, set_cell, etc.)
- **Complex modifications**: Use batch operations or session management
- **Tracked changes (Word)**: Use redlining workflow with OOXML editing (see DOCX skill)

### Document Analysis
- **Content extraction**: Use `analyze_structure()` methods
- **Formula validation (Excel)**: Use `validate_formulas()` with recalc.py integration
- **Format analysis**: Check formatting and style consistency

## Core Components

### CLI Wrapper (`LibreOfficeCLI`)
- Direct interface to `cli-anything-libreoffice` commands
- JSON output parsing and error handling
- Session management with undo/redo support

### Document Handlers
- **DocxHandler**: Word document operations with OOXML awareness
- **XlsxHandler**: Excel operations with formula preservation and validation
- **PptxHandler**: PowerPoint operations with slide deck management

### Integration Points
- **OpenClaw**: Use as standalone skill with markdown documentation
- **Claude Skill**: Package as Claude Code skill with skill.json
- **OpenCode**: Import as Python library for programmatic use

## Critical Rules

1. **Always validate Excel formulas**: After any spreadsheet modification, run formula validation to ensure zero errors.
2. **Preserve document integrity**: Use proper session management to avoid corruption.
3. **Leverage real LibreOffice backend**: Don't reimplement rendering; use the actual software.
4. **Support tracked changes properly**: For DOCX redlining, only mark changed text and preserve original RSIDs.

## Examples

### Creating a Financial Report
```python
from office_skill import XlsxHandler, DocxHandler

# Create financial model
xls = XlsxHandler("financial_model.xlsx")
xls.set_cell("Assumptions", "B5", "0.05", formula=False)  # Growth rate
xls.set_cell("Projections", "C10", "=B5*C9", formula=True)  # Formula

# Generate report document
doc = DocxHandler("report.docx")
doc.add_heading("Financial Analysis", level=1)
doc.add_table(5, 3)  # Results table
doc.export("report.pdf")  # Export to PDF
```

### Batch Processing
```python
from office_skill import LibreOfficeCLI

cli = LibreOfficeCLI()
# Create batch command file
commands = [
    "writer add-paragraph --text 'First paragraph'",
    "writer add-heading --text 'Section 1' --level 2",
    "export render --input document.docx --output document.pdf --preset pdf"
]
with open("commands.txt", "w") as f:
    f.write("\n".join(commands))

cli.batch("commands.txt")
```

## Troubleshooting

### Common Issues
1. **LibreOffice not found**: Ensure LibreOffice is installed and in PATH
2. **Document corruption**: Always use proper session management; don't modify files directly
3. **Formula errors in Excel**: Use `validate_formulas()` and fix all errors before delivery
4. **Performance issues**: Use batch operations for large documents

### Debugging
- Enable verbose logging: `cli-anything-libreoffice --verbose`
- Check session status: `cli-anything-libreoffice session status`
- Validate document structure with `analyze_structure()` methods

## Related Skills
- [DOCX Skill](02_DOCX_SKILL.md) - Detailed Word document workflows
- [XLSX Skill](03_XLSX_SKILL.md) - Excel spreadsheet operations
- [PPTX Skill](04_PPTX_SKILL.md) - PowerPoint presentation workflows
- [CLI Anything Methodology](05_CLI_ANYTHING.md) - Building CLI interfaces to GUI apps