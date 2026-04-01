---
name: office-skill
description: Comprehensive Microsoft Office document processing tool that leverages cli-anything-libreoffice as a backend to provide creation, editing, analysis, annotation, and template-based generation for Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) documents. Features formula validation for spreadsheets, tracked changes/redlining for Word, hierarchical template management with variable substitution, and enforces industry-standard formatting conventions for financial modeling.
license: MIT
author: office-skill-team
version: 0.1.0
---

# Office Skill - Microsoft Document Processing for AI Agents

## Overview

office-skill is an Open Claw Skill that provides comprehensive AI-assisted document processing workflows for Microsoft Office documents using the `cli-anything-libreoffice` backend. It enables AI coding agents to work confidently with Word, Excel, and PowerPoint files while respecting document integrity and following professional formatting standards.

**Key Capabilities:**

- **DOCX (Word)**: Full creation, editing, analysis, content management, and tracked changes (redlining) support with OOXML awareness
- **XLSX (Excel)**: Spreadsheet manipulation with automatic formula validation, enforces financial modeling color coding and formatting standards
- **PPTX (PowerPoint)**: Presentation creation, slide management, speaker notes, and element positioning
- **Template Management**: Hierarchical template storage (`domain.type.purpose.variant.version`), full-text search, and document generation with Jinja2 variable substitution
- **Unified CLI**: Command-line interface for all operations supports both interactive and batch processing

## Installation

```bash
# Install from source
pip install -e .

# Install dependencies
pip install cli-anything-libreoffice pyyaml jinja2

# Install system dependency (LibreOffice)
sudo apt install libreoffice  # Ubuntu/Debian
# brew install libreoffice    # macOS
```

## Quick Start

```python
from office_skill import DocxHandler, XlsxHandler, PptxHandler, TemplateManager

# --------------------------
# Word Document Example
# --------------------------
doc = DocxHandler("document.docx")
doc.add_heading("Document Title", level=1)
doc.add_paragraph("This is an introductory paragraph.")
doc.add_table(rows=3, cols=3)
doc.export("document.pdf")  # Export to PDF

# --------------------------
# Excel Spreadsheet Example
# --------------------------
xls = XlsxHandler("spreadsheet.xlsx")
xls.set_cell("Sheet1", "A1", "Revenue", formula=False)
xls.set_cell("Sheet1", "B1", "=SUM(B2:B10)", formula=True)
xls.set_cell_format("Sheet1", "B1", "$#,##0")

# Always validate after modifications!
validation = xls.validate_formulas()
if validation["status"] != "success":
    print(f"Formula errors: {validation['errors']}")

# --------------------------
# PowerPoint Example
# --------------------------
ppt = PptxHandler("presentation.pptx")
ppt.add_slide(layout="Title Slide")
ppt.set_content(0, title="Welcome", content="My Presentation")
ppt.set_speaker_notes(0, notes="Greet the audience and introduce yourself.")

# --------------------------
# Template Generation Example
# --------------------------
manager = TemplateManager()

# Generate document from existing template
result = manager.generate_from_template(
    template_name="business.report.quarterly.standard.v1",
    output_path="q1_report.docx",
    variables={
        "company": "Acme Inc",
        "quarter": "Q1 2026",
        "revenue": "$5.2M"
    }
)
```

## Module Overview

### 1. DOCX (Word Documents)
**Description**: Word document creation, editing, and analysis with support for tracked changes (redlining), comments, formatting preservation, and text extraction.

**Key Features:**
- Add/remove headings, paragraphs, lists, and tables
- Modify existing content and set table cell values
- Support for tracked changes (redlining) with proper OOXML tagging that preserves original RSIDs
- Export to PDF, ODT, HTML, and other formats
- Template-based generation with variable substitution

**Example Usage**:
```python
doc = DocxHandler("report.docx")
doc.add_heading("Quarterly Performance Review", level=1)
doc.add_paragraph("Executive summary...")
table_idx = doc.add_table(5, 3)["table_index"]
doc.set_table_cell(table_idx, 0, 0, "Metric")
```

### 2. XLSX (Excel Spreadsheets)
**Description**: Excel spreadsheet creation, editing, and analysis with mandatory formula validation and enforced financial modeling standards.

**Key Features:**
- Cell get/set operations with formula preservation
- Automatic formula validation ensures zero errors (#REF!, #DIV/0!, etc.)
- Sheet management (add, remove, rename)
- Cell formatting with compliance to financial standards:
  - Blue text: Inputs/hardcoded values
  - Black text: Formulas/calculations
  - Green text: Internal cross-sheet links
  - Red text: External links
  - Currency, percentage, zero, and negative number formatting standards
- Merged cell support

**Example Usage**:
```python
xls = XlsxHandler("budget.xlsx")
xls.add_sheet(name="2026 Budget")
xls.set_cell("2026 Budget", "A1", "Category", formula=False)
xls.set_cell("2026 Budget", "B1", "Amount", formula=False)
xls.set_cell("2026 Budget", "B10", "=SUM(B2:B9)", formula=True)
xls.set_cell_format("2026 Budget", "B10", "$#,##0")
xls.validate_formulas()  # Always validate!
```

### 3. PPTX (PowerPoint Presentations)
**Description**: PowerPoint presentation creation, editing, and analysis with support for slide decks, layouts, speaker notes, and export.

**Key Features:**
- Add/remove slides with standard layouts (Title Slide, Title and Content, etc.)
- Set slide content, titles, and speaker notes
- Add and modify text, shapes, and other elements with coordinate positioning
- Batch create complete slide decks from definitions
- Export to PDF and other formats

**Example Usage**:
```python
ppt = PptxHandler("pitch_deck.pptx")
slides = [
    {"layout": "Title Slide", "title": "Acme Inc", "content": "Revolutionizing the Industry"},
    {"layout": "Title and Content", "title": "The Problem", "content": "Current solutions are inefficient..."},
    {"layout": "Title and Content", "title": "Our Solution", "content": "10x efficiency improvement..."},
    {"layout": "Title Slide", "title": "Thank You", "content": "Contact: info@acme.com"}
]
ppt.create_slide_deck(slides)
ppt.export("pitch_deck.pdf")
```

### 4. Template Management
**Description**: Comprehensive document template management with hierarchical organization, full-text search, markdown conversion for AI analysis, and Jinja2 variable substitution.

**Key Features:**
- Hierarchical storage following `domain.type.purpose.variant.version` naming convention
- Automatic document structure analysis (pages, paragraphs, tables, sheets, slides)
- Conversion to markdown/YAML for AI-friendly content analysis
- Full-text search across template names, descriptions, and tags
- Jinja2 advanced templating with variable substitution, conditionals, and loops
- Batch generation from templates with CSV data

**Storage Structure:**
```
templates/
└── domain/
    └── type/
        └── purpose/
            └── variant/
                └── version/
                    ├── metadata.json
                    ├── template.md (DOCX) or template.yaml (XLSX/PPTX)
                    └── original.docx
```

**Example Usage**:
```python
from office_skill import TemplateManager

tm = TemplateManager()

# Add a new template
tm.add_template(
    source_path="meeting_nameplate.docx",
    name="meeting.signin.nameplate.chinese.v1",
    description="会议席签模板 - 中文",
    tags=["meeting", "signin", "nameplate"]
)

# Search templates
results = tm.search_templates("meeting")
for t in results:
    print(t["name"])

# Generate document from template
result = tm.generate_from_template(
    template_name="meeting.signin.nameplate.chinese.v1",
    output_path="zhangsan_nameplate.docx",
    variables={"name": "张三"}
)
```

## Critical Rules

1. **Always validate Excel formulas**: After any spreadsheet modification, run `validate_formulas()` to ensure zero errors.
2. **Delegated rendering**: Never reimplement Office document rendering; always rely on LibreOffice via `cli-anything-libreoffice`.
3. **DOCX redlining**: For tracked changes, only mark changed text with `<w:ins>`/`<w:del>` tags and preserve original RSIDs.
4. **Template naming**: All templates **must** follow the `domain.type.purpose.variant.version` naming format.
5. **Financial standards**: Follow established color coding and formatting conventions when working with financial spreadsheets.
6. **Preserve document integrity**: Use proper session management to avoid document corruption.
7. **Format preservation**: Maintain original formatting when generating documents from templates.
8. **Version control**: Always increment version when updating templates.

## Supported Formats

| Format | Operations Supported | Key Features |
|--------|----------------------|--------------|
| DOCX/DOC | Create, edit, analyze, export, redlining | Tracked changes, OOXML awareness |
| XLSX/XLS/CSV | Create, edit, analyze, formula validation | Financial standards compliance |
| PPTX/PPT | Create, edit, analyze, export | Slide management, speaker notes |
| PDF | Export from all formats | High-quality rendering via LibreOffice |

## Workflow Decision Tree

### Creating New Documents
- **From scratch**: Use direct handler methods (`add_heading`, `set_cell`, `add_slide`, etc.)
- **From template**: Use `TemplateManager.generate_from_template()` for consistent results
- **Batch creation**: Use `create_slide_deck()` (PPTX) or batch processing for multiple documents

### Editing Existing Documents
- **Simple changes**: Use direct handler methods
- **Legal/business documents**: Use redlining/tracked changes workflow for DOCX
- **Spreadsheets**: Always run formula validation after edits

### Analysis
- **Content extraction**: Use `analyze_structure()` for document structure overview
- **Template discovery**: Use `search_templates()` to find reusable patterns
- **Quality assurance**: Validate formulas (XLSX) and check exports (all formats)

## Development

### Code Quality Checks
```bash
black src/          # format code
ruff check src/     # lint
mypy src/          # type checking
pytest tests/      # run tests
```

### Project Structure
```
src/
├── office_main/
│   ├── cli/              # Command-line interface
│   └── core/             # DOCX/XLSX/PPTX handlers
└── template_manager/     # Template management system
templates/                # Template storage (hierarchical)
tests/                    # Test files and example documents
skills/                   # Individual skill markdown files
```

## Dependencies

- `cli-anything-libreoffice`: Backend for Office operations
- `LibreOffice`: System-level installation for document processing
- `pyyaml`: YAML configuration support
- `jinja2`: Advanced templating and variable substitution
- `pytest`: Testing framework (development)

## Examples

See the [`./examples/`](./examples/) directory for complete usage examples:
- [`basic_template_management.md`](./examples/basic_template_management.md) - Adding, listing, searching, and deleting templates
- [`generate_document_from_template.md`](./examples/generate_document_from_template.md) - Generating documents with Jinja2 variables
- [`analyze_document.md`](./examples/analyze_document.md) - Analyzing document structure metadata
- [`convert_to_markdown.md`](./examples/convert_to_markdown.md) - Converting documents to Markdown

## License

MIT
