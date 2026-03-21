---
name: office-skill
description: "An office processing tool built upon the CLI-Anything-generated cli-anything-libreoffice, which supports the creation, editing, modification, analysis, annotation, template management, and execution of any other relevant tasks for Word, Excel, and PowerPoint documents."
license: MIT
---

# Office Skill - Comprehensive Document Processing

## Overview

Office-skill is a powerful Python package and Claude Skill for processing Microsoft Office documents (Word, Excel, PowerPoint) using the `cli-anything-libreoffice` backend. It provides both programmatic APIs and command-line interfaces for professional document workflows.

## Key Features

### 1. **Multi-Format Support**
- **Word**: .docx, .doc (creation, editing, analysis, tracked changes)
- **Excel**: .xlsx, .xls, .csv (spreadsheet operations, formula validation)
- **PowerPoint**: .pptx, .ppt (presentation creation, slide management)

### 2. **Template Management System**
- Hierarchical template storage with `domain.type.purpose.variant.version` naming
- Office document to Markdown conversion
- Template search and filtering capabilities
- Variable-based document generation

### 3. **Integration Platforms**
- **OpenClaw**: Skill definitions in openclaw_skills/ directory
- **Claude Code**: Native skill integration via skill.json
- **OpenCode**: Direct Python API import

### 4. **Professional Workflows**
- Tracked changes and redlining for legal documents
- Financial modeling standards enforcement
- Presentation deck creation and export

## Installation

### Python Package
```bash
pip install office-skill
```

### System Dependencies
```bash
# Ubuntu/Debian
sudo apt install libreoffice

# macOS
brew install libreoffice

# Windows
# Download from https://www.libreoffice.org/
```

### Claude Skill Installation
```bash
# In Claude Code
claude skills add ./path/to/office-skill
```

## Quick Start

### Python API
```python
from office_skill import DocxHandler, XlsxHandler, PptxHandler, TemplateManager

# Word Document
doc = DocxHandler("document.docx")
doc.add_heading("Document Title", level=1)
doc.add_paragraph("Content paragraph...")

# Excel Spreadsheet
xls = XlsxHandler("spreadsheet.xlsx")
xls.set_cell("Sheet1", "A1", "Data", formula=False)

# PowerPoint Presentation
ppt = PptxHandler("presentation.pptx")
ppt.add_slide(layout="Title and Content")

# Template Management
manager = TemplateManager()
manager.add_template(
    source_path="report.docx",
    name="business.report.quarterly.standard.v1",
    description="Quarterly business report template",
    tags=["business", "report", "quarterly"]
)
```

### CLI Usage
```bash
# Create documents
office_cli.py create --output document.docx

# Analyze content
office_cli.py analyze --input spreadsheet.xlsx --json

# Manage templates
office_cli.py template add --input report.docx --name business.report.quarterly.standard.v1
office_cli.py template list --type report --verbose
office_cli.py template generate --template business.report.quarterly.standard.v1 --output new_report.docx
```

## Template Management System

### Naming Convention
Templates use a hierarchical naming system:
```
domain.type.purpose.variant.version
```

**Examples:**
- `business.report.quarterly.standard.v1`
- `legal.contract.nda.comprehensive.v2`
- `education.presentation.lecture.simple.v1`

### Directory Structure
```
templates/
├── business/
│   └── report/
│       └── quarterly/
│           └── standard/
│               └── v1/
│                   ├── template.md      # Markdown conversion
│                   ├── metadata.json    # Template metadata
│                   └── original.docx    # Original document copy
```

### Template Operations
1. **Add Template**: Convert Office document to markdown, store with metadata
2. **List Templates**: Browse with domain/type/purpose filtering
3. **Search Templates**: Find by name, description, or tags
4. **Generate Documents**: Create new documents from templates with variable substitution

## Integration with OpenClaw

The office-skill package includes comprehensive OpenClaw skill definitions:

### Available Skills
1. **Office Skill** (`01_OFFICE_SKILL.md`) - Comprehensive office document processing
2. **DOCX Skill** (`02_DOCX_SKILL.md`) - Word document workflows with tracked changes
3. **XLSX Skill** (`03_XLSX_SKILL.md`) - Excel operations with formula validation
4. **PPTX Skill** (`04_PPTX_SKILL.md`) - PowerPoint presentation management

### Template Integration
Template management is integrated across all skill types:
- **DOCX Templates**: Word document templates for reports, letters, contracts
- **XLSX Templates**: Spreadsheet templates for financial models, budgets, dashboards
- **PPTX Templates**: Presentation templates for business decks, training materials

## Usage Examples

### 1. Business Report Generation
```python
from office_skill import TemplateManager

manager = TemplateManager()
report_data = {
    "company_name": "Acme Inc",
    "quarter": "Q4 2023",
    "revenue": "$5.2M",
    "growth": "+15%"
}

manager.generate_from_template(
    template_name="business.report.quarterly.standard.v1",
    output_path="q4_report.docx",
    variables=report_data
)
```

### 2. Financial Model Template
```python
manager.add_template(
    source_path="financial_model.xlsx",
    name="finance.model.three_statement.standard.v1",
    description="Three-statement financial model template",
    tags=["finance", "model", "accounting"]
)
```

### 3. Presentation Deck Creation
```python
manager.add_template(
    source_path="pitch_deck.pptx",
    name="marketing.presentation.investor_pitch.modern.v1",
    description="Modern investor pitch deck template",
    tags=["marketing", "presentation", "pitch"]
)
```

## Template Naming Guidelines

### Domain Categories
- `business` - Corporate and enterprise documents
- `legal` - Contracts, agreements, legal documents
- `education` - Learning materials, course content
- `finance` - Financial models, budgets, reports
- `marketing` - Promotional materials, presentations
- `technology` - Technical documentation, specifications

### Document Types
- `report` - Analytical and summary documents
- `contract` - Legal agreements and terms
- `presentation` - Slide decks and visual materials
- `model` - Financial and analytical models
- `plan` - Strategic and operational plans
- `letter` - Formal and informal correspondence

### Purpose Examples
- `quarterly` - Regular periodic reports
- `annual` - Yearly summaries and plans
- `nda` - Non-disclosure agreements
- `proposal` - Business and project proposals
- `budget` - Financial planning documents
- `training` - Educational materials

## Critical Rules

### 1. **Template Naming Compliance**
- Always use 5-part format: `domain.type.purpose.variant.version`
- Follow established domain and type categories
- Version numbers must be sequential (v1, v2, v3)

### 2. **Document Integrity**
- Preserve original formatting and styles
- Validate formulas in Excel templates
- Maintain consistent layouts in PowerPoint templates

### 3. **Performance Considerations**
- Use batch operations for multiple template additions
- Implement caching for frequently accessed templates
- Clean up temporary files after conversions

## System Requirements

### Software
- **Python**: 3.8+
- **LibreOffice**: 7.0+ (for document conversion)
- **cli-anything-libreoffice**: 0.1.0+ (installed via pip)

### Storage
- Templates directory with hierarchical structure
- Sufficient space for original documents and markdown conversions

### Permissions
- Read/write access to templates directory
- Execution permissions for LibreOffice commands

## Troubleshooting

### Common Issues
1. **LibreOffice Not Found**: Ensure LibreOffice is installed and in PATH
2. **Template Conversion Errors**: Verify document is not corrupted
3. **Storage Issues**: Check available space and directory permissions

### Debug Mode
Enable verbose logging for troubleshooting:
```bash
office_cli.py template add --input document.docx --name test.template.debug.v1 --verbose
```

## Related Resources

### Documentation
- `AGENTS.md` - Development guide for AI agents
- `README.md` - User documentation
- `pyproject.toml` - Build configuration

### Source Code
- `src/office_skill/` - Core Python package
- `src/claude_skill/` - Claude Skill integration
- `examples/` - Usage examples and templates

### Skill Definitions
- `openclaw_skills/` - OpenClaw skill markdown files
- `skill.json` - Claude Skill configuration

## License

MIT License - See LICENSE file for details.

## Support

For issues, feature requests, or contributions:
- Report issues on GitHub repository
- Submit pull requests for improvements
- Contact maintainers for critical problems