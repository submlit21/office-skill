# Office Skill

A comprehensive office document processing tool compatible with OpenClaw, Claude Skill, and OpenCode, built upon the `cli-anything-libreoffice` backend.

## Features

- **Multi-format Support**: Word (.docx/.doc), Excel (.xlsx/.xls), PowerPoint (.pptx/.ppt)
- **Cross-platform Compatibility**: Works with OpenClaw, Claude Skill, and OpenCode
- **Real Backend Integration**: Leverages actual LibreOffice for rendering and conversion
- **Professional Workflows**: Supports tracked changes, formula validation, presentation decks
- **Python API**: Clean, typed Python interface for programmatic use

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

# Windows: Install LibreOffice from https://www.libreoffice.org/
```

### Claude Skill Installation
```bash
# In Claude Code, add the skill directory
claude skills add ./path/to/office-skill
```

## Quick Start

### Python API
```python
from office_skill import DocxHandler, XlsxHandler, PptxHandler

# Word document
doc = DocxHandler("document.docx")
doc.add_heading("Title", level=1)
doc.add_paragraph("Content")
doc.export("output.pdf")

# Excel spreadsheet
xls = XlsxHandler("spreadsheet.xlsx")
xls.set_cell("Sheet1", "A1", "Data")
xls.set_cell("Sheet1", "B1", "=SUM(B2:B10)", formula=True)

# PowerPoint presentation
ppt = PptxHandler("presentation.pptx")
ppt.add_slide(layout="Title and Content")
ppt.set_content(0, title="Welcome", content="Content")
```

### OpenClaw Usage
Load the skill files from `openclaw_skills/` directory in OpenClaw.

### Claude Skill Usage
```bash
# Using Claude Code with the skill
claude office create-document --type docx --output report.docx
claude office analyze-document --input data.xlsx --type xlsx
claude office export-document --input presentation.pptx --output slides.pdf --format pdf
```

## Project Structure

```
office-skill/
├── src/
│   ├── office_skill/          # Core Python package
│   │   ├── __init__.py
│   │   ├── cli_wrapper.py     # CLI interface
│   │   ├── docx_handler.py    # Word operations
│   │   ├── xlsx_handler.py    # Excel operations
│   │   └── pptx_handler.py    # PowerPoint operations
│   └── claude_skill/          # Claude Skill adapter
│       ├── index.js
│       └── skill.json
├── openclaw_skills/           # OpenClaw skill definitions
│   ├── 01_OFFICE_SKILL.md
│   ├── 02_DOCX_SKILL.md
│   ├── 03_XLSX_SKILL.md
│   ├── 04_PPTX_SKILL.md
│   └── _meta.json
├── examples/                  # Usage examples
├── tests/                     # Test suite
├── pyproject.toml            # Python build config
├── package.json              # Node.js/Claude Skill config
└── README.md
```

## Compatibility

### OpenClaw
- Skills defined in Markdown with frontmatter
- Load from `openclaw_skills/` directory
- Follows OpenClaw skill specification

### Claude Skill
- `skill.json` configuration
- JavaScript/TypeScript adapter
- Integrated with Claude Code

### OpenCode
- Python package (`office_skill`)
- Direct import and usage
- Programmatic API

## Development

### Setup Development Environment
```bash
git clone <repository>
cd office-skill
pip install -e .[dev]
pre-commit install
```

### Running Tests
```bash
pytest tests/ -v
pytest --cov=office_skill tests/
```

### Code Quality
```bash
ruff check src/
black src/
mypy src/
```

### Building
```bash
python -m build
pip install dist/office_skill-*.whl
```

## Examples

Check the `examples/` directory for:
- Basic document creation
- Financial modeling
- Presentation generation
- Batch processing

## Dependencies

- **Primary**: `cli-anything-libreoffice` (CLI interface to LibreOffice)
- **System**: LibreOffice (document rendering backend)
- **Python**: Python 3.8+

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make changes with tests
4. Run test suite
5. Submit pull request

## License

MIT License - see LICENSE file for details.

## Acknowledgments

- Built upon `cli-anything-libreoffice` project
- Inspired by OpenClaw office skills methodology
- Compatible with Claude Code skill system