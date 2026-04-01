# PROJECT KNOWLEDGE BASE

**Generated:** Wed Apr 01 2026
**Commit:** 217c2c2
**Branch:** main

## OVERVIEW
Open Claw Skill for Microsoft Office document processing (DOCX/XLSX/PPTX) that delegates to `cli-anything-libreoffice` as backend. Provides Python interfaces with template management, formula validation, and structured workflows.

## STRUCTURE
```
src/
├── office_main/    # Core document handlers + CLI
├── template_manager/  # Modular template system
└── ...
templates/          # Hierarchical template storage
tests/              # Test files and documents
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Document handlers | src/office_main/core/ | DOCX/XLSX/PPTX handlers |
| CLI interface | src/office_main/cli/ | Command-line entry |
| Template system | src/template_manager/ | Template loading and management |
| Templates | templates/ | Stored in hierarchical structure |
| Tests | tests/ | Test documents and cases |

## CONVENTIONS
- Line length: 100 characters (enforced by black/ruff)
- Indentation: 4 spaces, double quotes
- Use `pathlib.Path` over `os.path`
- Type hints required for all function signatures
- Naming: PascalCase for classes, snake_case for functions/variables
- Templates follow `domain.type.purpose.variant.version` naming
- Always use `json_output=True` for CLI commands when possible

## ANTI-PATTERNS (THIS PROJECT)
- Never reimplement Office rendering - delegate to backend
- Don't commit without running `black src/` first
- Don't leave Excel formula errors - always validate
- Don't corrupt documents - use proper session management
- Don't repeat generic Python advice here

## UNIQUE STYLES
- **Excel Financial Standards**: Zero formula errors mandatory; color coding (blue=inputs, black=formulas, green=internal, red=external); specific number formatting
- **DOCX Redlining**: Only mark changed text, preserve original RSIDs
- **Symmetric document handlers**: Same pattern across DOCX/XLSX/PPTX

## COMMANDS
```bash
# Setup
pip install -e .
pip install cli-anything-libreoffice pyyaml

# Code quality
black src/
ruff check src/
mypy src/

# Testing
pytest tests/
pytest -v tests/
pytest tests/test_file.py
```

## NOTES
- `cli-anything-libreoffice` required - install via pip if missing
- Always validate Excel formulas after modifications with `validate_formulas()`
- Verify template naming convention when template loading fails
- Check libreoffice/pandoc installation for conversion failures
