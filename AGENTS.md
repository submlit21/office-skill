# AGENTS.md - Office Skill Codebase Guidelines

This file provides guidelines for agentic coding agents working in the `office-skill` repository.

## Project Overview

`office-skill` is an Open Claw Skill for Microsoft Office document processing that delegates to `cli-anything-libreoffice` as its backend. It provides Python interfaces for Word (DOCX), Excel (XLSX), and PowerPoint (PPTX) operations with template management, formula validation, and structured workflows.

## Build, Lint, and Test Commands

### Setup and Installation
```bash
pip install -e .                          # development mode
pip install cli-anything-libreoffice pyyaml # all dependencies
```

### Code Quality
```bash
black src/                                # format (line-length: 100)
ruff check src/                           # lint (E, W, F, I, B, C4)
mypy src/                                 # type checking
ruff check --fix src/                     # auto-fix lint issues
black --check src/                        # check formatting
```

### Testing
```bash
pytest tests/                             # all tests
pytest -v tests/                          # verbose
pytest tests/test_file.py                 # specific file
pytest tests/test_file.py::test_function  # specific function
pytest -k "pattern" tests/                # pattern matching
pytest -x tests/                          # stop on first failure
pytest --lf tests/                        # re-run last failures
pytest --co tests/                        # collect tests (dry-run)
pytest --cov=src tests/                   # coverage
pytest --cov=src --cov-report html tests/ # HTML coverage report
```

**Note**: `pyproject.toml` configures `testpaths = ["tests"]`. Place tests in `tests/` directory.

## Code Style Guidelines

### Import Organization
```python
import json
import os
import subprocess
import tempfile
from pathlib import Path
from typing import Dict, List, Any, Optional, Union

import yaml

from .cli_wrapper import LibreOfficeCLI
```

### Formatting
- **Line length**: 100 characters (black & ruff)
- **Black formatting**: Run before commits
- **Python version**: 3.8+
- **Indentation**: 4 spaces, no tabs
- **Quotes**: Double quotes (black enforces)

### Type Annotations
- Use type hints for all function parameters and return values
- Import from `typing`: `Dict`, `List`, `Any`, `Optional`, `Union`
- Use `Optional[T]` for nullable parameters
- Use `Union[T1, T2]` for multiple types
- Use `Any` sparingly; prefer specific types
- Use `# type: ignore` with explanatory comment

### Naming Conventions
- **Classes**: `PascalCase` (`DocxHandler`, `LibreOfficeCLI`)
- **Functions/Methods**: `snake_case` (`add_paragraph`, `validate_formulas`)
- **Variables**: `snake_case` (`document_path`, `project_path`)
- **Constants**: `UPPER_SNAKE_CASE` (`DEFAULT_TIMEOUT`, `MAX_RETRIES`)
- **Private methods**: `_leading_underscore` (`_run_command`)

### Error Handling
- Use try/except for external command execution
- Raise specific exceptions with descriptive messages
- Handle `cli-anything-libreoffice` command failures gracefully
- Validate file paths and document existence before operations
- Use `pathlib.Path` over `os.path`

### Docstrings
- Use triple‑double‑quoted docstrings for public modules, classes, functions
- Follow pattern: brief description, `Args:` and `Returns:` sections
- Keep concise but informative

### Document Handler Patterns
When extending `DocxHandler`, `XlsxHandler`, `PptxHandler`:
1. Map methods to underlying `cli-anything-libreoffice` commands via `LibreOfficeCLI`
2. Include proper error handling and session management
3. Follow existing method signatures and return types
4. Use `json_output=True` for CLI commands when possible

### Template Management
- Follow hierarchical naming: `domain.type.purpose.variant.version`
- Store templates in `templates/` with appropriate structure
- Include `metadata.json` and content file (`template.md` for DOCX, `template.yaml` for XLSX/PPTX)
- Use modular template system in `src/template_manager/`

## Key Design Principles

1. **Delegate to Backend**: Never reimplement Office rendering; use `cli-anything-libreoffice`
2. **Formula Validation**: Excel files must have zero formula errors; use `validate_formulas()` after modifications
3. **Template Consistency**: Follow `domain.type.purpose.variant.version` naming for all templates
4. **Document Integrity**: Use proper session management to avoid corruption
5. **Excel Financial Modeling Standards**:
   - Zero formula errors (#REF!, #DIV/0!, etc.) mandatory
   - Color coding: Blue for inputs, Black for formulas, Green for internal links, Red for external links
   - Number formatting: Years as text, currency with units, zeros as "-", percentages with one decimal
   - Negative numbers in parentheses `(123)` not minus `-123`
6. **DOCX Redlining**: For tracked changes, only mark changed text and preserve original RSIDs

## File Structure
```
src/
├── office_main/
│   ├── cli/              # CLI interface
│   └── core/             # Core document handlers
└── template_manager/     # Modular template system
templates/                # Hierarchical template storage
tests/                    # Test files and documents
skills/                   # Claude skill definitions
```

## Development Workflow & Agent Instructions

1. **Before changes**: Run `ruff check src/` and `mypy src/` to ensure clean baseline
2. **Implement**: Follow existing patterns and conventions
3. **After changes**: Run `black src/` to format code
4. **Test**: Run `pytest tests/` to verify functionality
5. **Verify quality**: Re-run `ruff check src/` and `mypy src/`
6. **ALWAYS** run lint and typecheck after making changes
7. **NEVER** commit without running `black src/` first
8. **USE** template system for document generation
9. **VALIDATE** Excel formulas after any spreadsheet modifications
10. **TEST** with actual Office documents in `tests/` directory
11. **VERIFY** file existence and size after document operations
12. **PREFER** `pathlib.Path` over `os.path` for path manipulation
13. **USE** `typing` annotations consistently

## Common Issues and Solutions

- **`cli-anything-libreoffice` not found**: Install with `pip install cli-anything-libreoffice`
- **LibreOffice conversion failures**: Check `libreoffice` or `pandoc` installation
- **Formula errors in Excel**: Always run `validate_formulas()`; use recalc.py for complex models
- **Template loading failures**: Verify naming convention and directory structure