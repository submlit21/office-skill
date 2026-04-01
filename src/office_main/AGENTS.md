# PACKAGE KNOWLEDGE BASE - office_main

**Generated:** Wed Apr 01 2026
**Parent Project:** office-skill

## OVERVIEW
Top-level main package for the Office Skill. Provides entry points via CLI and exports core document handler classes.

Separates concerns between command-line interface and core document processing logic. All actual Office document manipulation delegates to the `cli-anything-libreoffice` backend through wrapper classes.

## STRUCTURE
```
office_main/
├── cli/                # Command-line interface entry point
│   └── office_cli.py   # Main CLI parser and command routing
├── core/               # Core document processing logic
│   ├── __init__.py     # Exports public handler classes
│   ├── cli_wrapper.py  # Wrapper for cli-anything-libreoffice
│   ├── docx_handler.py # Word document handler
│   ├── xlsx_handler.py # Excel workbook handler
│   ├── pptx_handler.py # PowerPoint presentation handler
│   └── template_handler.py # Template session management
├── __init__.py         # Package version and docstring
└── py.typed            # PEP 561 type hint marker
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Add new CLI command | `cli/office_cli.py` | Add subcommands and argument parsing here |
| Fix document handling | `core/[format]_handler.py` | Format-specific logic in individual handlers |
| Backend invocation | `core/cli_wrapper.py` | All LibreOffice CLI calls go through here |
| Template session | `core/template_handler.py` | Template loading and session management |
| Public API exports | `core/__init__.py` | Update exports when adding new handlers |

## CONVENTIONS (PACKAGE-SPECIFIC)
- All handlers follow symmetric interface pattern: `open()`, `save()`, `close()`
- CLI wrapper manages temporary working directories for backend sessions
- Handlers don't implement file format parsing - all I/O delegates to backend
- Keep CLI argument parsing separate from business logic
- Return structured dictionaries for all CLI operations when possible

## ANTI-PATTERNS (PACKAGE-SPECIFIC)
- Don't implement direct Office file parsing - always use backend via `LibreOfficeCLI`
- Don't add format-specific code in the CLI layer - keep CLI thin
- Don't leave temporary directories behind - always use `cleanup()` in wrapper
- Don't import `template_manager` into core __init__ - avoids circular imports
- Don't break the symmetric handler interface pattern across formats
