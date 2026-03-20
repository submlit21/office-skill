# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

This is the `office-skill` repository, a Claude skill for office document processing. The skill uses `cli-anything-libreoffice` to create, edit, modify, analyze, annotate, and execute tasks on Word (.doc/.docx), Excel (.xls/.xlsx), and PowerPoint (.ppt/.pptx/.pptx) documents.

## Development Environment

### Prerequisites
- `cli-anything-libreoffice` – the underlying CLI tool that must be installed and available in the PATH
- LibreOffice – required by `cli-anything-libreoffice` for document conversion and manipulation
- Runtime environment (Node.js, Python, or other) as defined by the skill's implementation language

### Initial Setup
1. Install the skill's runtime dependencies (check `package.json`, `requirements.txt`, or similar)
2. Ensure `cli-anything-libreoffice` is installed globally or locally and accessible
3. Verify LibreOffice is installed and functional

## Project Structure

- `src/` – Source code (language‑specific)
  - `handlers/` – Skill handlers for different document types
    - `docx/` – Word document handlers
    - `xlsx/` – Excel spreadsheet handlers
    - `pptx/` – PowerPoint presentation handlers
  - `lib/` – Shared utilities and CLI‑Anything integration
  - Entry point (e.g., `index.ts`, `main.py`, `skill.json`) – Skill manifest and registration
- `tests/` – Unit and integration tests
- `docx/`, `xlsx/`, `pptx/` – Example documents and templates (currently empty)

## Common Development Tasks

### Building the Skill
```bash
# If using Node.js
npm run build
# If using Python
python -m pip install -e .
# Consult the project's build documentation for other runtimes
```

### Running Tests
```bash
# If using Node.js
npm test
# If using Python
pytest
# For other runtimes, see the project's test runner
```

To run a single test file or module, use the appropriate test runner options.

### Linting & Formatting
```bash
# Use the project's configured linter and formatter (e.g., eslint, black, ruff)
npm run lint   # Node.js example
npm run format # Node.js example
```

### Local Skill Testing with Claude Code
1. Build the skill according to its runtime.
2. Use the skill locally via Claude Code's skill runner (see Claude Code documentation).

## Skill Architecture

The skill follows the Claude Skill pattern:
- Each document type (Word, Excel, PowerPoint) has its own handler module in `src/handlers/<format>/`
- Handlers expose actions (create, edit, modify, analyze, annotate, etc.) that call `cli-anything-libreoffice` commands
- Shared CLI‑Anything wrapper code lives in `src/lib/` (e.g., `cli-anything.ts`, `cli_anything.py`) to ensure consistent argument formatting and error handling
- The skill manifest (`skill.json`, `index.ts`, `skill.yaml`, etc.) declares available actions and their parameters

### Key Design Points
- **Stateless operations**: Each action should be idempotent and not rely on persistent state between calls.
- **File‑based I/O**: The skill reads input documents and writes output documents; temporary files should be cleaned up after use.
- **Error handling**: Fail gracefully with descriptive error messages when `cli-anything-libreoffice` or LibreOffice fails.
- **Cross‑platform support**: Ensure commands work on Linux, macOS, and Windows (where LibreOffice is available).

## Integration with `cli-anything-libreoffice`

The skill delegates all office‑document operations to `cli-anything-libreoffice`. Familiarity with its command‑line interface is essential for development. Typical usage:

```bash
cli-anything-libreoffice --input document.docx --output document.pdf --operation convert
```

The skill’s wrapper should map high‑level actions (e.g., “add watermark”, “extract tables”) to appropriate `cli-anything-libreoffice` command‑line arguments.

## Testing Strategy

- **Unit tests**: Mock `cli-anything-libreoffice` calls to verify argument generation and error handling.
- **Integration tests**: Require a working LibreOffice installation; run against sample documents in the `docx/`, `xlsx/`, `pptx/` directories.
- **End‑to‑end tests**: Use the skill via Claude Code’s skill runner to ensure the full workflow works.

## Notes for Contributors

- Add new sample documents to the appropriate `docx/`, `xlsx/`, or `pptx/` folder when creating tests.
- Keep the SKILL.md frontmatter (name, description, license) up‑to‑date.
- Follow the existing code style and naming conventions of the implementation language (e.g., camelCase for JavaScript/TypeScript, snake_case for Python).