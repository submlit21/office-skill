# AGENTS.md - Office Skill Development Guide

This document provides essential information for AI agents working on the office-skill project.

## Project Overview

The office-skill project is a Python package for office document processing (Word, Excel, PowerPoint) built on top of `cli-anything-libreoffice`. It provides:
- Python API for programmatic document manipulation
- Claude Skill integration
- OpenClaw compatibility
- Cross-platform support via LibreOffice backend

## Build, Test, and Lint Commands

### Development Setup
```bash
# Install with development dependencies
pip install -e .[dev]

# Install system dependencies (Ubuntu/Debian)
sudo apt install libreoffice
```

### Testing Commands
```bash
# Run all tests
pytest tests/ -v

# Run tests with coverage
pytest --cov=office_skill tests/

# Run a single test file
pytest tests/test_basic.py -v

# Run a single test class
pytest tests/test_basic.py::TestDocxHandler -v

# Run a single test method
pytest tests/test_basic.py::TestDocxHandler::test_handler_initialization -v

# Run integration tests (marked with @pytest.mark.integration)
pytest -m integration -v

# Skip tests requiring cli-anything-libreoffice
pytest -k "not CLI_AVAILABLE" -v
```

### Linting and Formatting
```bash
# Check code with ruff
ruff check src/

# Auto-fix ruff issues
ruff check --fix src/

# Format with black
black src/

# Type checking with mypy
mypy src/

# Run all quality checks
ruff check src/ && black --check src/ && mypy src/
```

### Building and Packaging
```bash
# Build package
python -m build

# Install from built package
pip install dist/office_skill-*.whl

# Clean build artifacts
rm -rf dist/ build/ src/office_skill.egg-info/
```

### Claude Skill Integration
```bash
# Install the skill in Claude Code
claude skills add ./path/to/office-skill

# Use office skill commands in Claude Code
claude office create-document --type docx --output report.docx
claude office analyze-document --input data.xlsx --type xlsx
```

## Code Style Guidelines

### Python Version
- Target Python 3.8+ (configured in pyproject.toml)
- Use type hints for all function signatures
- Support both positional and keyword arguments where appropriate

### Imports
```python
# Standard library imports first
import json
import subprocess
import tempfile
import os
from pathlib import Path
from typing import Dict, List, Any, Optional, Union

# Third-party imports
# (none beyond cli-anything-libreoffice)

# Local imports
from .cli_wrapper import LibreOfficeCLI
```

### Formatting
- Line length: 100 characters (configured in black/ruff)
- Use double quotes for docstrings, single quotes for strings
- Follow Black formatting rules
- Use trailing commas in multi-line collections

### Naming Conventions
- **Classes**: PascalCase (e.g., `DocxHandler`, `LibreOfficeCLI`)
- **Functions/Methods**: snake_case (e.g., `add_paragraph`, `analyze_structure`)
- **Variables**: snake_case (e.g., `document_path`, `json_output`)
- **Constants**: UPPER_SNAKE_CASE (e.g., `CLI_AVAILABLE`, `__version__`)
- **Private members**: leading underscore (e.g., `_run_command`, `_session_file`)

### Type Hints
- Always use type hints for function parameters and return values
- Use `Optional[T]` for nullable types
- Use `Union[T1, T2]` for multiple possible types
- Use `Dict[str, Any]` for JSON-like return values
- Import typing annotations from `typing` module

### Error Handling
- Use try/except blocks for external command execution
- Raise appropriate exceptions with descriptive messages
- Return `Dict[str, Any]` for command results with success/error status
- Use `Optional` return types for operations that may fail

### Documentation
- Include docstrings for all public classes and methods
- Use Google-style docstring format:
  ```python
  def add_paragraph(self, text: str, index: Optional[int] = None) -> Dict[str, Any]:
      """
      Add a paragraph to the document.

      Args:
          text: The text content
          index: Position to insert (append if None)

      Returns:
          Command result dictionary
      """
  ```
- Include `__init__.py` with module-level docstring
- Document exceptions that may be raised

### File Structure
```
src/office_skill/
├── __init__.py           # Package exports and version
├── cli_wrapper.py       # Low-level CLI interface
├── docx_handler.py      # Word document operations
├── xlsx_handler.py      # Excel spreadsheet operations
└── pptx_handler.py      # PowerPoint presentation operations
```

### Testing Conventions
- Test files in `tests/` directory
- Test class names: `Test*` (e.g., `TestDocxHandler`)
- Test method names: `test_*` (e.g., `test_handler_initialization`)
- Use `pytest.mark.skipif` for tests requiring external dependencies
- Use `tempfile` module for test file creation
- Clean up test files after execution

### Claude Skill Integration
- Claude Skill configuration in `src/claude_skill/skill.json`
- Python bridge implementation in `src/claude_skill/office_bridge.py`
- `skill.json` defines available actions and parameters
- Follows Claude Skill specification for action definitions

### Error Messages
- Use descriptive error messages
- Include relevant context (file paths, command arguments)
- Return structured error information in JSON responses
- Log errors at appropriate levels

### Performance Considerations
- Use `Path` objects for file path manipulation
- Cache expensive operations where appropriate
- Use context managers for resource cleanup
- Consider memory usage for large documents

## Project-Specific Patterns

### CLI Command Execution
```python
def _run_command(self, args: List[str]) -> Dict[str, Any]:
    """
    Run a cli-anything-libreoffice command and return parsed output.
    
    Returns parsed JSON output or raises exception on error
    """
    cmd = ["cli-anything-libreoffice"]
    if self.json_output:
        cmd.append("--json")
    # ... command execution logic
```

### Handler Class Pattern
```python
class DocxHandler:
    """Handler for Word document operations."""
    
    def __init__(self, document_path: str, project_path: Optional[str] = None):
        self.document_path = Path(document_path).absolute()
        self.cli = LibreOfficeCLI(project_path=project_path)
    
    def add_paragraph(self, text: str, index: Optional[int] = None) -> Dict[str, Any]:
        kwargs = {"text": text}
        if index is not None:
            kwargs["index"] = index
        return self.cli.writer("add-paragraph", **kwargs)
```

### Integration with External Tools
- **Primary dependency**: `cli-anything-libreoffice` must be installed and in PATH
- **System requirement**: LibreOffice for document rendering
- Install with: `pip install cli-anything-libreoffice`
- Handle missing dependencies gracefully in tests using `pytest.mark.skipif`

## Common Tasks for AI Agents

### Adding New Features
1. Add method to appropriate handler class
2. Implement CLI command execution via `_run_command`
3. Add type hints and docstring
4. Write tests in corresponding test file
5. Update `__init__.py` exports if needed
6. Run tests and linting

### Fixing Bugs
1. Reproduce issue with test case
2. Fix implementation
3. Ensure backward compatibility
4. Update tests if behavior changes
5. Run full test suite

### Refactoring
1. Maintain existing API surface
2. Update type hints and docstrings
3. Run tests to ensure no regressions
4. Update dependent code if needed

### Documentation Updates
1. Update docstrings for changed functionality
2. Update README.md if public API changes
3. Update examples if needed
4. Keep AGENTS.md up to date