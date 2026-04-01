# CORE SUBSPACKAGE KNOWLEDGE BASE - office_main/core

**Generated:** Wed Apr 01 2026
**Parent Package:** office_main

## OVERVIEW
Core document handler implementation following the **symmetric handler pattern**. One file per Office document format (DOCX/XLSX/PPTX), all sharing identical interface structure. All backend operations delegate to the shared `LibreOfficeCLI` wrapper.

## STRUCTURE
```
core/
├── __init__.py         # Exports all three format handler classes via __all__
├── cli_wrapper.py      # Shared wrapper for cli-anything-libreoffice backend
├── docx_handler.py     # Word document handler
├── xlsx_handler.py     # Excel workbook handler
├── pptx_handler.py     # PowerPoint presentation handler
└── template_handler.py # Template session management and validation
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Fix format-specific logic | `[format]_handler.py` | Each format in its own file |
| Backend invocation issues | `cli_wrapper.py` | All LibreOffice CLI calls go through here |
| Add new document format | Create `[fmt]_handler.py` following symmetric pattern | Update `__init__.py` exports |
| Template loading/validation | `template_handler.py` | Template initialization and session handling |
| Public API exports | `__init__.py` | Add new handlers to `__all__` list |

## CONVENTIONS (SYMMETRIC PATTERN)
- **Constructor**: All handlers accept `document_path` and optional `project_path` parameters
- **Context manager protocol**: All implement `__enter__`/`__exit__` for automatic cleanup
- **Required methods**: Every handler must implement:
  - `create_from_template()` - Initialize document from template
  - `analyze_structure()` - Parse document structure and return metadata
  - `export()` - Export document to target format
  - `close()` - Clean up resources and temporary files
- **Dependency**: All handlers depend on `LibreOfficeCLI` from `cli_wrapper.py`
- **Naming**: Format-specific classes: `DOCXHandler`, `XLSXHandler`, `PPTXHandler` (PascalCase)
- **Delegation**: All actual file operations delegate to backend - no direct parsing

## ANTI-PATTERNS (DIRECTORY-SPECIFIC)
- Don't break the symmetric interface pattern - every format must have same method signatures
- Don't implement direct document parsing - always delegate through `LibreOfficeCLI`
- Don't add format-specific code to `cli_wrapper.py` - keep it generic
- Don't forget to add new handlers to `__init__.py` `__all__` export list
- Don't leave temporary files behind - rely on context manager cleanup
- Don't duplicate method implementations across handlers - extract to shared wrapper
