# PACKAGE KNOWLEDGE BASE - template_manager

**Generated:** Wed Apr 01 2026
**Purpose:** Modular template management system for Office documents

## OVERVIEW
Independent top-level package implementing **single-responsibility principle** with strict modular design.
Manages template storage, analysis, search, and document generation with dependency injection.
Hierarchical filesystem mirrors 5-part naming convention exactly.

## STRUCTURE
```
template_manager/
├── storage.py    # Filesystem operations + naming enforcement
├── analyzer.py   # Document structure analysis + markdown conversion
├── search.py     # Template search/filtering
├── generator.py  # Document generation from templates
└── __init__.py   # Package exports
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Naming convention enforcement | `storage.py` | Storage layer validates all template names |
| Dependency injection examples | `generator.py` | All dependencies passed via constructor |
| Document analysis | `analyzer.py` | Multiple conversion fallbacks implemented |
| Template search logic | `search.py` | Depends on Storage abstraction |
| Template generation | `generator.py` | Variable substitution with LibreOffice integration |

## DESIGN PRINCIPLES
- **Single Responsibility**: Each module has exactly one responsibility
- **Dependency Injection**: All dependencies passed via constructor (no internal instantiation)
- **Abstraction**: Modules depend on abstractions, not concrete implementations
- **Enforcement**: Naming rules enforced at storage boundary, not spread across codebase

## CONVENTIONS
- 5-part naming: See root AGENTS.md for convention description
- Strict enforcement: Storage layer rejects any template that doesn't match `domain.type.purpose.variant.version`
- Hierarchical storage: Directory structure mirrors naming components exactly
- All public methods require type hints
- Preserve original documents - never modify stored templates in-place

## ANTI-PATTERNS (THIS PACKAGE)
- Don't bypass storage layer to access templates directly - always go through Storage class
- Don't add naming checks elsewhere - all enforcement belongs in storage.py
- Don't hardcode dependencies inside classes - use dependency injection
- Don't mix responsibilities - if it doesn't fit in an existing module, add a new one
- Don't modify original templates - always work on copies during generation
- Don't repeat root AGENTS.md content here
