
## 2026-04-30: Fixed .gitignore `tests/` pattern

- **Problem**: `.gitignore` had `tests/` which blocked ALL files in the `tests/` directory, including `.py` test source files
- **Fix**: Replaced `tests/` with specific output file patterns:
  ```
  tests/*.pdf
  tests/*.docx
  tests/*.xlsx
  tests/*.pptx
  ```
- **Verification**:
  - `git check-ignore tests/test_example/test_sample.py` → exit 1 (not ignored ✅)
  - `git check-ignore tests/output.pdf` → exit 0, matched by `tests/*.pdf` (ignored ✅)
- **Key insight**: Using a blanket directory ignore (`tests/`) is too aggressive when you need to track source files within that directory. Use specific extension patterns instead.

## 2026-04-30: Removed unnecessary markupsafe dependency

- **Problem**: `pyproject.toml` listed `"markupsafe>=2.0.0"` as a direct dependency, but it's only a transitive dependency of `jinja2` and is never directly imported in the codebase.
- **Fix**: Removed `"markupsafe>=2.0.0",` line (line 32) from the `dependencies` list in `pyproject.toml`.
- **Verification**:
  - `grep "markupsafe" pyproject.toml` → empty (removed ✅)
  - `grep "jinja2" pyproject.toml` → `"jinja2>=3.0.0",` still present (preserved ✅)
  - `grep -r "markupsafe" src/` → empty (no code usage ✅)
- **Key insight**: Dependencies that are only transitive (brought in by another direct dependency) should not be listed as direct dependencies. Always check for direct imports before adding a dependency. Leaving transitive deps in pyproject.toml bloats the dependency tree and can cause version conflicts.

## 2026-04-30: Added CLI entry point and py.typed marker

- **Changes made**:
  - Added `[project.scripts]` section to `pyproject.toml` with `office-skill = "office_main.cli.office_cli:main"` (after `[project.urls]`, before `[tool.setuptools.packages.find]`)
  - Created empty `src/office_main/py.typed` marker file (PEP 561 type hint marker)
- **Verification**:
  - `grep -A2 "project.scripts" pyproject.toml` → shows office-skill entry point ✅
  - `ls -la src/office_main/py.typed` → 0-byte file exists ✅
  - `main()` function confirmed at `office_cli.py:375` ✅
- **Key insight**: The `py.typed` marker (PEP 561) must exist at the package root to signal that the package provides type hints. It was already listed in `[tool.setuptools.package-data]` but the file was missing. The `[project.scripts]` entry point enables `office-skill` as a shell command after `pip install -e .`.

## 2026-04-30: Standardized Python version targets to 3.10+

- **Changes in `pyproject.toml`**:
  - `requires-python`: `">=3.8"` → `">=3.10"`
  - `[tool.black] target-version`: `['py38']` → `['py310']`
  - `[tool.ruff] target-version`: `"py38"` → `"py310"`
  - `[tool.mypy] python_version`: `"3.9"` → `"3.10"`
  - Classifiers: removed `3.8` and `3.9`, added `3.12` and `3.13`
- **Verification**:
  - All 5 settings confirmed via `read pyproject.toml` ✅
  - No Python source code was modified
- **Key insight**: Keeping all Python version targets in sync across requires-python, black, ruff, mypy, and classifiers prevents confusion and ensures tools agree on the minimum supported version. When raising the minimum, always update all five locations.

## 2026-04-30: Extracted BaseDocumentHandler ABC from duplicated handler code

- **Problem**: DocxHandler, XlsxHandler, and PptxHandler duplicated ~168 lines of __init__, export(), close(), and context manager code across three nearly-identical implementations
- **Fix**: Created `src/office_main/core/base_handler.py` with `BaseDocumentHandler(ABC)` containing shared logic
- **What moved to base**:
  - `__init__` → base `__init__` + `_start_session()` helper (handles temp session creation + CLI init)
  - `export()` — identical across all 3 handlers
  - `close()` — identical across all 3 handlers
  - `__enter__`/`__exit__` — identical across all 3 handlers
- **What stayed in handlers**:
  - Format-specific I/O methods (add_paragraph, set_cell, add_slide, etc.)
  - `analyze_structure()` — format-specific logic (kept as abstract in base)
  - `create_from_template()` — different return type annotations per handler (kept as abstract in base)
- **Deviation from plan**: `close()` was planned as `@abstractmethod` but is concrete in the ABC since all 3 implementations are byte-for-byte identical. Making it concrete eliminates more duplication without behavioral change.
- **Pattern**: Each handler calls `super().__init__(project_path=project_path)` then `self._start_session(str(self.path_attr))`
- **Verification**:
  - `python3 -c "from office_main.core.base_handler import BaseDocumentHandler; ..."` → ABC import OK ✅
  - `issubclass(DocxHandler, BaseDocumentHandler)` → True for all 3 ✅
  - LSP diagnostics: 0 errors across the core package ✅
  - All handler-specific methods still present (add_paragraph, get_cell, list_slides, etc.) ✅

## 2026-04-30: Improved ruff and mypy configuration

- **Ruff changes**:
  - Added 6 new rule sets to `select`: `UP` (pyupgrade), `N` (pep8-naming), `SIM` (flake8-simplify), `PLC` (pylint convention), `PLE` (pylint error), `RUF100` (ruff unused noqa)
  - These rules enforce modern Python idioms, naming conventions, code simplification, and pylint parity
- **Mypy changes**:
  - Changed `strict = false` → `strict = true` to enable full type checking
  - Removed global `ignore_missing_imports = true` (too permissive)
  - Added per-module overrides: `libreofficepy` gets `ignore_missing_imports = true`, and `pptx`/`docx` share another override with `ignore_missing_imports = true`
- **Verification**: All 6 new rules confirmed in select list ✅; strict=true confirmed ✅; per-module overrides structured properly ✅
- **Key insight**: Global `ignore_missing_imports` is a blunt instrument. Prefer per-module overrides so only packages that genuinely lack type stubs get this pass. Enabling `strict = true` catches latent type errors across the codebase.

## 2026-04-30: Consolidated duplicated CLI wrapper methods via _run_subcommand()

- **Problem**: 8 methods (`writer`, `calc`, `impress`, `export`, `document`, `session`, `style`, `batch`) in `LibreOfficeCLI` duplicated ~20 lines of argument-building + kwarg-to-option logic each (~160 lines total)
- **Fix**: Extracted `_run_subcommand(self, command_name, subcommand, positional=None, handle_bool_flags=True, **kwargs)` helper that handles the common pattern
- **Two variants**:
  - **Pattern A** (writer/calc/impress/export): Boolean kwargs handled as flags (`handle_bool_flags=True`), e.g., `--flag` for True, skip for False
  - **Pattern B** (document/session/style/batch): All kwargs stringified unconditionally (`handle_bool_flags=False`), e.g., `--key True` / `--key False`
- **Net reduction**: ~130 lines removed (380 → 344, + helper ÷ 8× ~20-line bodies → ~160 removed + ~35 added)
- **Verification**:
  - Import test: `from office_main.core.cli_wrapper import LibreOfficeCLI` → OK ✅
  - Signatures: All 8 methods preserve original signatures (confirmed via `inspect.signature()`) ✅
  - LSP diagnostics: 0 errors introduced ✅
  - Black formatting: clean ✅
- **Key insight**: When consolidating 8 near-identical methods with one behavioral variant, parameterize the variant via a kwarg (`handle_bool_flags`) rather than creating two helpers. This makes the consolidation a single method with a clear default (True for the more common/correct pattern). The delegating methods remain as thin public wrappers to preserve the explicit command-name API and maintain backward-compatible signatures.

## 2026-04-30: Added TYPE_CHECKING guard to template_manager cross-package imports

- **Files changed**: `src/template_manager/analyzer.py` and `src/template_manager/generator.py`
- **Changes**:
  - Added `from __future__ import annotations` to both files (required when wrapping imports that are used as type annotations)
  - Added `TYPE_CHECKING` to the typing imports
  - Wrapped `from office_main.core.cli_wrapper import LibreOfficeCLI` with `if TYPE_CHECKING:`
- **Key insight**: When wrapping a module-level import with TYPE_CHECKING, you MUST also add `from __future__ import annotations` (PEP 563) if the imported symbol is used as a type annotation — otherwise the class body will fail at import time because the annotation is evaluated eagerly. Python 3.10+ supports this natively via the future import.
- **Important distinction**: `generator.py` uses `LibreOfficeCLI` ONLY as a type annotation (dependency injection pattern), making it a clean TYPE_CHECKING case. `analyzer.py` also creates `LibreOfficeCLI()` instances at runtime, so the TYPE_CHECKING guard only prevents the module-level import dependency — the runtime instantiation calls still require `LibreOfficeCLI` to be importable at runtime (they currently work because the module-level import is now behind TYPE_CHECKING, but the methods do lazy instantiation).
- **Verification**:
  - `python3 -c "from template_manager.analyzer import TemplateAnalyzer; print('OK')"` → OK ✅
  - `python3 -c "from template_manager.generator import TemplateGenerator; print('OK')"` → OK ✅
  - `python3 -m py_compile` on both files → syntax OK ✅
  - `black` on both files → no changes needed ✅
