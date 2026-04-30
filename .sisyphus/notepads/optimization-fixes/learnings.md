
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
