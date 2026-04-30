
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
