
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
