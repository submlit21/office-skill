# Optimization Fixes for office-skill

## TL;DR

> **Quick Summary**: Fix critical configuration issues, reduce code duplication (~280 lines), and improve code quality tooling in the office-skill project.
>
> **Deliverables**:
> - Fixed .gitignore allowing test files
> - Added CLI entry point for pip installation
> - Removed unnecessary markupsafe dependency
> - Standardized Python version targets (3.10+)
> - Extracted BaseDocumentHandler ABC reducing duplication by ~280 lines
> - Consolidated CLI wrapper methods
> - Fixed cross-package import dependency
> - Improved ruff/mypy configuration
>
> **Estimated Effort**: Medium
> **Parallel Execution**: YES - 3 waves
> **Critical Path**: Task 1 → Task 5 → Task 7 → Task 9 → Task 10

---

## Context

### Original Request
"修复问题" (Fix issues) - referring to optimization issues identified in the office-skill project analysis.

### Interview Summary
**Key Discussions**:
- Test strategy: User chose NO automated tests (pytest configured but no test files)
- Scope: Focus on identified optimization issues only

**Research Findings** (from Metis review):
- Handler duplication confirmed: identical `__init__`, `close()`, `export()`, `create_from_template()` across 3 handlers
- CLI duplication confirmed: 8 methods in `cli_wrapper.py` plus 3 repeated type-dispatch chains
- .gitignore blocking tests confirmed at line 37
- markupsafe unnecessary confirmed: zero imports, transitive dep of jinja2
- Python version inconsistency confirmed: black/ruff=3.8, mypy=3.9, requires-python>=3.8
- Cross-package dependency confirmed: `template_manager` imports from `office_main.core.cli_wrapper`

### Metis Review
**Identified Gaps** (addressed):
- Missing CLI entry point: Added as Task 3 (console_scripts in pyproject.toml)
- `create_from_template()` stubs: Decision needed - will be addressed in ABC extraction
- `py.typed` marker missing: Added as Task 6
- Python 3.8 is EOL: Addressing by bumping to 3.10+
- TDD contradiction: Resolved - no tests per user decision, manual verification instead

---

## Work Objectives

### Core Objective
Fix critical configuration issues and reduce code duplication in the office-skill project to improve maintainability and usability.

### Concrete Deliverables
- `.gitignore` allows test files to be committed
- `pyproject.toml` has CLI entry point, consistent Python 3.10+ targets, improved ruff/mypy config
- `markupsafe` dependency removed
- `BaseDocumentHandler` ABC extracted and all handlers inherit from it
- CLI wrapper duplication consolidated
- Cross-package import uses TYPE_CHECKING guard

### Definition of Done
- [ ] `black src/ --check` passes
- [ ] `ruff check src/` passes with new rules
- [ ] `mypy src/` passes (or has documented ignores)
- [ ] `pip install -e .` succeeds
- [ ] `office-skill --help` works after installation
- [ ] All handler imports work without error

### Must Have
- All CLI commands produce identical behavior before/after refactoring
- Handler method signatures remain unchanged
- Import paths remain backward-compatible
- Each commit is independently revertable

### Must NOT Have (Guardrails)
- No new features or capabilities
- No automated tests (user decision)
- No documentation updates (separate task)
- No migration to Typer/click (just consolidate within argparse)
- No full codebase restructuring (minimal fix for cross-package deps)

---

## Verification Strategy (MANDATORY)

> **ZERO HUMAN INTERVENTION** - ALL verification is agent-executed. No exceptions.

### Test Decision
- **Infrastructure exists**: YES (pytest configured in pyproject.toml)
- **Automated tests**: NO (user decision)
- **Framework**: none

### QA Policy
Every task MUST include agent-executed QA scenarios.
Evidence saved to `.sisyphus/evidence/task-{N}-{scenario-slug}.{ext}`.

- **Library/Module**: Use Bash (python REPL) - Import, call functions, compare output
- **CLI**: Use Bash - Run commands, validate output

---

## Execution Strategy

### Parallel Execution Waves

```
Wave 1 (Foundation - config fixes, no code changes):
├── Task 1: Fix .gitignore to allow test files [quick]
├── Task 2: Remove unnecessary markupsafe dependency [quick]
├── Task 3: Add CLI entry point and py.typed marker [quick]
├── Task 4: Standardize Python version targets to 3.10+ [quick]
└── Task 5: Improve ruff and mypy configuration [quick]

Wave 2 (Core refactoring - after Wave 1):
├── Task 6: Extract BaseDocumentHandler ABC [deep]
├── Task 7: Consolidate CLI wrapper duplication [deep]
└── Task 8: Fix cross-package dependency with TYPE_CHECKING [quick]

Wave 3 (Cleanup - after Wave 2):
├── Task 9: Apply ruff auto-fixes for new rules [quick]
└── Task 10: Apply Black formatting [quick]

Wave FINAL (After ALL tasks — parallel reviews):
├── Task F1: Plan compliance audit (oracle)
├── Task F2: Code quality review (unspecified-high)
├── Task F3: Manual QA verification (unspecified-high)
└── Task F4: Scope fidelity check (deep)
-> Present results -> Get explicit user okay

Critical Path: Task 1 → Task 5 → Task 6 → Task 7 → Task 9 → F1-F4 → user okay
Parallel Speedup: ~50% faster than sequential
Max Concurrent: 5 (Wave 1)
```

### Dependency Matrix

| Task | Depends On | Blocks |
|------|------------|--------|
| 1 | - | 6, 7, 8 |
| 2 | - | 9, 10 |
| 3 | - | 9, 10 |
| 4 | - | 5, 6, 7 |
| 5 | 4 | 6, 7, 9, 10 |
| 6 | 1, 4, 5 | 9, 10 |
| 7 | 1, 4, 5 | 9, 10 |
| 8 | 1 | 9, 10 |
| 9 | 2, 3, 5, 6, 7, 8 | 10 |
| 10 | 9 | F1-F4 |

### Agent Dispatch Summary

- **Wave 1**: 5 tasks - T1-T5 → `quick`
- **Wave 2**: 3 tasks - T6 → `deep`, T7 → `deep`, T8 → `quick`
- **Wave 3**: 2 tasks - T9 → `quick`, T10 → `quick`
- **FINAL**: 4 tasks - F1 → `oracle`, F2 → `unspecified-high`, F3 → `unspecified-high`, F4 → `deep`

---

## TODOs

- [x] 1. Fix .gitignore to allow test source files

  **What to do**:
  - Edit `.gitignore` line 37: change `tests/` to allow Python files
  - Add specific patterns: `tests/*.pdf`, `tests/*.docx`, `tests/*.xlsx`, `tests/*.pptx`
  - Keep `__pycache__/` and `*.pyc` patterns for test directory

  **Must NOT do**:
  - Do not remove the `tests/` entry entirely (need to ignore test output files)
  - Do not modify any other .gitignore patterns

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Single file, minimal change, clear specification
  - **Skills**: []
    - No skills needed for this simple config change
  - **Skills Evaluated but Omitted**:
    - `test-driven-development`: No tests being written

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2, 3, 4, 5)
  - **Blocks**: Tasks 6, 7, 8 (can create test files after this)
  - **Blocked By**: None (can start immediately)

  **References**:

  **Pattern References**:
  - `.gitignore:37` - Current line blocking tests: `tests/`

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Git documentation on .gitignore patterns

  **WHY Each Reference Matters**:
  - `.gitignore:37`: This is the exact line that needs modification. Currently blocks all test files from version control.

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify .gitignore allows Python test files
    Tool: Bash
    Preconditions: .gitignore file exists
    Steps:
      1. Create test directory: mkdir -p tests/test_example
      2. Create test file: touch tests/test_example/test_sample.py
      3. Run: git status tests/
      4. Assert: test_sample.py appears in untracked files
    Expected Result: Python test files are now trackable by git
    Failure Indicators: test_sample.py does not appear in git status output
    Evidence: .sisyphus/evidence/task-1-gitignore-test-files.txt

  Scenario: Verify .gitignore still blocks test output files
    Tool: Bash
    Preconditions: .gitignore file exists
    Steps:
      1. Create test output: touch tests/output.pdf
      2. Run: git status tests/
      3. Assert: output.pdf does NOT appear in untracked files
    Expected Result: Test output files remain ignored
    Failure Indicators: output.pdf appears in git status output
    Evidence: .sisyphus/evidence/task-1-gitignore-output-blocked.txt
  ```

  **Commit**: YES
  - Message: `fix(gitignore): allow test source files`
  - Files: `.gitignore`
  - Pre-commit: None

- [ ] 2. Remove unnecessary markupsafe dependency

  **What to do**:
  - Edit `pyproject.toml`: remove `markupsafe>=2.0.0` from dependencies list (line 23)
  - Verify markupsafe is not imported anywhere in the codebase

  **Must NOT do**:
  - Do not remove jinja2 (markupsafe is its transitive dependency, will be installed automatically)
  - Do not modify any other dependencies

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Single line removal from config file
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 3, 4, 5)
  - **Blocks**: Tasks 9, 10 (clean dependency list before formatting)
  - **Blocked By**: None (can start immediately)

  **References**:

  **Pattern References**:
  - `pyproject.toml:23` - Current markupsafe dependency line

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - None

  **WHY Each Reference Matters**:
  - `pyproject.toml:23`: This is the exact line to remove. markupsafe is only a transitive dependency of jinja2.

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify markupsafe removed from dependencies
    Tool: Bash
    Preconditions: pyproject.toml exists
    Steps:
      1. Run: grep -n "markupsafe" pyproject.toml
      2. Assert: No output (markupsafe not found)
    Expected Result: markupsafe is no longer listed as a direct dependency
    Failure Indicators: grep returns a line containing markupsafe
    Evidence: .sisyphus/evidence/task-2-markupsafe-removed.txt

  Scenario: Verify jinja2 still present
    Tool: Bash
    Preconditions: pyproject.toml exists
    Steps:
      1. Run: grep -n "jinja2" pyproject.toml
      2. Assert: Returns line with jinja2 dependency
    Expected Result: jinja2 dependency remains
    Failure Indicators: grep returns no output
    Evidence: .sisyphus/evidence/task-2-jinja2-present.txt
  ```

  **Commit**: YES
  - Message: `fix(deps): remove unnecessary markupsafe dependency`
  - Files: `pyproject.toml`
  - Pre-commit: None

- [ ] 3. Add CLI entry point and py.typed marker

  **What to do**:
  - Add `[project.scripts]` section to `pyproject.toml` with `office-skill = "office_main.cli.office_cli:main"`
  - Create empty `src/office_main/py.typed` marker file

  **Must NOT do**:
  - Do not change the CLI implementation, only add entry point
  - Do not add any other script entry points

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Config change + empty file creation
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 2, 4, 5)
  - **Blocks**: Tasks 9, 10
  - **Blocked By**: None (can start immediately)

  **References**:

  **Pattern References**:
  - `pyproject.toml:44` - Declares `py.typed` in package-data but file doesn't exist
  - `src/office_main/cli/office_cli.py` - Contains the CLI main function

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - PEP 561 - py.typed marker specification
  - Python packaging documentation - console_scripts entry points

  **WHY Each Reference Matters**:
  - `pyproject.toml:44`: Shows the intended py.typed marker location
  - `src/office_main/cli/office_cli.py`: Contains the entry point function to reference

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify CLI entry point works after install
    Tool: Bash
    Preconditions: Python environment available
    Steps:
      1. Run: pip install -e .
      2. Run: office-skill --help
      3. Assert: Output shows help with subcommands
    Expected Result: CLI entry point works after pip install
    Failure Indicators: Command not found or error
    Evidence: .sisyphus/evidence/task-3-cli-entry-point.txt

  Scenario: Verify py.typed marker exists
    Tool: Bash
    Preconditions: None
    Steps:
      1. Run: ls -la src/office_main/py.typed
      2. Assert: File exists
    Expected Result: py.typed marker file is present
    Failure Indicators: File not found error
    Evidence: .sisyphus/evidence/task-3-py-typed.txt
  ```

  **Commit**: YES
  - Message: `feat(packaging): add CLI entry point and py.typed marker`
  - Files: `pyproject.toml`, `src/office_main/py.typed`
  - Pre-commit: None

- [ ] 4. Standardize Python version targets to 3.10+

  **What to do**:
  - Update `pyproject.toml` line 8: change `requires-python = ">=3.8"` to `requires-python = ">=3.10"`
  - Update `pyproject.toml` line 57: change `target-version = "py38"` to `target-version = "py310"`
  - Update `pyproject.toml` line 78: change `target-version = "py38"` to `target-version = "py310"`
  - Update `pyproject.toml` line 88: change `python_version = "3.9"` to `python_version = "3.10"`
  - Update Python classifiers to include 3.12, 3.13 and remove 3.8, 3.9

  **Must NOT do**:
  - Do not change any code to use 3.10+ features yet (separate task)
  - Do not modify any other configuration

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Multiple config changes but all in same file with clear specifications
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 2, 3, 5)
  - **Blocks**: Tasks 5, 6, 7 (need consistent Python version first)
  - **Blocked By**: None (can start immediately)

  **References**:

  **Pattern References**:
  - `pyproject.toml:8` - requires-python setting
  - `pyproject.toml:57` - black target-version
  - `pyproject.toml:78` - ruff target-version
  - `pyproject.toml:88` - mypy python_version

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Python version support schedule

  **WHY Each Reference Matters**:
  - All four lines need to be synchronized to the same Python version target

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify Python version targets are consistent
    Tool: Bash
    Preconditions: pyproject.toml exists
    Steps:
      1. Run: grep -E "requires-python|target-version|python_version" pyproject.toml
      2. Assert: All values show 3.10 or py310
    Expected Result: All Python version settings are consistent at 3.10
    Failure Indicators: Any line shows 3.8 or 3.9
    Evidence: .sisyphus/evidence/task-4-python-versions.txt

  Scenario: Verify classifiers updated
    Tool: Bash
    Preconditions: pyproject.toml exists
    Steps:
      1. Run: grep "Programming Language :: Python :: 3" pyproject.toml
      2. Assert: Shows 3.10, 3.11, 3.12, 3.13 but not 3.8 or 3.9
    Expected Result: Classifiers reflect supported Python versions
    Failure Indicators: Old versions present or new versions missing
    Evidence: .sisyphus/evidence/task-4-classifiers.txt
  ```

  **Commit**: YES
  - Message: `fix(config): standardize Python version to 3.10+`
  - Files: `pyproject.toml`
  - Pre-commit: None

- [ ] 5. Improve ruff and mypy configuration

  **What to do**:
  - Update `[tool.ruff.lint]` in `pyproject.toml` line 79-81: add `"UP"`, `"N"`, `"SIM"`, `"PLC"`, `"PLE"`, `"RUF100"` to select list
  - Update `[tool.mypy]` in `pyproject.toml`: change `strict = false` to `strict = true`
  - Add per-module mypy overrides for third-party packages (libreofficepy, pptx, docx)

  **Must NOT do**:
  - Do not run ruff fix yet (separate task)
  - Do not fix any existing type errors yet

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Config file changes with clear specifications
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 2, 3, 4)
  - **Blocks**: Tasks 6, 7, 9, 10
  - **Blocked By**: Task 4 (need Python version standardized first)

  **References**:

  **Pattern References**:
  - `pyproject.toml:79-81` - Current ruff select rules
  - `pyproject.toml:85-89` - Current mypy configuration

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Ruff documentation: rule codes
  - mypy documentation: strict mode

  **WHY Each Reference Matters**:
  - `pyproject.toml:79-81`: Current ruff rules need expansion
  - `pyproject.toml:85-89`: Current mypy config is too permissive

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify ruff rules expanded
    Tool: Bash
    Preconditions: pyproject.toml exists
    Steps:
      1. Run: grep -A 10 "\[tool.ruff.lint\]" pyproject.toml
      2. Assert: Output contains "UP", "N", "SIM", "PLC", "PLE", "RUF100"
    Expected Result: All new ruff rules are in the select list
    Failure Indicators: Any of the specified rules missing
    Evidence: .sisyphus/evidence/task-5-ruff-rules.txt

  Scenario: Verify mypy strict mode enabled
    Tool: Bash
    Preconditions: pyproject.toml exists
    Steps:
      1. Run: grep "strict = true" pyproject.toml
      2. Assert: Output shows strict = true
    Expected Result: mypy strict mode is enabled
    Failure Indicators: strict = false or no output
    Evidence: .sisyphus/evidence/task-5-mypy-strict.txt
  ```

  **Commit**: YES
  - Message: `fix(config): improve ruff and mypy configuration`
  - Files: `pyproject.toml`
  - Pre-commit: None

- [ ] 6. Extract BaseDocumentHandler ABC

  **What to do**:
  - Create new file `src/office_main/core/base_handler.py` with abstract base class
  - Define ABC with abstract methods: `close()`, `export()`, `__enter__()`, `__exit__()`
  - Move shared `__init__` logic to base class
  - Update `DocxHandler`, `XlsxHandler`, `PptxHandler` to inherit from `BaseDocumentHandler`
  - Update `src/office_main/core/__init__.py` to export `BaseDocumentHandler`
  - Remove duplicated code from each handler

  **Must NOT do**:
  - Do not change method signatures
  - Do not change return types
  - Do not modify the handler's public API

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: Complex refactoring requiring understanding of inheritance and ABC patterns
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - `test-driven-development`: No tests being written

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 7, 8)
  - **Blocks**: Tasks 9, 10
  - **Blocked By**: Tasks 1, 4, 5 (need config fixes first)

  **References**:

  **Pattern References**:
  - `src/office_main/core/docx_handler.py` - Handler implementation to refactor
  - `src/office_main/core/xlsx_handler.py` - Handler implementation to refactor
  - `src/office_main/core/pptx_handler.py` - Handler implementation to refactor
  - `src/office_main/core/__init__.py` - Module exports to update

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Python ABC documentation

  **WHY Each Reference Matters**:
  - All three handlers have identical `__init__`, `close()`, `export()` implementations that should be extracted
  - `__init__.py` needs to export the new base class

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify BaseDocumentHandler can be imported
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "from office_main.core.base_handler import BaseDocumentHandler; print('OK')"
      2. Assert: Output shows "OK"
    Expected Result: BaseDocumentHandler is importable
    Failure Indicators: ImportError
    Evidence: .sisyphus/evidence/task-6-import-abc.txt

  Scenario: Verify handlers inherit from ABC
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "from office_main.core import DocxHandler, XlsxHandler, PptxHandler, BaseDocumentHandler; print(issubclass(DocxHandler, BaseDocumentHandler), issubclass(XlsxHandler, BaseDocumentHandler), issubclass(PptxHandler, BaseDocumentHandler))"
      2. Assert: Output shows "True True True"
    Expected Result: All handlers are subclasses of BaseDocumentHandler
    Failure Indicators: Any False in output
    Evidence: .sisyphus/evidence/task-6-handler-inheritance.txt

  Scenario: Verify handlers still work
    Tool: Bash
    Preconditions: Package installed, test directory exists
    Steps:
      1. Run: python -c "from office_main.core import DocxHandler; h = DocxHandler('/tmp/test.docx'); h.close(); print('OK')"
      2. Assert: Output shows "OK"
    Expected Result: Handler instantiation and close works
    Failure Indicators: Any exception
    Evidence: .sisyphus/evidence/task-6-handler-works.txt
  ```

  **Commit**: YES
  - Message: `refactor(core): extract BaseDocumentHandler ABC`
  - Files: `src/office_main/core/base_handler.py`, `src/office_main/core/__init__.py`, `src/office_main/core/docx_handler.py`, `src/office_main/core/xlsx_handler.py`, `src/office_main/core/pptx_handler.py`
  - Pre-commit: None

- [ ] 7. Consolidate CLI wrapper duplication

  **What to do**:
  - In `src/office_main/core/cli_wrapper.py`, extract helper method `_run_subcommand()`
  - Consolidate the 8 duplicated methods (`writer`, `calc`, `impress`, `export`, `document`, `session`, `style`, `batch`) to use the helper
  - Keep method signatures unchanged for backward compatibility

  **Must NOT do**:
  - Do not change method signatures
  - Do not change return types
  - Do not migrate to different CLI framework

  **Recommended Agent Profile**:
  - **Category**: `deep`
    - Reason: Complex refactoring requiring understanding of subprocess patterns
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - `test-driven-development`: No tests being written

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 6, 8)
  - **Blocks**: Tasks 9, 10
  - **Blocked By**: Tasks 1, 4, 5 (need config fixes first)

  **References**:

  **Pattern References**:
  - `src/office_main/core/cli_wrapper.py:161-359` - Duplicated methods to consolidate

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - None

  **WHY Each Reference Matters**:
  - Lines 161-359 contain 8 methods with nearly identical subprocess invocation patterns

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify CLI wrapper methods still work
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "from office_main.core.cli_wrapper import LibreOfficeCLI; cli = LibreOfficeCLI(); print('OK')"
      2. Assert: Output shows "OK"
    Expected Result: LibreOfficeCLI can be instantiated
    Failure Indicators: ImportError or exception
    Evidence: .sisyphus/evidence/task-7-cli-wrapper.txt

  Scenario: Verify method signatures unchanged
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "import inspect; from office_main.core.cli_wrapper import LibreOfficeCLI; sig = inspect.signature(LibreOfficeCLI.writer); print(sig)"
      2. Assert: Output shows original signature
    Expected Result: Method signature preserved
    Failure Indicators: Different signature or error
    Evidence: .sisyphus/evidence/task-7-method-signatures.txt
  ```

  **Commit**: YES
  - Message: `refactor(cli): consolidate CLI wrapper duplication`
  - Files: `src/office_main/core/cli_wrapper.py`
  - Pre-commit: None

- [ ] 8. Fix cross-package dependency with TYPE_CHECKING

  **What to do**:
  - In `src/template_manager/analyzer.py`: wrap import of `office_main.core.cli_wrapper` with `TYPE_CHECKING` guard
  - In `src/template_manager/generator.py`: wrap import of `office_main.core.cli_wrapper` with `TYPE_CHECKING` guard
  - Add `from typing import TYPE_CHECKING` import at top of each file

  **Must NOT do**:
  - Do not restructure packages into common/office_main/template_manager
  - Do not change the actual import paths, only add guard

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Simple import guard pattern, well-defined change
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 2 (with Tasks 6, 7)
  - **Blocks**: Tasks 9, 10
  - **Blocked By**: Task 1 (need test files allowed first)

  **References**:

  **Pattern References**:
  - `src/template_manager/analyzer.py:13` - Current import to guard
  - `src/template_manager/generator.py:18` - Current import to guard

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Python TYPE_CHECKING documentation

  **WHY Each Reference Matters**:
  - Both files import from `office_main.core.cli_wrapper` creating cross-package dependency

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify template_manager imports work
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "from template_manager.analyzer import TemplateAnalyzer; print('OK')"
      2. Assert: Output shows "OK"
    Expected Result: TemplateAnalyzer is importable
    Failure Indicators: ImportError
    Evidence: .sisyphus/evidence/task-8-import-analyzer.txt

  Scenario: Verify TYPE_CHECKING guard present
    Tool: Bash
    Preconditions: Files exist
    Steps:
      1. Run: grep -n "TYPE_CHECKING" src/template_manager/analyzer.py src/template_manager/generator.py
      2. Assert: Both files show TYPE_CHECKING import and guard
    Expected Result: TYPE_CHECKING guard is implemented
    Failure Indicators: No TYPE_CHECKING found
    Evidence: .sisyphus/evidence/task-8-type-checking.txt
  ```

  **Commit**: YES
  - Message: `fix(deps): add TYPE_CHECKING guard for cross-package import`
  - Files: `src/template_manager/analyzer.py`, `src/template_manager/generator.py`
  - Pre-commit: None

- [ ] 9. Apply ruff auto-fixes for new rules

  **What to do**:
  - Run `ruff check --fix src/` to auto-fix violations from new rules (UP, N, SIM, PLC, PLE, RUF100)
  - Review and commit auto-fixed changes
  - Manually fix any remaining violations that can't be auto-fixed

  **Must NOT do**:
  - Do not add new rules beyond what was configured in Task 5
  - Do not refactor code beyond what ruff suggests

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Automated tool execution with clear specifications
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 3 (sequential after Tasks 6, 7, 8)
  - **Blocks**: Task 10
  - **Blocked By**: Tasks 2, 3, 5, 6, 7, 8 (need all refactoring done first)

  **References**:

  **Pattern References**:
  - `pyproject.toml:79-81` - Ruff rules configured in Task 5

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Ruff documentation: auto-fix

  **WHY Each Reference Matters**:
  - The ruff rules configured in Task 5 determine what auto-fixes will be applied

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify ruff passes after auto-fix
    Tool: Bash
    Preconditions: Ruff configured with new rules
    Steps:
      1. Run: ruff check src/
      2. Assert: Output shows "All checks passed!"
    Expected Result: No ruff violations remain
    Failure Indicators: Any violations reported
    Evidence: .sisyphus/evidence/task-9-ruff-check.txt

  Scenario: Verify code still imports correctly
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "from office_main.core import DocxHandler; print('OK')"
      2. Assert: Output shows "OK"
    Expected Result: Imports still work after ruff fixes
    Failure Indicators: ImportError
    Evidence: .sisyphus/evidence/task-9-imports-work.txt
  ```

  **Commit**: YES
  - Message: `style: apply ruff auto-fixes`
  - Files: various (determined by ruff)
  - Pre-commit: `ruff check src/`

- [ ] 10. Apply Black formatting

  **What to do**:
  - Run `black src/` to format all Python files
  - Verify formatting is consistent

  **Must NOT do**:
  - Do not change any logic, only formatting
  - Do not add or remove any code

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Automated formatting tool
  - **Skills**: []
    - No skills needed
  - **Skills Evaluated but Omitted**:
    - None applicable

  **Parallelization**:
  - **Can Run In Parallel**: NO
  - **Parallel Group**: Wave 3 (after Task 9)
  - **Blocks**: F1-F4
  - **Blocked By**: Task 9 (format after ruff fixes)

  **References**:

  **Pattern References**:
  - `pyproject.toml:57-60` - Black configuration

  **API/Type References**:
  - None

  **Test References**:
  - None

  **External References**:
  - Black documentation

  **WHY Each Reference Matters**:
  - Black configuration determines formatting style

  **Acceptance Criteria**:

  **QA Scenarios (MANDATORY):**

  ```
  Scenario: Verify black passes
    Tool: Bash
    Preconditions: Black configured
    Steps:
      1. Run: black src/ --check
      2. Assert: Output shows "All done!"
    Expected Result: All files are properly formatted
    Failure Indicators: Any files need formatting
    Evidence: .sisyphus/evidence/task-10-black-check.txt

  Scenario: Verify code still works after formatting
    Tool: Bash
    Preconditions: Package installed
    Steps:
      1. Run: python -c "from office_main.core import DocxHandler, XlsxHandler, PptxHandler; print('OK')"
      2. Assert: Output shows "OK"
    Expected Result: All imports work after formatting
    Failure Indicators: ImportError
    Evidence: .sisyphus/evidence/task-10-imports-work.txt
  ```

  **Commit**: YES
  - Message: `style: apply black formatting`
  - Files: various (determined by black)
  - Pre-commit: `black src/ --check`

---

## Final Verification Wave (MANDATORY — after ALL implementation tasks)

> 4 review agents run in PARALLEL. ALL must APPROVE. Present consolidated results to user and get explicit "okay" before completing.

- [ ] F1. **Plan Compliance Audit** — `oracle`
  Read the plan end-to-end. For each "Must Have": verify implementation exists (read file, run command). For each "Must NOT Have": search codebase for forbidden patterns — reject with file:line if found. Check evidence files exist in .sisyphus/evidence/. Compare deliverables against plan.
  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [ ] F2. **Code Quality Review** — `unspecified-high`
  Run `black src/ --check`, `ruff check src/`, `mypy src/`. Review all changed files for: `as any`/`@ts-ignore`, empty catches, console.log in prod, commented-out code, unused imports. Check AI slop: excessive comments, over-abstraction, generic names.
  Output: `Build [PASS/FAIL] | Lint [PASS/FAIL] | Files [N clean/N issues] | VERDICT`

- [ ] F3. **Real Manual QA** — `unspecified-high`
  Start from clean state. Execute EVERY QA scenario from EVERY task — follow exact steps, capture evidence. Test cross-task integration. Save to `.sisyphus/evidence/final-qa/`.
  Output: `Scenarios [N/N pass] | Integration [N/N] | VERDICT`

- [ ] F4. **Scope Fidelity Check** — `deep`
  For each task: read "What to do", read actual diff (git log/diff). Verify 1:1 — everything in spec was built (no missing), nothing beyond spec was built (no creep). Check "Must NOT do" compliance. Flag unaccounted changes.
  Output: `Tasks [N/N compliant] | Contamination [CLEAN/N issues] | Unaccounted [CLEAN/N files] | VERDICT`

---

## Commit Strategy

- **Task 1**: `fix(gitignore): allow test source files` - .gitignore
- **Task 2**: `fix(deps): remove unnecessary markupsafe dependency` - pyproject.toml
- **Task 3**: `feat(packaging): add CLI entry point and py.typed marker` - pyproject.toml, src/office_main/py.typed
- **Task 4**: `fix(config): standardize Python version to 3.10+` - pyproject.toml
- **Task 5**: `fix(config): improve ruff and mypy configuration` - pyproject.toml
- **Task 6**: `refactor(core): extract BaseDocumentHandler ABC` - src/office_main/core/base_handler.py, handlers/*.py
- **Task 7**: `refactor(cli): consolidate CLI wrapper duplication` - src/office_main/core/cli_wrapper.py
- **Task 8**: `fix(deps): add TYPE_CHECKING guard for cross-package import` - src/template_manager/analyzer.py, generator.py
- **Task 9**: `style: apply ruff auto-fixes` - various
- **Task 10**: `style: apply black formatting` - various

---

## Success Criteria

### Verification Commands
```bash
black src/ --check  # Expected: All done!
ruff check src/  # Expected: All checks passed!
mypy src/  # Expected: Success: no issues found
pip install -e .  # Expected: Successfully installed
office-skill --help  # Expected: Shows help with all subcommands
python -c "from office_main.core import DocxHandler, XlsxHandler, PptxHandler"  # Expected: No error
python -c "from office_main.core.base_handler import BaseDocumentHandler"  # Expected: No error
```

### Final Checklist
- [ ] All "Must Have" present
- [ ] All "Must NOT Have" absent
- [ ] All verification commands pass
- [ ] CLI entry point works after pip install
- [ ] All handlers inherit from BaseDocumentHandler
- [ ] No circular imports
