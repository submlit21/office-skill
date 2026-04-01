# Hierarchical AGENTS.md Generation Plan

## TL;DR

> **Quick Summary**: Generate hierarchical AGENTS.md knowledge base files at root and key subdirectories, following the `/init-deep` command specification.
> 
> **Deliverables**:
> - AGENTS.md at project root (update existing with structured knowledge base format)
> - AGENTS.md at src/office_main/
> - AGENTS.md at src/office_main/core/
> - AGENTS.md at src/office_main/cli/
> - AGENTS.md at src/template_manager/
> - AGENTS.md at skills/
> 
> **Estimated Effort**: Short
> **Parallel Execution**: YES - 2 waves
> **Critical Path**: Root → Core+CLI+Template → Final review

---

## Context

### Original Request
`/init-deep` - Generate hierarchical AGENTS.md files with root + complexity-scored subdirectories. This is an update mode operation (modify existing + create new where warranted).

### Interview Summary
**Key Discussions**:
- Project already has a root AGENTS.md with basic guidelines
- Codebase is small-to-medium with 2913 lines of Python across 12 files
- Clear modular structure with two top-level packages under src/
- Empty tests/ and templates/ directories already created but not populated

**Research Findings**:
- Project follows modern Python src-layout with pyproject.toml configuration
- All tools (black, ruff, mypy, pytest) properly configured
- Codebase has consistent patterns: symmetric document handlers, modular template system
- No oversized files (largest is 499 lines, well under 500)
- Hidden conventions found: symmetric handler pattern, dependency injection in template manager, two-stage Jinja2 substitution

### Metis Review
*(Not yet applied - will happen after planning)*

---

## Work Objectives

### Core Objective
Generate a complete hierarchical knowledge base of AGENTS.md files that:
1. Starts with root AGENTS.md containing high-level project overview
2. Adds AGENTS.md at key subdirectories with directory-specific conventions
3. Never repeats parent content in child files
4. Captures all discovered patterns and anti-patterns
5. Follows the `/init-deep` specification strictly

### Concrete Deliverables
- [Update] `./AGENTS.md` - Root project knowledge base
- [Create] `./src/office_main/AGENTS.md` - Main package overview
- [Create] `./src/office_main/core/AGENTS.md` - Core document handlers
- [Create] `./src/office_main/cli/AGENTS.md` - CLI interface
- [Create] `./src/template_manager/AGENTS.md` - Template management system
- [Create] `./skills/AGENTS.md` - AI skill definitions

### Definition of Done
- [ ] All files follow size limits (root: 50-150 lines, subdirs: 30-80 lines)
- [ ] No duplicate content between parent and child
- [ ] No generic advice - only project-specific information
- [ ] All scoring-based locations have AGENTS.md created
- [ ] Telegraphic style maintained

### Must Have
- All content must be project-specific - no generic Python advice
- Child files must NOT repeat parent content
- Every file must have WHERE TO LOOK, CONVENTIONS, ANTI-PATTERNS
- Quality gates on line count must be respected

### Must NOT Have (Guardrails)
- No generic "how to write Python" content
- No repeating parent guidelines in child files
- No over-documenting directories that score <8
- No verbose explanations - keep it telegraphic

---

## Verification Strategy

> **ZERO HUMAN INTERVENTION** — ALL verification is agent-executed. No exceptions.

### Test Decision
- **Infrastructure exists**: YES (root already has it)
- **Automated tests**: None needed for documentation
- **Framework**: None - just markdown file validation
- **Agent-Executed QA**: ALWAYS (mandatory)

### QA Policy
Every generated file will have agent-executed QA scenarios:

1. **File validation**: Verify file exists, check line count against limits
2. **Content validation**: Verify no duplication from parent, all required sections present
3. **Path validation**: Verify all referenced file paths actually exist
4. **Anti-pattern check**: Verify no generic content, only project-specific advice

---

## Execution Strategy

### Parallel Execution Waves

> Maximize throughput by grouping independent tasks into parallel waves.

Wave 1 (Can start immediately - all independent):
├── Task 1: Generate ./AGENTS.md (root update) [quick]
├── Task 2: Generate ./src/office_main/AGENTS.md [quick]
├── Task 3: Generate ./src/office_main/core/AGENTS.md [quick]
├── Task 4: Generate ./src/office_main/cli/AGENTS.md [quick]
├── Task 5: Generate ./src/template_manager/AGENTS.md [quick]
└── Task 6: Generate ./skills/AGENTS.md [quick]

Wave FINAL (After ALL Wave 1 tasks complete - parallel review):
├── Task F1: Validate all files (line counts, no duplicates, paths exist) [unspecified-low]
├── Task F2: Deduplicate content across hierarchy [unspecified-low]
└── Task F3: Final quality audit - all conventions followed [quick]

### Dependency Matrix

- **1-6**: — — All can run in parallel, no dependencies
- **F1-F3**: 1-6 completed — All can run in parallel

### Agent Dispatch Summary

- **Wave 1**: **6** — All tasks → `quick` category (6x quick in parallel)
- **Wave Final**: **3** — F1→`unspecified-low`, F2→`unspecified-low`, F3→`quick`

---

## TODOs

- [x] 1. Generate root AGENTS.md at ./AGENTS.md

  **What to do**:
  - Update existing root AGENTS.md to the PROJECT KNOWLEDGE BASE format
  - Include all required sections: OVERVIEW, STRUCTURE, WHERE TO LOOK, CODE MAP (skipped - LSP not used), CONVENTIONS, ANTI-PATTERNS, UNIQUE STYLES, COMMANDS, NOTES
  - Keep 50-150 lines total
  - Incorporate existing content from the original AGENTS.md
  - Remove generic advice already covered by Python standards

  **Must NOT do**:
  - Don't remove existing critical information
  - Don't increase size beyond 150 lines
  - Don't repeat subdirectory-specific content here

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Simple file update with structured format already defined
  - **Skills**: []
  - **Skills Evaluated but Omitted**:
    - `writing`: Not needed for simple structured update

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 2-6)
  - **Blocks**: Final verification wave (F1-F3)
  - **Blocked By**: None (can start immediately)

  **References** (CRITICAL - Be Exhaustive):

  **Existing Content**:
  - `/home/hs/workspace/opencode/office-skill/AGENTS.md:1-151` - Original guidelines to incorporate

  **Pattern References**:
  - `/init-deep` command specification - Root AGENTS.md structure template

  **Acceptance Criteria**:
  - [ ] File updated at ./AGENTS.md
  - [ ] All required sections present
  - [ ] Line count between 50-150
  - [ ] No generic Python advice

  **QA Scenarios**:

  ```
  Scenario: Root file format validation
    Tool: Bash
    Preconditions: File exists
    Steps:
      1. Count lines: `wc -l ./AGENTS.md`
      2. Verify all required section headers exist (grep for each section)
      3. Verify file paths in WHERE TO LOOK table all exist
    Expected Result: Line count between 50-150, all sections present, all paths verified
    Failure Indicators: Line count outside range, missing sections, broken paths
    Evidence: .sisyphus/evidence/task-1-root-validation.txt
  ```

  **Commit**: YES
  - Message: `docs: update root AGENTS.md to knowledge base format`
  - Files: ./AGENTS.md

- [x] 2. Generate AGENTS.md at ./src/office_main/

  **What to do**:
  - Create new AGENTS.md for the office_main top-level package
  - Include: OVERVIEW, STRUCTURE, WHERE TO LOOK, CONVENTIONS, ANTI-PATTERNS
  - Keep to 30-80 lines max
  - Don't repeat any content from root AGENTS.md
  - Focus on package-specific organization

  **Must NOT do**:
  - Don't repeat formatting/naming conventions already at root
  - Don't document individual handlers here - that's for core/ subdirectory

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Small file creation following template, straightforward content
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1 (with Tasks 1, 3-6)
  - **Blocks**: Final verification
  - **Blocked By**: None

  **References**:
  - Explore results: `src/office_main/` structure analysis - overall organization, separation cli/core, export pattern

  **Acceptance Criteria**:
  - [ ] File created at ./src/office_main/AGENTS.md
  - [ ] 30-80 lines
  - [ ] No content duplicated from root
  - [ ] All required sections present

  **QA Scenarios**:

  ```
  Scenario: src/office_main file creation
    Tool: Bash
    Preconditions: Directory exists
    Steps:
      1. Check file exists at ./src/office_main/AGENTS.md
      2. Count lines and verify within 30-80
      3. Check that OVERVIEW section describes the package purpose
    Expected Result: File created, line count correct, all sections present
    Failure Indicators: Missing file, wrong line count, missing sections
    Evidence: .sisyphus/evidence/task-2-office-main-validation.txt

  Scenario: No content duplication
    Tool: Grep
    Preconditions: Both root and this file exist
    Steps:
      1. Check that this file doesn't repeat the full command list from root
      2. Check that this file doesn't repeat naming/formatting conventions from root
    Expected Result: No duplication found - only package-specific content
    Failure Indicators: Duplicate content found
    Evidence: .sisyphus/evidence/task-2-no-duplication.txt
  ```

  **Commit**: YES
  - Message: `docs: add AGENTS.md for src/office_main`
  - Files: ./src/office_main/AGENTS.md

- [x] 3. Generate AGENTS.md at ./src/office_main/core/

  **What to do**:
  - Create new AGENTS.md for the core subpackage
  - Include sections: OVERVIEW, STRUCTURE, WHERE TO LOOK, CONVENTIONS, ANTI-PATTERNS
  - Document the symmetric document handler pattern that all three handlers follow
  - Focus on the specific conventions of this directory
  - Keep to 30-80 lines max
  - Don't repeat content from parent or root

  **Must NOT do**:
  - Don't repeat root-level formatting conventions
  - Don't repeat package-level structure from office_main/AGENTS.md

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: Straightforward documentation of discovered patterns
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1
  - **Blocks**: Final verification
  - **Blocked By**: None

  **References**:
  - Deep exploration: Symmetric handler pattern, one file per document format, shared LibreOfficeCLI

  **Acceptance Criteria**:
  - [ ] File created at ./src/office_main/core/AGENTS.md
  - [ ] 30-80 lines
  - [ ] Documents the symmetric handler pattern that must be preserved when adding new handlers
  - [ ] No duplication from root or parent

  **QA Scenarios**:
  ```
  Scenario: Core directory documentation complete
    Tool: Bash
    Preconditions: Directory exists
    Steps:
      1. Check file exists
      2. Count lines
      3. Verify symmetric handler pattern is documented
    Expected Result: File exists, line count within range, pattern documented
    Evidence: .sisyphus/evidence/task-3-core-validation.txt
  ```

  **Commit**: YES
  - Message: `docs: add AGENTS.md for src/office_main/core`
  - Files: ./src/office_main/core/AGENTS.md

- [x] 4. Generate AGENTS.md at ./src/office_main/cli/

  **What to do**:
  - Create new AGENTS.md for the CLI directory
  - Document the subcommand pattern and argument parsing conventions
  - Include how the CLI layer depends on core functionality
  - Keep 30-80 lines max
  - Don't repeat parent/root content

  **Must NOT do**:
  - Don't repeat overall project commands - those are at root

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - Reason: Single-file directory with clear pattern to document

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1
  - **Blocks**: Final verification
  - **Blocked By**: None

  **References**:
  - Analysis: All CLI in one file (office_cli.py), subcommand hierarchy, consistent --json flag pattern

  **Acceptance Criteria**:
  - [ ] File created at ./src/office_main/cli/AGENTS.md
  - [ ] 30-80 lines
  - [ ] Documents CLI subcommand pattern
  - [ ] No duplication from root/parent

  **QA Scenarios**:
  ```
  Scenario: CLI directory documentation
    Tool: Bash
    Steps:
      1. Verify file exists
      2. Check line count
      3. Confirm subcommand pattern and --json convention documented
    Expected Result: All checks pass
    Evidence: .sisyphus/evidence/task-4-cli-validation.txt
  ```

  **Commit**: YES
  - Message: `docs: add AGENTS.md for src/office_main/cli`
  - Files: ./src/office_main/cli/AGENTS.md

- [x] 5. Generate AGENTS.md at ./src/template_manager/

  **What to do**:
  - Create new AGENTS.md for the template_manager package
  - Document the modular single-responsibility design
  - Explain each module's responsibility (storage, analyzer, search, generator)
  - Document the dependency injection pattern used
  - Document the 5-part naming convention enforcement
  - Keep 30-80 lines max

  **Must NOT do**:
  - Don't repeat hierarchical naming from root - just reference it, specify how this package enforces it

  **Recommended Agent Profile**:
  - **Category**: `quick`
    - Reason: 5 modules with clear separation already analyzed, just document it
  - **Skills**: []

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1
  - **Blocks**: Final verification
  - **Blocked By**: None

  **References**:
  - Deep exploration: Single-responsibility modular design, dependency injection, hierarchical storage

  **Acceptance Criteria**:
  - [ ] File created at ./src/template_manager/AGENTS.md
  - [ ] 30-80 lines
  - [ ] Documents modular design and dependency injection pattern
  - [ ] No duplication from root/parent

  **QA Scenarios**:
  ```
  Scenario: Template manager documentation
    Tool: Bash
    Steps:
      1. Check file exists
      2. Verify line count within 30-80
      3. Confirm each module's responsibility is documented
      4. Verify dependency injection pattern is documented
    Expected Result: All checks pass
    Evidence: .sisyphus/evidence/task-5-template-manager-validation.txt
  ```

  **Commit**: YES
  - Message: `docs: add AGENTS.md for src/template_manager`
  - Files: ./src/template_manager/AGENTS.md

- [x] 6. Generate AGENTS.md at ./skills/

  **What to do**:
  - Create new AGENTS.md for the skills directory
  - Document what each skill file covers (01_OFFICE_SKILL through 05_TEMPLATE_SKILL)
  - Explain the purpose of these skill definitions (AI agent integration for Open Claw)
  - Keep 30-80 lines max

  **Must NOT do**:
  - Don't repeat content from the skill files themselves - just document the organization

  **Recommended Agent Profile**:
  - **Category**: `quick`
  - Reason: Simple directory documentation with 6 files, straightforward

  **Parallelization**:
  - **Can Run In Parallel**: YES
  - **Parallel Group**: Wave 1
  - **Blocks**: Final verification
  - **Blocked By**: None

  **References**:
  - Directory analysis: 6 markdown files, each for a specific document type/processing domain

  **Acceptance Criteria**:
  - [ ] File created at ./skills/AGENTS.md
  - [ ] 30-80 lines
  - [ ] Lists each skill file and its domain/purpose
  - [ ] Explains Open Claw skill integration purpose
  - [ ] No duplication from root

  **QA Scenarios**:
  ```
  Scenario: Skills directory documentation complete
    Tool: Bash
    Steps:
      1. Check file exists
      2. Verify line count within 30-80
      3. Confirm all 6 skill files (01-05) are listed with their purposes
    Expected Result: All checks pass
    Evidence: .sisyphus/evidence/task-6-skills-validation.txt
  ```

  **Commit**: YES
  - Message: `docs: add AGENTS.md for skills`
  - Files: ./skills/AGENTS.md

---

## Final Verification Wave

> 4 review agents run in PARALLEL. ALL must APPROVE. No auto-proceed - waits for completion.

- [x] F1. **Validate All Files** — `unspecified-low`
   Verify: 1) All expected files created/updated, 2) Line counts within limits, 3) All required section headers present, 4) All referenced file paths actually exist. Output: `Files [6/6] | Lines [within limits] | VERDICT: APPROVE`

- [x] F2. **Deduplicate Content** — `unspecified-low`
   Check: 1) Child files don't repeat parent content, 2) Root doesn't include subdirectory-specific details, 3) No generic Python advice that belongs to Python documentation. Remove any duplication found. Output: `Duplicates [0 found/removed] | VERDICT: APPROVE`

- [x] F3. **Final Quality Audit** — `quick`
   Verify: 1) Telegraphic style maintained (concise, no verbose explanations), 2) All anti-patterns captured, 3) All project-specific conventions documented, 4) File hierarchy matches scoring decisions. Output: `Quality [OK] | VERDICT: APPROVE`

---

## Commit Strategy

- Each task commits individually after task completes
- Message format: `docs: [update/add] AGENTS.md [for path]`

---

## Success Criteria

### Final Checklist
- [ ] Root AGENTS.md updated to PROJECT KNOWLEDGE BASE format
- [ ] All 5 subdirectory AGENTS.md files created
- [ ] All line count limits respected (root 50-150, subdirs 30-80)
- [ ] No duplicate content between parent and child
- [ ] No generic advice - all content project-specific
- [ ] All scoring-based locations have AGENTS.md created
- [ ] All files pass quality audit

### Verification Commands
```bash
# Check all files exist
ls -la ./AGENTS.md ./src/office_main/AGENTS.md ./src/office_main/core/AGENTS.md ./src/office_main/cli/AGENTS.md ./src/template_manager/AGENTS.md ./skills/AGENTS.md

# Verify line counts
wc -l ./AGENTS.md ./src/office_main/AGENTS.md ./src/office_main/core/AGENTS.md ./src/office_main/cli/AGENTS.md ./src/template_manager/AGENTS.md ./skills/AGENTS.md
```

