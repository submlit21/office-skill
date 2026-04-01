# SKILLS DIRECTORY KNOWLEDGE BASE

**Generated:** Wed Apr 01 2026
**Purpose:** Open Claw AI agent skill definitions for Office document processing

## OVERVIEW
This directory contains domain-specific skill files that provide guidance to AI agents working on Open Claw Office document processing tasks. Each file covers a specific document type or processing domain.

## STRUCTURE
```
skills/
├── 01_OFFICE_SKILL.md    # Overall project and core rules
├── 02_DOCX_SKILL.md      # Word document processing rules
├── 03_XLSX_SKILL.md      # Excel spreadsheet processing rules
├── 04_PPTX_SKILL.md      # PowerPoint presentation rules
├── 05_TEMPLATE_SKILL.md  # Template management system rules
└── _meta.json            # Skill metadata for agent loading
```

## WHERE TO LOOK
| Domain | Skill File | Purpose |
|--------|------------|---------|
| Overall project | 01_OFFICE_SKILL.md | Core integration rules and agent responsibilities |
| Word documents | 02_DOCX_SKILL.md | DOCX-specific processing, redlining, text manipulation |
| Excel spreadsheets | 03_XLSX_SKILL.md | XLSX formula validation, financial standards, formatting |
| PowerPoint | 04_PPTX_SKILL.md | PPTX slide manipulation, layout preservation |
| Template system | 05_TEMPLATE_SKILL.md | Template loading, naming conventions, management |

## CONVENTIONS
- Numbered prefix indicates loading order (lower numbers load first)
- Each skill is domain-specific and complements higher-level guidance
- Follow rules in order: specific skill rules override general project rules
- All rules are mandatory for AI agents when working in the corresponding domain

## ANTI-PATTERNS (THIS DIRECTORY)
- Don't repeat general project conventions already in root AGENTS.md
- Don't merge domains - keep each skill focused on one area
- Don't omit mandatory validation checks in domain-specific rules
- Don't change numbering scheme - it determines precedence
- Don't duplicate advice between skills