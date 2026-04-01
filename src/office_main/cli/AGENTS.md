# CLI SUBSYSTEM KNOWLEDGE BASE

**Generated:** Wed Apr 01 2026
**Parent Package:** office_main

## OVERVIEW
Single-file command-line interface for the Office Skill. Provides a structured CLI with hierarchical subcommands for all document operations and template management.

All business logic delegates to core handlers — this layer is purely for argument parsing and result output.

## STRUCTURE
```
cli/
└── office_cli.py   # Entire CLI implementation in one file (499 lines)
```

## WHERE TO LOOK
| Task | Location | Notes |
|------|----------|-------|
| Add new command | `office_cli.py` | Add handler function + register in subparser |
| Fix argument parsing | `office_cli.py` | All argparse setup here |
| Change JSON output format | `office_cli.py` | JSON serialization at output time |

## CONVENTIONS
- **Subcommand pattern**: Main top-level parser → subparsers for command categories (document commands, template commands)
- **Consistent `--json` flag**: Every command supports `--json` for machine-readable output
- **Thin layer**: Parse args → call core function → output results (no business logic)
- **Handler functions**: One function per top-level command (e.g., `create_document`, `analyze_document`)
- **Error handling**: Print errors to stderr, return non-zero exit code on failure
- **Return values**: All command handlers return structured dicts for JSON output

## COMMANDS AVAILABLE
11 top-level subcommands:
- Document: create, analyze, validate, export
- Template: add, get, list, search, delete, generate

## ANTI-PATTERNS
- Don't implement business logic in CLI layer — always delegate to core
- Don't break the subcommand hierarchy pattern
- Don't add new commands without supporting `--json` output
- Don't repeat argument definitions that can be shared
- Don't print success messages to stdout when `--json` is used
