#!/usr/bin/env python3
"""
Office CLI - Command-line interface for office-skill.

This script provides a command-line interface to the office-skill package,
making it easy to perform common office document operations from the terminal.
"""

import argparse
import json
import os
import sys
from pathlib import Path

# Add src to path for development
sys.path.insert(0, str(Path(__file__).parent.parent.parent))

try:
    from office_main.core import (
        DocxHandler,
        LibreOfficeCLI,
        PptxHandler,
        TemplateManager,
        XlsxHandler,
    )
except ImportError:
    print("Error: office_main package not found. Install with: pip install -e .")
    sys.exit(1)


def create_document(args):
    """Create a new document."""
    output_path = Path(args.output).absolute()

    if output_path.suffix.lower() == ".docx":
        handler = DocxHandler(str(output_path))
        print(f"Created Word document: {output_path}")
    elif output_path.suffix.lower() in [".xlsx", ".xls"]:
        handler = XlsxHandler(str(output_path))
        print(f"Created Excel spreadsheet: {output_path}")
    elif output_path.suffix.lower() in [".pptx", ".ppt"]:
        handler = PptxHandler(str(output_path))
        print(f"Created PowerPoint presentation: {output_path}")
    else:
        print(f"Unsupported file extension: {output_path.suffix}")
        return

    # Add basic content if template provided
    if args.template and os.path.exists(args.template):
        handler.create_from_template(args.template)
        print(f"Applied template: {args.template}")

    return {"status": "success", "output": str(output_path)}


def analyze_document(args):
    """Analyze a document."""
    input_path = Path(args.input).absolute()

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return {"status": "error", "message": "File not found"}

    if input_path.suffix.lower() == ".docx":
        handler = DocxHandler(str(input_path))
    elif input_path.suffix.lower() in [".xlsx", ".xls"]:
        handler = XlsxHandler(str(input_path))
    elif input_path.suffix.lower() in [".pptx", ".ppt"]:
        handler = PptxHandler(str(input_path))
    else:
        print(f"Unsupported file extension: {input_path.suffix}")
        return {"status": "error", "message": "Unsupported file type"}

    analysis = handler.analyze_structure()

    if args.json:
        print(json.dumps(analysis, indent=2))
    else:
        print(f"Document: {analysis['document']}")
        if "content_summary" in analysis:
            print(f"Content items: {len(analysis['content_summary'].get('items', []))}")
        if "sheet_count" in analysis:
            print(f"Sheets: {analysis['sheet_count']}")
        if "slide_count" in analysis:
            print(f"Slides: {analysis['slide_count']}")

    return analysis


def export_document(args):
    """Export a document to another format."""
    input_path = Path(args.input).absolute()
    output_path = Path(args.output).absolute()

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return {"status": "error", "message": "Input file not found"}

    # Determine handler type from input extension
    suffix = input_path.suffix.lower()
    if suffix == ".docx":
        handler = DocxHandler(str(input_path))
    elif suffix in [".xlsx", ".xls"]:
        handler = XlsxHandler(str(input_path))
    elif suffix in [".pptx", ".ppt"]:
        handler = PptxHandler(str(input_path))
    else:
        print(f"Unsupported input file extension: {suffix}")
        return {"status": "error", "message": "Unsupported input file type"}

    result = handler.export(str(output_path), args.format)

    print(f"Exported: {input_path} -> {output_path} ({args.format})")
    return {
        "status": "success",
        "input": str(input_path),
        "output": str(output_path),
        "format": args.format,
        "result": result,
    }


def validate_spreadsheet(args):
    """Validate spreadsheet formulas."""
    input_path = Path(args.input).absolute()

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return {"status": "error", "message": "File not found"}

    if input_path.suffix.lower() not in [".xlsx", ".xls"]:
        print(f"Error: Not a spreadsheet file: {input_path}")
        return {"status": "error", "message": "Not a spreadsheet file"}

    handler = XlsxHandler(str(input_path))
    result = handler.validate_formulas()

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        status = result.get("status", "unknown")
        print(f"Validation status: {status}")
        if "errors" in result and result["errors"]:
            print(f"Errors found: {len(result['errors'])}")
            for error in result["errors"][:5]:  # Show first 5 errors
                print(f"  - {error}")

    return result


def cli_info(args):
    """Show CLI information."""
    cli = LibreOfficeCLI()
    try:
        presets = cli.export("presets")
        print("CLI Information:")
        print("• cli-anything-libreoffice: Available")
        print(f"• Export presets: {len(presets.get('presets', []))}")
        print("• JSON output: Supported")
    except Exception as e:
        print(f"CLI test failed: {e}")
        print("Make sure cli-anything-libreoffice is installed and in PATH")

    return {"status": "success"}


# Template management functions
def template_add(args):
    """Add a new template."""
    manager = TemplateManager()

    try:
        result = manager.add_template(
            source_path=args.input,
            name=args.name,
            description=args.description,
            tags=args.tags.split(",") if args.tags else None,
            overwrite=args.overwrite,
        )

        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"Template added successfully: {args.name}")
            print(f"Location: {result.get('template_path', 'Unknown')}")
            print(f"Description: {result.get('description', 'No description')}")
            if result.get("tags"):
                print(f"Tags: {', '.join(result['tags'])}")

        return result

    except Exception as e:
        print(f"Error adding template: {e}", file=sys.stderr)
        return {"status": "error", "message": str(e)}


def template_get(args):
    """Get template details."""
    manager = TemplateManager()

    try:
        result = manager.get_template(args.name)

        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"Template: {result['name']}")
            print(f"Description: {result['description']}")
            print(f"Created: {result['created']}")
            print(f"Format: {result['format']}")
            print(f"Original: {result.get('original_name', 'Unknown')}")

            if result.get("analysis"):
                analysis = result["analysis"]
                print("\nAnalysis:")
                if analysis.get("type") == "word":
                    print("  Type: Word document")
                    print(f"  Pages: {analysis.get('pages', 0)}")
                    print(f"  Paragraphs: {analysis.get('paragraphs', 0)}")
                    print(f"  Tables: {analysis.get('tables', 0)}")
                elif analysis.get("type") == "excel":
                    print("  Type: Excel spreadsheet")
                    print(f"  Sheets: {analysis.get('sheets', 0)}")
                    print(f"  Cells: {analysis.get('cells', 0)}")
                elif analysis.get("type") == "powerpoint":
                    print("  Type: PowerPoint presentation")
                    print(f"  Slides: {analysis.get('slides', 0)}")
                    print(f"  Masters: {analysis.get('masters', 0)}")

            if result.get("tags"):
                print(f"\nTags: {', '.join(result['tags'])}")

        return result

    except Exception as e:
        print(f"Error getting template: {e}", file=sys.stderr)
        return {"status": "error", "message": str(e)}


def template_list(args):
    """List available templates."""
    manager = TemplateManager()

    try:
        templates = manager.list_templates(
            domain=args.domain, type_filter=args.type, purpose=args.purpose
        )

        if args.json:
            print(json.dumps(templates, indent=2))
        else:
            if not templates:
                print("No templates found.")
                return []

            print(f"Found {len(templates)} template(s):")
            print("-" * 80)

            for template in templates:
                print(f"Name: {template['name']}")
                print(f"Description: {template['description']}")
                print(f"Format: {template['format']}")
                print(f"Created: {template['created']}")

                if args.verbose:
                    if template.get("analysis"):
                        analysis = template["analysis"]
                        if analysis.get("type") == "word":
                            print(
                                f"  Pages: {analysis.get('pages', 0)}, Paragraphs: {analysis.get('paragraphs', 0)}"
                            )
                        elif analysis.get("type") == "excel":
                            print(
                                f"  Sheets: {analysis.get('sheets', 0)}, Cells: {analysis.get('cells', 0)}"
                            )
                        elif analysis.get("type") == "powerpoint":
                            print(
                                f"  Slides: {analysis.get('slides', 0)}, Masters: {analysis.get('masters', 0)}"
                            )

                print("-" * 80)

        return templates

    except Exception as e:
        print(f"Error listing templates: {e}", file=sys.stderr)
        return {"status": "error", "message": str(e)}


def template_search(args):
    """Search templates."""
    manager = TemplateManager()

    try:
        results = manager.search_templates(args.query)

        if args.json:
            print(json.dumps(results, indent=2))
        else:
            if not results:
                print(f"No templates found matching '{args.query}'")
                return []

            print(f"Found {len(results)} template(s) matching '{args.query}':")
            for template in results:
                print(f"• {template['name']}: {template['description']}")

        return results

    except Exception as e:
        print(f"Error searching templates: {e}", file=sys.stderr)
        return {"status": "error", "message": str(e)}


def template_delete(args):
    """Delete a template."""
    manager = TemplateManager()

    try:
        if not args.force:
            # Ask for confirmation
            response = input(f"Are you sure you want to delete template '{args.name}'? [y/N]: ")
            if response.lower() != "y":
                print("Deletion cancelled.")
                return {"status": "cancelled"}

        success = manager.delete_template(args.name, force=args.force)

        if success:
            print(f"Template '{args.name}' deleted successfully.")
            return {"status": "success", "template": args.name}
        else:
            print(f"Failed to delete template '{args.name}'.")
            return {"status": "error", "template": args.name}

    except Exception as e:
        print(f"Error deleting template: {e}", file=sys.stderr)
        return {"status": "error", "message": str(e)}


def template_generate(args):
    """Generate document from template."""
    manager = TemplateManager()

    try:
        # Parse variables if provided
        variables = {}
        if args.variables:
            for var_str in args.variables:
                if "=" in var_str:
                    key, value = var_str.split("=", 1)
                    variables[key.strip()] = value.strip()

        result = manager.generate_from_template(
            template_name=args.template,
            output_path=args.output,
            variables=variables if variables else None,
        )

        if args.json:
            print(json.dumps(result, indent=2))
        else:
            print(f"Generated document from template: {args.template}")
            print(f"Output: {args.output}")
            if variables:
                print(f"Variables applied: {', '.join(variables.keys())}")

        return result

    except Exception as e:
        print(f"Error generating from template: {e}", file=sys.stderr)
        return {"status": "error", "message": str(e)}


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Office document processing CLI",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s create --output document.docx
  %(prog)s analyze --input spreadsheet.xlsx --json
  %(prog)s export --input presentation.pptx --output slides.pdf --format pdf
  %(prog)s validate --input financial_model.xlsx
        """,
    )

    subparsers = parser.add_subparsers(dest="command", help="Command to execute")

    # Create command
    create_parser = subparsers.add_parser("create", help="Create a new document")
    create_parser.add_argument("--output", required=True, help="Output file path")
    create_parser.add_argument("--template", help="Template file path (optional)")
    create_parser.set_defaults(func=create_document)

    # Analyze command
    analyze_parser = subparsers.add_parser("analyze", help="Analyze a document")
    analyze_parser.add_argument("--input", required=True, help="Input file path")
    analyze_parser.add_argument("--json", action="store_true", help="Output as JSON")
    analyze_parser.set_defaults(func=analyze_document)

    # Export command
    export_parser = subparsers.add_parser("export", help="Export document to another format")
    export_parser.add_argument("--input", required=True, help="Input file path")
    export_parser.add_argument("--output", required=True, help="Output file path")
    export_parser.add_argument("--format", default="pdf", help="Export format (default: pdf)")
    export_parser.set_defaults(func=export_document)

    # Validate command
    validate_parser = subparsers.add_parser("validate", help="Validate spreadsheet formulas")
    validate_parser.add_argument("--input", required=True, help="Input spreadsheet path")
    validate_parser.add_argument("--json", action="store_true", help="Output as JSON")
    validate_parser.set_defaults(func=validate_spreadsheet)

    # Info command
    info_parser = subparsers.add_parser("info", help="Show CLI information")
    info_parser.set_defaults(func=cli_info)

    # Template management commands
    template_parser = subparsers.add_parser("template", help="Manage document templates")
    template_subparsers = template_parser.add_subparsers(
        dest="template_command", help="Template command"
    )

    # Template add
    add_parser = template_subparsers.add_parser("add", help="Add a new template")
    add_parser.add_argument("--input", required=True, help="Source document path")
    add_parser.add_argument(
        "--name", required=True, help="Template name (format: domain.type.purpose.variant.version)"
    )
    add_parser.add_argument("--description", help="Template description")
    add_parser.add_argument("--tags", help="Comma-separated tags")
    add_parser.add_argument("--overwrite", action="store_true", help="Overwrite existing template")
    add_parser.add_argument("--json", action="store_true", help="Output as JSON")
    add_parser.set_defaults(func=template_add)

    # Template get
    get_parser = template_subparsers.add_parser("get", help="Get template details")
    get_parser.add_argument("--name", required=True, help="Template name")
    get_parser.add_argument("--json", action="store_true", help="Output as JSON")
    get_parser.set_defaults(func=template_get)

    # Template list
    list_parser = template_subparsers.add_parser("list", help="List available templates")
    list_parser.add_argument("--domain", help="Filter by domain")
    list_parser.add_argument("--type", help="Filter by type")
    list_parser.add_argument("--purpose", help="Filter by purpose")
    list_parser.add_argument(
        "--verbose", "-v", action="store_true", help="Show detailed information"
    )
    list_parser.add_argument("--json", action="store_true", help="Output as JSON")
    list_parser.set_defaults(func=template_list)

    # Template search
    search_parser = template_subparsers.add_parser("search", help="Search templates")
    search_parser.add_argument("--query", required=True, help="Search query")
    search_parser.add_argument("--json", action="store_true", help="Output as JSON")
    search_parser.set_defaults(func=template_search)

    # Template delete
    delete_parser = template_subparsers.add_parser("delete", help="Delete a template")
    delete_parser.add_argument("--name", required=True, help="Template name")
    delete_parser.add_argument(
        "--force", "-f", action="store_true", help="Force deletion without confirmation"
    )
    delete_parser.set_defaults(func=template_delete)

    # Template generate
    generate_parser = template_subparsers.add_parser(
        "generate", help="Generate document from template"
    )
    generate_parser.add_argument("--template", required=True, help="Template name")
    generate_parser.add_argument("--output", required=True, help="Output document path")
    generate_parser.add_argument("--variables", nargs="*", help="Variables in format key=value")
    generate_parser.add_argument("--json", action="store_true", help="Output as JSON")
    generate_parser.set_defaults(func=template_generate)

    # Add version
    parser.add_argument("--version", action="version", version="office-skill 0.1.0")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    try:
        args.func(args)
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.command == "analyze" and args.json:
            print(json.dumps({"status": "error", "message": str(e)}, indent=2))
        return 1


if __name__ == "__main__":
    sys.exit(main())
