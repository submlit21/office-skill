#!/usr/bin/env python3
"""
Office CLI - Command-line interface for office-skill.

This script provides a command-line interface to the office-skill package,
making it easy to perform common office document operations from the terminal.
"""

import argparse
import json
import sys
import os
from pathlib import Path
from typing import Dict, Any

# Add src to path for development
sys.path.insert(0, str(Path(__file__).parent / "src"))

try:
    from office_skill import (
        LibreOfficeCLI,
        DocxHandler,
        XlsxHandler,
        PptxHandler
    )
except ImportError:
    print("Error: office_skill package not found. Install with: pip install -e .")
    sys.exit(1)


def create_document(args):
    """Create a new document."""
    output_path = Path(args.output).absolute()

    if output_path.suffix.lower() == '.docx':
        handler = DocxHandler(str(output_path))
        print(f"Created Word document: {output_path}")
    elif output_path.suffix.lower() in ['.xlsx', '.xls']:
        handler = XlsxHandler(str(output_path))
        print(f"Created Excel spreadsheet: {output_path}")
    elif output_path.suffix.lower() in ['.pptx', '.ppt']:
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

    if input_path.suffix.lower() == '.docx':
        handler = DocxHandler(str(input_path))
    elif input_path.suffix.lower() in ['.xlsx', '.xls']:
        handler = XlsxHandler(str(input_path))
    elif input_path.suffix.lower() in ['.pptx', '.ppt']:
        handler = PptxHandler(str(input_path))
    else:
        print(f"Unsupported file extension: {input_path.suffix}")
        return {"status": "error", "message": "Unsupported file type"}

    analysis = handler.analyze_structure()

    if args.json:
        print(json.dumps(analysis, indent=2))
    else:
        print(f"Document: {analysis['document']}")
        if 'content_summary' in analysis:
            print(f"Content items: {len(analysis['content_summary'].get('items', []))}")
        if 'sheet_count' in analysis:
            print(f"Sheets: {analysis['sheet_count']}")
        if 'slide_count' in analysis:
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
    if suffix == '.docx':
        handler = DocxHandler(str(input_path))
    elif suffix in ['.xlsx', '.xls']:
        handler = XlsxHandler(str(input_path))
    elif suffix in ['.pptx', '.ppt']:
        handler = PptxHandler(str(input_path))
    else:
        print(f"Unsupported input file extension: {suffix}")
        return {"status": "error", "message": "Unsupported input file type"}

    result = handler.export(str(output_path), args.format)

    print(f"Exported: {input_path} -> {output_path} ({args.format})")
    return {"status": "success", "input": str(input_path),
            "output": str(output_path), "format": args.format, "result": result}


def validate_spreadsheet(args):
    """Validate spreadsheet formulas."""
    input_path = Path(args.input).absolute()

    if not input_path.exists():
        print(f"Error: Input file not found: {input_path}")
        return {"status": "error", "message": "File not found"}

    if not input_path.suffix.lower() in ['.xlsx', '.xls']:
        print(f"Error: Not a spreadsheet file: {input_path}")
        return {"status": "error", "message": "Not a spreadsheet file"}

    handler = XlsxHandler(str(input_path))
    result = handler.validate_formulas()

    if args.json:
        print(json.dumps(result, indent=2))
    else:
        status = result.get('status', 'unknown')
        print(f"Validation status: {status}")
        if 'errors' in result and result['errors']:
            print(f"Errors found: {len(result['errors'])}")
            for error in result['errors'][:5]:  # Show first 5 errors
                print(f"  - {error}")

    return result


def cli_info(args):
    """Show CLI information."""
    cli = LibreOfficeCLI()
    try:
        presets = cli.export("presets")
        print(f"CLI Information:")
        print(f"• cli-anything-libreoffice: Available")
        print(f"• Export presets: {len(presets.get('presets', []))}")
        print(f"• JSON output: Supported")
    except Exception as e:
        print(f"CLI test failed: {e}")
        print("Make sure cli-anything-libreoffice is installed and in PATH")

    return {"status": "success"}


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
        """
    )

    subparsers = parser.add_subparsers(dest='command', help='Command to execute')

    # Create command
    create_parser = subparsers.add_parser('create', help='Create a new document')
    create_parser.add_argument('--output', required=True, help='Output file path')
    create_parser.add_argument('--template', help='Template file path (optional)')
    create_parser.set_defaults(func=create_document)

    # Analyze command
    analyze_parser = subparsers.add_parser('analyze', help='Analyze a document')
    analyze_parser.add_argument('--input', required=True, help='Input file path')
    analyze_parser.add_argument('--json', action='store_true', help='Output as JSON')
    analyze_parser.set_defaults(func=analyze_document)

    # Export command
    export_parser = subparsers.add_parser('export', help='Export document to another format')
    export_parser.add_argument('--input', required=True, help='Input file path')
    export_parser.add_argument('--output', required=True, help='Output file path')
    export_parser.add_argument('--format', default='pdf', help='Export format (default: pdf)')
    export_parser.set_defaults(func=export_document)

    # Validate command
    validate_parser = subparsers.add_parser('validate', help='Validate spreadsheet formulas')
    validate_parser.add_argument('--input', required=True, help='Input spreadsheet path')
    validate_parser.add_argument('--json', action='store_true', help='Output as JSON')
    validate_parser.set_defaults(func=validate_spreadsheet)

    # Info command
    info_parser = subparsers.add_parser('info', help='Show CLI information')
    info_parser.set_defaults(func=cli_info)

    # Add version
    parser.add_argument('--version', action='version', version='office-skill 0.1.0')

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    try:
        result = args.func(args)
        return 0
    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        if args.command == 'analyze' and args.json:
            print(json.dumps({"status": "error", "message": str(e)}, indent=2))
        return 1


if __name__ == "__main__":
    sys.exit(main())