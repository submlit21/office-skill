#!/usr/bin/env python3
"""
Bridge script for safe communication between Node.js and office_skill.

This script reads JSON from stdin, executes the requested operation,
and writes JSON result to stdout. This prevents command injection
by avoiding string interpolation in generated Python code.
"""

import json
import sys
import traceback
from pathlib import Path

# Add parent directory to Python path to find office_skill module
sys.path.insert(0, str(Path(__file__).parent.parent))

# Import office_skill modules
try:
    from office_skill import LibreOfficeCLI, DocxHandler, XlsxHandler, PptxHandler
except ImportError as e:
    print(json.dumps({
        "error": f"Failed to import office_skill: {e}",
        "type": "ImportError",
        "python_path": sys.path
    }))
    sys.exit(1)


def safe_path(path_str):
    """Validate and convert a path string to absolute Path."""
    if not path_str:
        return None

    # Basic safety checks
    if '\0' in path_str:
        raise ValueError("Null byte in path")

    # Convert to Path and resolve
    path = Path(path_str).resolve()

    # Additional checks (optional)
    # Reject paths with too many '..' components (before resolving)
    if path_str.count('..') > 5:
        raise ValueError("Too many parent directory references")

    return str(path)


def create_document(params):
    """Create a new document."""
    doc_type = params.get("type")
    output_path = safe_path(params.get("output"))
    template_path = safe_path(params.get("template"))

    if not output_path:
        return {"error": "Output path is required"}

    # Map type to handler class
    handler_classes = {
        "docx": DocxHandler,
        "xlsx": XlsxHandler,
        "pptx": PptxHandler
    }

    if doc_type not in handler_classes:
        return {"error": f"Unsupported document type: {doc_type}"}

    HandlerClass = handler_classes[doc_type]

    try:
        # Create handler and call create_from_template
        handler = HandlerClass(output_path)
        result = handler.create_from_template(template_path)

        # For now, just return success since create_from_template is a stub
        return {
            "status": "success",
            "output": output_path,
            "template_used": template_path is not None
        }
    except Exception as e:
        return {
            "error": f"Failed to create document: {e}",
            "traceback": traceback.format_exc()
        }


def edit_document(params):
    """Edit an existing document."""
    input_path = safe_path(params.get("input"))
    output_path = safe_path(params.get("output"))
    operations = params.get("operations", [])

    if not input_path:
        return {"error": "Input path is required"}
    if not output_path:
        return {"error": "Output path is required"}

    # Determine document type from extension
    ext = Path(input_path).suffix.lower()[1:]  # Remove leading dot
    type_map = {
        "docx": "docx", "doc": "docx",
        "xlsx": "xlsx", "xls": "xlsx",
        "pptx": "pptx", "ppt": "pptx"
    }
    doc_type = type_map.get(ext, "docx")

    # Map type to handler class
    handler_classes = {
        "docx": DocxHandler,
        "xlsx": XlsxHandler,
        "pptx": PptxHandler
    }

    if doc_type not in handler_classes:
        return {"error": f"Unsupported document type for file: {input_path}"}

    HandlerClass = handler_classes[doc_type]

    try:
        handler = HandlerClass(input_path)
        results = []

        for op in operations:
            action = op.get("action")
            result = {"action": action}

            if doc_type == "docx":
                if action == "add_paragraph":
                    result["value"] = handler.add_paragraph(
                        op.get("text", ""),
                        op.get("index")
                    )
                elif action == "add_heading":
                    result["value"] = handler.add_heading(
                        op.get("text", ""),
                        op.get("level", 1),
                        op.get("index")
                    )
                else:
                    result["error"] = f"Unknown DOCX action: {action}"

            elif doc_type == "xlsx":
                if action == "set_cell":
                    result["value"] = handler.set_cell(
                        op.get("sheet", ""),
                        op.get("cell", ""),
                        op.get("value", ""),
                        op.get("formula", False)
                    )
                else:
                    result["error"] = f"Unknown XLSX action: {action}"

            elif doc_type == "pptx":
                if action == "add_slide":
                    result["value"] = handler.add_slide(
                        op.get("layout", "Title and Content"),
                        op.get("index")
                    )
                else:
                    result["error"] = f"Unknown PPTX action: {action}"

            results.append(result)

        return {
            "status": "success",
            "operations": results,
            "output": output_path
        }
    except Exception as e:
        return {
            "error": f"Failed to edit document: {e}",
            "traceback": traceback.format_exc()
        }


def analyze_document(params):
    """Analyze a document."""
    input_path = safe_path(params.get("input"))
    doc_type = params.get("type")

    if not input_path:
        return {"error": "Input path is required"}

    # If type not provided, infer from extension
    if not doc_type:
        ext = Path(input_path).suffix.lower()[1:]  # Remove leading dot
        type_map = {
            "docx": "docx", "doc": "docx",
            "xlsx": "xlsx", "xls": "xlsx",
            "pptx": "pptx", "ppt": "pptx"
        }
        doc_type = type_map.get(ext, "docx")

    # Map type to handler class
    handler_classes = {
        "docx": DocxHandler,
        "xlsx": XlsxHandler,
        "pptx": PptxHandler
    }

    if doc_type not in handler_classes:
        return {"error": f"Unsupported document type: {doc_type}"}

    HandlerClass = handler_classes[doc_type]

    try:
        handler = HandlerClass(input_path)
        analysis = handler.analyze_structure()
        return analysis
    except Exception as e:
        return {
            "error": f"Failed to analyze document: {e}",
            "traceback": traceback.format_exc()
        }


def export_document(params):
    """Export a document to another format."""
    input_path = safe_path(params.get("input"))
    output_path = safe_path(params.get("output"))
    format_type = params.get("format", "pdf")

    if not input_path:
        return {"error": "Input path is required"}
    if not output_path:
        return {"error": "Output path is required"}

    # Determine document type from extension
    ext = Path(input_path).suffix.lower()[1:]  # Remove leading dot
    type_map = {
        "docx": "docx", "doc": "docx",
        "xlsx": "xlsx", "xls": "xlsx",
        "pptx": "pptx", "ppt": "pptx"
    }
    doc_type = type_map.get(ext, "docx")

    # Map type to handler class
    handler_classes = {
        "docx": DocxHandler,
        "xlsx": XlsxHandler,
        "pptx": PptxHandler
    }

    if doc_type not in handler_classes:
        return {"error": f"Unsupported document type for file: {input_path}"}

    HandlerClass = handler_classes[doc_type]

    try:
        handler = HandlerClass(input_path)
        result = handler.export(output_path, format_type)
        return {
            "status": "success",
            "output": output_path,
            "format": format_type,
            "result": result
        }
    except Exception as e:
        return {
            "error": f"Failed to export document: {e}",
            "traceback": traceback.format_exc()
        }


def validate_spreadsheet(params):
    """Validate spreadsheet formulas."""
    input_path = safe_path(params.get("input"))
    timeout = params.get("timeout", 60)

    if not input_path:
        return {"error": "Input path is required"}

    try:
        handler = XlsxHandler(input_path)
        result = handler.validate_formulas()
        return result
    except Exception as e:
        return {
            "error": f"Failed to validate spreadsheet: {e}",
            "traceback": traceback.format_exc()
        }


def main():
    """Main entry point."""
    try:
        # Read JSON from stdin with size limit (10MB)
        max_size = 10 * 1024 * 1024  # 10MB
        input_data = sys.stdin.read(max_size + 1)
        if len(input_data) > max_size:
            print(json.dumps({"error": f"Input too large (max {max_size} bytes)"}))
            sys.exit(1)

        if not input_data:
            print(json.dumps({"error": "No input provided"}))
            sys.exit(1)

        request = json.loads(input_data)
        command = request.get("command")
        params = request.get("params", {})

        # Dispatch to appropriate function
        handlers = {
            "create_document": create_document,
            "edit_document": edit_document,
            "analyze_document": analyze_document,
            "export_document": export_document,
            "validate_spreadsheet": validate_spreadsheet
        }

        if command not in handlers:
            result = {"error": f"Unknown command: {command}"}
        else:
            result = handlers[command](params)

        # Output result as JSON
        print(json.dumps(result, indent=2))

    except json.JSONDecodeError as e:
        print(json.dumps({
            "error": f"Invalid JSON input: {e}",
            "type": "JSONDecodeError"
        }))
        sys.exit(1)
    except Exception as e:
        print(json.dumps({
            "error": f"Unexpected error: {e}",
            "traceback": traceback.format_exc()
        }))
        sys.exit(1)


if __name__ == "__main__":
    main()