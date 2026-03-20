"""
Basic usage examples for Office Skill.

This script demonstrates the core functionality of the office-skill package.
"""

import os
import tempfile
from pathlib import Path

# Import the handlers
from office_skill import DocxHandler, XlsxHandler, PptxHandler


def example_docx_creation():
    """Create a simple Word document."""
    print("=== Creating Word Document ===")

    # Create a temporary file for the example
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        doc_path = tmp.name

    try:
        doc = DocxHandler(doc_path)

        # Add title
        doc.add_heading("Sample Document", level=1)

        # Add introduction
        doc.add_paragraph("This is a sample document created with office-skill.")
        doc.add_paragraph("It demonstrates basic document creation capabilities.")

        # Add a section
        doc.add_heading("Features", level=2)

        # Add a list
        doc.add_list([
            "Document creation",
            "Content editing",
            "Formatting",
            "Export to multiple formats"
        ])

        # Add a table
        doc.add_table(rows=3, cols=3)
        doc.set_table_cell(0, 0, 0, "Item")
        doc.set_table_cell(0, 0, 1, "Quantity")
        doc.set_table_cell(0, 0, 2, "Price")

        # Add some data
        data = [
            ("Product A", "10", "$100"),
            ("Product B", "5", "$250"),
            ("Product C", "8", "$75")
        ]

        for i, (item, qty, price) in enumerate(data, start=1):
            doc.set_table_cell(0, i, 0, item)
            doc.set_table_cell(0, i, 1, qty)
            doc.set_table_cell(0, i, 2, price)

        # Add conclusion
        doc.add_heading("Conclusion", level=2)
        doc.add_paragraph("This demonstrates the basic capabilities of the office-skill package.")

        # Export to PDF
        pdf_path = doc_path.replace(".docx", ".pdf")
        doc.export(pdf_path, format="pdf")

        print(f"Document created: {doc_path}")
        print(f"PDF exported: {pdf_path}")
        print(f"File sizes: DOCX={Path(doc_path).stat().st_size} bytes, "
              f"PDF={Path(pdf_path).stat().st_size} bytes")

        return doc_path, pdf_path

    finally:
        # Cleanup (in real usage, you might want to keep the files)
        if os.path.exists(doc_path):
            os.unlink(doc_path)
        pdf_path = doc_path.replace(".docx", ".pdf")
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)


def example_xlsx_creation():
    """Create a simple Excel spreadsheet."""
    print("\n=== Creating Excel Spreadsheet ===")

    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        xls_path = tmp.name

    try:
        xls = XlsxHandler(xls_path)

        # Create a budget sheet
        xls.add_sheet(name="Budget")

        # Set headers
        xls.set_cell("Budget", "A1", "Category", formula=False)
        xls.set_cell("Budget", "B1", "Budgeted", formula=False)
        xls.set_cell("Budget", "C1", "Actual", formula=False)
        xls.set_cell("Budget", "D1", "Variance", formula=False)

        # Add data
        categories = [
            ("Rent", 1500, 1500),
            ("Utilities", 200, 185),
            ("Groceries", 400, 420),
            ("Transportation", 150, 145),
            ("Entertainment", 100, 120)
        ]

        for i, (category, budgeted, actual) in enumerate(categories, start=2):
            xls.set_cell("Budget", f"A{i}", category, formula=False)
            xls.set_cell("Budget", f"B{i}", str(budgeted), formula=False)
            xls.set_cell("Budget", f"C{i}", str(actual), formula=False)
            xls.set_cell("Budget", f"D{i}", f"=C{i}-B{i}", formula=True)

        # Add totals
        total_row = len(categories) + 2
        xls.set_cell("Budget", f"A{total_row}", "TOTAL", formula=False)
        xls.set_cell("Budget", f"B{total_row}", f"=SUM(B2:B{total_row-1})", formula=True)
        xls.set_cell("Budget", f"C{total_row}", f"=SUM(C2:C{total_row-1})", formula=True)
        xls.set_cell("Budget", f"D{total_row}", f"=SUM(D2:D{total_row-1})", formula=True)

        # Apply formatting
        xls.set_cell_format("Budget", "B2:D10", "$#,##0")

        # Create summary sheet
        xls.add_sheet(name="Summary")
        xls.set_cell("Summary", "A1", "Budget Summary", formula=False)
        xls.merge_cells("Summary", "A1:D1")

        xls.set_cell("Summary", "A3", "Total Budgeted", formula=False)
        xls.set_cell("Summary", "B3", f"=Budget!B{total_row}", formula=True)

        xls.set_cell("Summary", "A4", "Total Actual", formula=False)
        xls.set_cell("Summary", "B4", f"=Budget!C{total_row}", formula=True)

        xls.set_cell("Summary", "A5", "Overall Variance", formula=False)
        xls.set_cell("Summary", "B5", f"=Budget!D{total_row}", formula=True)

        # Format summary
        xls.set_cell_format("Summary", "B3:B5", "$#,##0")

        print(f"Spreadsheet created: {xls_path}")
        print("Sheets created: Budget, Summary")
        print("Formulas added: Variance calculations, totals")

        return xls_path

    finally:
        if os.path.exists(xls_path):
            os.unlink(xls_path)


def example_pptx_creation():
    """Create a simple PowerPoint presentation."""
    print("\n=== Creating PowerPoint Presentation ===")

    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        ppt_path = tmp.name

    try:
        ppt = PptxHandler(ppt_path)

        # Create title slide
        ppt.add_slide(layout="Title Slide")
        ppt.set_content(0, title="Office Skill Demo", content="Presentation Example")
        ppt.set_speaker_notes(0, "Welcome to the Office Skill demonstration.")

        # Create agenda slide
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(1, title="Agenda", content="• Introduction\n• Features\n• Examples\n• Conclusion")
        ppt.set_speaker_notes(1, "Briefly outline what we'll cover.")

        # Create content slides
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(2, title="Features", content="• Document creation\n• Spreadsheet manipulation\n• Presentation design\n• Multi-format export")
        ppt.set_speaker_notes(2, "Highlight key features.")

        ppt.add_slide(layout="Title and Content")
        ppt.set_content(3, title="Examples", content="• Word: Reports and letters\n• Excel: Budgets and analysis\n• PowerPoint: Presentations and decks")
        ppt.set_speaker_notes(3, "Show practical applications.")

        # Create conclusion slide
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(4, title="Conclusion", content="• Powerful office automation\n• Easy integration\n• Professional results")
        ppt.set_speaker_notes(4, "Summarize benefits and next steps.")

        # Export to PDF
        pdf_path = ppt_path.replace(".pptx", ".pdf")
        ppt.export(pdf_path, format="pdf")

        print(f"Presentation created: {ppt_path}")
        print(f"PDF exported: {pdf_path}")
        print(f"Slides created: 5")
        print(f"Speaker notes added: Yes")

        return ppt_path, pdf_path

    finally:
        if os.path.exists(ppt_path):
            os.unlink(ppt_path)
        pdf_path = ppt_path.replace(".pptx", ".pdf")
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)


def example_cli_wrapper():
    """Demonstrate direct CLI wrapper usage."""
    print("\n=== CLI Wrapper Example ===")

    from office_skill import LibreOfficeCLI

    # Create a CLI instance
    cli = LibreOfficeCLI()

    # Execute some commands
    print("Testing CLI connection...")

    try:
        # Get export presets
        presets = cli.export("presets")
        print(f"Available export presets: {len(presets.get('presets', []))}")

        # Test writer list command
        writer_info = cli.writer("list")
        print(f"Writer commands available: {len(writer_info.get('commands', []))}")

        print("CLI wrapper test successful")

    except Exception as e:
        print(f"CLI test failed: {e}")
        print("Make sure cli-anything-libreoffice is installed and in PATH")


def main():
    """Run all examples."""
    print("Office Skill - Basic Usage Examples")
    print("=" * 50)

    # Run examples
    example_docx_creation()
    example_xlsx_creation()
    example_pptx_creation()
    example_cli_wrapper()

    print("\n" + "=" * 50)
    print("All examples completed successfully!")
    print("\nNote: Files were created in temporary locations and cleaned up.")
    print("In real usage, specify your own file paths.")


if __name__ == "__main__":
    main()