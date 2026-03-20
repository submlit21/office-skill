"""
Financial modeling example using office-skill.

This example demonstrates creating a financial model with proper
formatting, formulas, and validation.
"""

import tempfile
from pathlib import Path
from office_skill import XlsxHandler


def create_income_statement(xls, year):
    """Create an income statement sheet."""
    xls.add_sheet(name=f"IS {year}")

    # Header
    xls.set_cell(f"IS {year}", "A1", f"Income Statement {year}", formula=False)
    xls.merge_cells(f"IS {year}", "A1:D1")

    # Column headers
    headers = ["", "Q1", "Q2", "Q3", "Q4", "Total"]
    for col, header in enumerate(headers, start=1):
        xls.set_cell(f"IS {year}", f"{chr(64+col)}2", header, formula=False)

    # Revenue section
    xls.set_cell(f"IS {year}", "A3", "Revenue", formula=False)
    revenue_data = [1000000, 1100000, 1200000, 1300000]
    for q, value in enumerate(revenue_data, start=1):
        xls.set_cell(f"IS {year}", f"{chr(64+q+1)}3", str(value), formula=False)

    # COGS (60% of revenue)
    xls.set_cell(f"IS {year}", "A4", "Cost of Goods Sold", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}4", f"={col}3*0.6", formula=True)

    # Gross Profit
    xls.set_cell(f"IS {year}", "A5", "Gross Profit", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}5", f"={col}3-{col}4", formula=True)

    # Operating Expenses (fixed)
    xls.set_cell(f"IS {year}", "A6", "Operating Expenses", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}6", "200000", formula=False)

    # EBITDA
    xls.set_cell(f"IS {year}", "A7", "EBITDA", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}7", f"={col}5-{col}6", formula=True)

    # Depreciation
    xls.set_cell(f"IS {year}", "A8", "Depreciation", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}8", "50000", formula=False)

    # EBIT
    xls.set_cell(f"IS {year}", "A9", "EBIT", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}9", f"={col}7-{col}8", formula=True)

    # Interest
    xls.set_cell(f"IS {year}", "A10", "Interest Expense", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}10", "30000", formula=False)

    # EBT
    xls.set_cell(f"IS {year}", "A11", "EBT", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}11", f"={col}9-{col}10", formula=True)

    # Taxes (25%)
    xls.set_cell(f"IS {year}", "A12", "Taxes", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}12", f"={col}11*0.25", formula=True)

    # Net Income
    xls.set_cell(f"IS {year}", "A13", "Net Income", formula=False)
    for q in range(1, 5):
        col = chr(64 + q + 1)
        xls.set_cell(f"IS {year}", f"{col}13", f"={col}11-{col}12", formula=True)

    # Totals column
    xls.set_cell(f"IS {year}", "F3", "=SUM(B3:E3)", formula=True)
    xls.set_cell(f"IS {year}", "F4", "=SUM(B4:E4)", formula=True)
    xls.set_cell(f"IS {year}", "F5", "=SUM(B5:E5)", formula=True)
    xls.set_cell(f"IS {year}", "F6", "=SUM(B6:E6)", formula=True)
    xls.set_cell(f"IS {year}", "F7", "=SUM(B7:E7)", formula=True)
    xls.set_cell(f"IS {year}", "F8", "=SUM(B8:E8)", formula=True)
    xls.set_cell(f"IS {year}", "F9", "=SUM(B9:E9)", formula=True)
    xls.set_cell(f"IS {year}", "F10", "=SUM(B10:E10)", formula=True)
    xls.set_cell(f"IS {year}", "F11", "=SUM(B11:E11)", formula=True)
    xls.set_cell(f"IS {year}", "F12", "=SUM(B12:E12)", formula=True)
    xls.set_cell(f"IS {year}", "F13", "=SUM(B13:E13)", formula=True)

    # Apply formatting
    # Revenue and COGS (blue - inputs)
    for row in [3, 4]:
        for q in range(1, 5):
            col = chr(64 + q + 1)
            xls.set_cell_format(f"IS {year}", f"{col}{row}", "$#,##0")

    # Calculations (black - formulas)
    for row in [5, 7, 9, 11, 12, 13]:
        for q in range(1, 6):  # Include total column
            col = chr(64 + q + 1)
            xls.set_cell_format(f"IS {year}", f"{col}{row}", "$#,##0")

    # Fixed expenses (blue)
    for row in [6, 8, 10]:
        for q in range(1, 5):
            col = chr(64 + q + 1)
            xls.set_cell_format(f"IS {year}", f"{col}{row}", "$#,##0")

    return xls


def create_balance_sheet(xls, year):
    """Create a balance sheet."""
    xls.add_sheet(name=f"BS {year}")

    # Header
    xls.set_cell(f"BS {year}", "A1", f"Balance Sheet {year}", formula=False)
    xls.merge_cells(f"BS {year}", "A1:D1")

    # Structure
    xls.set_cell(f"BS {year}", "A2", "ASSETS", formula=False)
    xls.set_cell(f"BS {year}", "C2", "LIABILITIES & EQUITY", formula=False)

    # Assets
    assets = [
        ("Cash", 500000),
        ("Accounts Receivable", 300000),
        ("Inventory", 400000),
        ("PP&E", 2000000),
    ]

    for i, (account, value) in enumerate(assets, start=3):
        xls.set_cell(f"BS {year}", f"A{i}", account, formula=False)
        xls.set_cell(f"BS {year}", f"B{i}", str(value), formula=False)

    # Total Assets
    total_assets_row = len(assets) + 3
    xls.set_cell(f"BS {year}", f"A{total_assets_row}", "Total Assets", formula=False)
    xls.set_cell(f"BS {year}", f"B{total_assets_row}",
                 f"=SUM(B3:B{total_assets_row-1})", formula=True)

    # Liabilities
    liabilities = [
        ("Accounts Payable", 200000),
        ("Debt", 1000000),
    ]

    for i, (account, value) in enumerate(liabilities, start=3):
        xls.set_cell(f"BS {year}", f"C{i}", account, formula=False)
        xls.set_cell(f"BS {year}", f"D{i}", str(value), formula=False)

    # Equity (link to income statement)
    equity_start = len(liabilities) + 3
    xls.set_cell(f"BS {year}", f"C{equity_start}", "Retained Earnings", formula=False)
    xls.set_cell(f"BS {year}", f"D{equity_start}",
                 f"='IS {year}'!F13", formula=True)  # Link to net income

    # Total Liabilities & Equity
    total_row = equity_start + 1
    xls.set_cell(f"BS {year}", f"C{total_row}", "Total L&E", formula=False)
    xls.set_cell(f"BS {year}", f"D{total_row}",
                 f"=SUM(D3:D{total_row-1})", formula=True)

    # Check balance
    xls.set_cell(f"BS {year}", "A10", "Check:", formula=False)
    xls.set_cell(f"BS {year}", "B10", f"=B{total_assets_row}-D{total_row}", formula=True)

    # Formatting
    for row in range(3, total_assets_row + 1):
        xls.set_cell_format(f"BS {year}", f"B{row}", "$#,##0")

    for row in range(3, total_row + 1):
        xls.set_cell_format(f"BS {year}", f"D{row}", "$#,##0")

    return xls


def create_assumptions_sheet(xls):
    """Create assumptions sheet."""
    xls.add_sheet(name="Assumptions")

    # Header
    xls.set_cell("Assumptions", "A1", "Model Assumptions", formula=False)
    xls.merge_cells("Assumptions", "A1:B1")

    # Assumptions
    assumptions = [
        ("Revenue Growth QoQ", "5%"),
        ("COGS % of Revenue", "60%"),
        ("Operating Expenses", "200,000"),
        ("Depreciation", "50,000"),
        ("Interest Expense", "30,000"),
        ("Tax Rate", "25%"),
    ]

    for i, (label, value) in enumerate(assumptions, start=3):
        xls.set_cell("Assumptions", f"A{i}", label, formula=False)
        xls.set_cell("Assumptions", f"B{i}", value, formula=False)

    # Format assumptions as blue (inputs)
    for i in range(3, 3 + len(assumptions)):
        xls.set_cell_format("Assumptions", f"B{i}", "0%"
                            if "%" in assumptions[i-3][1] else "#,##0")

    return xls


def create_summary_sheet(xls, year):
    """Create summary sheet with key metrics."""
    xls.add_sheet(name="Summary")

    # Header
    xls.set_cell("Summary", "A1", f"Financial Summary {year}", formula=False)
    xls.merge_cells("Summary", "A1:C1")

    # Key Metrics
    metrics = [
        ("Revenue", f"='IS {year}'!F3"),
        ("Gross Profit", f"='IS {year}'!F5"),
        ("EBITDA", f"='IS {year}'!F7"),
        ("EBIT", f"='IS {year}'!F9"),
        ("Net Income", f"='IS {year}'!F13"),
        ("Gross Margin", f"='IS {year}'!F5/'IS {year}'!F3"),
        ("EBITDA Margin", f"='IS {year}'!F7/'IS {year}'!F3"),
        ("Net Margin", f"='IS {year}'!F13/'IS {year}'!F3"),
    ]

    for i, (label, formula) in enumerate(metrics, start=3):
        xls.set_cell("Summary", f"A{i}", label, formula=False)
        xls.set_cell("Summary", f"B{i}", formula, formula=True)

        # Format based on type
        if "Margin" in label:
            xls.set_cell_format("Summary", f"B{i}", "0.0%")
        else:
            xls.set_cell_format("Summary", f"B{i}", "$#,##0")

    return xls


def main():
    """Create a complete financial model."""
    print("Creating Financial Model Example")
    print("=" * 50)

    # Create temporary file
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
        model_path = tmp.name

    try:
        print(f"Creating model at: {model_path}")

        # Initialize handler
        xls = XlsxHandler(model_path)

        # Remove default sheet
        xls.remove_sheet("Sheet1")

        # Create sheets
        print("• Creating Income Statement...")
        create_income_statement(xls, 2026)

        print("• Creating Balance Sheet...")
        create_balance_sheet(xls, 2026)

        print("• Creating Assumptions sheet...")
        create_assumptions_sheet(xls)

        print("• Creating Summary sheet...")
        create_summary_sheet(xls, 2026)

        # Validate formulas
        print("\nValidating formulas...")
        validation = xls.validate_formulas()
        print(f"Validation status: {validation.get('status', 'unknown')}")

        # Save
        print(f"\nModel saved to: {model_path}")
        print(f"File size: {Path(model_path).stat().st_size:,} bytes")

        # Export to PDF
        pdf_path = model_path.replace(".xlsx", ".pdf")
        xls.export(pdf_path, format="pdf")
        print(f"PDF exported to: {pdf_path}")

        return model_path, pdf_path

    except Exception as e:
        print(f"Error creating model: {e}")
        raise
    finally:
        # In production, you would keep the files
        # For this example, we'll clean up
        import os
        if os.path.exists(model_path):
            os.unlink(model_path)
        pdf_path = model_path.replace(".xlsx", ".pdf")
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)
        print("\nTemporary files cleaned up.")


if __name__ == "__main__":
    main()