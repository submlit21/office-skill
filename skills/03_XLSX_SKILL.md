---
name: xlsx
description: "Excel spreadsheet creation, editing, and analysis with support for formulas, formatting, data analysis, and visualization using cli-anything-libreoffice backend with formula validation."
license: MIT
---

# XLSX creation, editing, and analysis

## Overview

A comprehensive skill for working with Excel spreadsheets (.xlsx, .xls, .csv) using the `cli-anything-libreoffice` backend with formula preservation and validation. This skill enforces financial modeling standards and ensures zero formula errors.

## Requirements for Outputs

### All Excel Files
- **Zero Formula Errors**: Every Excel model MUST be delivered with ZERO formula errors (#REF!, #DIV/0!, #VALUE!, #N/A, #NAME?)
- **Formula Preservation**: Always use `data_only=False` when loading to preserve formulas
- **Validation Mandatory**: Always run `validate_formulas()` or external recalc.py after modifications

### Financial Models
When working with financial models, follow industry-standard conventions unless otherwise specified:

#### Color Coding Standards
- **Blue text (RGB: 0,0,255)**: Hardcoded inputs, and numbers users will change for scenarios
- **Black text (RGB: 0,0,0)**: ALL formulas and calculations
- **Green text (RGB: 0,128,0)**: Links pulling from other worksheets within same workbook
- **Red text (RGB: 255,0,0)**: External links to other files
- **Yellow background (RGB: 255,255,0)**: Key assumptions needing attention or cells that need to be updated

#### Number Formatting Standards
- **Years**: Format as text strings (e.g., "2024" not "2,024")
- **Currency**: Use `$#,##0` format; ALWAYS specify units in headers ("Revenue ($mm)")
- **Zeros**: Use number formatting to make all zeros "-", including percentages
- **Percentages**: Default to `0.0%` format (one decimal)
- **Multiples**: Format as `0.0x` for valuation multiples (EV/EBITDA, P/E)
- **Negative numbers**: Use parentheses `(123)` not minus `-123`

## Core Operations

### Basic Spreadsheet Operations
```python
from office_skill import XlsxHandler

xls = XlsxHandler("spreadsheet.xlsx")

# Get cell value
value = xls.get_cell("Sheet1", "A1")
print(f"A1 value: {value}")

# Set cell value
xls.set_cell("Sheet1", "B5", "Revenue", formula=False)
xls.set_cell("Sheet1", "C5", "1000000", formula=False)

# Set cell with formula
xls.set_cell("Sheet1", "D5", "SUM(C5:C10)", formula=True)
```

### Sheet Management
```python
# List all sheets
sheets = xls.list_sheets()
print(f"Workbook has {len(sheets['sheets'])} sheets")

# Add new sheet
xls.add_sheet(name="Analysis", index=1)

# Rename sheet
xls.rename_sheet(old_name="Sheet1", new_name="Data")

# Remove sheet
xls.remove_sheet("Sheet2")
```

### Formatting
```python
# Apply financial formatting
xls.set_cell_format("Sheet1", "B5", "$#,##0")  # Currency
xls.set_cell_format("Sheet1", "C5", "0.0%")    # Percentage

# Apply to multiple cells
currency_cells = ["B5", "B6", "B7", "B8"]
xls.apply_financial_formatting("Sheet1", currency_cells=currency_cells)
```

### Cell Merging
```python
# Merge cells
xls.merge_cells("Sheet1", "A1:D1")  # Merge first row for title

# Unmerge cells
xls.unmerge_cells("Sheet1", "A1")
```

## Advanced Workflows

### Financial Model Creation
```python
def create_financial_model():
    """Create a basic three-statement financial model."""
    xls = XlsxHandler("financial_model.xlsx")

    # Income Statement sheet
    xls.add_sheet(name="Income Statement")
    xls.set_cell("Income Statement", "A1", "Income Statement", formula=False)
    xls.set_cell("Income Statement", "A3", "Revenue", formula=False)
    xls.set_cell("Income Statement", "B3", "1000000", formula=False)
    xls.set_cell("Income Statement", "A4", "COGS", formula=False)
    xls.set_cell("Income Statement", "B4", "=B3*0.6", formula=True)  # 60% of revenue
    xls.set_cell("Income Statement", "A5", "Gross Profit", formula=False)
    xls.set_cell("Income Statement", "B5", "=B3-B4", formula=True)

    # Apply formatting
    xls.set_cell_format("Income Statement", "B3", "$#,##0")
    xls.set_cell_format("Income Statement", "B5", "$#,##0")

    # Validate formulas
    validation = xls.validate_formulas()
    if validation["status"] != "success":
        print(f"Formula errors: {validation}")

    return xls
```

### Data Analysis and Reporting
```python
def analyze_and_report(data_path):
    """Load data, analyze, and create summary report."""
    # Load data (simplified)
    xls = XlsxHandler(data_path)

    # Create analysis sheet
    xls.add_sheet(name="Analysis")

    # Calculate summary statistics
    xls.set_cell("Analysis", "A1", "Data Analysis", formula=False)
    xls.set_cell("Analysis", "A3", "Count", formula=False)
    xls.set_cell("Analysis", "B3", "=COUNTA(Data!A:A)", formula=True)
    xls.set_cell("Analysis", "A4", "Average", formula=False)
    xls.set_cell("Analysis", "B4", "=AVERAGE(Data!B:B)", formula=True)
    xls.set_cell("Analysis", "A5", "Std Dev", formula=False)
    xls.set_cell("Analysis", "B5", "=STDEV(Data!B:B)", formula=True)

    return xls
```

### Template Management for Excel Spreadsheets

#### Template Naming Convention
Excel templates use the format: `domain.type.purpose.variant.version`

**Examples:**
- `finance.model.three_statement.standard.v1` - Three-statement financial model
- `business.budget.annual.comprehensive.v2` - Annual business budget
- `sales.dashboard.quarterly.analytics.v1` - Sales performance dashboard

#### Adding XLSX Templates
```python
from office_skill import TemplateManager

manager = TemplateManager()

# Add a financial model template
template_data = manager.add_template(
    source_path="financial_model.xlsx",
    name="finance.model.three_statement.standard.v1",
    description="Three-statement financial model with income statement, balance sheet, cash flow",
    tags=["finance", "model", "accounting", "financial"]
)

print(f"Template added: {template_data['name']}")
print(f"Sheets: {template_data['analysis'].get('sheets', 0)}")
print(f"Formulas: {template_data['analysis'].get('formulas', 0)}")
```

#### Generating Spreadsheets from Templates
```python
# Generate a budget from template
budget_data = {
    "company_name": "Acme Inc",
    "fiscal_year": "2026",
    "department": "Marketing",
    "currency": "USD",
    "scenario": "Conservative"
}

result = manager.generate_from_template(
    template_name="business.budget.annual.comprehensive.v1",
    output_path="acme_marketing_budget_2026.xlsx",
    variables=budget_data
)

print(f"Generated: {result['output_path']}")
```

#### Template Analysis for Excel
```python
# Analyze an Excel document for template suitability
analysis = manager.analyze_document_structure("spreadsheet.xlsx")
print(f"Document type: {analysis['type']}")
print(f"Sheets: {analysis['sheets']}")
print(f"Cells: {analysis['cells']}")
print(f"Formulas: {analysis['formulas']}")

# Check if suitable for template library
if analysis['sheets'] > 0 and analysis['formulas'] > 10:
    print("Spreadsheet suitable for template library")
```

#### Example: Financial Model Generation
```python
def generate_financial_model(company, year, scenario, output_path):
    """Generate a financial model from template."""
    from office_skill import TemplateManager
    
    manager = TemplateManager()
    
    model_data = {
        "company_name": company,
        "model_year": year,
        "scenario_type": scenario,
        "currency": "USD",
        "units": "thousands",
        "prepared_by": "Financial Analyst",
        "as_of_date": "2026-03-20"
    }
    
    result = manager.generate_from_template(
        template_name="finance.model.three_statement.standard.v1",
        output_path=output_path,
        variables=model_data
    )
    
    # Validate formulas after generation
    from office_skill import XlsxHandler
    xls = XlsxHandler(output_path)
    validation = xls.validate_formulas()
    
    if validation.get("status") == "success":
        print(f"Financial model generated with zero formula errors")
    else:
        print(f"Warning: Formula errors detected: {validation.get('errors', [])}")
    
    return result
```

### Formula Validation
```python
# Always validate after modifications
xls = XlsxHandler("model.xlsx")

# Make changes
xls.set_cell("Sheet1", "C10", "=B10/0", formula=True)  # Potential DIV/0 error

# Validate
validation = xls.validate_formulas()
if "errors" in validation and validation["errors"]:
    print(f"Formula errors found: {validation['errors']}")
    # Fix errors before proceeding
```

#### Budget Template Generation
```python
def generate_department_budget(department, year, manager_name):
    """Generate a department budget from template."""
    budget_vars = {
        "department_name": department,
        "fiscal_year": str(year),
        "budget_manager": manager_name,
        "currency": "USD",
        "version": "1.0",
        "prepared_date": "2026-03-20"
    }
    
    output_file = f"{department.lower()}_budget_{year}.xlsx"
    
    result = manager.generate_from_template(
        template_name="business.budget.department.comprehensive.v1",
        output_path=output_file,
        variables=budget_vars
    )
    
    print(f"Budget template generated: {output_file}")
    print(f"Variables: {list(budget_vars.keys())}")
    
    return result
```

## Integration with cli-anything-libreoffice

### Direct CLI Usage
```bash
# Basic calc commands
cli-anything-libreoffice calc set-cell --sheet "Sheet1" --cell "A1" --value "Data"
cli-anything-libreoffice calc get-cell --sheet "Sheet1" --cell "A1"
cli-anything-libreoffice calc list-sheets

# Sheet operations
cli-anything-libreoffice calc add-sheet --name "NewSheet"
cli-anything-libreoffice calc rename-sheet --sheet "Sheet1" --name "Data"

# Formatting
cli-anything-libreoffice calc set-cell-format --sheet "Sheet1" --cell "B5" --format "$#,##0"
```

### Batch Operations
```bash
# Create batch file
cat > operations.txt << EOF
calc set-cell --sheet Sheet1 --cell A1 --value "Revenue"
calc set-cell --sheet Sheet1 --cell B1 --value "=SUM(B2:B10)"
calc add-sheet --name Analysis
EOF

# Execute batch
cli-anything-libreoffice batch operations.txt
```

## Critical Rules for Spreadsheet Processing

1. **Always use formulas for calculations**: Never hardcode calculated values. Use cell references in formulas.

2. **Validate all formulas**: After ANY modification, run validation to ensure zero errors.

3. **Follow financial modeling standards**: Use standard color coding and formatting unless instructed otherwise.

4. **Preserve existing templates**: Study and EXACTLY match existing format, style, and conventions.

5. **Document assumptions**: Place all assumptions in separate, clearly labeled cells or sheets.

## Examples

### Creating a Sales Dashboard
```python
def create_sales_dashboard():
    """Create a sales performance dashboard."""
    xls = XlsxHandler("sales_dashboard.xlsx")

    # Dashboard sheet
    xls.add_sheet(name="Dashboard")
    xls.set_cell("Dashboard", "A1", "Sales Dashboard", formula=False)
    xls.merge_cells("Dashboard", "A1:D1")

    # Key metrics
    metrics = [
        ("Total Revenue", "=SUM(Data!B:B)"),
        ("Average Sale", "=AVERAGE(Data!C:C)"),
        ("Top Product", "=INDEX(Data!A:A, MATCH(MAX(Data!B:B), Data!B:B, 0))"),
        ("Growth Rate", "=(SUM(Data!B:B)-SUM(Data!D:D))/SUM(Data!D:D)")
    ]

    for i, (label, formula) in enumerate(metrics, start=3):
        xls.set_cell("Dashboard", f"A{i}", label, formula=False)
        xls.set_cell("Dashboard", f"B{i}", formula, formula=True)
        xls.set_cell_format("Dashboard", f"B{i}", "$#,##0" if "Revenue" in label else "0.0%")

    return xls
```

### Budget Template
```python
def create_budget_template():
    """Create a monthly budget template."""
    xls = XlsxHandler("budget_template.xlsx")

    # Income section
    xls.set_cell("Sheet1", "A1", "Monthly Budget", formula=False)
    xls.set_cell("Sheet1", "A3", "Income", formula=False)
    xls.set_cell("Sheet1", "A4", "Salary", formula=False)
    xls.set_cell("Sheet1", "B4", "5000", formula=False)  # Input cell (blue)
    xls.set_cell("Sheet1", "A5", "Other Income", formula=False)
    xls.set_cell("Sheet1", "B5", "200", formula=False)   # Input cell

    # Expenses section
    xls.set_cell("Sheet1", "A7", "Expenses", formula=False)
    expenses = [
        ("Rent", "1500"),
        ("Utilities", "200"),
        ("Groceries", "400"),
        ("Transportation", "150"),
        ("Entertainment", "100")
    ]

    for i, (category, amount) in enumerate(expenses, start=8):
        xls.set_cell("Sheet1", f"A{i}", category, formula=False)
        xls.set_cell("Sheet1", f"B{i}", amount, formula=False)

    # Totals
    xls.set_cell("Sheet1", "A15", "Total Income", formula=False)
    xls.set_cell("Sheet1", "B15", "=SUM(B4:B5)", formula=True)  # Formula (black)
    xls.set_cell("Sheet1", "A16", "Total Expenses", formula=False)
    xls.set_cell("Sheet1", "B16", "=SUM(B8:B12)", formula=True)  # Formula
    xls.set_cell("Sheet1", "A17", "Net", formula=False)
    xls.set_cell("Sheet1", "B17", "=B15-B16", formula=True)  # Formula

    # Formatting
    xls.set_cell_format("Sheet1", "B15:B17", "$#,##0")

    return xls
```

## Troubleshooting

### Common Formula Errors
- **#REF!**: Broken cell reference (check sheet names, cell ranges)
- **#DIV/0!**: Division by zero (add IFERROR or validation)
- **#VALUE!**: Wrong data type in formula
- **#N/A**: Lookup value not found
- **#NAME?**: Unrecognized function name

### Performance Tips
- Use named ranges for complex formulas
- Avoid volatile functions (OFFSET, INDIRECT, etc.) when possible
- Use Excel Tables for dynamic ranges
- Minimize cross-sheet references

## Related Resources
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/)
- [Excel Formula Reference](https://support.microsoft.com/en-us/office/excel-functions-alphabetical-b3944572-255d-4efb-bb96-c6d90033e188)
- [Financial Modeling Standards](https://corporatefinanceinstitute.com/resources/financial-modeling/financial-modeling-best-practices/)