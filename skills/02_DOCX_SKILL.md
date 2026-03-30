---
name: docx
description: "Word document creation, editing, and analysis with support for tracked changes, comments, formatting preservation, and text extraction using cli-anything-libreoffice backend."
license: MIT
---

# DOCX creation, editing, and analysis

## Overview

A comprehensive skill for working with Word documents (.docx, .doc) using the `cli-anything-libreoffice` backend with OOXML awareness. This skill supports both simple document operations and complex tracked changes workflows.

## Workflow Decision Tree

### Reading/Analyzing Content
- **Quick text extraction**: Use `DocxHandler.analyze_structure()` for overview
- **Detailed analysis**: Unpack with OOXML and examine raw XML structure
- **Tracked changes**: Convert to markdown with pandoc to see all changes

### Creating New Document
- **From scratch**: Use `DocxHandler()` with sequence of add_* methods
- **From template**: Use `create_from_template()` with existing document
- **Structured creation**: Build document outline first, then fill content

### Editing Existing Document
- **Your own document + simple changes**: Direct handler methods
- **Someone else's document**: Redlining workflow (recommended default)
- **Legal/academic/business docs**: Redlining workflow (required)

## Core Operations

### Document Creation and Basic Editing
```python
from office_skill import DocxHandler

doc = DocxHandler("document.docx")

# Add document structure
doc.add_heading("Document Title", level=1)
doc.add_paragraph("Introduction paragraph...")

# Add lists
doc.add_list(["First item", "Second item", "Third item"])

# Add tables
doc.add_table(rows=3, cols=4)
doc.set_table_cell(table_index=0, row=0, col=0, value="Header 1")

# Page management
doc.add_page_break()
```

### Content Management
```python
# List all content
content = doc.list_content()
print(f"Document has {len(content['items'])} content items")

# Modify existing content
doc.set_text(index=2, text="Updated paragraph text")

# Remove content
doc.remove(index=5)
```

### Export and Conversion
```python
# Export to PDF
doc.export("document.pdf", format="pdf")

# Export to other formats
doc.export("document.odt", format="odt")
doc.export("document.html", format="html")
```

## Advanced Workflows

### Redlining (Tracked Changes)
For professional documents with tracked changes:

1. **Extract changes**:
```bash
pandoc --track-changes=all document.docx -o changes.md
```

2. **Analyze change batches**:
```python
# Group changes into logical batches (3-10 changes each)
# Implement changes systematically
```

3. **Apply changes with OOXML awareness**:
```python
# Use unpack/pack workflow from reference project
# Or implement with cli-anything-libreoffice session management
```

### Template Management for Word Documents

#### Template Naming Convention
Word templates use the format: `domain.type.purpose.variant.version`

**Examples:**
- `business.report.quarterly.standard.v1` - Quarterly business report
- `legal.contract.nda.comprehensive.v2` - Non-disclosure agreement
- `education.syllabus.course.basic.v1` - Course syllabus template

#### Adding DOCX Templates
```python
from office_skill import TemplateManager

manager = TemplateManager()

# Add a Word document template
template_data = manager.add_template(
    source_path="business_report.docx",
    name="business.report.quarterly.standard.v1",
    description="Standard quarterly business report template",
    tags=["business", "report", "quarterly", "financial"]
)

print(f"Template added: {template_data['name']}")
print(f"Analysis: {template_data['analysis']}")
```

#### Generating Documents from Templates
```python
# Generate a new report from template
report_data = {
    "company_name": "Acme Corporation",
    "quarter": "Q4 2026",
    "fiscal_year": "2026",
    "prepared_by": "Jane Smith",
    "approval_date": "March 20, 2026"
}

result = manager.generate_from_template(
    template_name="business.report.quarterly.standard.v1",
    output_path="acme_q4_report.docx",
    variables=report_data
)

print(f"Generated: {result['output_path']}")
```

#### Template Analysis
```python
# Analyze a Word document for template suitability
analysis = manager.analyze_document_structure("document.docx")
print(f"Document type: {analysis['type']}")
print(f"Pages: {analysis['pages']}")
print(f"Paragraphs: {analysis['paragraphs']}")
print(f"Tables: {analysis['tables']}")

# Check if suitable for template library
if analysis['pages'] > 0 and analysis['paragraphs'] > 5:
    print("Document suitable for template library")
```

### Template-Based Document Generation
```python
def generate_report(template_name, data, output_path):
    """Generate document from template with data."""
    from office_skill import TemplateManager
    
    manager = TemplateManager()
    
    # Generate from template
    result = manager.generate_from_template(
        template_name=template_name,
        output_path=output_path,
        variables=data
    )
    
    # Log generation details
    print(f"Generated {output_path} from template {template_name}")
    print(f"Variables applied: {list(data.keys())}")
    
    return result
```

#### Example: Business Letter Generation
```python
def generate_business_letter(recipient, position, company, address, date):
    """Generate a formal business letter."""
    letter_data = {
        "recipient_name": recipient,
        "recipient_position": position,
        "recipient_company": company,
        "recipient_address": address,
        "letter_date": date,
        "sender_name": "John Doe",
        "sender_position": "Director of Operations",
        "sender_company": "Acme Inc"
    }
    
    return generate_report(
        template_name="business.letter.formal.standard.v1",
        data=letter_data,
        output_path=f"letter_to_{recipient.replace(' ', '_')}.docx"
    )
```

### Document Analysis
```python
analysis = doc.analyze_structure()
print(f"Document: {analysis['document']}")
print(f"Content items: {len(analysis['content_summary']['items'])}")
print(f"Has tables: {analysis['has_tables']}")
print(f"Has headings: {analysis['has_headings']}")
```

## Integration with cli-anything-libreoffice

### Direct CLI Usage
```bash
# Basic writer commands
cli-anything-libreoffice writer add-paragraph --text "Paragraph text"
cli-anything-libreoffice writer add-heading --text "Heading" --level 2
cli-anything-libreoffice writer list

# Table operations
cli-anything-libreoffice writer add-table --rows 5 --cols 3
cli-anything-libreoffice writer set-table-cell --table-index 0 --row 1 --col 1 --value "Data"

# Export
cli-anything-libreoffice export render --input document.docx --output document.pdf --preset pdf
```

### Session Management
```bash
# Start session
cli-anything-libreoffice --project session.json document open document.docx

# Execute operations
cli-anything-libreoffice --project session.json writer add-paragraph --text "New content"

# Save and close
cli-anything-libreoffice --project session.json document save
```

## Critical Rules for DOCX Processing

1. **Minimal edits for redlining**: Only changed text gets `<w:ins>` and `<w:del>` tags. Preserve original RSIDs for unchanged runs.

2. **Preserve document structure**: Always analyze existing structure before making changes. Maintain consistent heading hierarchy.

3. **Use proper formatting**: Apply styles consistently. Don't use direct formatting when styles are available.

4. **Validate output**: Always check exported documents for formatting preservation.

## Examples

### Creating a Business Letter
```python
letter = DocxHandler("business_letter.docx")
letter.add_paragraph("Company Name", index=0)
letter.add_paragraph("123 Business Street", index=1)
letter.add_paragraph("City, State ZIP", index=2)
letter.add_paragraph("", index=3)  # Empty line
letter.add_paragraph("Date: March 20, 2026", index=4)
letter.add_paragraph("", index=5)
letter.add_paragraph("Dear Recipient,", index=6)
letter.add_paragraph("", index=7)
letter.add_paragraph("This is the body of the business letter...", index=8)
letter.export("business_letter.pdf")
```

### Generating a Report with Sections
```python
report = DocxHandler("monthly_report.docx")

# Title
report.add_heading("Monthly Performance Report", level=1)

# Executive Summary
report.add_heading("Executive Summary", level=2)
report.add_paragraph("Summary of key metrics...")

# Financials
report.add_heading("Financial Performance", level=2)
table_idx = report.add_table(rows=6, cols=3)["table_index"]
report.set_table_cell(table_idx, 0, 0, "Metric")
report.set_table_cell(table_idx, 0, 1, "Current Month")
report.set_table_cell(table_idx, 0, 2, "YTD")

# Appendices
report.add_heading("Appendices", level=2)
report.add_paragraph("Detailed data available upon request.")
```

## Troubleshooting

### Common Issues
- **Document corruption**: Always use session management for multiple operations
- **Formatting loss**: Use styles instead of direct formatting
- **Tracked changes not visible**: Ensure proper OOXML tagging with `<w:ins>`/`<w:del>`

### Performance Tips
- Use batch operations for multiple changes
- Close sessions when done to free resources
- For large documents, work in sections

## Related Resources
- [OOXML Documentation](https://learn.microsoft.com/en-us/office/open-xml/open-xml-sdk)
- [pandoc for document conversion](https://pandoc.org/)
- [LibreOffice Writer Guide](https://documentation.libreoffice.org/en/english-documentation/)