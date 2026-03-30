---
name: template
description: "Document template management, storage, search, and generation with hierarchical organization, markdown conversion, and variable substitution."
license: MIT
---

# Template Management for Office Documents

## Overview

This skill provides comprehensive template management for Microsoft Office documents (DOCX, XLSX, PPTX) with hierarchical storage, intelligent search, and document generation. It enables creating reusable document patterns, converting documents to analyzable formats (markdown/YAML), and generating new documents with advanced Jinja2 template variable substitution.

## Quick Start

### Installation Requirements
```bash
# Install office-skill with template dependencies
pip install office-skill

# Ensure cli-anything-libreoffice is available
pip install cli-anything-libreoffice

# Jinja2 and dependencies are automatically installed with office-skill
# For advanced templating features
```

### Basic Template Operations
```python
from office_main.core import TemplateManager

# Initialize template manager
tm = TemplateManager()

# Add a template (follows naming convention: domain.type.purpose.variant.version)
result = tm.add_template(
    source_path="席签模板.docx",
    name="meeting.signin.nameplate.chinese.v1",
    description="会议席签模板 - 中文",
    tags=["meeting", "signin", "nameplate", "chinese", "席签"]
)

# List all templates
templates = tm.list_templates()
for t in templates:
    print(f"{t['name']}: {t['description']}")

# Search templates
results = tm.search_templates("meeting")
for t in results:
    print(f"Found: {t['name']}")

# Generate document from template
generation_result = tm.generate_from_template(
    template_name="meeting.signin.nameplate.chinese.v1",
    output_path="张三_席签.docx",
    variables={"name": "张三"}
)
```

## Workflow Decision Tree

### When to Use Templates
- **Reusable document patterns**: Create templates for commonly used document structures
- **Consistent formatting**: Ensure all documents follow the same style and layout
- **Batch document generation**: Produce multiple documents with similar structure but different content
- **Document analysis**: Convert Office documents to markdown/YAML for AI analysis

### Template Management Workflows
- **Adding new templates**: Use `add_template()` with proper naming convention
- **Organizing templates**: Hierarchical structure based on domain.type.purpose.variant.version
- **Finding templates**: Use `list_templates()` with filtering or `search_templates()` for text search
- **Template analysis**: Convert documents to markdown for content review
- **Document generation**: Use `generate_from_template()` with variable substitution

### Naming Convention Decisions
- **Simple documents**: Use minimal hierarchy (e.g., `report.financial.quarterly.v1`)
- **Domain-specific templates**: Use full hierarchy (e.g., `meeting.signin.nameplate.chinese.v1`)
- **Version control**: Always include version suffix for template updates
- **Multi-language support**: Include language in variant (e.g., `chinese`, `english`)

## Core Components

### Template Storage System
- **Hierarchical Organization**: Templates stored as `domain/type/purpose/variant/version/` directories
- **Metadata Management**: Each template has `metadata.json` with analysis, tags, and creation info
- **Content Conversion**:
  - DOCX → Markdown for text analysis
  - XLSX/PPTX → YAML for structured content
- **Original Preservation**: Source document saved as `original.{ext}` for regeneration

### Template Analyzer
- **Document Structure Analysis**: Extracts pages, paragraphs, tables, sheets, slides, etc.
- **Markdown Conversion**: Uses LibreOffice or pandoc to convert documents to readable format
- **Format Detection**: Identifies document type and extracts key metadata
- **Content Extraction**: Converts Office formatting to plain text for AI processing

### Template Search Engine
- **Full-text Search**: Searches template names, descriptions, and tags
- **Hierarchical Filtering**: Filter by domain, type, or purpose
- **Fuzzy Matching**: Case-insensitive search with partial matches
- **Result Ranking**: Prioritizes exact matches then partial matches

### Template Generator
- **Document Copying**: Creates new documents from template originals
- **Variable Substitution**: Supports `{{variable}}` placeholders in generated documents
- **Format Preservation**: Maintains all original formatting and styles
- **Batch Processing**: Can generate multiple documents from same template

## Implementation Details

### API Reference

The `TemplateManager` class provides the main interface for template operations. Below are its public methods:

| Method | Parameters | Returns | Description |
|--------|------------|---------|-------------|
| `add_template(source_path, name, description=None, tags=None, overwrite=False)` | `source_path`: Path to source document<br>`name`: Template name in `domain.type.purpose.variant.version` format<br>`description`: Optional description<br>`tags`: Optional list of tags<br>`overwrite`: Whether to overwrite existing template (default: False) | Dictionary with template metadata | Adds a new template to the library, analyzes document structure, converts to markdown/YAML, and stores with metadata |
| `get_template(name)` | `name`: Template name | Complete template data including metadata, analysis, and markdown content | Retrieves a template by its full name |
| `list_templates(domain=None, type_filter=None, purpose=None)` | `domain`: Filter by domain<br>`type_filter`: Filter by type<br>`purpose`: Filter by purpose | List of template metadata dictionaries | Lists templates with optional hierarchical filtering |
| `search_templates(query)` | `query`: Search string | List of matching template metadata dictionaries | Searches templates by name, description, or tags (case-insensitive) |
| `delete_template(name, force=False)` | `name`: Template name<br>`force`: Skip confirmation (default: False) | True if deleted | Removes a template from storage, cleaning up empty parent directories |
| `generate_from_template(template_name, output_path, variables=None)` | `template_name`: Name of template to use<br>`output_path`: Path for generated document<br>`variables`: Optional dictionary of `{placeholder: value}` pairs | Dictionary with generation results | Creates a new document from a template, optionally substituting placeholders |
| `convert_to_markdown(source_path)` | `source_path`: Path to Office document | Markdown content as string | Converts any supported Office document to markdown format |
| `analyze_document_structure(source_path)` | `source_path`: Path to Office document | Dictionary with document analysis | Analyzes document structure (pages, paragraphs, sheets, slides, etc.) |

### Storage Structure

Each template is stored in a hierarchical directory under the templates root (default: `templates/`):

```
templates/
└── domain/
    └── type/
        └── purpose/
            └── variant/
                └── version/
                    ├── metadata.json
                    ├── template.md (for DOCX) or template.yaml (for XLSX/PPTX)
                    └── original.{ext}
```

**metadata.json** contains:
- `name`: Full template name
- `original_name`: Original filename
- `description`: Template description
- `tags`: List of tags
- `created`: ISO timestamp
- `format`: Document extension (docx, xlsx, pptx)
- `analysis`: Document structure analysis (type-specific metrics)
- `components`: Parsed name components (domain, type, purpose, variant, version)

**template.md** (for Word documents) contains markdown conversion of the document content using pandoc or LibreOffice.

**template.yaml** (for Excel and PowerPoint) contains structured YAML with:
- `markdown_content`: Converted markdown/text content
- `source_format`: Original file format

**original.{ext}** is a copy of the source document used for generation.

### Variable Substitution

Placeholder syntax: `{{variable_name}}` (double curly braces).

- **Word documents (DOCX/DOC)**: Placeholders in the document content are replaced directly using `cli-anything-libreoffice` session commands. The system searches text items and updates them in the actual document file.
- **Excel and PowerPoint**: Variable substitution is only applied to the cached markdown/YAML content, not the original document. The generated file is an exact copy of the original; placeholders in the original file are not modified.
- **Multiple occurrences**: All occurrences of a placeholder are replaced.
- **Unmatched placeholders**: If no matching placeholder is found, the variable is ignored (no error).

### Conversion Pipeline

Document conversion to markdown uses the following priority:

1. **pandoc** (if installed): Preferred for Word documents, produces cleaner markdown.
2. **LibreOffice** (via `libreoffice --headless --convert-to md`): Fallback for all document types.
3. **Simple fallback**: If both conversion tools fail, a minimal markdown placeholder is generated.

The conversion process is handled by `TemplateAnalyzer.convert_to_markdown()`, which catches exceptions and ensures a string is always returned.

### Error Handling and Fallbacks

- **Missing conversion tools**: If neither pandoc nor LibreOffice is available, a fallback markdown message is generated.
- **Document analysis failures**: If `cli-anything-libreoffice` analysis commands fail, partial analysis data with error information is stored.
- **Variable substitution failures**: If placeholder replacement in Word documents fails, the system falls back to copying the template without substitution and logs the error.
- **Storage errors**: File operations are wrapped in try-except blocks; errors are propagated to the caller.

## Critical Rules

1. **Follow Naming Convention**: All templates MUST use `domain.type.purpose.variant.version` format
2. **Preserve Original Files**: Always keep source documents for regeneration capability
3. **Validate Excel Formulas**: When adding Excel templates, ensure zero formula errors
4. **Convert for Analysis**: All templates should have markdown/YAML representation for AI processing
5. **Document Template Purpose**: Include clear description and tags for discoverability
6. **Version Control**: Update version number when modifying existing templates
7. **Consistent Metadata**: Ensure all templates have complete metadata.json files

## Template Operations

### Adding Templates
```python
from office_main.core import TemplateManager

tm = TemplateManager()

# Add Word template with markdown conversion
word_result = tm.add_template(
    source_path="report_template.docx",
    name="business.report.quarterly.standard.v1",
    description="Standard quarterly business report template",
    tags=["business", "report", "quarterly", "financial"]
)

# Add Excel template (formulas validated automatically)
excel_result = tm.add_template(
    source_path="financial_model.xlsx",
    name="finance.model.budget.annual.v1",
    description="Annual budget financial model template",
    tags=["finance", "model", "budget", "excel"]
)

# Add PowerPoint template
ppt_result = tm.add_template(
    source_path="presentation_template.pptx",
    name="marketing.presentation.product.launch.v1",
    description="Product launch presentation template",
    tags=["marketing", "presentation", "product", "launch"]
)
```

### Searching and Filtering
```python
# List all templates
all_templates = tm.list_templates()

# Filter by domain
business_templates = tm.list_templates(domain="business")

# Filter by domain and type
financial_reports = tm.list_templates(domain="finance", type_filter="report")

# Search by keyword
search_results = tm.search_templates("quarterly")

# Complex search with filtering
filtered_search = [
    t for t in tm.list_templates(domain="business")
    if "report" in t["description"].lower()
]
```

### Template Analysis
```python
# Get template details
template_details = tm.get_template("business.report.quarterly.standard.v1")
print(f"Template: {template_details['name']}")
print(f"Description: {template_details['description']}")
print(f"Format: {template_details['format']}")
print(f"Created: {template_details['created']}")

# View markdown content (for DOCX templates)
if template_details.get("markdown_content"):
    print(f"Markdown preview (first 200 chars):")
    print(template_details["markdown_content"][:200])

# View analysis data
analysis = template_details.get("analysis", {})
if analysis.get("type") == "word":
    print(f"Pages: {analysis.get('pages', 0)}")
    print(f"Paragraphs: {analysis.get('paragraphs', 0)}")
    print(f"Tables: {analysis.get('tables', 0)}")
```

### Document Generation
```python
# Simple generation (copy template)
result = tm.generate_from_template(
    template_name="business.report.quarterly.standard.v1",
    output_path="Q1_2024_Report.docx"
)

# Generation with variables
result = tm.generate_from_template(
    template_name="meeting.signin.nameplate.chinese.v1",
    output_path="李四_席签.docx",
    variables={"name": "李四"}
)

# Multiple variable substitution
result = tm.generate_from_template(
    template_name="contract.template.services.standard.v1",
    output_path="Client_XYZ_Contract.docx",
    variables={
        "client_name": "XYZ Corporation",
        "effective_date": "2024-03-26",
        "service_description": "Cloud hosting services",
        "amount": "$10,000"
    }
)
 ```

### Advanced Jinja2 Templating

Office-skill now supports Jinja2 templating engine for advanced variable substitution:

```python
# Jinja2 expressions in variable values
result = tm.generate_from_template(
    template_name="invoice.template.standard.v1",
    output_path="Invoice_123.docx",
    variables={
        "customer": "ABC Company",
        "items": [
            {"name": "Product A", "quantity": 2, "price": 100},
            {"name": "Product B", "quantity": 1, "price": 250}
        ],
        # Variables can reference other variables using Jinja2 syntax
        "subtotal": "{{ items[0].quantity * items[0].price + items[1].quantity * items[1].price }}",
        "tax_rate": "0.08",
        "tax": "{{ subtotal|float * tax_rate|float }}",
        "total": "{{ subtotal|float + tax|float }}",
        "formatted_total": "${{ '%0.2f'|format(total|float) }}",
        # Conditional formatting
        "discount_note": "{% if subtotal|float > 500 %}Volume discount applied{% else %}{% endif %}"
    }
)

# Jinja2 control structures in template markdown
# Template content can include:
# - Conditionals: {% if condition %}...{% endif %}
# - Loops: {% for item in items %}...{% endfor %}
# - Filters: {{ value|upper }}, {{ number|round(2) }}
# - Math: {{ a + b }}, {{ total / count }}

# The system automatically detects Jinja2 syntax and renders variables accordingly.
# If Jinja2 rendering fails, it falls back to simple {{variable}} replacement.
```

**Key Features:**
1. **Variable Interpolation**: Variables can reference other variables using `{{var_name}}` syntax
2. **Expressions**: Support for mathematical operations, comparisons, and logical expressions
3. **Filters**: Use Jinja2 filters like `|upper`, `|lower`, `|length`, `|float`, `|int`
4. **Control Structures**: Conditionals (`if/else`) and loops (`for`) in template content
5. **Error Handling**: Falls back to simple replacement if Jinja2 syntax errors occur

**Note**: For DOCX documents, variable substitution happens in the actual document content.
For other formats or cached markdown, Jinja2 rendering is applied to the markdown content.

## CLI Operations

### Template Management via CLI
```bash
# Add template
python src/office_main/cli/office_cli.py template add \
  --input 席签模板.docx \
  --name meeting.signin.nameplate.chinese.v1 \
  --description "会议席签模板 - 中文" \
  --tags "meeting,signin,nameplate,chinese,席签"

# List templates
python src/office_main/cli/office_cli.py template list

# List with filtering
python src/office_main/cli/office_cli.py template list --domain meeting

# Search templates
python src/office_main/cli/office_cli.py template search --query "nameplate"

# Get template details
python src/office_main/cli/office_cli.py template get --name meeting.signin.nameplate.chinese.v1

# Generate from template
python src/office_main/cli/office_cli.py template generate \
  --template meeting.signin.nameplate.chinese.v1 \
  --output 张三_席签.docx \
  --variables "name=张三"

# Delete template
python src/office_main/cli/office_cli.py template delete --name meeting.signin.nameplate.chinese.v1
```

## Advanced Use Cases

### Batch Template Processing
```python
import pandas as pd
from office_main.core import TemplateManager

tm = TemplateManager()

# Read data from CSV
data = pd.read_csv("attendees.csv")

# Generate documents for each attendee
for index, row in data.iterrows():
    result = tm.generate_from_template(
        template_name="meeting.signin.nameplate.chinese.v1",
        output_path=f"{row['name']}_席签.docx",
        variables={"name": row["name"], "title": row["title"]}
    )
    print(f"Generated: {result['output_path']}")
```

### Template Quality Assurance
```python
def validate_template_system(tm):
    """Validate all templates in the system."""
    templates = tm.list_templates()

    for template in templates:
        try:
            details = tm.get_template(template["name"])

            # Check metadata completeness
            required_fields = ["name", "description", "format", "analysis"]
            for field in required_fields:
                if field not in details:
                    print(f"WARNING: Template {template['name']} missing {field}")

            # Check markdown conversion
            if not details.get("markdown_content"):
                print(f"WARNING: Template {template['name']} has no markdown content")

            # Check original file exists
            if not details.get("original_file"):
                print(f"WARNING: Template {template['name']} has no original file")

        except Exception as e:
            print(f"ERROR: Failed to validate template {template['name']}: {e}")
```

### Custom Template Workflows
```python
class CustomTemplateWorkflow:
    def __init__(self, tm):
        self.tm = tm

    def create_document_family(self, base_template, variations):
        """Create multiple variations from a base template."""
        results = []

        for var_name, variables in variations.items():
            result = self.tm.generate_from_template(
                template_name=base_template,
                output_path=f"{var_name}_document.docx",
                variables=variables
            )
            results.append(result)

        return results

    def export_template_catalog(self, output_file="template_catalog.md"):
        """Export all templates as a markdown catalog."""
        templates = self.tm.list_templates()

        with open(output_file, "w", encoding="utf-8") as f:
            f.write("# Template Catalog\n\n")

            # Group by domain
            domains = {}
            for template in templates:
                domain = template["name"].split(".")[0]
                if domain not in domains:
                    domains[domain] = []
                domains[domain].append(template)

            # Write by domain
            for domain, domain_templates in domains.items():
                f.write(f"## {domain.title()}\n\n")
                for template in domain_templates:
                    f.write(f"### {template['name']}\n")
                    f.write(f"- **Description**: {template['description']}\n")
                    f.write(f"- **Format**: {template['format']}\n")
                    f.write(f"- **Created**: {template['created']}\n\n")

        return output_file
```

## Troubleshooting

### Common Issues

1. **Template Name Format Error**
   ```
   ValueError: Template name must be in format 'domain.type.purpose.variant.version'
   ```
   **Solution**: Ensure template name has exactly 5 parts separated by dots.

2. **Markdown Conversion Failure**
   ```
   Conversion failed, no markdown preview available.
   ```
   **Solution**: Install LibreOffice (`libreoffice --headless`) or pandoc for better conversion.

3. **Template Already Exists**
   ```
   FileExistsError: Template 'name' already exists.
   ```
   **Solution**: Use `overwrite=True` parameter or choose a different name/version.

4. **Missing Original File**
   ```
   FileNotFoundError: Source file not found
   ```
   **Solution**: Verify source file path is correct and accessible.

5. **Variable Substitution Not Working**
   **Solution**: Ensure template document contains `{{variable_name}}` placeholders.

### Debug Mode
```python
# Enable detailed logging
import logging
logging.basicConfig(level=logging.DEBUG)

# Test template conversion
from office_main.core import TemplateManager
tm = TemplateManager()

# Convert individual document to markdown
markdown = tm.convert_to_markdown("document.docx")
print(f"Converted markdown length: {len(markdown)} characters")

# Analyze document structure
analysis = tm.analyze_document_structure("document.docx")
print(f"Document analysis: {analysis}")
```

## Integration with Other Skills

### Combining with DOCX Skill
```python
from office_main.core import TemplateManager, DocxHandler

tm = TemplateManager()

# Generate document from template
result = tm.generate_from_template(
    template_name="business.report.quarterly.standard.v1",
    output_path="Q1_Report.docx",
    variables={"quarter": "Q1", "year": "2024"}
)

# Further edit with DOCX skill
doc = DocxHandler("Q1_Report.docx")
doc.add_paragraph("Additional analysis specific to Q1...")
```

### Combining with XLSX Skill
```python
from office_main.core import TemplateManager, XlsxHandler

tm = TemplateManager()

# Generate financial model from template
result = tm.generate_from_template(
    template_name="finance.model.budget.annual.v1",
    output_path="2024_Budget.xlsx"
)

# Validate formulas after generation
xls = XlsxHandler("2024_Budget.xlsx")
validation = xls.validate_formulas()
if validation.get("status") == "error":
    print(f"Formula errors found: {validation.get('errors')}")
```

## Best Practices

### Template Design
1. **Use Placeholders**: Include `{{variable}}` placeholders for dynamic content
2. **Standardize Styles**: Use consistent formatting across similar templates
3. **Include Instructions**: Add comments or notes in templates for users
4. **Test Generation**: Always test template generation with sample variables
5. **Version Control**: Increment version number for template updates

### Organization Strategy
1. **Domain Hierarchy**: Organize by business domain (finance, marketing, hr, etc.)
2. **Type Classification**: Classify by document type (report, contract, presentation, etc.)
3. **Purpose Tags**: Add descriptive tags for searchability
4. **Variant Management**: Create variants for different languages or formats
5. **Version History**: Maintain version history for template evolution

### Performance Optimization
1. **Batch Operations**: Use batch processing for multiple document generation
2. **Caching**: Cache template metadata for frequent access
3. **Indexing**: Periodically rebuild search indexes for large template libraries
4. **Cleanup**: Remove unused or obsolete templates
5. **Backup**: Regularly backup template directory