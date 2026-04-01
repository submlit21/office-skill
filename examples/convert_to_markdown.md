# Example 4: Convert Office Document to Markdown

This example demonstrates converting an Office document to Markdown format for documentation purposes.

## Usage

```bash
# The convert_to_markdown function is available via Python API
python3 -c "
from pathlib import Path
from office_main.core import DocxHandler
from template_manager.analyzer import TemplateAnalyzer
from office_main.core.cli_wrapper import LibreOfficeCLI

# Analyze and convert
doc_path = Path('document.docx')
handler = DocxHandler(str(doc_path))
analysis = handler.analyze_structure()

# Get markdown content
markdown = analysis.get('markdown', '')
print(markdown)

# Save to file
with open('document.md', 'w') as f:
    f.write(markdown)
"
```

## What it does

1. First tries `cli-anything-libreoffice` to export the document as text
2. If that fails or produces too little content, falls back to pandoc
3. If pandoc isn't found, tries direct libreoffice conversion
4. Returns the Markdown content for use in templates or documentation

## Example Output

```markdown
# My Document.docx

# Introduction

This is the first paragraph of my document.

## Second Section

This is the content of the second section.

- Bullet point 1
- Bullet point 2
```

## Notes

- Conversion quality depends on the source document and the conversion method used
- pandoc generally produces the best Markdown conversion from DOCX
- Temporary files are automatically cleaned up after conversion
