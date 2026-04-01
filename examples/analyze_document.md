# Example 3: Analyze Document Structure

This example demonstrates how to analyze an existing Office document to extract structural metadata.

## Commands

### Analyze a Word document

```bash
office-cli analyze --input document.docx
```

Example output:
```
Document: document.docx
Content items: 15
```

With JSON output:
```bash
office-cli analyze --input document.docx --json
```

JSON output includes:
- `type`: "word"
- `pages`: number of pages
- `paragraphs`: number of paragraphs
- `tables`: number of tables
- `markdown_content`: Markdown conversion of document content

### Analyze an Excel spreadsheet

```bash
office-cli analyze --input spreadsheet.xlsx
```

Output includes:
```
Document: spreadsheet.xlsx
Sheets: 3
```

JSON output includes:
- `type`: "excel"
- `sheet_count`: number of sheets
- `sheets`: list of sheet names
- `cells`: estimated total cells

### Analyze a PowerPoint presentation

```bash
office-cli analyze --input presentation.pptx
```

Output includes:
```
Document: presentation.pptx
Slides: 12
```

JSON output includes:
- `type`: "powerpoint"
- `slide_count`: number of slides
- `slide_titles`: detected slide titles
- `layouts`: layout usage statistics
