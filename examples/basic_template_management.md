# Example 1: Basic Template Management

This example demonstrates how to add, list, search, and delete templates using the Office Skill CLI.

## Prerequisites

```bash
# Ensure office-skill is installed
pip install -e .
# Verify CLI works
office-cli --help
```

## Step 1: Add a template

Create or obtain a Word document (e.g., `meeting_notes.docx`) and add it to the template library:

```bash
office-cli template add \
  --input meeting_notes.docx \
  --name "meeting.notes.basic.standard.v1" \
  --description "Standard meeting notes template" \
  --tags "meeting,notes,documentation"
```

Output:
```
Template added successfully: meeting.notes.basic.standard.v1
Location: ~/.office_skill/templates/meeting/notes/basic/standard/v1
Description: Standard meeting notes template
Tags: meeting, notes, documentation
```

## Step 2: List all templates

```bash
office-cli template list --verbose
```

Output:
```
Found 1 template(s):
--------------------------------------------------------------------------------
Name: meeting.notes.basic.standard.v1
Description: Standard meeting notes template
Format: docx
Created: 2024-03-30T12:34:56
  Type: Word document
  Pages: 1, Paragraphs: 8
--------------------------------------------------------------------------------
```

## Step 3: Search for templates

```bash
office-cli template search meeting
```

Output:
```
Found 1 template(s) matching 'meeting':
• meeting.notes.basic.standard.v1: Standard meeting notes template
```

## Step 4: Get template details

```bash
office-cli template get --name meeting.notes.basic.standard.v1
```

## Step 5: Delete a template (when no longer needed)

```bash
office-cli template delete --name meeting.notes.basic.standard.v1 --force
```
