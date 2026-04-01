# Example 2: Generate Document from Template

This example demonstrates how to generate a new document from an existing template with variable substitution.

## Prerequisites

- Complete Example 1: Basic Template Management (you need a template in the library)

## Step 1: Create/add the template with placeholders

Create a Word document `invoice_template.docx` with Jinja2 placeholders like:

```
Invoice Number: {{ invoice_number }}
Date: {{ invoice_date }}
Client: {{ client_name }}

Items:
{% for item in items %}
  - {{ item.name }}: ${{ item.price }}
{% endfor %}

Total: ${{ total }}
```

Add it to the template library:

```bash
office-cli template add \
  --input invoice_template.docx \
  --name "business.invoice.billing.standard.v1" \
  --description "Standard business invoice template" \
  --tags "business,invoice,billing"
```

## Step 2: Generate the invoice with variables

```bash
office-cli template generate \
  --template business.invoice.billing.standard.v1 \
  --output invoice_2024_001.docx \
  invoice_number=INV-2024-001 \
  invoice_date="2024-03-30" \
  client_name="Acme Corporation"
```

Output:
```
Generated document from template: business.invoice.billing.standard.v1
Output: invoice_2024_001.docx
Variables applied: invoice_number, invoice_date, client_name
```

## Step 3: Open the generated document

The output `invoice_2024_001.docx` will have all variables substituted. Jinja2 supports loops, conditionals, and other advanced templating features.
