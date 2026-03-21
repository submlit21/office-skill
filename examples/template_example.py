#!/usr/bin/env python3
"""
Example script demonstrating template management functionality.

This example shows how to use the TemplateManager to:
1. Add templates from Office documents
2. List and search templates
3. Generate new documents from templates
"""

import sys
import os
from pathlib import Path

# Add parent directory to path for imports
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from src.office_skill import TemplateManager

    print("✓ TemplateManager imported successfully")
except ImportError as e:
    print(f"✗ Import error: {e}")
    print("Please install the package first: pip install -e .")
    sys.exit(1)


def demonstrate_template_management():
    """Demonstrate template management features."""

    print("=" * 60)
    print("Template Management Example")
    print("=" * 60)

    # 1. Create template manager
    manager = TemplateManager()
    print("\n1. Created TemplateManager")
    print(f"   Templates directory: {manager.templates_root}")

    # 2. List existing templates
    print("\n2. Listing existing templates:")
    templates = manager.list_templates()
    if templates:
        print(f"   Found {len(templates)} template(s):")
        for template in templates:
            print(f"   • {template['name']} ({template['format']})")
    else:
        print("   No templates found (expected for first run)")

    # 3. Demonstrate template name validation
    print("\n3. Template name validation examples:")
    test_names = [
        "business.report.quarterly.standard.v1",
        "legal.contract.nda.simple.v2",
        "education.lesson.plan.basic.v1",
        "invalid.name",  # Too few parts
        "too.many.parts.in.this.name.extra.v1",  # Too many parts
    ]

    for name in test_names:
        is_valid = manager._validate_template_name(name)
        status = "✓ Valid" if is_valid else "✗ Invalid"
        print(f"   {status}: {name}")

    # 4. Show directory structure for a template
    print("\n4. Template directory structure example:")
    example_name = "business.report.quarterly.standard.v1"
    try:
        template_dir = manager._get_template_path(example_name)
        rel_path = template_dir.relative_to(manager.templates_root)
        print(f"   Template: {example_name}")
        print(f"   Storage path: {template_dir}")
        print(f"   Hierarchical structure: {' → '.join(rel_path.parts)}")
    except Exception as e:
        print(f"   Error: {e}")

    # 5. Template search demonstration (conceptual)
    print("\n5. Template search capabilities:")
    print("   • Search by name (partial match)")
    print("   • Search by description")
    print("   • Search by tags")
    print("   • Filter by domain, type, purpose")

    # 6. CLI usage examples
    print("\n6. CLI Command Examples:")
    print("   # Add a template")
    print("   office_cli.py template add \\")
    print("     --input report.docx \\")
    print("     --name business.report.quarterly.standard.v1 \\")
    print("     --description 'Quarterly business report' \\")
    print("     --tags 'business,report,quarterly'")

    print("\n   # List templates")
    print("   office_cli.py template list --type report --verbose")

    print("\n   # Search templates")
    print("   office_cli.py template search --query 'business'")

    print("\n   # Generate from template")
    print("   office_cli.py template generate \\")
    print("     --template business.report.quarterly.standard.v1 \\")
    print("     --output q4_2023_report.docx \\")
    print("     --variables company_name='Acme Inc' quarter='Q4 2023'")

    # 7. Python API usage
    print("\n7. Python API Examples:")
    print("   ```python")
    print("   from office_skill import TemplateManager")
    print("   ")
    print("   # Initialize")
    print("   manager = TemplateManager()")
    print("   ")
    print("   # Add template")
    print("   manager.add_template(")
    print("       source_path='document.docx',")
    print("       name='domain.type.purpose.variant.version',")
    print("       description='Template description',")
    print("       tags=['tag1', 'tag2']")
    print("   )")
    print("   ")
    print("   # List with filtering")
    print("   templates = manager.list_templates(")
    print("       domain='business',")
    print("       type='report',")
    print("       purpose='quarterly'")
    print("   )")
    print("   ")
    print("   # Generate document")
    print("   manager.generate_from_template(")
    print("       template_name='business.report.quarterly.standard.v1',")
    print("       output_path='new_document.docx',")
    print("       variables={'key': 'value'}")
    print("   )")
    print("   ```")

    print("\n" + "=" * 60)
    print("Template naming convention:")
    print("=" * 60)
    print("Format: domain.type.purpose.variant.version")
    print("\nExamples:")
    print("  • business.report.quarterly.standard.v1")
    print("  • legal.contract.nda.comprehensive.v2")
    print("  • education.presentation.lecture.simple.v1")
    print("  • marketing.promo.social_media.instagram.v3")

    print("\nComponents:")
    print("  • Domain: Broad category (business, legal, education, etc.)")
    print("  • Type: Document type (report, contract, presentation, etc.)")
    print("  • Purpose: Specific use case (quarterly, nda, lecture, etc.)")
    print("  • Variant: Style or complexity (standard, simple, comprehensive)")
    print("  • Version: Version number (v1, v2, v3)")

    print("\n" + "=" * 60)
    print("Note: To actually add templates, you need real Office documents")
    print("(docx, xlsx, pptx files). The system will convert them to")
    print("markdown and store them in the hierarchical template directory.")
    print("=" * 60)


if __name__ == "__main__":
    demonstrate_template_management()
