---
name: pptx
description: "PowerPoint presentation creation, editing, and analysis with support for slide decks, layouts, speaker notes, and export using cli-anything-libreoffice backend."
license: MIT
---

# PPTX creation, editing, and analysis

## Overview

A comprehensive skill for working with PowerPoint presentations (.pptx, .ppt) using the `cli-anything-libreoffice` backend. This skill supports professional presentation creation with consistent layouts, speaker notes, and export capabilities.

## Workflow Decision Tree

### Creating New Presentation
- **From scratch**: Build slide-by-slide with `PptxHandler` methods
- **From template**: Use `create_from_template()` with existing presentation
- **Structured deck**: Use `create_slide_deck()` for batch slide creation

### Editing Existing Presentation
- **Content updates**: Modify slides with `set_content()` and `modify_element()`
- **Structure changes**: Add/remove slides, change layouts
- **Visual enhancements**: Add elements, adjust formatting

### Presentation Analysis
- **Structure analysis**: Use `analyze_structure()` for slide count and layout distribution
- **Content review**: Extract speaker notes and slide content
- **Export validation**: Check exported formats for quality

## Core Operations

### Slide Management
```python
from office_skill import PptxHandler

ppt = PptxHandler("presentation.pptx")

# Add slides
ppt.add_slide(layout="Title and Content")
ppt.add_slide(layout="Title Only")
ppt.add_slide(layout="Blank")

# List slides
slides = ppt.list_slides()
print(f"Presentation has {len(slides['slides'])} slides")

# Remove slide
ppt.remove_slide(index=2)  # Remove third slide
```

### Slide Content
```python
# Set slide title and content
ppt.set_content(slide_index=0, title="Welcome", content="Introduction to the presentation")

# Modify existing content
ppt.set_content(slide_index=1, title="Updated Title")

# Change slide layout
ppt.set_layout(slide_index=0, layout="Section Header")
```

### Speaker Notes
```python
# Add speaker notes
ppt.set_speaker_notes(slide_index=0, notes="Welcome the audience. Introduce yourself.")

# Get speaker notes
notes = ppt.get_speaker_notes(slide_index=0)
print(f"Speaker notes: {notes}")
```

### Adding Elements
```python
# Add text element
ppt.add_element(slide_index=0, element_type="text",
                text="Additional point",
                x=100, y=200, width=400, height=50)

# Add shape element
ppt.add_element(slide_index=1, element_type="rectangle",
                x=50, y=50, width=200, height=100,
                fill_color="blue", line_color="black")

# Modify existing element
ppt.modify_element(slide_index=0, element_index=0,
                   text="Updated text",
                   font_size=18)
```

### Export and Conversion
```python
# Export to PDF
ppt.export("presentation.pdf", format="pdf")

# Export to other formats
ppt.export("presentation.odp", format="odp")
ppt.export("presentation.html", format="html")
ppt.export("slides.pptx", format="pptx")  # Copy
```

## Advanced Workflows

### Creating a Complete Slide Deck
```python
def create_business_presentation():
    """Create a standard business presentation."""
    ppt = PptxHandler("business_presentation.pptx")

    slide_definitions = [
        {
            "layout": "Title Slide",
            "title": "Quarterly Business Review",
            "content": "Q1 2026",
            "notes": "Start with welcome. Thank attendees."
        },
        {
            "layout": "Title and Content",
            "title": "Agenda",
            "content": "1. Performance Overview\n2. Financial Results\n3. Key Initiatives\n4. Q2 Outlook",
            "notes": "Briefly outline each section."
        },
        {
            "layout": "Title and Content",
            "title": "Performance Overview",
            "content": "• Revenue growth: +15% YoY\n• Customer acquisition: +25%\n• Market share: 12%",
            "notes": "Highlight key metrics. Use data from slide 5."
        },
        {
            "layout": "Title and Content",
            "title": "Financial Results",
            "content": "• Total revenue: $5.2M\n• Operating margin: 18%\n• Net income: $1.1M",
            "notes": "Emphasize profitability improvements."
        },
        {
            "layout": "Title and Content",
            "title": "Key Initiatives",
            "content": "1. Product expansion\n2. Market penetration\n3. Operational efficiency",
            "notes": "Detail each initiative in following slides."
        },
        {
            "layout": "Title and Content",
            "title": "Q2 Outlook",
            "content": "• Projected growth: +12%\n• New market entry\n• Product launch",
            "notes": "End with strong forward-looking statement."
        },
        {
            "layout": "Title Slide",
            "title": "Thank You",
            "content": "Questions?",
            "notes": "Open floor for questions. Provide contact info."
        }
    ]

    result = ppt.create_slide_deck(slide_definitions)
    print(f"Created {len(result['created_slides'])} slides")

    # Export for distribution
    ppt.export("business_review.pdf")

    return ppt
```

### Presentation Analysis and Reporting
```python
def analyze_presentation(presentation_path):
    """Analyze presentation structure and content."""
    ppt = PptxHandler(presentation_path)

    # Basic analysis
    analysis = ppt.analyze_structure()
    print(f"Presentation: {analysis['presentation']}")
    print(f"Slide count: {analysis['slide_count']}")
    print(f"Layout distribution: {analysis['layout_distribution']}")
    print(f"Has speaker notes: {analysis['has_speaker_notes']}")

    # Detailed slide analysis
    slides = ppt.list_slides()
    for slide in slides.get("slides", []):
        idx = slide.get("index", "N/A")
        layout = slide.get("layout", "Unknown")
        has_notes = slide.get("has_notes", False)
        print(f"Slide {idx}: {layout} (notes: {has_notes})")

    return analysis
```

### Template Management for PowerPoint Presentations

#### Template Naming Convention
PowerPoint templates use the format: `domain.type.purpose.variant.version`

**Examples:**
- `business.presentation.investor_pitch.modern.v1` - Investor pitch deck
- `education.presentation.training.comprehensive.v2` - Training presentation
- `marketing.presentation.product_launch.creative.v1` - Product launch deck

#### Adding PPTX Templates
```python
from office_skill import TemplateManager

manager = TemplateManager()

# Add a presentation template
template_data = manager.add_template(
    source_path="investor_pitch.pptx",
    name="business.presentation.investor_pitch.modern.v1",
    description="Modern investor pitch deck with startup focus",
    tags=["business", "presentation", "pitch", "investor", "startup"]
)

print(f"Template added: {template_data['name']}")
print(f"Slides: {template_data['analysis'].get('slides', 0)}")
print(f"Layouts: {template_data['analysis'].get('layouts', 0)}")
```

#### Generating Presentations from Templates
```python
# Generate a pitch deck from template
pitch_data = {
    "company_name": "TechStart Inc",
    "founding_year": "2023",
    "industry": "SaaS",
    "target_market": "Enterprise",
    "funding_round": "Series A",
    "funding_amount": "$5M",
    "use_of_funds": "Product development, Market expansion",
    "contact_email": "ceo@techstart.com"
}

result = manager.generate_from_template(
    template_name="business.presentation.investor_pitch.modern.v1",
    output_path="techstart_pitch_deck.pptx",
    variables=pitch_data
)

print(f"Generated: {result['output_path']}")
```

#### Template Analysis for Presentations
```python
# Analyze a presentation for template suitability
analysis = manager.analyze_document_structure("presentation.pptx")
print(f"Document type: {analysis['type']}")
print(f"Slides: {analysis['slides']}")
print(f"Masters: {analysis['masters']}")
print(f"Layouts: {analysis['layouts']}")

# Check if suitable for template library
if analysis['slides'] >= 5 and analysis['layouts'] >= 3:
    print("Presentation suitable for template library")
```

#### Example: Investor Pitch Generation
```python
def generate_investor_pitch(company_details, output_path):
    """Generate an investor pitch deck from template."""
    from office_skill import TemplateManager
    
    manager = TemplateManager()
    
    pitch_variables = {
        "company_name": company_details.get("name", "Company"),
        "tagline": company_details.get("tagline", "Revolutionizing the industry"),
        "founding_year": company_details.get("founding_year", "2023"),
        "team_size": company_details.get("team_size", "10"),
        "traction": company_details.get("traction", "$2M ARR"),
        "market_size": company_details.get("market_size", "$50B TAM"),
        "funding_ask": company_details.get("funding_ask", "$5M"),
        "use_of_funds": company_details.get("use_of_funds", "Product & Market Expansion"),
        "contact_info": company_details.get("contact", "contact@company.com")
    }
    
    result = manager.generate_from_template(
        template_name="business.presentation.investor_pitch.modern.v1",
        output_path=output_path,
        variables=pitch_variables
    )
    
    print(f"Investor pitch deck generated: {output_path}")
    print(f"Company: {company_details.get('name')}")
    
    return result
```

### Template-Based Presentation Generation
```python
def generate_from_template(template_name, data, output_path):
    """Generate presentation from template with dynamic content."""
    from office_skill import TemplateManager
    
    manager = TemplateManager()
    
    result = manager.generate_from_template(
        template_name=template_name,
        output_path=output_path,
        variables=data
    )
    
    print(f"Generated {output_path} from template {template_name}")
    print(f"Variables applied: {list(data.keys())}")
    
    return result
```

#### Training Presentation Generation
```python
def generate_training_presentation(course_name, modules, trainer, duration):
    """Generate a training presentation from template."""
    training_data = {
        "course_title": course_name,
        "trainer_name": trainer,
        "course_duration": duration,
        "module_count": str(len(modules)),
        "modules_list": "\n".join([f"• {module}" for module in modules]),
        "target_audience": "Employees, Managers",
        "learning_objectives": "Understand key concepts, Apply skills, Achieve certification",
        "contact_info": "training@company.com"
    }
    
    output_file = f"{course_name.lower().replace(' ', '_')}_training.pptx"
    
    return generate_from_template(
        template_name="education.presentation.training.comprehensive.v1",
        data=training_data,
        output_path=output_file
    )
```

#### Product Launch Deck Generation
```python
def generate_product_launch(product_details, launch_date, target_audience):
    """Generate a product launch presentation."""
    launch_data = {
        "product_name": product_details.get("name", "New Product"),
        "product_category": product_details.get("category", "Software"),
        "key_features": "\n".join([f"• {feature}" for feature in product_details.get("features", [])]),
        "unique_value": product_details.get("value_prop", "Industry-leading solution"),
        "target_customers": target_audience,
        "launch_date": launch_date,
        "pricing_tier": product_details.get("pricing", "Contact for pricing"),
        "next_steps": "Schedule demo, Request trial, Contact sales"
    }
    
    output_file = f"{product_details.get('name', 'product').lower().replace(' ', '_')}_launch.pptx"
    
    return generate_from_template(
        template_name="marketing.presentation.product_launch.creative.v1",
        data=launch_data,
        output_path=output_file
    )
```

## Integration with cli-anything-libreoffice

### Direct CLI Usage
```bash
# Basic impress commands
cli-anything-libreoffice impress add-slide --layout "Title and Content"
cli-anything-libreoffice impress list-slides
cli-anything-libreoffice impress set-content --slide 0 --title "Slide Title" --content "Slide content"

# Speaker notes
cli-anything-libreoffice impress set-speaker-notes --slide 0 --notes "Speaker notes text"
cli-anything-libreoffice impress get-speaker-notes --slide 0

# Elements
cli-anything-libreoffice impress add-element --slide 0 --type text --text "Element text" --x 100 --y 100
```

### Batch Creation
```bash
# Create batch file for presentation
cat > presentation_ops.txt << EOF
impress add-slide --layout "Title Slide"
impress set-content --slide 0 --title "Welcome"
impress add-slide --layout "Title and Content"
impress set-content --slide 1 --title "Agenda" --content "1. Item 1\n2. Item 2"
impress add-slide --layout "Title and Content"
impress set-content --slide 2 --title "Content" --content "Detailed content here"
EOF

cli-anything-libreoffice batch presentation_ops.txt
```

## Critical Rules for Presentation Processing

1. **Maintain consistent design**: Use the same layouts, fonts, and colors throughout the presentation.

2. **Keep slides focused**: One main idea per slide. Avoid overcrowding with text.

3. **Use speaker notes effectively**: Notes should complement slides, not duplicate them.

4. **Test exports**: Always check PDF and other exports for formatting issues.

5. **Follow presentation structure**: Standard structure: Title, Agenda, Content, Summary, Q&A.

## Examples

### Creating a Training Presentation
```python
def create_training_presentation():
    """Create a software training presentation."""
    ppt = PptxHandler("software_training.pptx")

    # Title slide
    ppt.add_slide(layout="Title Slide")
    ppt.set_content(0, title="Software Training", content="Getting Started with New System")
    ppt.set_speaker_notes(0, "Welcome participants. Explain training objectives.")

    # Introduction
    ppt.add_slide(layout="Title and Content")
    ppt.set_content(1, title="Training Objectives", content="• Understand system architecture\n• Learn key workflows\n• Practice common tasks\n• Know where to get help")

    # Modules
    modules = [
        ("System Overview", "Architecture, components, and integration points"),
        ("Basic Operations", "Daily tasks and common workflows"),
        ("Advanced Features", "Power user tips and tricks"),
        ("Troubleshooting", "Common issues and solutions")
    ]

    for i, (title, content) in enumerate(modules, start=2):
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(i, title=title, content=content)
        ppt.set_speaker_notes(i, f"Detailed explanation of {title.lower()}.")

    # Summary
    ppt.add_slide(layout="Title and Content")
    ppt.set_content(len(modules)+2, title="Key Takeaways", content="• System is intuitive and powerful\n• Support is always available\n• Continuous learning resources")
    ppt.set_speaker_notes(len(modules)+2, "Summarize main points. Open for questions.")

    # Q&A
    ppt.add_slide(layout="Title Slide")
    ppt.set_content(len(modules)+3, title="Questions?", content="Thank you!")
    ppt.set_speaker_notes(len(modules)+3, "Answer questions. Distribute contact info.")

    return ppt
```

### Investor Pitch Deck
```python
def create_investor_pitch():
    """Create an investor pitch presentation."""
    ppt = PptxHandler("pitch_deck.pptx")

    slides = [
        {"layout": "Title Slide", "title": "Company Name", "content": "Revolutionizing Industry", "notes": "Start strong. Capture attention."},
        {"layout": "Title and Content", "title": "The Problem", "content": "• Current solutions are inefficient\n• High costs for customers\n• Limited scalability", "notes": "Establish need for your solution."},
        {"layout": "Title and Content", "title": "Our Solution", "content": "• Innovative technology\n• 10x efficiency improvement\n• Cost-effective scaling", "notes": "Present your unique value proposition."},
        {"layout": "Title and Content", "title": "Market Opportunity", "content": "• $50B total addressable market\n• Growing at 20% annually\n• Limited competition", "notes": "Show market size and growth potential."},
        {"layout": "Title and Content", "title": "Business Model", "content": "• Subscription-based revenue\n• $100/month per user\n• 80% gross margins", "notes": "Explain how you make money."},
        {"layout": "Title and Content", "title": "Traction", "content": "• 100+ enterprise customers\n• $2M ARR\n• 150% YoY growth", "notes": "Demonstrate progress and validation."},
        {"layout": "Title and Content", "title": "Team", "content": "• Experienced founders\n• Industry experts\n• Technical talent", "notes": "Highlight team credentials."},
        {"layout": "Title and Content", "title": "Financial Projections", "content": "• $10M revenue in Year 3\n• Profitability in Year 2\n• 40% EBITDA margins", "notes": "Show financial potential."},
        {"layout": "Title and Content", "title": "Funding Ask", "content": "• Seeking $5M Series A\n• Product development\n• Market expansion", "notes": "State what you need and how you'll use it."},
        {"layout": "Title Slide", "title": "Thank You", "content": "Contact: email@company.com", "notes": "End professionally. Provide contact."}
    ]

    ppt.create_slide_deck(slides)
    ppt.export("pitch_deck.pdf")

    return ppt
```

## Troubleshooting

### Common Issues
- **Layout mismatch**: Ensure layout names match available templates in LibreOffice
- **Element positioning**: Coordinates are in points (1/72 inch); test positioning
- **Export formatting**: Check PDF exports for font and layout consistency
- **Speaker notes visibility**: Notes may not appear in some export formats

### Performance Tips
- Use slide masters for consistent formatting
- Reuse layouts and elements where possible
- For large presentations, create in batches
- Test exports early in the process

## Related Resources
- [LibreOffice Impress Guide](https://documentation.libreoffice.org/en/english-documentation/impress/)
- [Presentation Design Principles](https://www.garrreynolds.com/preso-tips/design/)
- [Slide Design Best Practices](https://speakerhub.com/skillcamp/slide-design-best-practices)