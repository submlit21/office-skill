"""
Presentation deck creation example.

This example demonstrates creating a professional presentation deck
with consistent formatting, speaker notes, and proper structure.
"""

import tempfile
from pathlib import Path
from office_skill import PptxHandler


def create_business_presentation():
    """Create a comprehensive business presentation."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        ppt_path = tmp.name

    try:
        print(f"Creating presentation at: {ppt_path}")
        ppt = PptxHandler(ppt_path)

        # Title Slide
        print("• Adding title slide...")
        ppt.add_slide(layout="Title Slide")
        ppt.set_content(0,
                       title="Quarterly Business Review",
                       content="Q1 2026")
        ppt.set_speaker_notes(0,
                             "Welcome everyone. Thank you for attending our quarterly review.")

        # Agenda
        print("• Adding agenda slide...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(1,
                       title="Agenda",
                       content="1. Performance Overview\n"
                               "2. Financial Results\n"
                               "3. Key Initiatives\n"
                               "4. Market Analysis\n"
                               "5. Q2 Outlook\n"
                               "6. Q&A")
        ppt.set_speaker_notes(1,
                             "Briefly outline each section. Keep it concise.")

        # Performance Overview
        print("• Adding performance overview...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(2,
                       title="Performance Overview",
                       content="• Revenue Growth: +15% YoY\n"
                               "• Customer Acquisition: +25%\n"
                               "• Market Share: 12% (up from 10%)\n"
                               "• Customer Satisfaction: 4.8/5.0\n"
                               "• Employee Retention: 92%")
        ppt.set_speaker_notes(2,
                             "Highlight key metrics. Emphasize growth and market share gains.")

        # Financial Results
        print("• Adding financial results...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(3,
                       title="Financial Results",
                       content="• Total Revenue: $5.2M\n"
                               "• Operating Margin: 18%\n"
                               "• Net Income: $1.1M\n"
                               "• Cash Flow: $800K\n"
                               "• ROI: 22%")
        ppt.set_speaker_notes(3,
                             "Focus on profitability and cash flow. These are strong results.")

        # Financial Chart Slide
        print("• Adding financial chart slide...")
        ppt.add_slide(layout="Title Only")
        ppt.set_content(4, title="Revenue Growth Trend")
        # In a real scenario, you would add a chart element here
        ppt.set_speaker_notes(4,
                             "Show quarterly revenue growth. Highlight Q1 performance.")

        # Key Initiatives
        print("• Adding key initiatives...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(5,
                       title="Key Initiatives",
                       content="1. Product Expansion\n"
                               "   • New feature releases\n"
                               "   • Platform integration\n\n"
                               "2. Market Penetration\n"
                               "   • New geographic markets\n"
                               "   • Partnership development\n\n"
                               "3. Operational Efficiency\n"
                               "   • Process automation\n"
                               "   • Cost optimization")
        ppt.set_speaker_notes(5,
                             "Detail each initiative. Explain strategic importance.")

        # Market Analysis
        print("• Adding market analysis...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(6,
                       title="Market Analysis",
                       content="• Total Addressable Market: $50B\n"
                               "• Current Penetration: 0.1%\n"
                               "• Competitor Analysis:\n"
                               "   - Competitor A: 40% share\n"
                               "   - Competitor B: 25% share\n"
                               "   - Others: 35% share\n"
                               "• Growth Rate: 20% annually")
        ppt.set_speaker_notes(6,
                             "Emphasize growth opportunity. Show competitive landscape.")

        # SWOT Analysis
        print("• Adding SWOT analysis...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(7,
                       title="SWOT Analysis",
                       content="Strengths:\n"
                               "• Strong brand recognition\n"
                               "• Innovative technology\n\n"
                               "Weaknesses:\n"
                               "• Limited international presence\n"
                               "• High customer acquisition cost\n\n"
                               "Opportunities:\n"
                               "• Emerging markets\n"
                               "• Product diversification\n\n"
                               "Threats:\n"
                               "• Regulatory changes\n"
                               "• New market entrants")
        ppt.set_speaker_notes(7,
                             "Be honest about weaknesses. Frame threats as challenges to address.")

        # Q2 Outlook
        print("• Adding Q2 outlook...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(8,
                       title="Q2 Outlook",
                       content="• Revenue Target: $5.8M (+12%)\n"
                               "• New Product Launch: Q2 Week 8\n"
                               "• Market Expansion: 3 new countries\n"
                               "• Team Growth: +10 FTEs\n"
                               "• R&D Investment: $500K")
        ppt.set_speaker_notes(8,
                             "Set ambitious but achievable goals. Show clear roadmap.")

        # Roadmap
        print("• Adding roadmap...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(9,
                       title="2026 Roadmap",
                       content="Q1: Foundation & Planning ✓\n"
                               "Q2: Market Expansion & Product Launch\n"
                               "Q3: Scaling & Optimization\n"
                               "Q4: International Growth & Year-End Review")
        ppt.set_speaker_notes(9,
                             "Show progress to date. Build confidence in execution.")

        # Team Slide
        print("• Adding team slide...")
        ppt.add_slide(layout="Title and Content")
        ppt.set_content(10,
                       title="Our Team",
                       content="• Leadership: 10+ years average experience\n"
                               "• Engineering: 50+ developers\n"
                               "• Sales & Marketing: 25+ professionals\n"
                               "• Operations: 15+ specialists\n"
                               "• Advisors: Industry experts")
        ppt.set_speaker_notes(10,
                             "Highlight team strength. Build credibility.")

        # Thank You Slide
        print("• Adding thank you slide...")
        ppt.add_slide(layout="Title Slide")
        ppt.set_content(11,
                       title="Thank You",
                       content="Questions & Discussion")
        ppt.set_speaker_notes(11,
                             "Open floor for questions. Provide contact information.")

        # Analyze presentation
        print("\nAnalyzing presentation structure...")
        analysis = ppt.analyze_structure()
        print(f"• Total slides: {analysis['slide_count']}")
        print(f"• Layout distribution: {analysis['layout_distribution']}")
        print(f"• Has speaker notes: {analysis['has_speaker_notes']}")

        # Export to PDF
        pdf_path = ppt_path.replace(".pptx", ".pdf")
        print(f"\nExporting to PDF: {pdf_path}")
        ppt.export(pdf_path, format="pdf")

        print(f"\nPresentation created successfully!")
        print(f"• PPTX: {ppt_path}")
        print(f"• PDF: {pdf_path}")
        print(f"• File size: {Path(ppt_path).stat().st_size:,} bytes")

        return ppt_path, pdf_path

    except Exception as e:
        print(f"Error creating presentation: {e}")
        raise
    finally:
        # Cleanup for example purposes
        import os
        if os.path.exists(ppt_path):
            os.unlink(ppt_path)
        pdf_path = ppt_path.replace(".pptx", ".pdf")
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)
        print("\nTemporary files cleaned up.")


def create_training_presentation():
    """Create a training presentation with interactive elements."""
    with tempfile.NamedTemporaryFile(suffix=".pptx", delete=False) as tmp:
        training_path = tmp.name

    try:
        print(f"\nCreating training presentation at: {training_path}")
        ppt = PptxHandler(training_path)

        slides = [
            {
                "layout": "Title Slide",
                "title": "Software Training",
                "content": "Advanced Features Workshop",
                "notes": "Welcome participants. Set expectations."
            },
            {
                "layout": "Title and Content",
                "title": "Learning Objectives",
                "content": "• Master advanced functionality\n"
                          "• Improve productivity\n"
                          "• Troubleshoot common issues\n"
                          "• Apply best practices",
                "notes": "Outline what participants will learn."
            },
            {
                "layout": "Title and Content",
                "title": "Agenda",
                "content": "1. Advanced Navigation\n"
                          "2. Power Features\n"
                          "3. Customization\n"
                          "4. Integration\n"
                          "5. Troubleshooting\n"
                          "6. Q&A",
                "notes": "Provide structure for the session."
            },
            {
                "layout": "Title and Content",
                "title": "Advanced Navigation",
                "content": "• Keyboard shortcuts\n"
                          "• Quick access toolbar\n"
                          "• Navigation pane\n"
                          "• Search functionality",
                "notes": "Demo each navigation method."
            },
            {
                "layout": "Title and Content",
                "title": "Power Features",
                "content": "• Macros and automation\n"
                          "• Advanced formatting\n"
                          "• Template management\n"
                          "• Collaboration tools",
                "notes": "Show real-world examples."
            },
            {
                "layout": "Title and Content",
                "title": "Customization",
                "content": "• Interface customization\n"
                          "• Shortcut creation\n"
                          "• Theme development\n"
                          "• Add-in integration",
                "notes": "Allow time for hands-on practice."
            },
            {
                "layout": "Title and Content",
                "title": "Integration",
                "content": "• API connectivity\n"
                          "• Data import/export\n"
                          "• Third-party tools\n"
                          "• Cloud integration",
                "notes": "Demonstrate integration workflows."
            },
            {
                "layout": "Title and Content",
                "title": "Troubleshooting",
                "content": "• Common error messages\n"
                          "• Performance issues\n"
                          "• Compatibility problems\n"
                          "• Recovery procedures",
                "notes": "Share troubleshooting checklist."
            },
            {
                "layout": "Title and Content",
                "title": "Best Practices",
                "content": "• Regular backups\n"
                          "• Version control\n"
                          "• Documentation\n"
                          "• Security protocols",
                "notes": "Emphasize importance of good practices."
            },
            {
                "layout": "Title and Content",
                "title": "Resources",
                "content": "• Online documentation\n"
                          "• Community forums\n"
                          "• Support channels\n"
                          "• Training materials",
                "notes": "Provide all necessary resources."
            },
            {
                "layout": "Title Slide",
                "title": "Thank You",
                "content": "Questions?",
                "notes": "Open Q&A session. Collect feedback."
            }
        ]

        print(f"Creating {len(slides)} slides...")
        result = ppt.create_slide_deck(slides)
        print(f"Created {len(result['created_slides'])} slides successfully")

        # Export
        pdf_path = training_path.replace(".pptx", "_training.pdf")
        ppt.export(pdf_path, format="pdf")
        print(f"Training PDF exported: {pdf_path}")

        return training_path, pdf_path

    finally:
        import os
        if os.path.exists(training_path):
            os.unlink(training_path)
        pdf_path = training_path.replace(".pptx", "_training.pdf")
        if os.path.exists(pdf_path):
            os.unlink(pdf_path)


def main():
    """Run presentation examples."""
    print("Presentation Deck Examples")
    print("=" * 50)

    print("\n1. Business Presentation Example")
    print("-" * 30)
    create_business_presentation()

    print("\n2. Training Presentation Example")
    print("-" * 30)
    create_training_presentation()

    print("\n" + "=" * 50)
    print("All presentation examples completed successfully!")
    print("\nNote: Files were created in temporary locations.")
    print("In real usage, specify your own file paths.")


if __name__ == "__main__":
    main()