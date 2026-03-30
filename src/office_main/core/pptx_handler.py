"""
PPTX (PowerPoint) presentation handler.

Provides high-level operations for PowerPoint presentation manipulation.
"""

from pathlib import Path
from typing import Any, Dict, List, Optional

from .cli_wrapper import LibreOfficeCLI


class PptxHandler:
    """Handler for PowerPoint presentation operations."""

    def __init__(self, presentation_path: str, project_path: Optional[str] = None):
        """
        Initialize handler for a PowerPoint presentation.

        Args:
            presentation_path: Path to the .pptx presentation
            project_path: Optional path to .lo-cli.json project file.
                         If None, a temporary session file will be created.
        """
        self.presentation_path = Path(presentation_path).absolute()
        self.project_path = project_path
        self._temp_session = False

        if self.project_path is None:
            # Create a temporary session file
            temp_cli = LibreOfficeCLI()
            self.project_path = temp_cli.start_session(str(self.presentation_path))
            # Do NOT end session here - keep file for CLI operations
            self._temp_session = True

        self.cli = LibreOfficeCLI(project_path=self.project_path)

    def add_slide(
        self, layout: str = "Title and Content", index: Optional[int] = None
    ) -> Dict[str, Any]:
        """
        Add a slide to the presentation.

        Args:
            layout: Slide layout (e.g., "Title and Content", "Title Only")
            index: Position to insert (append if None)

        Returns:
            Command result
        """
        kwargs = {"layout": layout}
        if index is not None:
            kwargs["index"] = str(index)
        return self.cli.impress("add-slide", positional=None, **kwargs)

    def remove_slide(self, index: int) -> Dict[str, Any]:
        """
        Remove a slide by index.

        Args:
            index: Slide index (0-based)

        Returns:
            Command result
        """
        return self.cli.impress("remove-slide", positional=[str(index)])

    def list_slides(self) -> Dict[str, Any]:
        """
        List all slides in the presentation.

        Returns:
            List of slides with indices and layouts
        """
        return self.cli.impress("list-slides")

    def set_content(
        self, slide_index: int, title: Optional[str] = None, content: Optional[str] = None
    ) -> Dict[str, Any]:
        """
        Update a slide's title and/or content.

        Args:
            slide_index: Slide index (0-based)
            title: New title text (optional)
            content: New content text (optional)

        Returns:
            Command result
        """
        kwargs = {}
        if title is not None:
            kwargs["title"] = title
        if content is not None:
            kwargs["content"] = content
        return self.cli.impress("set-content", positional=[str(slide_index)], **kwargs)

    def set_layout(self, slide_index: int, layout: str) -> Dict[str, Any]:
        """
        Set slide layout.

        Args:
            slide_index: Slide index (0-based)
            layout: New layout name

        Returns:
            Command result
        """
        return self.cli.impress("set-layout", positional=[str(slide_index), layout])

    def add_element(self, slide_index: int, element_type: str, **kwargs) -> Dict[str, Any]:
        """
        Add an element to a slide.

        Args:
            slide_index: Slide index (0-based)
            element_type: Type of element (text, image, shape, etc.)
            **kwargs: Additional element parameters

        Returns:
            Command result
        """
        # This is a generic wrapper - actual parameters depend on element_type
        all_kwargs = {"slide": str(slide_index), "type": element_type}
        all_kwargs.update(kwargs)
        return self.cli.impress("add-element", positional=None, **all_kwargs)

    def modify_element(self, slide_index: int, element_index: int, **kwargs) -> Dict[str, Any]:
        """
        Modify an existing element on a slide.

        Args:
            slide_index: Slide index (0-based)
            element_index: Element index on the slide
            **kwargs: Modification parameters

        Returns:
            Command result
        """
        all_kwargs = {"slide": str(slide_index), "element": str(element_index)}
        all_kwargs.update(kwargs)
        return self.cli.impress("modify-element", positional=None, **all_kwargs)

    def set_speaker_notes(self, slide_index: int, notes: str) -> Dict[str, Any]:
        """
        Set speaker notes for a slide.

        Args:
            slide_index: Slide index (0-based)
            notes: Speaker notes text

        Returns:
            Command result
        """
        return self.cli.impress("set-speaker-notes", positional=[str(slide_index), notes])

    def get_speaker_notes(self, slide_index: int) -> Dict[str, Any]:
        """
        Get speaker notes for a slide.

        Args:
            slide_index: Slide index (0-based)

        Returns:
            Speaker notes text
        """
        return self.cli.impress("get-speaker-notes", positional=[str(slide_index)])

    def export(self, output_path: str, format: str = "pdf") -> Dict[str, Any]:
        """
        Export the presentation to another format.

        Args:
            output_path: Path for the exported file
            format: Export format (pdf, pptx, odp, etc.)

        Returns:
            Command result
        """
        return self.cli.export("render", positional=[output_path], preset=format)

    def create_from_template(self, template_path: Optional[str] = None) -> "PptxHandler":
        """
        Create a new presentation from template.

        Note: This method is not yet implemented. For presentation creation,
        use the LibreOfficeCLI.document() method directly.

        Args:
            template_path: Path to template presentation (optional)

        Returns:
            New PptxHandler instance

        Raises:
            NotImplementedError: This method is not implemented
        """
        raise NotImplementedError(
            "create_from_template is not implemented. "
            "Use LibreOfficeCLI.document('new', ...) for presentation creation."
        )

    def analyze_structure(self) -> Dict[str, Any]:
        """
        Analyze presentation structure.

        Returns:
            Presentation analysis including slides, layouts, etc.
        """
        slides = self.list_slides()
        slide_count = len(slides.get("slides", []))

        # Analyze layouts
        layouts: Dict[str, int] = {}
        for slide in slides.get("slides", []):
            layout = slide.get("layout", "Unknown")
            layouts[layout] = layouts.get(layout, 0) + 1

        return {
            "presentation": str(self.presentation_path),
            "slide_count": slide_count,
            "layout_distribution": layouts,
            "has_speaker_notes": any(
                slide.get("has_notes", False) for slide in slides.get("slides", [])
            ),
        }

    def create_slide_deck(self, slides_data: List[Dict[str, Any]]) -> Dict[str, Any]:
        """
        Create a complete slide deck from structured data.

        Args:
            slides_data: List of slide definitions, each with:
                - title: Slide title
                - content: Slide content
                - layout: Slide layout (optional)
                - notes: Speaker notes (optional)

        Returns:
            Creation results
        """
        results = []
        for i, slide_def in enumerate(slides_data):
            # Add slide
            layout = slide_def.get("layout", "Title and Content")
            add_result = self.add_slide(layout=layout, index=i)
            slide_index = add_result.get("slide_index", i)

            # Set content
            if "title" in slide_def or "content" in slide_def:
                content_result = self.set_content(
                    slide_index=slide_index,
                    title=slide_def.get("title"),
                    content=slide_def.get("content"),
                )
                add_result["content"] = content_result

            # Set speaker notes
            if "notes" in slide_def:
                notes_result = self.set_speaker_notes(
                    slide_index=slide_index, notes=slide_def["notes"]
                )
                add_result["notes"] = notes_result

            results.append({"slide_index": slide_index, "layout": layout, "results": add_result})

        return {"created_slides": results}

    def close(self):
        """Close the handler and clean up temporary session files."""
        if self._temp_session and self.project_path:
            try:
                import os

                if os.path.exists(self.project_path):
                    os.unlink(self.project_path)
            except Exception:
                pass  # Ignore cleanup errors
        self.cli.end_session()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()
