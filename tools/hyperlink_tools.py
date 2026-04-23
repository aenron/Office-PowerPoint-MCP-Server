"""
Hyperlink management tools for PowerPoint MCP Server.
Implements hyperlink operations for text shapes and runs.
"""

from typing import Dict, List
from mcp.types import ToolAnnotations


def register_hyperlink_tools(app, presentations, get_current_presentation_id, validate_parameters,
                             is_positive, is_non_negative, is_in_range, is_valid_rgb):
    """Register hyperlink management tools with the FastMCP app."""

    def get_slide(presentation_id: str, slide_index: int):
        if presentation_id not in presentations:
            return None, {"error": "Presentation not found"}

        pres = presentations[presentation_id]
        if not (0 <= slide_index < len(pres.slides)):
            return None, {"error": f"Slide index {slide_index} out of range"}

        return pres.slides[slide_index], None

    def get_text_shape(presentation_id: str, slide_index: int, shape_index: int):
        slide, error = get_slide(presentation_id, slide_index)
        if error:
            return None, error

        if not (0 <= shape_index < len(slide.shapes)):
            return None, {"error": f"Shape index {shape_index} out of range"}

        shape = slide.shapes[shape_index]
        if not hasattr(shape, 'text_frame') or not shape.text_frame:
            return None, {"error": "Shape does not contain text"}

        return shape, None

    @app.tool(
        annotations=ToolAnnotations(
            title="Manage Hyperlinks",
        ),
    )
    def manage_hyperlinks(
        operation: str,
        presentation_id: str,
        slide_index: int,
        shape_index: int,
        run_index: int,
        text: str,
        url: str,
    ) -> Dict:
        """List, add, update, or remove hyperlinks on slides and text shapes."""
        try:
            if operation == "list_slide":
                slide, error = get_slide(presentation_id, slide_index)
                if error:
                    return error

                hyperlinks = []
                for shape_idx, shape in enumerate(slide.shapes):
                    if hasattr(shape, 'text_frame') and shape.text_frame:
                        for para_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                            for current_run_index, run in enumerate(paragraph.runs):
                                if run.hyperlink.address:
                                    hyperlinks.append({
                                        "shape_index": shape_idx,
                                        "paragraph_index": para_idx,
                                        "run_index": current_run_index,
                                        "text": run.text,
                                        "url": run.hyperlink.address
                                    })

                return {
                    "message": f"Found {len(hyperlinks)} hyperlinks on slide {slide_index}",
                    "hyperlinks": hyperlinks
                }

            shape, error = get_text_shape(
                presentation_id, slide_index, shape_index)
            if error:
                return error

            if operation == "add_shape":
                paragraph = shape.text_frame.paragraphs[0]
                run = paragraph.add_run()
                run.text = text
                run.hyperlink.address = url

                return {
                    "message": f"Added hyperlink '{text}' -> '{url}' to shape {shape_index}",
                    "text": text,
                    "url": url
                }

            paragraphs = shape.text_frame.paragraphs
            if run_index >= len(paragraphs[0].runs):
                return {"error": f"Run index {run_index} out of range"}

            run = paragraphs[0].runs[run_index]

            if operation == "add_run":
                run.hyperlink.address = url
                return {
                    "message": f"Added hyperlink '{url}' to run {run_index}",
                    "run_index": run_index,
                    "text": run.text,
                    "url": url
                }

            if operation == "update_shape":
                old_url = run.hyperlink.address
                run.hyperlink.address = url
                return {
                    "message": f"Updated hyperlink from '{old_url}' to '{url}'",
                    "old_url": old_url,
                    "new_url": url,
                    "text": run.text
                }

            if operation == "remove_shape":
                old_url = run.hyperlink.address
                run.hyperlink.address = None
                return {
                    "message": f"Removed hyperlink '{old_url}' from text '{run.text}'",
                    "removed_url": old_url,
                    "text": run.text
                }

            return {"error": "Invalid hyperlink operation. Use: list_slide, add_shape, add_run, update_shape, remove_shape"}

        except Exception as e:
            return {"error": f"Failed to manage hyperlinks: {str(e)}"}
