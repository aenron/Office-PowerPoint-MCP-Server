"""
Slide transition management tools for PowerPoint MCP Server.
Implements slide transition and timing capabilities.
"""

from typing import Dict
from mcp.types import ToolAnnotations


def register_transition_tools(app, presentations, get_current_presentation_id, validate_parameters,
                              is_positive, is_non_negative, is_in_range, is_valid_rgb):
    """Register slide transition management tools with the FastMCP app."""

    def get_slide_and_presentation(presentation_id: str, slide_index: int):
        if presentation_id not in presentations:
            return None, None, {"error": "Presentation not found"}

        pres = presentations[presentation_id]

        if not (0 <= slide_index < len(pres.slides)):
            return None, None, {"error": f"Slide index {slide_index} out of range"}

        return pres, pres.slides[slide_index], None

    @app.tool(
        annotations=ToolAnnotations(
            title="Manage Slide Transition",
        ),
    )
    def manage_slide_transition(
        operation: str,
        presentation_id: str,
        slide_index: int,
        transition_type: str = "",
        duration: float = 0.0,
    ) -> Dict:
        """Get, set, or remove slide transition information."""
        try:
            _, _, error = get_slide_and_presentation(
                presentation_id, slide_index)
            if error:
                return error

            if operation == "get":
                return {
                    "message": f"Transition info for slide {slide_index}",
                    "slide_index": slide_index,
                    "note": "Transition reading has limited support in python-pptx"
                }

            if operation == "set":
                return {
                    "message": f"Transition setting requested for slide {slide_index}",
                    "slide_index": slide_index,
                    "transition_type": transition_type,
                    "duration": duration,
                    "note": "Transition setting has limited support in python-pptx - this is a placeholder for future enhancement"
                }

            if operation == "remove":
                return {
                    "message": f"Transition removal requested for slide {slide_index}",
                    "slide_index": slide_index,
                    "note": "Transition removal has limited support in python-pptx - this is a placeholder for future enhancement"
                }

            return {"error": "Invalid transition operation. Use: get, set, remove"}

        except Exception as e:
            return {"error": f"Failed to manage slide transition: {str(e)}"}
