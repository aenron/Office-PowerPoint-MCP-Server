"""
Slide master management tools for PowerPoint MCP Server.
Implements slide master and layout access capabilities.
"""

from typing import Dict
from mcp.types import ToolAnnotations


def register_master_tools(app, presentations, get_current_presentation_id, validate_parameters,
                          is_positive, is_non_negative, is_in_range, is_valid_rgb):
    """Register slide master management tools with the FastMCP app."""

    def get_presentation(presentation_id: str):
        if presentation_id not in presentations:
            return None, {"error": "Presentation not found"}
        return presentations[presentation_id], None

    def get_master(presentation_id: str, master_index: int):
        pres, error = get_presentation(presentation_id)
        if error:
            return None, None, error

        if not (0 <= master_index < len(pres.slide_masters)):
            return pres, None, {"error": f"Master index {master_index} out of range"}

        return pres, pres.slide_masters[master_index], None

    @app.tool(
        annotations=ToolAnnotations(
            title="Manage Slide Masters",
        ),
    )
    def manage_slide_masters(
        operation: str,
        presentation_id: str,
        master_index: int = 0,
    ) -> Dict:
        """List slide masters, list layouts, or get master details."""
        try:
            if operation == "list_masters":
                pres, error = get_presentation(presentation_id)
                if error:
                    return error

                masters_info = []
                for idx, master in enumerate(pres.slide_masters):
                    masters_info.append({
                        "index": idx,
                        "layout_count": len(master.slide_layouts),
                        "name": getattr(master, 'name', f"Master {idx}")
                    })

                return {
                    "message": f"Found {len(masters_info)} slide masters",
                    "masters": masters_info,
                    "total_masters": len(pres.slide_masters)
                }

            if operation == "list_layouts":
                _, master, error = get_master(presentation_id, master_index)
                if error:
                    return error

                layouts_info = []
                for idx, layout in enumerate(master.slide_layouts):
                    layouts_info.append({
                        "index": idx,
                        "name": layout.name,
                        "placeholder_count": len(layout.placeholders) if hasattr(layout, 'placeholders') else 0
                    })

                return {
                    "message": f"Master {master_index} has {len(layouts_info)} layouts",
                    "master_index": master_index,
                    "layouts": layouts_info
                }

            if operation == "get_details":
                _, master, error = get_master(presentation_id, master_index)
                if error:
                    return error

                layouts_info = []
                for idx, layout in enumerate(master.slide_layouts):
                    placeholders_info = []
                    if hasattr(layout, 'placeholders'):
                        for placeholder in layout.placeholders:
                            placeholders_info.append({
                                "idx": placeholder.placeholder_format.idx,
                                "type": str(placeholder.placeholder_format.type),
                                "name": getattr(placeholder, 'name', 'Unnamed')
                            })

                    layouts_info.append({
                        "index": idx,
                        "name": layout.name,
                        "placeholder_count": len(layout.placeholders) if hasattr(layout, 'placeholders') else 0,
                        "placeholders": placeholders_info
                    })

                return {
                    "message": f"Master {master_index} information",
                    "master_index": master_index,
                    "layout_count": len(master.slide_layouts),
                    "name": getattr(master, 'name', f"Master {master_index}"),
                    "layouts": layouts_info
                }

            return {"error": "Invalid master operation. Use: list_masters, list_layouts, get_details"}

        except Exception as e:
            return {"error": f"Failed to manage slide masters: {str(e)}"}
