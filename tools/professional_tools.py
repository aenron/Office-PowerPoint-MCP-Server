"""
Professional design tools for PowerPoint MCP Server.
Handles themes, effects, fonts, and advanced formatting.
"""
from typing import Dict, List
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
import utils as ppt_utils


def register_professional_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register professional design tools with the FastMCP app"""

    def get_required_presentation(presentation_id: str):
        if presentation_id not in presentations:
            return None, {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        return presentations[presentation_id], None

    def get_required_slide(presentation_id: str, slide_index: int):
        pres, error = get_required_presentation(presentation_id)
        if error:
            return None, None, error

        if slide_index < 0 or slide_index >= len(pres.slides):
            return pres, None, {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }

        return pres, pres.slides[slide_index], None

    @app.tool(
        annotations=ToolAnnotations(
            title="Apply Professional Design",
        ),
    )
    def apply_professional_design(
        operation: str,
        presentation_id: str,
        slide_type: str = "title_content",
        color_scheme: str = "modern_blue",
        title: str = "",
        apply_to_existing: bool = False,
        slide_index: int = 0,
        enhance_title: bool = True,
        enhance_content: bool = True,
        enhance_shapes: bool = True,
        enhance_charts: bool = True,
        content: List[str] = [],
    ) -> Dict:
        """Manage professional design operations."""
        if operation == "get_color_schemes":
            return ppt_utils.get_color_schemes()

        if operation == "create_slide":
            pres, error = get_required_presentation(presentation_id)
            if error:
                return error

            try:
                ppt_utils.add_professional_slide(
                    pres,
                    slide_type=slide_type,
                    color_scheme=color_scheme,
                    title=title,
                    content=content
                )
                return {
                    "message": f"Added professional {slide_type} slide",
                    "slide_index": len(pres.slides) - 1,
                    "color_scheme": color_scheme,
                    "slide_type": slide_type
                }
            except Exception as e:
                return {
                    "error": f"Failed to create professional slide: {str(e)}"
                }

        if operation == "apply_theme":
            pres, error = get_required_presentation(presentation_id)
            if error:
                return error

            try:
                ppt_utils.apply_professional_theme(
                    pres,
                    color_scheme=color_scheme,
                    apply_to_existing=apply_to_existing
                )
                return {
                    "message": f"Applied {color_scheme} theme to presentation",
                    "color_scheme": color_scheme,
                    "applied_to_existing": apply_to_existing
                }
            except Exception as e:
                return {
                    "error": f"Failed to apply professional theme: {str(e)}"
                }

        if operation == "enhance_slide":
            _, slide, error = get_required_slide(presentation_id, slide_index)
            if error:
                return error

            try:
                result = ppt_utils.enhance_existing_slide(
                    slide,
                    color_scheme=color_scheme,
                    enhance_title=enhance_title,
                    enhance_content=enhance_content,
                    enhance_shapes=enhance_shapes,
                    enhance_charts=enhance_charts
                )
                return {
                    "message": f"Enhanced slide {slide_index} with {color_scheme} scheme",
                    "slide_index": slide_index,
                    "color_scheme": color_scheme,
                    "enhancements_applied": result.get("enhancements_applied", [])
                }
            except Exception as e:
                return {
                    "error": f"Failed to enhance existing slide: {str(e)}"
                }

        return {"error": "Invalid design operation. Use: get_color_schemes, create_slide, apply_theme, enhance_slide"}

    @app.tool(
        annotations=ToolAnnotations(
            title="Apply Picture Effects",
        ),
    )
    def apply_picture_effects(
        presentation_id: str,
        slide_index: int,
        shape_index: int,
        effects: Dict[str, Dict],
    ) -> Dict:
        """Apply multiple picture effects in combination."""
        _, slide, error = get_required_slide(presentation_id, slide_index)
        if error:
            return error

        if shape_index < 0 or shape_index >= len(slide.shapes):
            return {
                "error": f"Invalid shape index: {shape_index}. Available shapes: 0-{len(slide.shapes) - 1}"
            }

        shape = slide.shapes[shape_index]

        try:
            applied_effects = []
            warnings = []

            for effect_type, effect_params in effects.items():
                try:
                    if effect_type == "shadow":
                        ppt_utils.apply_picture_shadow(
                            shape,
                            shadow_type=effect_params.get(
                                "shadow_type", "outer"),
                            blur_radius=effect_params.get("blur_radius", 4.0),
                            distance=effect_params.get("distance", 3.0),
                            direction=effect_params.get("direction", 315.0),
                            color=effect_params.get("color", [0, 0, 0]),
                            transparency=effect_params.get("transparency", 0.6)
                        )
                        applied_effects.append("shadow")
                    elif effect_type == "reflection":
                        ppt_utils.apply_picture_reflection(
                            shape,
                            size=effect_params.get("size", 0.5),
                            transparency=effect_params.get(
                                "transparency", 0.5),
                            distance=effect_params.get("distance", 0.0),
                            blur=effect_params.get("blur", 4.0)
                        )
                        applied_effects.append("reflection")
                    elif effect_type == "glow":
                        ppt_utils.apply_picture_glow(
                            shape,
                            size=effect_params.get("size", 5.0),
                            color=effect_params.get("color", [0, 176, 240]),
                            transparency=effect_params.get("transparency", 0.4)
                        )
                        applied_effects.append("glow")
                    elif effect_type == "soft_edges":
                        ppt_utils.apply_picture_soft_edges(
                            shape,
                            radius=effect_params.get("radius", 2.5)
                        )
                        applied_effects.append("soft_edges")
                    elif effect_type == "rotation":
                        ppt_utils.apply_picture_rotation(
                            shape,
                            rotation=effect_params.get("rotation", 0.0)
                        )
                        applied_effects.append("rotation")
                    elif effect_type == "transparency":
                        ppt_utils.apply_picture_transparency(
                            shape,
                            transparency=effect_params.get("transparency", 0.0)
                        )
                        applied_effects.append("transparency")
                    elif effect_type == "bevel":
                        ppt_utils.apply_picture_bevel(
                            shape,
                            bevel_type=effect_params.get(
                                "bevel_type", "circle"),
                            width=effect_params.get("width", 6.0),
                            height=effect_params.get("height", 6.0)
                        )
                        applied_effects.append("bevel")
                    elif effect_type == "filter":
                        ppt_utils.apply_picture_filter(
                            shape,
                            filter_type=effect_params.get(
                                "filter_type", "none"),
                            intensity=effect_params.get("intensity", 0.5)
                        )
                        applied_effects.append("filter")
                    else:
                        warnings.append(f"Unknown effect type: {effect_type}")
                except Exception as e:
                    warnings.append(
                        f"Failed to apply {effect_type} effect: {str(e)}")

            result = {
                "message": f"Applied {len(applied_effects)} effects to shape {shape_index} on slide {slide_index}",
                "applied_effects": applied_effects
            }

            if warnings:
                result["warnings"] = warnings

            return result
        except Exception as e:
            return {
                "error": f"Failed to apply picture effects: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Manage Fonts",
        ),
    )
    def manage_fonts(
        operation: str,
        font_path: str,
        output_path: str,
        text_content: str,
        presentation_type: str,
    ) -> Dict:
        """Manage font analysis, optimization, and recommendation operations."""
        if operation == "analyze":
            try:
                return ppt_utils.analyze_font_file(font_path)
            except Exception as e:
                return {
                    "error": f"Failed to analyze font: {str(e)}"
                }

        if operation == "optimize":
            try:
                optimized_path = ppt_utils.optimize_font_for_presentation(
                    font_path,
                    output_path=output_path,
                    text_content=text_content
                )
                return {
                    "message": f"Optimized font: {font_path}",
                    "original_path": font_path,
                    "optimized_path": optimized_path
                }
            except Exception as e:
                return {
                    "error": f"Failed to optimize font: {str(e)}"
                }

        if operation == "recommend":
            try:
                return ppt_utils.get_font_recommendations(
                    font_path,
                    presentation_type=presentation_type
                )
            except Exception as e:
                return {
                    "error": f"Failed to get font recommendations: {str(e)}"
                }

        return {"error": "Invalid font operation. Use: analyze, optimize, recommend"}
