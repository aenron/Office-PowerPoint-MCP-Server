"""
Enhanced template-based slide creation tools for PowerPoint MCP Server.
Handles template application, template management, automated slide generation,
and advanced features like dynamic sizing, auto-wrapping, and visual effects.
"""
from typing import Dict, List, Any
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
import utils.template_utils as template_utils


def register_template_tools(app: FastMCP, presentations: Dict, get_current_presentation_id):
    """Register template-based tools with the FastMCP app"""

    @app.tool(
        annotations=ToolAnnotations(
            title="Manage Slide Templates",
        ),
    )
    def manage_slide_templates(
        operation: str,
        presentation_id: str,
        slide_index: int,
        template_id: str,
        color_scheme: str = "modern_blue",
        content_mapping: Dict[str, str] = {},
        image_paths: Dict[str, str] = {},
        template_sequence: List[Dict[str, Any]] = [],
        presentation_title: str = "",
        layout_index: int = -1,
    ) -> Dict:
        """List, inspect, apply, and create slide templates."""
        try:
            if operation == "list":
                available_templates = template_utils.get_available_templates()
                usage_examples = template_utils.get_template_usage_examples()

                return {
                    "available_templates": available_templates,
                    "total_templates": len(available_templates),
                    "usage_examples": usage_examples,
                    "message": "Use manage_slide_templates with operation='apply_to_slide' to apply templates to slides"
                }

            if operation == "get_info":
                templates_data = template_utils.load_slide_templates()

                if template_id not in templates_data.get('templates', {}):
                    available_templates = list(
                        templates_data.get('templates', {}).keys())
                    return {
                        "error": f"Template '{template_id}' not found",
                        "available_templates": available_templates
                    }

                template = templates_data['templates'][template_id]

                # Extract element information
                elements_info = []
                for element in template.get('elements', []):
                    element_info = {
                        "type": element.get('type'),
                        "role": element.get('role'),
                        "position": element.get('position'),
                        "placeholder_text": element.get('placeholder_text', ''),
                        "styling_options": list(element.get('styling', {}).keys())
                    }
                    elements_info.append(element_info)

                return {
                    "template_id": template_id,
                    "name": template.get('name'),
                    "description": template.get('description'),
                    "layout_type": template.get('layout_type'),
                    "elements": elements_info,
                    "element_count": len(elements_info),
                    "has_background": 'background' in template,
                    "background_type": template.get('background', {}).get('type'),
                    "color_schemes": list(templates_data.get('color_schemes', {}).keys()),
                    "usage_tip": f"Use manage_slide_templates with operation='create_slide' and template_id='{template_id}' to create a slide with this layout"
                }

            if presentation_id not in presentations:
                return {
                    "error": "No presentation is currently loaded or the specified ID is invalid"
                }

            pres = presentations[presentation_id]

            if operation == "apply_to_slide":
                if slide_index < 0 or slide_index >= len(pres.slides):
                    return {
                        "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
                    }

                slide = pres.slides[slide_index]
                result = template_utils.apply_slide_template(
                    slide, template_id, color_scheme,
                    content_mapping, image_paths
                )

                if result['success']:
                    return {
                        "message": f"Applied template '{template_id}' to slide {slide_index}",
                        "slide_index": slide_index,
                        "template_applied": result
                    }

                return {
                    "error": f"Failed to apply template: {result.get('error', 'Unknown error')}"
                }

            if operation == "create_slide":
                # Validate layout index when explicitly provided
                if layout_index != -1 and (layout_index < 0 or layout_index >= len(pres.slide_layouts)):
                    return {
                        "error": f"Invalid layout index: {layout_index}. Available layouts: 0-{len(pres.slide_layouts) - 1}"
                    }

                # Add new slide
                layout = pres.slide_layouts[layout_index] if layout_index != -1 else template_utils.get_template_base_layout(
                    pres)
                slide = pres.slides.add_slide(layout)
                created_slide_index = len(pres.slides) - 1

                # Apply template
                result = template_utils.apply_slide_template(
                    slide, template_id, color_scheme,
                    content_mapping, image_paths
                )

                if result['success']:
                    return {
                        "message": f"Created slide {created_slide_index} using template '{template_id}'",
                        "slide_index": created_slide_index,
                        "template_applied": result
                    }

                return {
                    "error": f"Failed to apply template to new slide: {result.get('error', 'Unknown error')}"
                }

            if operation == "create_presentation":
                if not template_sequence:
                    return {
                        "error": "Template sequence cannot be empty"
                    }

                # Set presentation title if provided
                if presentation_title:
                    pres.core_properties.title = presentation_title

                # Create slides from template sequence
                result = template_utils.create_presentation_from_template_sequence(
                    pres, template_sequence, color_scheme
                )

                if result['success']:
                    return {
                        "message": f"Created presentation with {result['total_slides']} slides",
                        "presentation_id": presentation_id,
                        "creation_result": result,
                        "total_slides": len(pres.slides)
                    }

                return {
                    "warning": "Presentation created with some errors",
                    "presentation_id": presentation_id,
                    "creation_result": result,
                    "total_slides": len(pres.slides)
                }

            return {
                "error": "Invalid template operation. Use: list, get_info, apply_to_slide, create_slide, create_presentation"
            }

        except Exception as e:
            return {
                "error": f"Failed to manage slide templates: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Auto Generate Presentation",
        ),
    )
    def auto_generate_presentation(
        presentation_id: str,
        topic: str,
        slide_count: int = 5,
        presentation_type: str = "business",
        color_scheme: str = "modern_blue",
        include_charts: bool = True,
        include_images: bool = False,
    ) -> Dict:
        """
        Automatically generate a presentation based on topic and preferences.

        Args:
            topic: Main topic/theme for the presentation
            slide_count: Number of slides to generate (3-20)
            presentation_type: Type of presentation ('business', 'academic', 'creative')
            color_scheme: Color scheme to use
            include_charts: Whether to include chart slides
            include_images: Whether to include image placeholders
            presentation_id: Presentation ID
        """
        if presentation_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        if slide_count < 3 or slide_count > 20:
            return {
                "error": "Slide count must be between 3 and 20"
            }

        try:
            # Define presentation structures based on type
            if presentation_type == "business":
                base_templates = [
                    ("title_slide", {
                     "title": f"{topic}", "subtitle": "Executive Presentation", "author": "Business Team"}),
                    ("agenda_slide", {
                     "agenda_items": "1. Executive Summary\n\n2. Current Situation\n\n3. Analysis & Insights\n\n4. Recommendations\n\n5. Next Steps"}),
                    ("key_metrics_dashboard", {
                     "title": "Key Performance Indicators"}),
                    ("text_with_image", {
                     "title": "Current Situation", "content": f"Overview of {topic}:\n• Current status\n• Key challenges\n• Market position"}),
                    ("two_column_text", {"title": "Analysis", "content_left": "Strengths:\n• Advantage 1\n• Advantage 2\n• Advantage 3",
                     "content_right": "Opportunities:\n• Opportunity 1\n• Opportunity 2\n• Opportunity 3"}),
                ]
                if include_charts:
                    base_templates.append(
                        ("chart_comparison", {"title": "Performance Comparison"}))
                base_templates.append(("thank_you_slide", {
                                      "contact": "Thank you for your attention\nQuestions & Discussion"}))

            elif presentation_type == "academic":
                base_templates = [
                    ("title_slide", {"title": f"Research on {topic}",
                     "subtitle": "Academic Study", "author": "Research Team"}),
                    ("agenda_slide", {
                     "agenda_items": "1. Introduction\n\n2. Literature Review\n\n3. Methodology\n\n4. Results\n\n5. Conclusions"}),
                    ("text_with_image", {
                     "title": "Introduction", "content": f"Research focus on {topic}:\n• Background\n• Problem statement\n• Research questions"}),
                    ("two_column_text", {"title": "Methodology", "content_left": "Approach:\n• Method 1\n• Method 2\n• Method 3",
                     "content_right": "Data Sources:\n• Source 1\n• Source 2\n• Source 3"}),
                    ("data_table_slide", {"title": "Results Summary"}),
                ]
                if include_charts:
                    base_templates.append(
                        ("chart_comparison", {"title": "Data Analysis"}))
                base_templates.append(("thank_you_slide", {
                                      "contact": "Questions & Discussion\nContact: research@university.edu"}))

            else:  # creative
                base_templates = [
                    ("title_slide", {"title": f"Creative Vision: {topic}",
                     "subtitle": "Innovative Concepts", "author": "Creative Team"}),
                    ("full_image_slide", {
                     "overlay_title": f"Exploring {topic}", "overlay_subtitle": "Creative possibilities"}),
                    ("three_column_layout", {"title": "Creative Concepts"}),
                    ("quote_testimonial", {
                     "quote_text": f"Innovation in {topic} requires thinking beyond conventional boundaries", "attribution": "— Creative Director"}),
                    ("process_flow", {"title": "Creative Process"}),
                ]
                if include_charts:
                    base_templates.append(
                        ("key_metrics_dashboard", {"title": "Impact Metrics"}))
                base_templates.append(("thank_you_slide", {
                                      "contact": "Let's create something amazing together\ncreative@studio.com"}))

            # Adjust templates to match requested slide count
            template_sequence = []
            templates_to_use = base_templates[:slide_count]

            # If we need more slides, add content slides
            while len(templates_to_use) < slide_count:
                if include_images:
                    templates_to_use.insert(-1, ("text_with_image", {
                                            "title": f"{topic} - Additional Topic", "content": "• Key point\n• Supporting detail\n• Additional insight"}))
                else:
                    templates_to_use.insert(-1, ("two_column_text", {
                                            "title": f"{topic} - Analysis", "content_left": "Key Points:\n• Point 1\n• Point 2", "content_right": "Details:\n• Detail 1\n• Detail 2"}))

            # Convert to proper template sequence format
            for i, (template_id, content) in enumerate(templates_to_use):
                template_config = {
                    "template_id": template_id,
                    "content": content
                }
                template_sequence.append(template_config)

            # Create the presentation
            result = template_utils.create_presentation_from_template_sequence(
                presentations[presentation_id], template_sequence, color_scheme
            )

            return {
                "message": f"Auto-generated {slide_count}-slide presentation on '{topic}'",
                "topic": topic,
                "presentation_type": presentation_type,
                "color_scheme": color_scheme,
                "slide_count": slide_count,
                "generation_result": result,
                "templates_used": [t[0] for t in templates_to_use]
            }

        except Exception as e:
            return {
                "error": f"Failed to auto-generate presentation: {str(e)}"
            }

    # Text optimization tools

    @app.tool(
        annotations=ToolAnnotations(
            title="Optimize Slide Text",
        ),
    )
    def optimize_slide_text(
        presentation_id: str,
        slide_index: int,
        auto_resize: bool = True,
        auto_wrap: bool = True,
        optimize_spacing: bool = True,
        min_font_size: int = 8,
        max_font_size: int = 36,
    ) -> Dict:
        """
        Optimize text elements on a slide for better readability and fit.

        Args:
            slide_index: Index of the slide to optimize
            auto_resize: Whether to automatically resize fonts to fit containers
            auto_wrap: Whether to apply intelligent text wrapping
            optimize_spacing: Whether to optimize line spacing
            min_font_size: Minimum allowed font size
            max_font_size: Maximum allowed font size
            presentation_id: Presentation ID
        """
        if presentation_id not in presentations:
            return {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }

        pres = presentations[presentation_id]

        if slide_index < 0 or slide_index >= len(pres.slides):
            return {
                "error": f"Invalid slide index: {slide_index}. Available slides: 0-{len(pres.slides) - 1}"
            }

        slide = pres.slides[slide_index]

        try:
            optimizations_applied = []
            manager = template_utils.get_enhanced_template_manager()

            # Analyze each text shape on the slide
            for i, shape in enumerate(slide.shapes):
                if hasattr(shape, 'text_frame') and shape.text_frame.text:
                    text = shape.text_frame.text

                    # Calculate container dimensions
                    container_width = shape.width.inches
                    container_height = shape.height.inches

                    shape_optimizations = []

                    # Apply auto-resize if enabled
                    if auto_resize:
                        optimal_size = template_utils.calculate_dynamic_font_size(
                            text, container_width, container_height
                        )
                        optimal_size = max(min_font_size, min(
                            max_font_size, optimal_size))

                        # Apply the calculated font size
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                run.font.size = template_utils.Pt(optimal_size)

                        shape_optimizations.append(
                            f"Font resized to {optimal_size}pt")

                    # Apply auto-wrap if enabled
                    if auto_wrap:
                        current_font_size = 14  # Default assumption
                        if shape.text_frame.paragraphs and shape.text_frame.paragraphs[0].runs:
                            if shape.text_frame.paragraphs[0].runs[0].font.size:
                                current_font_size = shape.text_frame.paragraphs[0].runs[0].font.size.pt

                        wrapped_text = template_utils.wrap_text_automatically(
                            text, container_width, current_font_size
                        )

                        if wrapped_text != text:
                            shape.text_frame.text = wrapped_text
                            shape_optimizations.append(
                                "Text wrapped automatically")

                    # Optimize spacing if enabled
                    if optimize_spacing:
                        text_length = len(text)
                        if text_length > 300:
                            line_spacing = 1.4
                        elif text_length > 150:
                            line_spacing = 1.3
                        else:
                            line_spacing = 1.2

                        for paragraph in shape.text_frame.paragraphs:
                            paragraph.line_spacing = line_spacing

                        shape_optimizations.append(
                            f"Line spacing set to {line_spacing}")

                    if shape_optimizations:
                        optimizations_applied.append({
                            "shape_index": i,
                            "optimizations": shape_optimizations
                        })

            return {
                "message": f"Optimized {len(optimizations_applied)} text elements on slide {slide_index}",
                "slide_index": slide_index,
                "optimizations_applied": optimizations_applied,
                "settings": {
                    "auto_resize": auto_resize,
                    "auto_wrap": auto_wrap,
                    "optimize_spacing": optimize_spacing,
                    "font_size_range": f"{min_font_size}-{max_font_size}pt"
                }
            }

        except Exception as e:
            return {
                "error": f"Failed to optimize slide text: {str(e)}"
            }
