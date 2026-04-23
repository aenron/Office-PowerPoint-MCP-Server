"""
Presentation management tools for PowerPoint MCP Server.
Handles presentation creation, opening, saving, and core properties.
"""
from typing import Dict, List, Any
import os
from starlette.requests import Request
from starlette.responses import FileResponse, JSONResponse
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
import utils as ppt_utils


def register_presentation_tools(app: FastMCP, presentations: Dict, get_current_presentation_id, get_template_search_directories):
    """Register presentation management tools with the FastMCP app"""
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    download_dir = os.path.join(project_root, "ppt")
    download_url = os.environ.get("DOWNLOAD_URL")
    os.makedirs(download_dir, exist_ok=True)

    def get_effective_template_directories(template_directory: str = "") -> List[str]:
        if template_directory:
            expanded_directory = os.path.abspath(
                os.path.expanduser(template_directory))
            if not os.path.isdir(expanded_directory):
                return []
            return [expanded_directory]

        directories: List[str] = []
        for directory in get_template_search_directories():
            expanded_directory = os.path.abspath(os.path.expanduser(directory))
            if os.path.isdir(expanded_directory) and expanded_directory not in directories:
                directories.append(expanded_directory)

        return directories

    def resolve_template_path(template_name: str, template_directory: str = "") -> Dict[str, Any]:
        if os.path.exists(template_name) and os.path.isfile(template_name):
            return {
                "found": True,
                "template_path": os.path.abspath(template_name),
                "searched_directories": []
            }

        search_directories = get_effective_template_directories(
            template_directory)
        if not search_directories:
            return {
                "found": False,
                "error": f"Template directory not found: {template_directory}",
                "searched_directories": []
            }

        template_filename = os.path.basename(template_name)
        normalized_template_name = template_filename.lower()

        for directory in search_directories:
            for root, _, files in os.walk(directory):
                for file_name in files:
                    if file_name.lower() == normalized_template_name:
                        return {
                            "found": True,
                            "template_path": os.path.join(root, file_name),
                            "searched_directories": search_directories
                        }

        return {
            "found": False,
            "error": f"Template file not found: {template_name}",
            "searched_directories": search_directories
        }

    def get_required_presentation(presentation_id: str):
        if presentation_id not in presentations:
            return None, {
                "error": "No presentation is currently loaded or the specified ID is invalid"
            }
        return presentations[presentation_id], None

    def update_core_property(presentation_id: str, property_name: str, property_value: str) -> Dict:
        pres, error = get_required_presentation(presentation_id)
        if error:
            return error

        try:
            ppt_utils.set_core_properties(
                pres,
                **{property_name: property_value}
            )
            return {
                "message": f"Core property '{property_name}' updated successfully",
                "presentation_id": presentation_id,
                property_name: property_value
            }
        except Exception as e:
            return {
                "error": f"Failed to set core property '{property_name}': {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="List Presentation Templates",
            readOnlyHint=True,
        ),
    )
    def list_presentation_templates() -> Dict:
        """List available PowerPoint template files from the default template directories."""
        search_directories = get_effective_template_directories()

        templates: List[Dict[str, str]] = []
        seen_paths = set()

        for directory in search_directories:
            for root, _, files in os.walk(directory):
                for file_name in files:
                    if not file_name.lower().endswith(".pptx"):
                        continue

                    file_path = os.path.join(root, file_name)
                    normalized_path = os.path.normcase(
                        os.path.abspath(file_path))
                    if normalized_path in seen_paths:
                        continue

                    seen_paths.add(normalized_path)
                    templates.append({
                        "template_name": file_name,
                        "template_path": os.path.abspath(file_path),
                        "template_directory": os.path.abspath(root)
                    })

        templates.sort(key=lambda item: item["template_name"].lower())

        return {
            "searched_directories": search_directories,
            "templates": templates,
            "total_templates": len(templates)
        }

    @app.custom_route("/downloads/{filename}", methods=["GET"])
    async def download_presentation(request: Request):
        filename = os.path.basename(request.path_params["filename"])
        file_path = os.path.join(download_dir, filename)

        if not os.path.exists(file_path):
            return JSONResponse({"error": f"File not found: {filename}"}, status_code=404)

        return FileResponse(
            path=file_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation",
        ),
    )
    def create_presentation(id: str) -> Dict:
        """Create a new PowerPoint presentation."""
        # Create a new presentation
        pres = ppt_utils.create_presentation()

        # Store the presentation
        presentations[id] = pres
        # Set as current presentation (this would need to be handled by caller)

        return {
            "presentation_id": id,
            "message": f"Created new presentation with ID: {id}",
            "slide_count": len(pres.slides)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Create Presentation from Template",
        ),
    )
    def create_presentation_from_template(template_path: str, id: str) -> Dict:
        """Create a new PowerPoint presentation from a template file."""
        resolved_template = resolve_template_path(template_path)
        if not resolved_template["found"]:
            env_path_info = f" (PPT_TEMPLATE_PATH: {os.environ.get('PPT_TEMPLATE_PATH', 'not set')})" if os.environ.get(
                'PPT_TEMPLATE_PATH') else ""
            return {
                "error": f"{resolved_template['error']}. Searched in {', '.join(resolved_template['searched_directories'])}{env_path_info}"
            }

        template_path = resolved_template["template_path"]

        try:
            pres = ppt_utils.create_presentation_from_template(template_path)
        except Exception as e:
            return {
                "error": f"Failed to create presentation from template: {str(e)}"
            }

        # Store the presentation
        presentations[id] = pres

        return {
            "presentation_id": id,
            "message": f"Created new presentation from template '{template_path}' with ID: {id}",
            "template_path": template_path,
            "resolved_template_path": template_path,
            "slide_count": len(pres.slides),
            "layout_count": len(pres.slide_layouts)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Open Presentation",
            readOnlyHint=True,
        ),
    )
    def open_presentation(file_path: str, id: str) -> Dict:
        """Open an existing PowerPoint presentation from a file."""
        # Check if file exists
        if not os.path.exists(file_path):
            return {
                "error": f"File not found: {file_path}"
            }

        # Open the presentation
        try:
            pres = ppt_utils.open_presentation(file_path)
        except Exception as e:
            return {
                "error": f"Failed to open presentation: {str(e)}"
            }

        # Store the presentation
        presentations[id] = pres

        return {
            "presentation_id": id,
            "message": f"Opened presentation from {file_path} with ID: {id}",
            "slide_count": len(pres.slides)
        }

    @app.tool(
        annotations=ToolAnnotations(
            title="Save Presentation",
            destructiveHint=True,
        ),
    )
    def save_presentation(file_path: str, presentation_id: str) -> Dict:
        """Save a presentation to a file."""
        pres, error = get_required_presentation(presentation_id)
        if error:
            return error

        file_name = os.path.basename(
            file_path) if file_path else f"{presentation_id}.pptx"
        if not file_name.lower().endswith('.pptx'):
            file_name = f"{file_name}.pptx"

        saved_path = os.path.join(download_dir, file_name)

        try:
            saved_path = ppt_utils.save_presentation(pres, saved_path)
            port = getattr(app.settings, "port", 8000)
            download_base_url = (
                download_url or f"http://localhost:{port}").rstrip("/")
            return {
                "message": f"Presentation saved to {saved_path}",
                "file_path": saved_path,
                "download_url": f"{download_base_url}/downloads/{file_name}"
            }
        except Exception as e:
            return {
                "error": f"Failed to save presentation: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Presentation Info",
            readOnlyHint=True,
        ),
    )
    def get_presentation_info(presentation_id: str) -> Dict:
        """Get information about a presentation."""
        pres, error = get_required_presentation(presentation_id)
        if error:
            return error

        try:
            info = ppt_utils.get_presentation_info(pres)
            info["presentation_id"] = presentation_id
            return info
        except Exception as e:
            return {
                "error": f"Failed to get presentation info: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Get Template File Info",
            readOnlyHint=True,
        ),
    )
    def get_template_file_info(template_path: str) -> Dict:
        """Get information about a template file including layouts and properties."""
        resolved_template = resolve_template_path(template_path)
        if not resolved_template["found"]:
            return {
                "error": f"{resolved_template['error']}. Searched in {', '.join(resolved_template['searched_directories'])}"
            }

        template_path = resolved_template["template_path"]

        try:
            template_info = ppt_utils.get_template_info(template_path)
            template_info["resolved_template_path"] = template_path
            return template_info
        except Exception as e:
            return {
                "error": f"Failed to get template info: {str(e)}"
            }

    @app.tool(
        annotations=ToolAnnotations(
            title="Set Presentation Core Property",
        ),
    )
    def set_presentation_core_property(
        presentation_id: str,
        property_name: str,
        property_value: str,
    ) -> Dict:
        """Set a presentation core property."""
        valid_properties = ["title", "subject",
                            "author", "keywords", "comments"]
        if property_name not in valid_properties:
            return {
                "error": f"Invalid core property '{property_name}'. Valid properties are: {', '.join(valid_properties)}"
            }

        return update_core_property(presentation_id, property_name, property_value)
