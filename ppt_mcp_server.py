#!/usr/bin/env python
"""
MCP Server for PowerPoint manipulation using python-pptx.
Consolidated server with low-level and simplified workflow tools.
"""
import os
import argparse
from typing import Dict, Any
from mcp.server.fastmcp import FastMCP
from mcp.server.transport_security import TransportSecuritySettings

# import utils  # Currently unused
from tools import (
    register_presentation_tools,
    register_content_tools,
    register_structural_tools,
    register_professional_tools,
    register_template_tools,
    register_hyperlink_tools,
    register_chart_tools,
    register_connector_tools,
    register_master_tools,
    register_transition_tools,
    register_workflow_tools,
)

# Initialize the FastMCP server
app = FastMCP(
    name="ppt-mcp-server"
)

# Global state to store presentations in memory
presentations = {}
projects = {}
current_presentation_id = None
advanced_tools_enabled = False

# Template configuration


def get_template_search_directories():
    """
    Get list of directories to search for templates.
    Uses environment variable PPT_TEMPLATE_PATH if set, otherwise uses default directories.

    Returns:
        List of directories to search for templates
    """
    default_template_directories = ['./templates']
    template_env_path = os.environ.get('PPT_TEMPLATE_PATH')

    if template_env_path:
        # If environment variable is set, use it as the primary template directory
        # Support multiple paths separated by colon (Unix) or semicolon (Windows)
        import platform
        separator = ';' if platform.system() == "Windows" else ':'
        env_dirs = [path.strip()
                    for path in template_env_path.split(separator) if path.strip()]

        # Verify that the directories exist
        valid_env_dirs = []
        for dir_path in env_dirs:
            expanded_path = os.path.expanduser(dir_path)
            if os.path.exists(expanded_path) and os.path.isdir(expanded_path):
                valid_env_dirs.append(expanded_path)

        if valid_env_dirs:
            # Add default fallback directories
            return valid_env_dirs + default_template_directories
        else:
            print(
                f"Warning: PPT_TEMPLATE_PATH directories not found: {template_env_path}")

    # Default search directories when no environment variable or invalid paths
    return default_template_directories


def is_truthy_env(value: str) -> bool:
    """Interpret common truthy environment variable values."""
    return (value or "").strip().lower() in {"1", "true", "yes", "on"}

# ---- Helper Functions ----


def get_current_presentation():
    """Get the current presentation object or raise an error if none is loaded."""
    if current_presentation_id is None or current_presentation_id not in presentations:
        raise ValueError(
            "No presentation is currently loaded. Please create or open a presentation first.")
    return presentations[current_presentation_id]


def get_current_presentation_id():
    """Get the current presentation ID."""
    return current_presentation_id


def set_current_presentation_id(pres_id):
    """Set the current presentation ID."""
    global current_presentation_id
    current_presentation_id = pres_id


def set_advanced_tools_enabled(enabled: bool):
    """Store whether low-level tools are exposed."""
    global advanced_tools_enabled
    advanced_tools_enabled = enabled


def validate_parameters(params):
    """
    Validate parameters against constraints.

    Args:
        params: Dictionary of parameter name: (value, constraints) pairs

    Returns:
        (True, None) if all valid, or (False, error_message) if invalid
    """
    for param_name, (value, constraints) in params.items():
        for constraint_func, error_msg in constraints:
            if not constraint_func(value):
                return False, f"Parameter '{param_name}': {error_msg}"
    return True, None


def is_positive(value):
    """Check if a value is positive."""
    return value > 0


def is_non_negative(value):
    """Check if a value is non-negative."""
    return value >= 0


def is_in_range(min_val, max_val):
    """Create a function that checks if a value is in a range."""
    return lambda x: min_val <= x <= max_val


def is_in_list(valid_list):
    """Create a function that checks if a value is in a list."""
    return lambda x: x in valid_list


def is_valid_rgb(color_list):
    """Check if a color list is a valid RGB tuple."""
    if not isinstance(color_list, list) or len(color_list) != 3:
        return False
    return all(isinstance(c, int) and 0 <= c <= 255 for c in color_list)


def add_shape_direct(slide, shape_type: str, left: float, top: float, width: float, height: float) -> Any:
    """
    Add an auto shape to a slide using direct integer values instead of enum objects.

    This implementation provides a reliable alternative that bypasses potential 
    enum-related issues in the python-pptx library.

    Args:
        slide: The slide object
        shape_type: Shape type string (e.g., 'rectangle', 'oval', 'triangle')
        left: Left position in inches
        top: Top position in inches
        width: Width in inches
        height: Height in inches

    Returns:
        The created shape
    """
    from pptx.util import Inches

    # Direct mapping of shape types to their integer values
    # Values from MSO_AUTO_SHAPE_TYPE enum: https://github.com/scanny/python-pptx/blob/master/src/pptx/enum/shapes.py
    shape_type_map = {
        'rectangle': 1,              # RECTANGLE
        'rounded_rectangle': 5,      # ROUNDED_RECTANGLE
        'oval': 9,                   # OVAL
        'diamond': 4,                # DIAMOND
        'triangle': 7,               # ISOSCELES_TRIANGLE
        'right_triangle': 8,         # RIGHT_TRIANGLE
        'pentagon': 51,              # PENTAGON
        'hexagon': 10,               # HEXAGON
        'heptagon': 145,             # HEPTAGON
        'octagon': 6,                # OCTAGON
        'star': 92,                  # STAR_5_POINT
        'arrow': 33,                 # RIGHT_ARROW
        'cloud': 179,                # CLOUD
        'heart': 21,                 # HEART
        'lightning_bolt': 22,        # LIGHTNING_BOLT
        'sun': 23,                   # SUN
        'moon': 24,                  # MOON
        'smiley_face': 17,           # SMILEY_FACE
        'no_symbol': 19,             # NO_SYMBOL
        'flowchart_process': 61,     # FLOWCHART_PROCESS
        'flowchart_decision': 63,    # FLOWCHART_DECISION
        'flowchart_data': 64,        # FLOWCHART_DATA
        'flowchart_document': 67     # FLOWCHART_DOCUMENT
    }

    # Check if shape type is valid before trying to use it
    shape_type_lower = str(shape_type).lower()
    if shape_type_lower not in shape_type_map:
        available_shapes = ', '.join(sorted(shape_type_map.keys()))
        raise ValueError(
            f"Unsupported shape type: '{shape_type}'. Available shape types: {available_shapes}")

    # Get the integer value for the shape type
    shape_value = shape_type_map[shape_type_lower]

    # Create the shape using the direct integer value
    try:
        # The integer value is passed directly to add_shape
        shape = slide.shapes.add_shape(
            shape_value, Inches(left), Inches(
                top), Inches(width), Inches(height)
        )
        return shape
    except Exception as e:
        raise ValueError(
            f"Failed to create '{shape_type}' shape using direct value {shape_value}: {str(e)}")

# ---- Custom presentation management wrapper ----


class PresentationManager:
    """Wrapper to handle presentation state updates."""

    def __init__(self, presentations_dict):
        self.presentations = presentations_dict

    def store_presentation(self, pres, pres_id):
        """Store a presentation and set it as current."""
        self.presentations[pres_id] = pres
        set_current_presentation_id(pres_id)
        return pres_id

# ---- Register Tools ----


# Create presentation manager wrapper
presentation_manager = PresentationManager(presentations)

# Wrapper functions to handle state management


def create_presentation_wrapper(original_func):
    """Wrapper to handle presentation creation with state management."""
    def wrapper(*args, **kwargs):
        result = original_func(*args, **kwargs)
        if "presentation_id" in result and result["presentation_id"] in presentations:
            set_current_presentation_id(result["presentation_id"])
        return result
    return wrapper


def open_presentation_wrapper(original_func):
    """Wrapper to handle presentation opening with state management."""
    def wrapper(*args, **kwargs):
        result = original_func(*args, **kwargs)
        if "presentation_id" in result and result["presentation_id"] in presentations:
            set_current_presentation_id(result["presentation_id"])
        return result
    return wrapper


register_workflow_tools(
    app,
    presentations,
    projects,
    get_current_presentation_id,
    set_current_presentation_id,
    get_template_search_directories,
)


def register_advanced_tools():
    """Register low-level and compatibility tools for advanced mode."""
    register_presentation_tools(
        app,
        presentations,
        get_current_presentation_id,
        get_template_search_directories
    )

    register_content_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb
    )

    register_structural_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb,
        add_shape_direct
    )

    register_professional_tools(
        app,
        presentations,
        get_current_presentation_id
    )

    register_template_tools(
        app,
        presentations,
        get_current_presentation_id
    )

    register_hyperlink_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb
    )

    register_chart_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb
    )

    register_connector_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb
    )

    register_master_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb
    )

    register_transition_tools(
        app,
        presentations,
        get_current_presentation_id,
        validate_parameters,
        is_positive,
        is_non_negative,
        is_in_range,
        is_valid_rgb
    )

    register_advanced_utility_tools()


def register_advanced_utility_tools():
    """Register compatibility and diagnostic tools for advanced mode."""

    @app.tool()
    def list_presentations() -> Dict:
        """List all loaded presentations."""
        return {
            "presentations": [
                {
                    "id": pres_id,
                    "slide_count": len(pres.slides),
                    "is_current": pres_id == current_presentation_id
                }
                for pres_id, pres in presentations.items()
            ],
            "current_presentation_id": current_presentation_id,
            "total_presentations": len(presentations)
        }


    @app.tool()
    def switch_presentation(presentation_id: str) -> Dict:
        """Switch to a different loaded presentation."""
        if presentation_id not in presentations:
            return {
                "error": f"Presentation '{presentation_id}' not found. Available presentations: {list(presentations.keys())}"
            }

        global current_presentation_id
        old_id = current_presentation_id
        current_presentation_id = presentation_id

        return {
            "message": f"Switched from presentation '{old_id}' to '{presentation_id}'",
            "previous_presentation_id": old_id,
            "current_presentation_id": current_presentation_id
        }


    @app.tool()
    def get_server_info() -> Dict:
        """Get information about the MCP server."""
        total_tools = len(getattr(app._tool_manager, "_tools", {}))
        return {
            "name": "PowerPoint MCP Server - Enhanced Edition",
            "version": "2.1.0",
            "total_tools": total_tools,
            "loaded_presentations": len(presentations),
            "current_presentation": current_presentation_id,
            "active_projects": len(projects),
            "advanced_tools_enabled": advanced_tools_enabled,
            "features": [
                "Simplified Workflow (6 tools)",
                "Optional advanced editing surface for low-level operations",
            ],
            "improvements": [
                f"{total_tools} registered tools in the current exposure mode",
                "Default simplified MCP interface focused on end-to-end PPT generation",
                "Optional advanced mode for detailed editing, diagnostics, and compatibility",
                "Template-aware project, planning, build, revise, and export workflow",
            ],
        }


if is_truthy_env(os.environ.get("PPT_ENABLE_ADVANCED_TOOLS")):
    set_advanced_tools_enabled(True)
    register_advanced_tools()

# ---- Main Function ----


def env_int(name: str, default: int) -> int:
    """Read a positive integer from the environment."""
    try:
        value = int((os.environ.get(name) or "").strip())
        return value if value > 0 else default
    except ValueError:
        return default


def run_uvicorn_app(starlette_app, host: str, port: int, timeout_keep_alive: int, timeout_graceful_shutdown: int):
    """Run a FastMCP Starlette app with explicit uvicorn timeout settings."""
    import asyncio
    import uvicorn

    config = uvicorn.Config(
        starlette_app,
        host=host,
        port=port,
        log_level=app.settings.log_level.lower(),
        timeout_keep_alive=timeout_keep_alive,
        timeout_graceful_shutdown=timeout_graceful_shutdown,
    )
    server = uvicorn.Server(config)
    try:
        asyncio.run(server.serve())
    except asyncio.exceptions.CancelledError:
        print("Server stopped by user.")
    except KeyboardInterrupt:
        print("Server stopped by user.")
    except Exception as e:
        print(f"Error starting server: {e}")


def main(
    transport: str = "stdio",
    port: int = 8000,
    host: str = "127.0.0.1",
    allowed_hosts: str = "",
    timeout_keep_alive: int = None,
    timeout_graceful_shutdown: int = None,
):
    timeout_keep_alive = timeout_keep_alive or env_int(
        "MCP_TIMEOUT_KEEP_ALIVE", 120)
    timeout_graceful_shutdown = timeout_graceful_shutdown or env_int(
        "MCP_TIMEOUT_GRACEFUL_SHUTDOWN", 300)
    parsed_allowed_hosts = [item.strip()
                            for item in allowed_hosts.split(",") if item.strip()]

    if transport == "http":
        # Set the port for HTTP transport
        app.settings.host = host
        app.settings.port = port
        if parsed_allowed_hosts:
            app.settings.transport_security = TransportSecuritySettings(
                enable_dns_rebinding_protection=True,
                allowed_hosts=parsed_allowed_hosts,
            )
        # Start the FastMCP server with HTTP transport
        run_uvicorn_app(
            app.streamable_http_app(),
            host,
            port,
            timeout_keep_alive,
            timeout_graceful_shutdown,
        )

    elif transport == "sse":
        app.settings.host = host
        app.settings.port = port
        if parsed_allowed_hosts:
            app.settings.transport_security = TransportSecuritySettings(
                enable_dns_rebinding_protection=True,
                allowed_hosts=parsed_allowed_hosts,
            )
        # Run the FastMCP server in SSE (Server Side Events) mode
        run_uvicorn_app(
            app.sse_app(),
            host,
            port,
            timeout_keep_alive,
            timeout_graceful_shutdown,
        )

    else:
        # Run the FastMCP server
        app.run(transport='stdio')


if __name__ == "__main__":
    # Parse command line arguments
    parser = argparse.ArgumentParser(
        description="MCP Server for PowerPoint manipulation using python-pptx")

    parser.add_argument(
        "-t",
        "--transport",
        type=str,
        default="stdio",
        choices=["stdio", "http", "sse"],
        help="Transport method for the MCP server (default: stdio)"
    )

    parser.add_argument(
        "-p",
        "--port",
        type=int,
        default=8000,
        help="Port to run the MCP server on (default: 8000)"
    )

    parser.add_argument(
        "--host",
        type=str,
        default="127.0.0.1",
        help="Host to bind the MCP server to (default: 127.0.0.1)"
    )

    parser.add_argument(
        "--allowed-hosts",
        type=str,
        default="",
        help="Comma-separated Host header allowlist, e.g. '59.110.150.224:7380,example.com:7380'"
    )
    parser.add_argument(
        "--timeout-keep-alive",
        type=int,
        default=None,
        help="Uvicorn keep-alive timeout in seconds for HTTP/SSE transports (default: MCP_TIMEOUT_KEEP_ALIVE or 120)"
    )
    parser.add_argument(
        "--timeout-graceful-shutdown",
        type=int,
        default=None,
        help="Uvicorn graceful shutdown timeout in seconds for HTTP/SSE transports (default: MCP_TIMEOUT_GRACEFUL_SHUTDOWN or 300)"
    )
    args = parser.parse_args()
    main(
        args.transport,
        args.port,
        args.host,
        args.allowed_hosts,
        args.timeout_keep_alive,
        args.timeout_graceful_shutdown,
    )
