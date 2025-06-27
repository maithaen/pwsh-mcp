import asyncio
import json
import sys
import time
import logging
import os
from typing import Dict, Any, Optional
import pyperclip
from powershell_mcp.windows_terminal_controller import WindowsTerminalController

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class PowerShellMCPServer:
    """MCP Server for PowerShell automation"""

    def __init__(self):
        self.terminal_controller = WindowsTerminalController()
        self.tools = self._define_tools()

    def _define_tools(self) -> Dict[str, Dict[str, Any]]:
        """Define all available tools"""
        return {
            "execute_pwsh_script": {
                "name": "execute_pwsh_script",
                "description": "Execute PowerShell script by pasting into terminal (supports both single-line and multi-line scripts)",  # noqa E501
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "script": {"type": "string", "description": "PowerShell script to execute (single-line or multi-line)"},
                        "timeout": {"type": "integer", "description": "Timeout in seconds (default: 30)", "default": 30},
                    },
                    "required": ["script"],
                },
            },
            "get_clipboard": {
                "name": "get_clipboard",
                "description": "Get current clipboard content",
                "inputSchema": {
                    "type": "object",
                    "properties": {},
                },
            },
            "capture_pwsh_response": {
                "name": "capture_pwsh_response",
                "description": "Capture Windows PowerShell terminal output as screenshot",
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "save_path": {"type": "string", "description": "Optional path to save screenshot", "default": None},
                        "exclude_titlebar": {
                            "type": "boolean",
                            "description": "Exclude window title bar from capture (default: true)",
                            "default": True,
                        },
                    },
                },
            },
        }

    async def handle_request(self, request: Dict[str, Any]) -> Dict[str, Any]:
        """Handle incoming MCP requests"""
        try:
            method = request.get("method")
            params = request.get("params", {})
            request_id = request.get("id", 1)

            logger.debug(f"Handling request: {method}")

            if method == "initialize":
                return self._handle_initialize(request_id)
            elif method == "tools/list":
                return self._handle_tools_list(request_id)
            elif method == "tools/call":
                return await self._handle_tools_call(request_id, params)
            else:
                return self._create_error_response(request_id, -32601, f"Unknown method: {method}")

        except Exception as e:
            logger.error(f"Error handling request: {e}")
            error_id = request.get("id", 1)
            return self._create_error_response(error_id, -32603, str(e))

    def _handle_initialize(self, request_id: int) -> Dict[str, Any]:
        """Handle initialize request"""
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": "2024-11-05",
                "capabilities": {"tools": {}},
                "serverInfo": {"name": "powershell-mcp-server", "version": "1.1.0"},
            },
        }

    def _handle_tools_list(self, request_id: int) -> Dict[str, Any]:
        """Handle tools list request"""
        return {"jsonrpc": "2.0", "id": request_id, "result": {"tools": list(self.tools.values())}}

    async def _handle_tools_call(self, request_id: int, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle tools call request"""
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        logger.info(f"Calling tool: {tool_name}")

        try:
            if tool_name == "execute_pwsh_script":
                result = await self.execute_pwsh_script(arguments.get("script", ""), arguments.get("timeout", 30))
            elif tool_name == "get_clipboard":
                result = await self.get_clipboard()
            elif tool_name == "capture_pwsh_response":
                result = await self.capture_pwsh_response(arguments.get("save_path"), arguments.get("exclude_titlebar", True))
            else:
                return self._create_error_response(request_id, -32601, f"Unknown tool: {tool_name}")

            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": {"content": [{"type": "text", "text": json.dumps(result, indent=2)}]},
            }

        except Exception as e:
            logger.error(f"Error executing tool {tool_name}: {e}")
            return self._create_error_response(request_id, -32603, f"Tool execution failed: {str(e)}")

    def _create_error_response(self, request_id: int, code: int, message: str) -> Dict[str, Any]:
        """Create standardized error response"""
        return {"jsonrpc": "2.0", "id": request_id, "error": {"code": code, "message": message}}

    async def execute_pwsh_script(self, script: str, timeout: int = 30) -> Dict[str, Any]:
        """Execute PowerShell script by pasting into terminal (supports both single-line and multi-line scripts)"""
        try:
            if not script.strip():
                return {"success": False, "error": "Empty script provided"}

            # Count non-empty lines
            lines = [line.strip() for line in script.split("\n") if line.strip()]
            is_multiline = len(lines) > 1

            # Ensure terminal is running and window is available
            if not self.terminal_controller.is_terminal_running() or not self.terminal_controller.find_terminal_window():
                logger.info("Terminal not running or window not found, launching...")
                if not self.terminal_controller.launch_terminal():
                    return {"success": False, "error": "Failed to launch Windows Terminal"}

            # Double-check window is available and focused
            for _ in range(10):
                if self.terminal_controller.find_terminal_window():
                    if self.terminal_controller.focus_terminal():
                        break
                await asyncio.sleep(0.5)

            self.terminal_controller.timeout = timeout

            # Paste the script (works for both single-line and multi-line)
            script_type = "multi-line script" if is_multiline else "single command"
            logger.info(f"Pasting {script_type} ({len(lines)} line{'s' if len(lines) != 1 else ''})...")
            success = self.terminal_controller.paste_content(script, execute=True)

            if success:
                # Give more time for multi-line scripts
                await asyncio.sleep(2 if is_multiline else 1)
                return {
                    "success": True,
                    "lines_count": len(lines),
                    "is_multiline": is_multiline,
                    "message": f"{script_type.capitalize()} with {len(lines)} line{'s' if len(lines) != 1 else ''} pasted and executed successfully",  # noqa E501
                }
            else:
                return {"success": False, "error": f"Failed to paste {script_type}"}

        except Exception as e:
            logger.error(f"Error executing PowerShell script: {e}")
            return {"success": False, "error": str(e)}

    async def get_clipboard(self) -> Dict[str, Any]:
        """Get current clipboard content"""
        try:
            clipboard_content = pyperclip.paste()

            return {
                "success": True,
                "content": clipboard_content,
                "length": len(clipboard_content),
                "message": "Clipboard content retrieved successfully",
            }

        except Exception as e:
            logger.error(f"Error getting clipboard content: {e}")
            return {"success": False, "error": str(e)}

    async def capture_pwsh_response(self, save_path: Optional[str] = None, exclude_titlebar: bool = True) -> Dict[str, Any]:
        """Capture terminal output"""
        try:
            screenshot = self.terminal_controller.capture_terminal_output(exclude_titlebar)

            if screenshot is None:
                return {"success": False, "error": "Failed to capture terminal window"}

            # Determine save path
            if save_path:
                final_path = save_path
            else:
                # Save to temp/ directory
                temp_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "..", "temp")
                os.makedirs(temp_dir, exist_ok=True)
                final_path = os.path.join(temp_dir, f"terminal_capture_{int(time.time())}.png")

            # Save screenshot
            screenshot.save(final_path)

            return {
                "success": True,
                "saved_to": final_path,
                "size": {"width": screenshot.size[0], "height": screenshot.size[1]},
                "exclude_titlebar": exclude_titlebar,
                "message": f"Screenshot saved to {'specified path' if save_path else 'temp directory'}",
            }

        except Exception as e:
            logger.error(f"Error capturing terminal response: {e}")
            return {"success": False, "error": str(e)}

    async def run_server(self):
        """Run the MCP server"""
        logger.info("Starting PowerShell MCP Server v1.1.0...")
        logger.info("Available tools: execute_pwsh_script, get_clipboard, capture_pwsh_response")

        while True:
            try:
                # Read request from stdin
                line = sys.stdin.readline()
                if not line:
                    logger.info("No more input, shutting down server")
                    break

                line = line.strip()
                if not line:
                    continue

                request = json.loads(line)
                response = await self.handle_request(request)

                # Write response to stdout
                print(json.dumps(response))
                sys.stdout.flush()

            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON received: {e}")
                continue
            except Exception as e:
                logger.error(f"Error processing request: {e}")
                continue

        logger.info("PowerShell MCP Server stopped")
