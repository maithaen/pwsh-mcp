import asyncio
import json
import sys
import time
import logging
import os
from dataclasses import dataclass
from typing import Dict, Any, Optional
import pyperclip
from powershell_mcp.windows_terminal_controller import WindowsTerminalController

# Configure logging
logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

# Constants
DEFAULT_TIMEOUT = 30
DEFAULT_RETRY_ATTEMPTS = 10
DEFAULT_RETRY_DELAY = 0.5
MULTILINE_EXECUTION_DELAY = 2
SINGLELINE_EXECUTION_DELAY = 1
PROTOCOL_VERSION = "2024-11-05"
SERVER_VERSION = "1.1.0"


@dataclass
class ToolResult:
    """Standardized tool execution result"""

    success: bool
    data: Optional[Dict[str, Any]] = None
    error: Optional[str] = None

    def to_dict(self) -> Dict[str, Any]:
        result: Dict[str, Any] = {"success": self.success}
        if self.data:
            result.update(self.data)
        if self.error is not None:
            result["error"] = self.error
        return result


class PowerShellMCPError(Exception):
    """Base exception for PowerShell MCP operations"""

    pass


class TerminalNotAvailableError(PowerShellMCPError):
    """Raised when terminal is not available or cannot be launched"""

    pass


class ScriptExecutionError(PowerShellMCPError):
    """Raised when script execution fails"""

    pass


class PowerShellMCPServer:
    """MCP Server for PowerShell automation"""

    def __init__(self, timeout: int = DEFAULT_TIMEOUT):
        self.terminal_controller = WindowsTerminalController(timeout=timeout)
        self.tools = self._define_tools()
        self.default_timeout = timeout
        logger.info(f"PowerShell MCP Server initialized with timeout: {timeout}s")

    def _define_tools(self) -> Dict[str, Dict[str, Any]]:
        """Define all available tools with comprehensive schemas"""
        return {
            "execute_pwsh_script": {
                "name": "execute_pwsh_script",
                "description": "Execute PowerShell script by pasting into terminal (supports both single-line and multi-line scripts)",  # noqa: E501
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "script": {
                            "type": "string",
                            "description": "PowerShell script to execute (single-line or multi-line)",
                            "minLength": 1,
                        },
                        "timeout": {
                            "type": "integer",
                            "description": f"Timeout in seconds (default: {DEFAULT_TIMEOUT})",
                            "default": DEFAULT_TIMEOUT,
                            "minimum": 1,
                            "maximum": 300,
                        },
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
        logger.info("Handling initialize request")
        return {
            "jsonrpc": "2.0",
            "id": request_id,
            "result": {
                "protocolVersion": PROTOCOL_VERSION,
                "capabilities": {"tools": {}},
                "serverInfo": {"name": "powershell-mcp-server", "version": SERVER_VERSION},
            },
        }

    def _handle_tools_list(self, request_id: int) -> Dict[str, Any]:
        """Handle tools list request"""
        logger.debug(f"Listing {len(self.tools)} available tools")
        return {"jsonrpc": "2.0", "id": request_id, "result": {"tools": list(self.tools.values())}}

    async def _handle_tools_call(self, request_id: int, params: Dict[str, Any]) -> Dict[str, Any]:
        """Handle tools call request with improved error handling"""
        tool_name = params.get("name")
        arguments = params.get("arguments", {})

        if not tool_name:
            return self._create_error_response(request_id, -32602, "Tool name is required")

        if tool_name not in self.tools:
            return self._create_error_response(request_id, -32601, f"Unknown tool: {tool_name}")

        logger.info(f"Executing tool: {tool_name} with arguments: {list(arguments.keys())}")

        try:
            # Validate arguments based on tool schema
            validation_error = self._validate_tool_arguments(tool_name, arguments)
            if validation_error:
                return self._create_error_response(request_id, -32602, validation_error)

            # Execute the appropriate tool
            if tool_name == "execute_pwsh_script":
                result = await self.execute_pwsh_script(
                    arguments.get("script", ""), arguments.get("timeout", self.default_timeout)
                )
            elif tool_name == "get_clipboard":
                result = await self.get_clipboard()
            elif tool_name == "capture_pwsh_response":
                result = await self.capture_pwsh_response(arguments.get("save_path"), arguments.get("exclude_titlebar", True))
            else:
                return self._create_error_response(request_id, -32601, f"Tool not implemented: {tool_name}")

            logger.info(f"Tool {tool_name} executed successfully: {result.get('success', False)}")
            return {
                "jsonrpc": "2.0",
                "id": request_id,
                "result": {"content": [{"type": "text", "text": json.dumps(result, indent=2)}]},
            }

        except PowerShellMCPError as e:
            logger.error(f"PowerShell MCP error in tool {tool_name}: {e}")
            return self._create_error_response(request_id, -32603, str(e))
        except Exception as e:
            logger.error(f"Unexpected error executing tool {tool_name}: {e}", exc_info=True)
            return self._create_error_response(request_id, -32603, f"Tool execution failed: {str(e)}")

    def _validate_tool_arguments(self, tool_name: str, arguments: Dict[str, Any]) -> Optional[str]:
        """Validate tool arguments against schema"""
        tool_schema = self.tools.get(tool_name, {}).get("inputSchema", {})
        required_fields = tool_schema.get("required", [])
        properties = tool_schema.get("properties", {})  # noqa: F841

        # Check required fields
        for field in required_fields:
            if field not in arguments:
                return f"Missing required argument: {field}"

        # Validate specific constraints
        if tool_name == "execute_pwsh_script":
            script = arguments.get("script", "")
            if not script or not script.strip():
                return "Script cannot be empty"

            timeout = arguments.get("timeout", self.default_timeout)
            if not isinstance(timeout, int) or timeout < 1 or timeout > 300:
                return "Timeout must be an integer between 1 and 300 seconds"

        return None

    def _create_error_response(self, request_id: int, code: int, message: str) -> Dict[str, Any]:
        """Create standardized error response"""
        logger.warning(f"Creating error response - Code: {code}, Message: {message}")
        return {"jsonrpc": "2.0", "id": request_id, "error": {"code": code, "message": message}}

    async def execute_pwsh_script(self, script: str, timeout: int = DEFAULT_TIMEOUT) -> Dict[str, Any]:
        """Execute PowerShell script by pasting into terminal with improved error handling"""
        try:
            # Validate input
            if not script or not script.strip():
                raise ScriptExecutionError("Empty script provided")

            # Analyze script structure
            lines = [line.strip() for line in script.split("\n") if line.strip()]
            is_multiline = len(lines) > 1
            script_type = "multi-line script" if is_multiline else "single command"

            logger.info(f"Preparing to execute {script_type} with {len(lines)} line(s)")

            # Ensure terminal availability
            await self._ensure_terminal_available()

            # Configure timeout
            self.terminal_controller.timeout = timeout

            # Execute script
            logger.info(f"Executing {script_type}...")
            success = self.terminal_controller.paste_content(script, execute=True)

            if not success:
                raise ScriptExecutionError(f"Failed to paste {script_type}")

            # Wait for execution to complete
            execution_delay = MULTILINE_EXECUTION_DELAY if is_multiline else SINGLELINE_EXECUTION_DELAY
            await asyncio.sleep(execution_delay)

            result = ToolResult(
                success=True,
                data={
                    "lines_count": len(lines),
                    "is_multiline": is_multiline,
                    "script_type": script_type,
                    "timeout_used": timeout,
                    "message": f"{script_type.capitalize()} with {len(lines)} line{'s' if len(lines) != 1 else ''} executed successfully",  # noqa: E501
                },
            )

            logger.info("Script execution completed successfully")
            return result.to_dict()

        except PowerShellMCPError:
            raise
        except Exception as e:
            logger.error(f"Unexpected error executing PowerShell script: {e}", exc_info=True)
            raise ScriptExecutionError(f"Script execution failed: {str(e)}")

    async def _ensure_terminal_available(self) -> None:
        """Ensure terminal is running and available with retry logic"""
        if self.terminal_controller.is_terminal_running() and self.terminal_controller.find_terminal_window():
            if self.terminal_controller.focus_terminal():
                return

        logger.info("Terminal not available, attempting to launch...")
        if not self.terminal_controller.launch_terminal():
            raise TerminalNotAvailableError("Failed to launch Windows Terminal")

        # Retry logic for window availability
        for attempt in range(DEFAULT_RETRY_ATTEMPTS):
            if self.terminal_controller.find_terminal_window():
                if self.terminal_controller.focus_terminal():
                    logger.info(f"Terminal ready after {attempt + 1} attempt(s)")
                    return

            logger.debug(f"Terminal not ready, attempt {attempt + 1}/{DEFAULT_RETRY_ATTEMPTS}")
            await asyncio.sleep(DEFAULT_RETRY_DELAY)

        raise TerminalNotAvailableError("Terminal window not available after multiple attempts")

    async def get_clipboard(self) -> Dict[str, Any]:
        """Get current clipboard content with enhanced error handling"""
        try:
            clipboard_content = pyperclip.paste()
            content_length = len(clipboard_content)

            logger.debug(f"Retrieved clipboard content: {content_length} characters")

            result = ToolResult(
                success=True,
                data={
                    "content": clipboard_content,
                    "length": content_length,
                    "is_empty": content_length == 0,
                    "message": f"Clipboard content retrieved successfully ({content_length} characters)",
                },
            )

            return result.to_dict()

        except Exception as e:
            logger.error(f"Error getting clipboard content: {e}", exc_info=True)
            return ToolResult(success=False, error=f"Failed to get clipboard content: {str(e)}").to_dict()

    async def capture_pwsh_response(self, save_path: Optional[str] = None, exclude_titlebar: bool = True) -> Dict[str, Any]:
        """Capture terminal output with improved path handling and error reporting"""
        try:
            logger.info(f"Capturing terminal output (exclude_titlebar: {exclude_titlebar})")
            screenshot = self.terminal_controller.capture_terminal_output(exclude_titlebar)

            if screenshot is None:
                raise PowerShellMCPError("Failed to capture terminal window - window may not be visible or accessible")

            # Determine and validate save path
            final_path = self._get_screenshot_path(save_path)

            # Ensure directory exists
            os.makedirs(os.path.dirname(final_path), exist_ok=True)

            # Save screenshot
            screenshot.save(final_path)
            file_size = os.path.getsize(final_path)

            logger.info(f"Screenshot saved: {final_path} ({file_size} bytes)")

            result = ToolResult(
                success=True,
                data={
                    "saved_to": final_path,
                    "size": {"width": screenshot.size[0], "height": screenshot.size[1]},
                    "file_size_bytes": file_size,
                    "exclude_titlebar": exclude_titlebar,
                    "custom_path": save_path is not None,
                    "message": f"Screenshot saved to {'specified path' if save_path else 'temp directory'}",
                },
            )

            return result.to_dict()

        except PowerShellMCPError:
            raise
        except Exception as e:
            logger.error(f"Error capturing terminal response: {e}", exc_info=True)
            return ToolResult(success=False, error=f"Screenshot capture failed: {str(e)}").to_dict()

    def _get_screenshot_path(self, save_path: Optional[str]) -> str:
        """Generate appropriate screenshot save path"""
        if save_path:
            # Validate custom path
            if not save_path.lower().endswith(".png"):
                save_path += ".png"
            return os.path.abspath(save_path)
        else:
            # Generate temp path
            temp_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "..", "temp")
            timestamp = int(time.time())
            return os.path.join(temp_dir, f"terminal_capture_{timestamp}.png")

    async def run_server(self):
        """Run the MCP server with enhanced logging and error handling"""
        logger.info(f"Starting PowerShell MCP Server v{SERVER_VERSION}...")
        logger.info(f"Available tools: {', '.join(self.tools.keys())}")
        logger.info(f"Default timeout: {self.default_timeout}s")

        request_count = 0

        try:
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

                    request_count += 1
                    logger.debug(f"Processing request #{request_count}")

                    request = json.loads(line)
                    response = await self.handle_request(request)

                    # Write response to stdout
                    print(json.dumps(response))
                    sys.stdout.flush()

                except json.JSONDecodeError as e:
                    logger.error(f"Invalid JSON received in request #{request_count}: {e}")
                    # Send error response for malformed JSON
                    error_response = self._create_error_response(1, -32700, "Parse error")
                    print(json.dumps(error_response))
                    sys.stdout.flush()
                    continue
                except Exception as e:
                    logger.error(f"Error processing request #{request_count}: {e}", exc_info=True)
                    continue

        except KeyboardInterrupt:
            logger.info("Server interrupted by user")
        except Exception as e:
            logger.error(f"Fatal server error: {e}", exc_info=True)
        finally:
            logger.info(f"PowerShell MCP Server stopped after processing {request_count} requests")
