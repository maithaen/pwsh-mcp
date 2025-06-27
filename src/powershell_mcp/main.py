#!/usr/bin/env python3
"""
PowerShell MCP Server
A Model Context Protocol server for executing PowerShell commands using pyautogui
"""

import asyncio
import sys
import logging
import pyautogui

from powershell_mcp.powershell_server import PowerShellMCPServer

# Configure pyautogui
pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.1

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def main():
    """Main entry point"""
    server = PowerShellMCPServer()

    try:
        asyncio.run(server.run_server())
    except KeyboardInterrupt:
        logger.info("Server stopped by user (Ctrl+C)")
    except Exception as e:
        logger.error(f"Server error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
