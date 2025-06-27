# PowerShell MCP Server

A Model Context Protocol (MCP) server for automating Windows PowerShell tasks using Python. This server enables programmatic execution of PowerShell scripts, clipboard operations, and terminal output capture via a JSON-RPC interface.

## Features

- **Execute PowerShell Scripts:** Paste and run PowerShell scripts in Windows Terminal using the clipboard.
- **Clipboard Access:** Retrieve the current clipboard content.
- **Terminal Screenshot:** Capture the output of the PowerShell terminal as an image.

## Requirements

- Windows OS ( for windows only)
- Python 3.8+
- [pyautogui](https://pypi.org/project/pyautogui/)
- [pygetwindow](https://pypi.org/project/PyGetWindow/)
- [pyperclip](https://pypi.org/project/pyperclip/)
- [psutil](https://pypi.org/project/psutil/)
- [Pillow](https://pypi.org/project/Pillow/)

Install dependencies:
```pwsh
pip install uv && uv sync && uv pip install -e .
```

## Usage

Start the server:
```pwsh
uv run powershell-mcp
```

The server communicates via JSON-RPC over stdin/stdout.

### Available Tools

- `execute_pwsh_script`: Execute a PowerShell script by pasting it into the terminal.
- `get_clipboard`: Get the current clipboard content.
- `capture_pwsh_response`: Capture a screenshot of the PowerShell terminal output.

### Example Request

```json
{
  "jsonrpc": "2.0",
  "id": 1,
  "method": "tools/call",
  "params": {
    "name": "execute_pwsh_script",
    "arguments": {
      "script": "Get-Process",
      "timeout": 30
    }
  }
}
```

## Basic configuration on claude desktop

1. Open claude desktop
2. Go to file -> settings -> developer -> edit config

```json
{
  "globalShortcut": "",
  "mcpServers": {
    "powershell-mcp": {
      "command": "uv",
      "args": [
        "--directory",
        "D:\\MCP antropic\\powershell-mcp", # change to your path
        "run",
        "powershell-mcp"
      ]
    }
  }
}
```

## Project Structure

```
pyproject.toml
README.md
uv.lock
config/
src/
  powershell_mcp/
    main.py
temp/
tests/
  __init__.py
```

## License

MIT License
