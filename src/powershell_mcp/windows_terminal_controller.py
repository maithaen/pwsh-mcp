import subprocess
from PIL import Image, ImageGrab
import psutil
import pygetwindow as gw
from typing import Dict, Optional
import time
import pyautogui
import pyperclip
import logging

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class WindowsTerminalController:
    """Controller for Windows Terminal PowerShell operations"""

    def __init__(self, timeout: int = 30):
        self.timeout = timeout
        self.terminal_process = None
        self.terminal_window_titles = ["Windows PowerShell", "PowerShell", "Windows Terminal", "Command Prompt", "cmd"]

    def is_terminal_running(self) -> bool:
        """Check if Windows Terminal with PowerShell is running"""
        for proc in psutil.process_iter(["pid", "name", "cmdline"]):
            try:
                if proc.info["name"] and "WindowsTerminal" in proc.info["name"]:
                    return True
                if proc.info["cmdline"]:
                    cmdline = " ".join(proc.info["cmdline"])
                    if any(cmd in cmdline for cmd in ["wt.exe", "pwsh.exe", "powershell.exe"]):
                        return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False

    def launch_terminal(self) -> bool:
        """Launch Windows Terminal with PowerShell and wait for window to be ready"""
        try:
            commands = [["wt.exe", "-p", "PowerShell"], ["wt.exe", "pwsh.exe"], ["wt.exe"], ["pwsh.exe"], ["powershell.exe"]]

            # Try each command until one succeeds
            for cmd in commands:
                try:
                    cmd_str = " ".join(cmd)
                    logger.info(f"Attempting to launch with command: {cmd_str}")

                    # Fix: Use shell=False when passing a list of arguments
                    # For single-command executables like pwsh.exe or powershell.exe
                    if len(cmd) == 1:
                        self.terminal_process = subprocess.Popen(
                            cmd[0],
                            shell=True,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            creationflags=subprocess.CREATE_NEW_CONSOLE,
                        )
                    else:
                        # For wt.exe with parameters
                        self.terminal_process = subprocess.Popen(
                            cmd,
                            shell=False,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.PIPE,
                            creationflags=subprocess.CREATE_NEW_CONSOLE,
                        )

                    # Check process state immediately
                    if self.terminal_process.poll() is not None:
                        # Process terminated immediately
                        stdout, stderr = self.terminal_process.communicate(timeout=1)
                        exit_code = self.terminal_process.returncode
                        logger.warning(f"Command {cmd_str} terminated immediately with exit code {exit_code}")
                        logger.warning(f"Stdout: {stdout.decode('utf-8', errors='ignore')}")
                        logger.warning(f"Stderr: {stderr.decode('utf-8', errors='ignore')}")
                        continue

                    # Wait for process to start and window to appear
                    max_wait = 15  # seconds
                    poll_interval = 0.5
                    waited = 0

                    logger.info(f"Waiting for terminal window to appear (PID: {self.terminal_process.pid})")
                    while waited < max_wait:
                        # Check if process is still running
                        if self.terminal_process.poll() is not None:
                            logger.warning(f"Terminal process exited prematurely with code {self.terminal_process.returncode}")
                            break

                        # Check if terminal window is found
                        if self.is_terminal_running() and self.find_terminal_window():
                            logger.info(f"Terminal launched and window found with command: {cmd_str}")
                            # Try to focus the window as well
                            self.focus_terminal()
                            return True

                        time.sleep(poll_interval)
                        waited += poll_interval
                        logger.debug(f"Waited {waited}s for terminal window to appear")

                    logger.warning(f"Terminal process started but window not found after {max_wait}s for command: {cmd_str}")

                    # Try to terminate the process if it's still running
                    if self.terminal_process.poll() is None:
                        logger.info("Terminating failed terminal process")
                        try:
                            self.terminal_process.terminate()
                            self.terminal_process.wait(timeout=3)
                        except (subprocess.TimeoutExpired, Exception) as e:
                            logger.warning(f"Failed to terminate process: {e}")
                            try:
                                self.terminal_process.kill()
                            except Exception as e:
                                logger.warning(f"Failed to kill process: {e}")

                except FileNotFoundError:
                    logger.warning(f"Command not found: {' '.join(cmd)}")
                    continue
                except Exception as e:
                    logger.warning(f"Failed to launch with {' '.join(cmd)}: {e}")
                    continue

            # All attempts failed
            logger.error("All terminal launch attempts failed")
            return False

        except Exception as e:
            logger.error(f"Error launching terminal: {e}")
            return False

    def find_terminal_window(self) -> Optional[Dict[str, int]]:
        """Find Windows Terminal window coordinates"""
        try:
            for title in self.terminal_window_titles:
                try:
                    windows = gw.getWindowsWithTitle(title)
                    if windows:
                        window = windows[0]
                        if window.isMinimized:
                            window.restore()
                            time.sleep(0.5)

                        return {"left": window.left, "top": window.top, "width": window.width, "height": window.height}
                except Exception as e:
                    logger.debug(f"Error checking window title '{title}': {e}")
                    continue

            logger.warning("No terminal windows found")
            return None

        except Exception as e:
            logger.error(f"Error finding terminal window: {e}")
            return None

    def focus_terminal(self) -> bool:
        """Focus on the terminal window"""
        try:
            windows = []

            # Collect all potential terminal windows
            for title in self.terminal_window_titles:
                try:
                    wins = gw.getWindowsWithTitle(title)
                    if wins:
                        windows.extend(wins)
                        logger.debug(f"Found window with title: {title}")
                except Exception as e:
                    logger.debug(f"Error searching for window with title '{title}': {e}")

            if not windows:
                logger.error("No terminal windows found")
                return False

            # Try to activate the first found window
            window = windows[0]
            logger.info(f"Focusing window: {window.title}")

            # Multiple activation attempts
            try:
                if window.isMinimized:
                    window.restore()
                    time.sleep(0.5)

                window.activate()
                time.sleep(0.5)

            except Exception as e:
                logger.warning(f"Standard activate failed, trying alternative: {e}")
                try:
                    window.minimize()
                    time.sleep(0.2)
                    window.restore()
                    time.sleep(0.5)
                except Exception as e2:
                    logger.warning(f"Alternative activation also failed: {e2}")

            # Verify focus
            try:
                active_window = gw.getActiveWindow()
                if active_window and window.title in active_window.title:
                    logger.info("Successfully focused terminal window")
                    return True
                else:
                    logger.warning(f"Focus verification failed. Active: {getattr(active_window, 'title', 'None')}")
                    return True  # Still return True as command might work
            except Exception:
                logger.warning("Could not verify focus, but continuing...")
                return True

        except Exception as e:
            logger.error(f"Error focusing terminal: {e}")
            return False

    def capture_terminal_output(self, exclude_titlebar: bool = True) -> Optional[Image.Image]:
        """Capture terminal window output"""
        try:
            window_coords = self.find_terminal_window()
            if not window_coords:
                logger.error("Could not find terminal window for capture")
                return None

            left = window_coords["left"]
            top = window_coords["top"]
            width = window_coords["width"]
            height = window_coords["height"]

            # Optionally exclude title bar
            if exclude_titlebar:
                title_bar_height = 35
                top += title_bar_height
                height -= title_bar_height

            # Ensure coordinates are valid
            if width <= 0 or height <= 0:
                logger.error(f"Invalid window dimensions: {width}x{height}")
                return None

            # Capture the specified region
            screenshot = ImageGrab.grab(bbox=(left, top, left + width, top + height))
            logger.info(f"Successfully captured terminal screenshot: {screenshot.size}")
            return screenshot

        except Exception as e:
            logger.error(f"Error capturing terminal output: {e}")
            return None

    def type_command(self, command: str, execute: bool = True) -> bool:
        """Type command into terminal with optional execution"""
        try:
            logger.info(f"Typing command: {command[:50]}{'...' if len(command) > 50 else ''}")

            if not self.focus_terminal():
                logger.error("Failed to focus terminal window")
                return False

            # Clear any existing input
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.2)

            # Type the command with slight delay between keystrokes
            pyautogui.typewrite(command, interval=0.03)

            if execute:
                logger.info("Executing command...")
                pyautogui.press("enter")
            else:
                logger.info("Command typed but not executed")

            return True

        except Exception as e:
            logger.error(f"Error typing command: {e}")
            return False

    def paste_content(self, content: str, execute: bool = True) -> bool:
        """Paste content to terminal using clipboard"""
        try:
            logger.info(f"Pasting content of length: {len(content)}")

            if not self.focus_terminal():
                logger.error("Failed to focus terminal window")
                return False

            # Clear any existing input
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.2)

            # Set clipboard content and paste
            pyperclip.copy(content)
            time.sleep(0.1)
            pyautogui.hotkey("ctrl", "v")

            if execute:
                logger.info("Executing pasted content...")
                pyautogui.press("enter")
            else:
                logger.info("Content pasted but not executed")

            return True

        except Exception as e:
            logger.error(f"Error pasting content: {e}")
            return False
