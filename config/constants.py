"""
Constants for Excel Session Manager

This module contains all constant values used throughout the application.
These values should not be changed during runtime and represent fixed
application behaviors and limits.
"""

# =============================================================================
# APPLICATION METADATA
# =============================================================================

# Application information
APP_NAME = "Excel Session Manager"
APP_VERSION = "2.8.0"
APP_AUTHOR = "Excel Session Manager Team"
APP_DESCRIPTION = "Manage Excel workbook sessions and external links"

# =============================================================================
# UI CONSTANTS
# =============================================================================

# Available monospace fonts for consistent text alignment
MONO_FONTS = [
    "Consolas",          # Windows default monospace font
    "Courier New",       # Cross-platform monospace font
    "Fira Code",         # Modern programming font with ligatures
    "JetBrains Mono",    # Popular developer font
    "DejaVu Sans Mono",  # Good Unicode support
    "Source Code Pro",   # Adobe's open source font
    "Monaco"             # macOS default monospace font
]

# =============================================================================
# UI LAYOUT CONSTANTS
# =============================================================================

# Default window settings
DEFAULT_WINDOW_SIZE = "1200x850"
DEFAULT_WINDOW_TITLE = "Excel Session Manager"
MIN_WINDOW_SIZE = "600x650"
NORMAL_GEOMETRY = "1200x770"

# Mini widget settings
MINI_WIDGET_SIZE = 180
MINI_WIDGET_POSITION = "150+150"
MINI_WIDGET_ICON = "maximize_full_screen.png"

# Default directories
DEFAULT_LOG_DIRECTORY = r"D:\Pzone\Log"
DEFAULT_LOG_DIR = r"D:\Pzone\Log"  # Alias for compatibility
DEFAULT_SESSION_DIRECTORY = r"D:\Pzone\Sessions"
DEFAULT_SESSION_DIR = r"D:\Pzone\Sessions"  # Alias for compatibility

# Button properties
BUTTON_WIDTH = 20
BUTTON_HEIGHT = 2
BUTTON_WRAP_LENGTH = 140

# TreeView settings
TREEVIEW_COLUMN_WIDTH = 320
TREEVIEW_PADDING = 8

# Font settings
DEFAULT_FONT_SIZE = 12
HEADER_FONT_SIZE = 12
CONSOLE_FONT_SIZE = 11

# Timing settings
DEFAULT_CHECK_DAYS = 14
MAX_SAVE_RETRIES = 3
DEFAULT_MAX_RETRIES = 3  # Alias for compatibility
RETRY_DELAY_SECONDS = 0.1
DEFAULT_RETRY_DELAY = 0.1  # Alias for compatibility