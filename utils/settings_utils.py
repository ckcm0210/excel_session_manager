"""
Settings utility functions for Excel Session Manager

This module provides helper functions for managing application settings
and user preferences.
"""

from config.settings import settings, save_settings


def save_user_preferences(font_family, font_size, show_console, window_geometry=None):
    """
    Save user preferences to settings.
    
    Args:
        font_family: Selected font family
        font_size: Selected font size
        show_console: Whether to show console by default
        window_geometry: Current window geometry (optional)
    """
    # Update settings
    settings.set("ui.fonts.default_family", font_family)
    settings.set("ui.fonts.default_size", font_size)
    settings.set("progress.behavior.show_console_by_default", show_console)
    
    if window_geometry:
        settings.set("ui.window.default_size", window_geometry)
    
    # Save to file
    save_settings()


def load_user_preferences():
    """
    Load user preferences from settings.
    
    Returns:
        dict: Dictionary containing user preferences
    """
    return {
        "font_family": settings.default_font_family,
        "font_size": settings.default_font_size,
        "show_console": settings.show_console_by_default,
        "window_size": settings.window_size,
        "log_directory": settings.log_directory
    }


def reset_to_defaults():
    """Reset all settings to default values."""
    from config.constants import MONO_FONTS, DEFAULT_FONT_SIZE, DEFAULT_WINDOW_SIZE, DEFAULT_LOG_DIRECTORY
    
    settings.set("ui.fonts.default_family", MONO_FONTS[0])
    settings.set("ui.fonts.default_size", DEFAULT_FONT_SIZE)
    settings.set("ui.window.default_size", DEFAULT_WINDOW_SIZE)
    settings.set("external_links.logging.log_directory", DEFAULT_LOG_DIRECTORY)
    settings.set("progress.behavior.show_console_by_default", True)
    
    save_settings()


def get_current_theme():
    """
    Get current UI theme settings.
    
    Returns:
        dict: Theme configuration
    """
    return {
        "console_bg": settings.get("progress.console.background_color", "#111"),
        "console_fg": settings.get("progress.console.text_color", "#f9f9f9"),
        "console_font": settings.get("progress.console.font_family", "Consolas"),
        "console_font_size": settings.get("progress.console.font_size", 11)
    }