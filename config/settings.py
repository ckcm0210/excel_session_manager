"""
Settings loader for Excel Session Manager

This module handles loading and accessing configuration settings from YAML files.
It provides a centralized way to access all application settings with fallback
to default values when configuration is missing or invalid.
"""

import os
import yaml
from typing import Any, Dict, Optional
from .constants import (
    APP_NAME, MONO_FONTS, DEFAULT_CHECK_DAYS, DEFAULT_LOG_DIR, 
    DEFAULT_SESSION_DIR, DEFAULT_MAX_RETRIES, DEFAULT_RETRY_DELAY
)

class Settings:
    """
    Configuration settings manager for Excel Session Manager.
    
    This class loads settings from YAML configuration files and provides
    easy access to configuration values throughout the application.
    Includes fallback to default values when settings are missing.
    """
    
    def __init__(self, config_file: Optional[str] = None):
        """
        Initialize settings loader.
        
        Args:
            config_file: Path to YAML configuration file. If None, uses default location.
        """
        self._config = {}
        self._config_file = config_file or self._get_default_config_path()
        self._load_config()
    
    def _get_default_config_path(self) -> str:
        """Get the default path for configuration file."""
        current_dir = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(current_dir, "settings.yaml")
    
    def _load_config(self) -> None:
        """Load configuration from YAML file with error handling."""
        try:
            if os.path.exists(self._config_file):
                with open(self._config_file, 'r', encoding='utf-8') as file:
                    self._config = yaml.safe_load(file) or {}
            else:
                print(f"Warning: Configuration file not found at {self._config_file}")
                print("Using default settings.")
                self._config = {}
        except yaml.YAMLError as e:
            print(f"Error parsing YAML configuration: {e}")
            print("Using default settings.")
            self._config = {}
        except Exception as e:
            print(f"Error loading configuration: {e}")
            print("Using default settings.")
            self._config = {}
    
    def get(self, key_path: str, default: Any = None) -> Any:
        """
        Get configuration value using dot notation.
        
        Args:
            key_path: Dot-separated path to configuration value (e.g., "ui.window.default_size")
            default: Default value to return if key is not found
            
        Returns:
            Configuration value or default if not found
            
        Example:
            settings.get("ui.fonts.default_size", 12)
        """
        keys = key_path.split('.')
        value = self._config
        
        try:
            for key in keys:
                value = value[key]
            return value
        except (KeyError, TypeError):
            return default
    
    def set(self, key_path: str, value: Any) -> None:
        """
        Set configuration value using dot notation.
        
        Args:
            key_path: Dot-separated path to configuration value
            value: Value to set
        """
        keys = key_path.split('.')
        config = self._config
        
        # Navigate to the parent of the target key
        for key in keys[:-1]:
            if key not in config:
                config[key] = {}
            config = config[key]
        
        # Set the final value
        config[keys[-1]] = value
    
    def save(self) -> bool:
        """
        Save current configuration to YAML file.
        
        Returns:
            True if save successful, False otherwise
        """
        try:
            with open(self._config_file, 'w', encoding='utf-8') as file:
                yaml.dump(self._config, file, default_flow_style=False, indent=2)
            return True
        except Exception as e:
            print(f"Error saving configuration: {e}")
            return False
    
    # Convenience methods for commonly used settings
    
    @property
    def window_size(self) -> str:
        """Get default window size."""
        return self.get("ui.window.default_size", "1200x800")
    
    @property
    def min_window_size(self) -> str:
        """Get minimum window size."""
        return self.get("ui.window.min_size", "600x500")
    
    @property
    def window_title(self) -> str:
        """Get window title."""
        return self.get("ui.window.title", APP_NAME)
    
    @property
    def default_font_family(self) -> str:
        """Get default font family."""
        available_fonts = self.get("ui.fonts.monospace_options", MONO_FONTS)
        return self.get("ui.fonts.default_family", available_fonts[0] if available_fonts else "Consolas")
    
    @property
    def default_font_size(self) -> int:
        """Get default font size."""
        return self.get("ui.fonts.default_size", 12)
    
    @property
    def monospace_fonts(self) -> list:
        """Get list of available monospace fonts."""
        return self.get("ui.fonts.monospace_options", MONO_FONTS)
    
    @property
    def mini_widget_size(self) -> int:
        """Get mini widget size."""
        return self.get("ui.mini_widget.size", 180)
    
    @property
    def mini_widget_position(self) -> str:
        """Get mini widget default position."""
        return self.get("ui.mini_widget.default_position", "150+150")
    
    @property
    def check_days(self) -> int:
        """Get default check days for external links."""
        return self.get("external_links.defaults.check_days", DEFAULT_CHECK_DAYS)
    
    @property
    def log_directory(self) -> str:
        """Get log directory path."""
        return self.get("external_links.logging.log_directory", DEFAULT_LOG_DIR)
    
    @property
    def session_save_directory(self) -> str:
        """Get session save directory."""
        return self.get("session.directories.default_save", DEFAULT_SESSION_DIR)
    
    @property
    def session_load_directory(self) -> str:
        """Get session load directory."""
        return self.get("session.directories.default_load", DEFAULT_SESSION_DIR)
    
    @property
    def session_load_directory(self) -> str:
        """Get session load directory."""
        return self.get("session.directories.default_load", DEFAULT_SESSION_DIR)
    
    @property
    def include_timestamp_in_session(self) -> bool:
        """Whether to include timestamp in session filenames."""
        return self.get("session.file_format.include_timestamp", True)
    
    @property
    def show_console_by_default(self) -> bool:
        """Whether to show progress console by default."""
        return self.get("progress.behavior.show_console_by_default", True)
    
    @property
    def console_size(self) -> str:
        """Get console window size."""
        return self.get("progress.console.default_size", "1400x1200")
    
    @property
    def max_save_retries(self) -> int:
        """Get maximum save retry attempts."""
        return self.get("excel.operations.max_save_retries", DEFAULT_MAX_RETRIES)
    
    @property
    def retry_delay(self) -> float:
        """Get retry delay in seconds."""
        return self.get("excel.operations.retry_delay", DEFAULT_RETRY_DELAY)


# Global settings instance
# This can be imported and used throughout the application
settings = Settings()


def reload_settings(config_file: Optional[str] = None) -> None:
    """
    Reload settings from configuration file.
    
    Args:
        config_file: Optional path to different configuration file
    """
    global settings
    settings = Settings(config_file)


def get_setting(key_path: str, default: Any = None) -> Any:
    """
    Convenience function to get setting value.
    
    Args:
        key_path: Dot-separated path to setting
        default: Default value if setting not found
        
    Returns:
        Setting value or default
    """
    return settings.get(key_path, default)


def set_setting(key_path: str, value: Any) -> None:
    """
    Convenience function to set setting value.
    
    Args:
        key_path: Dot-separated path to setting
        value: Value to set
    """
    settings.set(key_path, value)


def save_settings() -> bool:
    """
    Convenience function to save current settings.
    
    Returns:
        True if save successful, False otherwise
    """
    return settings.save()