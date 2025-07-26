"""
Error Handler for Excel Session Manager

This module provides centralized error handling, logging, and user notification
functionality for the entire application.
"""

import logging
import traceback
from datetime import datetime
from tkinter import messagebox
from typing import Optional, Callable, Any
from enum import Enum


class ErrorSeverity(Enum):
    """Error severity levels."""
    INFO = "INFO"
    WARNING = "WARNING"
    ERROR = "ERROR"
    CRITICAL = "CRITICAL"


class ErrorCategory(Enum):
    """Error categories for better classification."""
    EXCEL_COM = "EXCEL_COM"
    FILE_OPERATION = "FILE_OPERATION"
    PERMISSION = "PERMISSION"
    VALIDATION = "VALIDATION"
    NETWORK = "NETWORK"
    UI = "UI"
    CONFIGURATION = "CONFIGURATION"
    UNKNOWN = "UNKNOWN"


class ErrorHandler:
    """
    Centralized error handling and logging system.
    
    Provides consistent error handling, logging, and user notification
    across the entire application.
    """
    
    def __init__(self, log_file: Optional[str] = None):
        """
        Initialize the error handler.
        
        Args:
            log_file: Optional path to log file
        """
        self.log_file = log_file
        self.logger = self._setup_logger()
        self.error_callbacks = []
        
    def _setup_logger(self) -> logging.Logger:
        """Setup the logging system."""
        logger = logging.getLogger("ExcelSessionManager")
        logger.setLevel(logging.DEBUG)
        
        # Clear existing handlers
        logger.handlers.clear()
        
        # Console handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_format = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
        )
        console_handler.setFormatter(console_format)
        logger.addHandler(console_handler)
        
        # File handler (if log file specified)
        if self.log_file:
            try:
                file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
                file_handler.setLevel(logging.DEBUG)
                file_format = logging.Formatter(
                    '%(asctime)s - %(name)s - %(levelname)s - %(funcName)s:%(lineno)d - %(message)s'
                )
                file_handler.setFormatter(file_format)
                logger.addHandler(file_handler)
            except Exception as e:
                print(f"Warning: Could not setup file logging: {e}")
        
        return logger
    
    def handle_error(
        self, 
        error: Exception, 
        severity: ErrorSeverity = ErrorSeverity.ERROR,
        category: ErrorCategory = ErrorCategory.UNKNOWN,
        context: str = "",
        show_user: bool = True,
        user_message: Optional[str] = None
    ) -> None:
        """
        Handle an error with appropriate logging and user notification.
        
        Args:
            error: The exception that occurred
            severity: Severity level of the error
            category: Category of the error
            context: Additional context information
            show_user: Whether to show error to user
            user_message: Custom user-friendly message
        """
        # Generate error details
        error_details = self._generate_error_details(error, severity, category, context)
        
        # Log the error
        self._log_error(error_details)
        
        # Notify user if requested
        if show_user:
            self._notify_user(error_details, user_message)
        
        # Call registered callbacks
        self._call_error_callbacks(error_details)
    
    def _generate_error_details(
        self, 
        error: Exception, 
        severity: ErrorSeverity, 
        category: ErrorCategory, 
        context: str
    ) -> dict:
        """Generate detailed error information."""
        return {
            'timestamp': datetime.now().isoformat(),
            'error_type': type(error).__name__,
            'error_message': str(error),
            'severity': severity.value,
            'category': category.value,
            'context': context,
            'traceback': traceback.format_exc(),
            'function': traceback.extract_tb(error.__traceback__)[-1].name if error.__traceback__ else 'unknown'
        }
    
    def _log_error(self, error_details: dict) -> None:
        """Log error details."""
        log_message = (
            f"[{error_details['category']}] {error_details['error_type']}: "
            f"{error_details['error_message']}"
        )
        
        if error_details['context']:
            log_message += f" | Context: {error_details['context']}"
        
        severity = error_details['severity']
        if severity == ErrorSeverity.INFO.value:
            self.logger.info(log_message)
        elif severity == ErrorSeverity.WARNING.value:
            self.logger.warning(log_message)
        elif severity == ErrorSeverity.ERROR.value:
            self.logger.error(log_message)
            self.logger.debug(f"Traceback:\n{error_details['traceback']}")
        elif severity == ErrorSeverity.CRITICAL.value:
            self.logger.critical(log_message)
            self.logger.debug(f"Traceback:\n{error_details['traceback']}")
    
    def _notify_user(self, error_details: dict, user_message: Optional[str]) -> None:
        """Show error notification to user."""
        severity = error_details['severity']
        
        if user_message:
            message = user_message
        else:
            message = self._generate_user_message(error_details)
        
        title = f"{error_details['category']} {severity}"
        
        try:
            if severity in [ErrorSeverity.CRITICAL.value, ErrorSeverity.ERROR.value]:
                messagebox.showerror(title, message)
            elif severity == ErrorSeverity.WARNING.value:
                messagebox.showwarning(title, message)
            else:
                messagebox.showinfo(title, message)
        except Exception:
            # Fallback if messagebox fails
            print(f"{title}: {message}")
    
    def _generate_user_message(self, error_details: dict) -> str:
        """Generate user-friendly error message."""
        category = error_details['category']
        error_type = error_details['error_type']
        error_message = error_details['error_message']
        
        # Category-specific user messages
        if category == ErrorCategory.EXCEL_COM.value:
            return (
                f"Excel operation failed: {error_message}\n\n"
                "This might be due to:\n"
                "• Excel not being installed or accessible\n"
                "• Excel files being locked by another process\n"
                "• Insufficient permissions\n\n"
                "Try closing Excel and running the operation again."
            )
        elif category == ErrorCategory.FILE_OPERATION.value:
            return (
                f"File operation failed: {error_message}\n\n"
                "This might be due to:\n"
                "• File being locked or in use\n"
                "• Insufficient permissions\n"
                "• File path not existing\n\n"
                "Please check the file and try again."
            )
        elif category == ErrorCategory.PERMISSION.value:
            return (
                f"Permission denied: {error_message}\n\n"
                "Try running the application as administrator or "
                "check file/folder permissions."
            )
        elif category == ErrorCategory.VALIDATION.value:
            return (
                f"Validation error: {error_message}\n\n"
                "Please check your input and try again."
            )
        else:
            return f"An error occurred: {error_message}"
    
    def _call_error_callbacks(self, error_details: dict) -> None:
        """Call registered error callbacks."""
        for callback in self.error_callbacks:
            try:
                callback(error_details)
            except Exception as e:
                self.logger.error(f"Error in error callback: {e}")
    
    def register_error_callback(self, callback: Callable[[dict], None]) -> None:
        """Register a callback to be called when errors occur."""
        self.error_callbacks.append(callback)
    
    def unregister_error_callback(self, callback: Callable[[dict], None]) -> None:
        """Unregister an error callback."""
        if callback in self.error_callbacks:
            self.error_callbacks.remove(callback)
    
    def log_info(self, message: str, context: str = "") -> None:
        """Log an info message."""
        full_message = f"{message} | Context: {context}" if context else message
        self.logger.info(full_message)
    
    def log_warning(self, message: str, context: str = "") -> None:
        """Log a warning message."""
        full_message = f"{message} | Context: {context}" if context else message
        self.logger.warning(full_message)
    
    def safe_execute(
        self, 
        func: Callable, 
        *args, 
        context: str = "",
        category: ErrorCategory = ErrorCategory.UNKNOWN,
        show_user: bool = True,
        default_return: Any = None,
        **kwargs
    ) -> Any:
        """
        Safely execute a function with error handling.
        
        Args:
            func: Function to execute
            *args: Function arguments
            context: Context information
            category: Error category
            show_user: Whether to show errors to user
            default_return: Value to return on error
            **kwargs: Function keyword arguments
            
        Returns:
            Function result or default_return on error
        """
        try:
            return func(*args, **kwargs)
        except Exception as e:
            self.handle_error(
                e, 
                severity=ErrorSeverity.ERROR,
                category=category,
                context=context,
                show_user=show_user
            )
            return default_return


# Global error handler instance
_global_error_handler: Optional[ErrorHandler] = None


def get_error_handler() -> ErrorHandler:
    """Get the global error handler instance."""
    global _global_error_handler
    if _global_error_handler is None:
        from config.settings import settings
        log_file = settings.get("debug.logging.debug_log_file", None)
        _global_error_handler = ErrorHandler(log_file)
    return _global_error_handler


def handle_error(
    error: Exception, 
    severity: ErrorSeverity = ErrorSeverity.ERROR,
    category: ErrorCategory = ErrorCategory.UNKNOWN,
    context: str = "",
    show_user: bool = True,
    user_message: Optional[str] = None
) -> None:
    """Convenience function for error handling."""
    get_error_handler().handle_error(
        error, severity, category, context, show_user, user_message
    )


def safe_execute(
    func: Callable, 
    *args, 
    context: str = "",
    category: ErrorCategory = ErrorCategory.UNKNOWN,
    show_user: bool = True,
    default_return: Any = None,
    **kwargs
) -> Any:
    """Convenience function for safe execution."""
    return get_error_handler().safe_execute(
        func, *args, 
        context=context, 
        category=category, 
        show_user=show_user, 
        default_return=default_return, 
        **kwargs
    )