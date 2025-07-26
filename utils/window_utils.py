"""
Window utility functions for Excel Session Manager

This module contains utility functions for window operations,
positioning, and display management.
"""

import win32gui
import win32con


def bring_window_to_front(hwnd):
    """
    Bring a window to the front.
    
    Args:
        hwnd: Window handle
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        return True
    except Exception as e:
        print(f"Error bringing window to front: {e}")
        return False


def minimize_window(hwnd):
    """
    Minimize a window.
    
    Args:
        hwnd: Window handle
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        return True
    except Exception as e:
        print(f"Error minimizing window: {e}")
        return False


def restore_window(hwnd):
    """
    Restore a minimized window.
    
    Args:
        hwnd: Window handle
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        return True
    except Exception as e:
        print(f"Error restoring window: {e}")
        return False


def get_window_position(hwnd):
    """
    Get window position and size.
    
    Args:
        hwnd: Window handle
        
    Returns:
        tuple: (left, top, right, bottom) or None if failed
    """
    try:
        return win32gui.GetWindowRect(hwnd)
    except Exception as e:
        print(f"Error getting window position: {e}")
        return None


def set_window_position(hwnd, x, y, width, height):
    """
    Set window position and size.
    
    Args:
        hwnd: Window handle
        x: Left position
        y: Top position
        width: Window width
        height: Window height
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        win32gui.SetWindowPos(
            hwnd, 
            win32con.HWND_TOP, 
            x, y, width, height, 
            win32con.SWP_SHOWWINDOW
        )
        return True
    except Exception as e:
        print(f"Error setting window position: {e}")
        return False


def center_window_on_screen(hwnd):
    """
    Center a window on the screen.
    
    Args:
        hwnd: Window handle
        
    Returns:
        bool: True if successful, False otherwise
    """
    try:
        # Get window size
        rect = win32gui.GetWindowRect(hwnd)
        if not rect:
            return False
            
        window_width = rect[2] - rect[0]
        window_height = rect[3] - rect[1]
        
        # Get screen size
        screen_width = win32gui.GetSystemMetrics(win32con.SM_CXSCREEN)
        screen_height = win32gui.GetSystemMetrics(win32con.SM_CYSCREEN)
        
        # Calculate center position
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        return set_window_position(hwnd, x, y, window_width, window_height)
        
    except Exception as e:
        print(f"Error centering window: {e}")
        return False


def is_window_minimized(hwnd):
    """
    Check if a window is minimized.
    
    Args:
        hwnd: Window handle
        
    Returns:
        bool: True if minimized, False otherwise
    """
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        return placement[1] == win32con.SW_SHOWMINIMIZED
    except Exception:
        return False


def is_window_maximized(hwnd):
    """
    Check if a window is maximized.
    
    Args:
        hwnd: Window handle
        
    Returns:
        bool: True if maximized, False otherwise
    """
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        return placement[1] == win32con.SW_SHOWMAXIMIZED
    except Exception:
        return False