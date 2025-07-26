"""
Mini Widget for Excel Session Manager

This module contains the mini widget functionality that provides
a floating, compact interface for the application.
"""

import tkinter as tk
from PIL import Image, ImageTk
import os


class MiniWidget:
    """
    Mini widget for floating compact interface.
    
    Provides a small floating window with essential controls
    when the main window is minimized to mini mode.
    """
    
    def __init__(self, parent_window, geometry, icon_path):
        """
        Initialize the mini widget.
        
        Args:
            parent_window: Reference to the main window
            geometry: Geometry string for mini widget size and position
            icon_path: Path to the icon file
        """
        self.parent_window = parent_window
        self.geometry = geometry
        self.icon_path = icon_path
        self.mini_window = None
        self.floating_icon_btn = None
    
    def create_mini_window(self):
        """Create the mini widget window."""
        if self.mini_window:
            return
            
        self.mini_window = tk.Toplevel(self.parent_window.root)
        self.mini_window.title("Excel Session Manager - Mini")
        self.mini_window.geometry(self.geometry)
        self.mini_window.resizable(False, False)
        self.mini_window.attributes("-topmost", True)
        
        # Try to load icon, fallback to text
        try:
            if os.path.exists(self.icon_path):
                img = Image.open(self.icon_path)
                img = img.resize((64, 64), Image.Resampling.LANCZOS)
                photo = ImageTk.PhotoImage(img)
                
                self.floating_icon_btn = tk.Button(
                    self.mini_window,
                    image=photo,
                    command=self.restore_main_window,
                    relief="flat",
                    bd=0
                )
                self.floating_icon_btn.image = photo  # Keep a reference
            else:
                raise FileNotFoundError("Icon file not found")
                
        except Exception:
            # Fallback to text button
            self.floating_icon_btn = tk.Button(
                self.mini_window,
                text="📖",
                command=self.restore_main_window,
                font=("Arial", 48),
                relief="flat",
                bd=0
            )
        
        self.floating_icon_btn.pack(expand=True, fill="both")
        
        # Bind close event
        self.mini_window.protocol("WM_DELETE_WINDOW", self.restore_main_window)
    
    def destroy_mini_window(self):
        """Destroy the mini widget window."""
        if self.mini_window:
            self.mini_window.destroy()
            self.mini_window = None
            self.floating_icon_btn = None
    
    def restore_main_window(self):
        """Restore the main window from mini mode."""
        self.parent_window.toggle_mini_mode()
    
    def is_active(self):
        """Check if mini widget is currently active."""
        return self.mini_window is not None