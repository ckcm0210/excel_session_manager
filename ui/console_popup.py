"""
Console Popup component for Excel Session Manager

This module contains the ConsolePopup class that provides a dark-themed
console window for displaying progress information and logs.
"""

import tkinter as tk
import datetime


class ConsolePopup(tk.Toplevel):
    """
    Console popup window for displaying progress and log information.
    
    Features a dark theme with scrollable text area for showing
    real-time progress updates during long-running operations.
    """
    
    def __init__(self, parent, title="Console Output"):
        """
        Initialize the console popup window.
        
        Args:
            parent: Parent window
            title: Window title
        """
        super().__init__(parent)
        self.title(title)
        self.geometry("1400x1200")
        self.resizable(True, True)
        
        # Setup the UI
        self._setup_ui()
        
        # Window close handling
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.closed = False

    def _setup_ui(self):
        """Setup the console UI components."""
        # Main frame
        frm = tk.Frame(self)
        frm.pack(fill="both", expand=True, padx=8, pady=8)
        
        # Text widget with dark theme
        self.text = tk.Text(
            frm, 
            font=("Consolas", 11), 
            state="disabled", 
            wrap="word", 
            bg="#111", 
            fg="#f9f9f9"
        )
        self.text.pack(side="left", fill="both", expand=True)
        
        # Scrollbar
        yscrollbar = tk.Scrollbar(frm, command=self.text.yview)
        yscrollbar.pack(side="right", fill="y")
        self.text["yscrollcommand"] = yscrollbar.set
        
        # Configure text widget
        self.text.config(undo=True, autoseparators=True, maxundo=-1)

    def print(self, msg):
        """
        Print a message to the console with timestamp.
        
        Args:
            msg: Message to display
        """
        if self.closed:
            return
            
        # Add timestamp
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Update text widget
        self.text.config(state="normal")
        self.text.insert("end", f"{ts} | {msg}\n")
        self.text.see("end")
        self.text.config(state="disabled")
        
        # Update display
        self.update_idletasks()

    def on_close(self):
        """Handle window close event."""
        self.closed = True
        self.destroy()