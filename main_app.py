"""
Main Application Entry Point for Excel Session Manager

This module contains the simplified main application class that uses
the new modular MainWindow component.
"""

import tkinter as tk
from ui.main_window import MainWindow


class ExcelSessionManagerApp:
    """
    Main application class for Excel Session Manager.
    
    This class serves as the entry point and coordinates the main window.
    """
    
    def __init__(self, root):
        """
        Initialize the application.
        
        Args:
            root: Tkinter root window
        """
        self.root = root
        self.main_window = MainWindow(root)
    
    def run(self):
        """Run the application main loop."""
        self.root.mainloop()


def main():
    """Main entry point for the application."""
    root = tk.Tk()
    app = ExcelSessionManagerApp(root)
    app.run()


if __name__ == "__main__":
    main()