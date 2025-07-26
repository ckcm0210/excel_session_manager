"""
Main Window for Excel Session Manager

This module contains the main application window class that handles
the primary user interface and coordinates between different components.
"""

import tkinter as tk
from tkinter import ttk, messagebox, font
import threading
from datetime import datetime

# Import UI components
from ui.components.drag_treeview import DragSelectTreeview
from ui.console_popup import ConsolePopup
from ui.dialogs.link_options import LinkUpdateOptionsDialog
from ui.mini_widget import MiniWidget
from ui.performance_dialog import PerformanceDialog

# Import core components
from core.session_manager import SessionManager
from core.excel_manager import ExcelManager
from core.process_manager import ProcessManager

# Import utility functions
from utils.file_utils import get_file_mtime_str, parse_mtime
from utils.ui_utils import calc_row_height, calc_col2_width

# Import constants and settings
from config.constants import (
    MONO_FONTS, DEFAULT_WINDOW_SIZE, DEFAULT_WINDOW_TITLE, MIN_WINDOW_SIZE,
    NORMAL_GEOMETRY, MINI_WIDGET_SIZE, MINI_WIDGET_POSITION, MINI_WIDGET_ICON,
    DEFAULT_LOG_DIRECTORY, BUTTON_WIDTH, BUTTON_HEIGHT, BUTTON_WRAP_LENGTH,
    TREEVIEW_COLUMN_WIDTH, TREEVIEW_PADDING, DEFAULT_FONT_SIZE, HEADER_FONT_SIZE,
    DEFAULT_CHECK_DAYS
)
from config.settings import settings


class MainWindow:
    """
    Main application window for Excel Session Manager.
    
    This class handles the primary user interface, coordinates between
    different managers, and provides the main application functionality.
    """
    
    def __init__(self, root):
        """
        Initialize the main window.
        
        Args:
            root: Tkinter root window
        """
        self.root = root
        self._setup_window()
        self._initialize_managers()
        self._initialize_variables()
        self._setup_ui()
        self._initial_refresh()
    
    def _setup_window(self):
        """Setup the main window properties."""
        self.root.title(settings.window_title)
        self.root.geometry(settings.window_size)
        min_w, min_h = map(int, settings.min_window_size.split('x'))
        self.root.minsize(min_w, min_h)
        self.root.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())
    
    def _initialize_managers(self):
        """Initialize all manager components."""
        self.session_manager = SessionManager(self.root)
        self.excel_manager = ExcelManager()
        self.process_manager = ProcessManager()
    
    def _initialize_variables(self):
        """Initialize application variables."""
        self.last_log_dir = settings.log_directory
        self.is_mini = False
        self.normal_geometry = settings.get("ui.window.normal_geometry", NORMAL_GEOMETRY)
        self.mini_side = settings.mini_widget_size
        self.mini_geometry = f"{self.mini_side}x{self.mini_side}+{settings.mini_widget_position}"
        self.mini_widget = MiniWidget(self, self.mini_geometry, settings.get("ui.mini_widget.icon_file", MINI_WIDGET_ICON))
    
    def _setup_ui(self):
        """Setup the user interface components."""
        self._create_main_layout()
        self._create_font_controls()
        self._create_button_panel()
        self._create_file_list()
        self._create_status_controls()
    
    def _create_main_layout(self):
        """Create the main layout structure."""
        # Main container
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Top frame for controls
        self.top_frame = tk.Frame(self.main_frame)
        self.top_frame.pack(fill="x", pady=(0, 5))
        
        # Middle frame for file list and buttons
        self.middle_frame = tk.Frame(self.main_frame)
        self.middle_frame.pack(fill="both", expand=True)
        
        # Bottom frame for status
        self.bottom_frame = tk.Frame(self.main_frame)
        self.bottom_frame.pack(fill="x", pady=(5, 0))
    
    def _create_font_controls(self):
        """Create font selection controls."""
        font_frame = tk.Frame(self.top_frame)
        font_frame.pack(side=tk.LEFT)
        
        tk.Label(font_frame, text="Font:", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=(0, 5))
        
        # Font family selection
        self.font_family_var = tk.StringVar(value=settings.default_font_family)
        font_family_combo = ttk.Combobox(
            font_frame, 
            textvariable=self.font_family_var, 
            values=settings.monospace_fonts, 
            state="readonly", 
            width=15
        )
        font_family_combo.pack(side=tk.LEFT, padx=(0, 5))
        font_family_combo.bind("<<ComboboxSelected>>", self._on_font_change)
        
        # Font size selection
        self.font_size_var = tk.IntVar(value=settings.default_font_size)
        font_size_spinbox = tk.Spinbox(
            font_frame, 
            from_=8, 
            to=32, 
            textvariable=self.font_size_var, 
            width=5, 
            command=self._on_font_change
        )
        font_size_spinbox.pack(side=tk.LEFT, padx=(0, 10))
        font_size_spinbox.bind("<KeyRelease>", self._on_font_change)
    
    def _create_button_panel(self):
        """Create the button panel."""
        # Button frame on the right side
        side_btn_frame = tk.Frame(self.middle_frame)
        side_btn_frame.pack(side=tk.RIGHT, fill="y", padx=(10, 0))
        
        # Button properties
        btn_props = {
            'width': BUTTON_WIDTH, 
            'height': BUTTON_HEIGHT, 
            'wraplength': BUTTON_WRAP_LENGTH, 
            'font': ("Arial", 10, "bold")
        }
        
        # Create buttons
        self._create_action_buttons(side_btn_frame, btn_props)
        self._create_utility_buttons(side_btn_frame, btn_props)
    
    def _create_action_buttons(self, parent, btn_props):
        """Create action buttons (Save, Load, etc.)."""
        refresh_btn = tk.Button(parent, text="Refresh List", command=self.show_names, **btn_props)
        refresh_btn.pack(pady=5, anchor='n')
        
        save_btn = tk.Button(parent, text="Save Selected", command=self.save_selected_workbooks, **btn_props)
        save_btn.pack(pady=5, anchor='n')
        
        save_close_btn = tk.Button(parent, text="Save and Close\nSelected", command=self.save_close_selected_workbooks, **btn_props)
        save_close_btn.pack(pady=5, anchor='n')
        
        activate_btn = tk.Button(parent, text="Activate Selected", command=self.activate_selected, **btn_props)
        activate_btn.pack(pady=5, anchor='n')
        
        minimize_btn = tk.Button(parent, text="Minimize All", command=self.minimize_all, **btn_props)
        minimize_btn.pack(pady=5, anchor='n')
        
        save_session_btn = tk.Button(parent, text="Save Session", command=self.save_session, **btn_props)
        save_session_btn.pack(pady=5, anchor='n')
        
        load_session_btn = tk.Button(parent, text="Load Session", command=self.load_session, **btn_props)
        load_session_btn.pack(pady=5, anchor='n')
    
    def _create_utility_buttons(self, parent, btn_props):
        """Create utility buttons (Update Links, Cleanup, etc.)."""
        update_links_btn = tk.Button(parent, text="Update Recent\nExternal Links", command=self.open_link_update_options, **btn_props)
        update_links_btn.pack(pady=5, anchor='n')
        
        cleanup_btn = tk.Button(parent, text="Cleanup Excel\nProcesses", command=self.cleanup_excel_processes, **btn_props)
        cleanup_btn.pack(pady=5, anchor='n')
        
        performance_btn = tk.Button(parent, text="Performance\nMonitor", command=self.show_performance_monitor, **btn_props)
        performance_btn.pack(pady=5, anchor='n')
        
        mini_btn = tk.Button(parent, text="Mini Mode", command=self.toggle_mini_mode, **btn_props)
        mini_btn.pack(pady=5, anchor='n')
    
    def _create_file_list(self):
        """Create the file list TreeView."""
        # TreeView frame
        tree_frame = tk.Frame(self.middle_frame)
        tree_frame.pack(side=tk.LEFT, fill="both", expand=True)
        
        # Create TreeView
        self.tree = DragSelectTreeview(tree_frame, columns=("col1", "col2"), show="headings", selectmode="extended")
        self.tree.heading("col1", text="File Name", command=lambda: self.sort_column("col1"))
        self.tree.heading("col2", text="Last Modified", command=lambda: self.sort_column("col2"))
        
        # Configure columns
        self.tree.column("col1", anchor="w", stretch=True)
        self.tree.column("col2", anchor="e", width=settings.get("ui.treeview.column_width", TREEVIEW_COLUMN_WIDTH), stretch=False)
        
        # Pack TreeView
        padding = settings.get("ui.treeview.padding", TREEVIEW_PADDING)
        self.tree.pack(expand=True, fill='both', padx=padding, pady=padding)
        
        # Bind events
        self.tree.bind("<<TreeviewSelect>>", self.on_selection_change)
        self.tree.bind("<Double-1>", self.on_double_click)
        
        # Initialize sorting
        self.sort_states = {"col1": "none", "col2": "none"}
    
    def _create_status_controls(self):
        """Create status and control elements."""
        # Select All checkbox
        self.select_all_var = tk.BooleanVar()
        select_all_cb = tk.Checkbutton(
            self.bottom_frame, 
            text="Select All", 
            variable=self.select_all_var, 
            command=self.toggle_select_all,
            font=("Arial", 10, "bold")
        )
        select_all_cb.pack(side=tk.LEFT)
        
        # Console progress checkbox
        self.show_console_progress_var = tk.BooleanVar(value=settings.show_console_by_default)
        console_cb = tk.Checkbutton(
            self.bottom_frame, 
            text="Show Console Progress", 
            variable=self.show_console_progress_var,
            font=("Arial", 10)
        )
        console_cb.pack(side=tk.RIGHT)
    
    def _initial_refresh(self):
        """Perform initial refresh of the file list."""
        self.show_names()
    
    def _on_font_change(self, event=None):
        """Handle font change events."""
        try:
            family = self.font_family_var.get()
            size = self.font_size_var.get()
            
            # Update TreeView font
            style = ttk.Style()
            style.configure("Treeview", font=(family, size))
            style.configure("Treeview.Heading", font=(family, HEADER_FONT_SIZE, "bold"))
            
            # Update row height
            row_height = calc_row_height(size)
            style.configure("Treeview", rowheight=row_height)
            
            # Update column width
            col2_width = calc_col2_width(size)
            self.tree.column("col2", width=col2_width)
            
        except Exception as e:
            print(f"Error updating font: {e}")
    
    # Delegate methods to managers
    def get_open_excel_files(self):
        """Get open Excel files using ExcelManager."""
        return self.excel_manager.get_open_workbooks()
    
    def save_selected_workbooks(self):
        """Save selected workbooks using ExcelManager."""
        selected = self.get_selected_workbooks()
        if not selected:
            return
            
        show_console = self.show_console_progress_var.get()
        popup = ConsolePopup(self.root, title="Save Selected Progress") if show_console else None
        
        def print_to_popup(msg):
            if popup:
                self.root.after(0, lambda: popup.print(msg))
                
        def thread_job():
            self.excel_manager.save_workbooks(selected, print_to_popup)
            self.root.after(0, lambda: messagebox.showinfo("Complete", f"{len(selected)} file(s) saved."))
            self.root.after(0, self.show_names)
            
        threading.Thread(target=thread_job, daemon=True).start()
    
    def save_close_selected_workbooks(self):
        """Save and close selected workbooks using ExcelManager."""
        selected = self.get_selected_workbooks()
        if not selected:
            return
            
        show_console = self.show_console_progress_var.get()
        popup = ConsolePopup(self.root, title="Save and Close Selected Progress") if show_console else None
        
        def print_to_popup(msg):
            if popup:
                self.root.after(0, lambda: popup.print(msg))
                
        def thread_job():
            self.excel_manager.save_and_close_workbooks(selected, print_to_popup)
            self.root.after(0, lambda: messagebox.showinfo("Complete", f"{len(selected)} file(s) saved and closed."))
            self.root.after(0, self.show_names)
            
        threading.Thread(target=thread_job, daemon=True).start()
    
    def activate_selected(self):
        """Activate selected workbooks using ExcelManager."""
        selected = self.get_selected_workbooks()
        if not selected:
            return
        self.excel_manager.activate_workbooks(selected)
    
    def minimize_all(self):
        """Minimize all Excel windows using ExcelManager."""
        self.excel_manager.minimize_all_excel()
    
    def save_session(self):
        """Save current session using SessionManager."""
        selected = self.get_selected_workbooks()
        if not selected:
            return
        saved_path = self.session_manager.save_session(selected)
        if saved_path:
            self.load_session_from_path(saved_path)
    
    def load_session(self):
        """Load session using SessionManager."""
        self.session_manager.load_session(
            self.get_open_excel_files, 
            self.show_console_progress_var
        )
        self.root.after(200, self.show_names)
    
    def open_link_update_options(self):
        """Open the link update options dialog."""
        dialog = LinkUpdateOptionsDialog(self.root, self.last_log_dir, getattr(self, 'last_summary_dir', None))
        self.last_log_dir, self.last_summary_dir = dialog.show()
    
    def cleanup_excel_processes(self):
        """Cleanup Excel processes using ProcessManager."""
        show_console = self.show_console_progress_var.get()
        popup = ConsolePopup(self.root, title="Excel Process Cleanup") if show_console else None
        
        def print_to_popup(msg):
            if popup:
                self.root.after(0, lambda: popup.print(msg))
                
        def thread_job():
            print_to_popup("=== Excel Process Health Check ===")
            health_report = self.process_manager.monitor_excel_health(print_to_popup)
            print_to_popup("")
            
            if health_report['zombie_processes'] > 0 or health_report['total_processes'] > 3:
                print_to_popup("=== Starting Cleanup ===")
                self.process_manager.cleanup_zombie_excel_processes(print_to_popup)
            else:
                print_to_popup("No cleanup needed - Excel processes are healthy")
            
            print_to_popup("")
            print_to_popup("=== Final Status ===")
            final_processes = self.process_manager.get_excel_process_info()
            if final_processes:
                print_to_popup(f"Remaining Excel processes: {len(final_processes)}")
                for proc in final_processes:
                    print_to_popup(f"  PID {proc['pid']}: {proc['name']} ({proc['memory_mb']} MB)")
            else:
                print_to_popup("No Excel processes running")
            
            self.root.after(0, lambda: messagebox.showinfo("Complete", "Excel process cleanup completed"))
            
        threading.Thread(target=thread_job, daemon=True).start()
    
    def show_performance_monitor(self):
        """Show the performance monitor dialog."""
        try:
            dialog = PerformanceDialog(self.root)
            dialog.show()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open performance monitor:\n{str(e)}")
    
    # UI helper methods
    def show_names(self):
        """Refresh the file list display."""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        file_list, sheet_list, cell_list, path_list = self.get_open_excel_files()
        workbook_details = self.excel_manager.get_workbook_details(file_list, sheet_list, cell_list, path_list)
        
        for name, path, sheet, cell, mtime_str in workbook_details:
            self.tree.insert("", "end", values=(name, mtime_str), tags=(name, path, sheet, cell))
        
        self.update_select_all_state()
    
    def get_selected_workbooks(self):
        """Get currently selected workbooks."""
        selected_items = self.tree.selection()
        selected_workbooks = []
        
        for item in selected_items:
            tags = self.tree.item(item, "tags")
            if len(tags) >= 4:
                name, path, sheet, cell = tags[:4]
                selected_workbooks.append((name, path, sheet, cell))
        
        return selected_workbooks
    
    def on_selection_change(self, event):
        """Handle selection change events."""
        self.update_select_all_state()
    
    def update_select_all_state(self):
        """Update the select all checkbox state."""
        total_items = len(self.tree.get_children())
        selected_items = len(self.tree.selection())
        
        if total_items == 0:
            self.select_all_var.set(False)
        elif selected_items == total_items:
            self.select_all_var.set(True)
        else:
            self.select_all_var.set(False)
    
    def on_double_click(self, event):
        """Handle double-click events."""
        selected = self.get_selected_workbooks()
        if selected:
            self.excel_manager.activate_workbooks(selected)
    
    def sort_column(self, col):
        """Sort TreeView by column."""
        items = [self.tree.item(iid, "values") for iid in self.tree.get_children()]
        if not hasattr(self, 'original_data') or len(self.original_data) != len(items):
            self.original_data = list(items)
            
        if self.sort_states[col] == "none":
            self.sort_states[col] = "asc"
        elif self.sort_states[col] == "asc":
            self.sort_states[col] = "desc"
        else:
            self.sort_states[col] = "none"
            
        # Reset other columns
        for other_col in self.sort_states:
            if other_col != col:
                self.sort_states[other_col] = "none"
                
        if self.sort_states[col] == "none":
            sorted_items = self.original_data
        else:
            reverse = self.sort_states[col] == "desc"
            if col == "col1":
                key_func = lambda x: (x[0] or "").lower()
            elif col == "col2":
                key_func = lambda x: (parse_mtime(x[1]) or datetime.min)
            else:
                key_func = lambda x: x
            sorted_items = sorted(items, key=key_func, reverse=reverse)
            
        # Clear and repopulate
        for item in self.tree.get_children():
            self.tree.delete(item)
        for values in sorted_items:
            self.tree.insert("", "end", values=values)
            
        self._on_font_change()
    
    def toggle_select_all(self):
        """Toggle select all checkbox."""
        if self.select_all_var.get():
            for item in self.tree.get_children():
                self.tree.selection_add(item)
        else:
            self.tree.selection_remove(self.tree.selection())
    
    def toggle_mini_mode(self):
        """Toggle mini mode."""
        if self.is_mini:
            # Restore from mini mode
            self.mini_widget.destroy_mini_window()
            self.root.deiconify()
            self.root.geometry(self.normal_geometry)
            self.is_mini = False
        else:
            # Enter mini mode
            self.normal_geometry = self.root.geometry()
            self.root.withdraw()
            self.mini_widget.create_mini_window()
            self.is_mini = True
    
    def load_session_from_path(self, path):
        """Load session from specific path."""
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            
            # Display session contents in a simple dialog
            session_info = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0]:
                    session_info.append(f"File: {row[0]}")
                    if len(row) > 1 and row[1]:
                        session_info.append(f"  Sheet: {row[1]}")
                    if len(row) > 2 and row[2]:
                        session_info.append(f"  Cell: {row[2]}")
                    session_info.append("")
            
            info_text = "\n".join(session_info[:20])  # Limit to first 20 lines
            if len(session_info) > 20:
                info_text += "\n... (truncated)"
                
            messagebox.showinfo("Session Saved", f"Session saved successfully!\n\nContents:\n{info_text}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not read session file:\n{str(e)}")