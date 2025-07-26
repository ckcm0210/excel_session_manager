"""
External Links Manager Dialog for Excel Session Manager

This module provides a dialog for managing external links in Excel workbooks,
including searching for specific links and batch updating functionality.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime
from core.external_links_manager import ExternalLinksManager
from utils.external_links_utils import ExcelNavigator, ExternalLinksExporter, DataFormatter


class ExternalLinksDialog:
    """
    Dialog for managing external links in Excel workbooks.
    
    Provides functionality to search for external links by keyword,
    view link details, and perform batch updates.
    """
    
    def __init__(self, parent):
        """
        Initialize the external links dialog.
        
        Args:
            parent: Parent window
        """
        self.parent = parent
        self.dialog = None
        self.links_manager = ExternalLinksManager()
        self.current_search_results = {}
        self.display_mode = 'grouped'  # 'grouped' or 'flat'
        
    def show(self):
        """Show the external links manager dialog."""
        if self.dialog:
            self.dialog.lift()
            return
            
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("External Links Manager")
        self.dialog.geometry("900x700")
        self.dialog.resizable(True, True)
        
        # Make dialog non-modal
        self.dialog.transient(self.parent)
        
        self._setup_ui()
        self._scan_external_links()
        
        # Handle close event
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # Center the dialog
        self._center_dialog()
    
    def _setup_ui(self):
        """Setup the user interface."""
        # Main container
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Search frame
        search_frame = ttk.LabelFrame(main_frame, text="Search External Links", padding="10")
        search_frame.pack(fill="x", pady=(0, 10))
        
        # Search input row 1
        search_input_frame = ttk.Frame(search_frame)
        search_input_frame.pack(fill="x", pady=(0, 5))
        
        ttk.Label(search_input_frame, text="Keyword:").pack(side="left", padx=(0, 5))
        
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_input_frame, textvariable=self.search_var, width=30)
        self.search_entry.pack(side="left", padx=(0, 10))
        self.search_entry.bind("<Return>", lambda e: self._search_links())
        
        # Search field selection
        ttk.Label(search_input_frame, text="Search in:").pack(side="left", padx=(10, 5))
        self.search_field_var = tk.StringVar(value="external_file")
        search_field_combo = ttk.Combobox(search_input_frame, textvariable=self.search_field_var, 
                                        values=["external_file", "formula", "source_workbook", "all"],
                                        state="readonly", width=15)
        search_field_combo.pack(side="left", padx=(0, 10))
        
        # Search buttons row 2
        search_btn_frame = ttk.Frame(search_frame)
        search_btn_frame.pack(fill="x")
        
        search_btn = ttk.Button(search_btn_frame, text="Search", command=self._search_links)
        search_btn.pack(side="left", padx=(0, 5))
        
        clear_btn = ttk.Button(search_btn_frame, text="Clear", command=self._clear_search)
        clear_btn.pack(side="left", padx=(0, 10))
        
        # Display mode toggle
        ttk.Label(search_btn_frame, text="View:").pack(side="left", padx=(10, 5))
        self.display_mode_var = tk.StringVar(value="grouped")
        mode_combo = ttk.Combobox(search_btn_frame, textvariable=self.display_mode_var,
                                values=["grouped", "flat"], state="readonly", width=10)
        mode_combo.pack(side="left")
        mode_combo.bind("<<ComboboxSelected>>", self._on_display_mode_change)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Search Results", padding="10")
        results_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Results tree (using tree view for hierarchical display)
        self.results_tree = ttk.Treeview(results_frame, show="tree headings", height=15)
        
        # Configure columns for tree view
        self.results_tree["columns"] = ("type", "details", "formula")
        self.results_tree.heading("#0", text="External File / Source")
        self.results_tree.heading("type", text="Type")
        self.results_tree.heading("details", text="Location")
        self.results_tree.heading("formula", text="Formula")
        
        self.results_tree.column("#0", width=250)
        self.results_tree.column("type", width=80)
        self.results_tree.column("details", width=150)
        self.results_tree.column("formula", width=350)
        
        # Scrollbars for tree
        tree_scroll_y = ttk.Scrollbar(results_frame, orient="vertical", command=self.results_tree.yview)
        tree_scroll_x = ttk.Scrollbar(results_frame, orient="horizontal", command=self.results_tree.xview)
        self.results_tree.configure(yscrollcommand=tree_scroll_y.set, xscrollcommand=tree_scroll_x.set)
        
        # Pack tree and scrollbars
        self.results_tree.grid(row=0, column=0, sticky="nsew")
        tree_scroll_y.grid(row=0, column=1, sticky="ns")
        tree_scroll_x.grid(row=1, column=0, sticky="ew")
        
        results_frame.grid_rowconfigure(0, weight=1)
        results_frame.grid_columnconfigure(0, weight=1)
        
        # Bind double-click event
        self.results_tree.bind("<Double-1>", self._on_result_double_click)
        
        # Status frame
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill="x", pady=(0, 10))
        
        self.status_label = ttk.Label(status_frame, text="Ready")
        self.status_label.pack(side="left")
        
        # Progress bar
        self.progress = ttk.Progressbar(status_frame, mode='indeterminate')
        self.progress.pack(side="right", padx=(10, 0))
        
        # Control buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill="x")
        
        refresh_btn = ttk.Button(button_frame, text="Rescan", command=self._scan_external_links)
        refresh_btn.pack(side="left", padx=(0, 10))
        
        export_btn = ttk.Button(button_frame, text="Export Results", command=self._export_results)
        export_btn.pack(side="left", padx=(0, 10))
        
        batch_update_btn = ttk.Button(button_frame, text="Batch Update", command=self._show_batch_update)
        batch_update_btn.pack(side="left", padx=(0, 10))
        
        close_btn = ttk.Button(button_frame, text="Close", command=self._on_close)
        close_btn.pack(side="right")
    
    def _scan_external_links(self):
        """Scan all open workbooks for external links."""
        self.status_label.config(text="Scanning external links...")
        self.progress.start()
        
        def scan_thread():
            try:
                external_links, statistics = self.links_manager.scan_open_workbooks()
                self.dialog.after(0, lambda: self._update_status_after_scan(statistics))
            except Exception as e:
                self.dialog.after(0, lambda: self._show_error(f"Scan failed: {str(e)}"))
        
        threading.Thread(target=scan_thread, daemon=True).start()
    
    
    def _update_status_after_scan(self, statistics):
        """Update status after scanning is complete."""
        self.progress.stop()
        
        self.status_label.config(
            text=f"Scan complete: {statistics['total_workbooks']} workbooks, "
                 f"{statistics['workbooks_with_links']} with links, "
                 f"{statistics['total_links']} total links, "
                 f"{statistics['unique_external_files']} unique external files"
        )
        
        # Show all links initially
        self._display_all_links()
    
    def _display_all_links(self):
        """Display all external links in the tree."""
        # Clear existing items
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        if self.display_mode_var.get() == "grouped":
            self._display_grouped_links(self.links_manager.grouped_links)
        else:
            self._display_flat_links(self.links_manager.external_links)
    
    def _display_grouped_links(self, grouped_links):
        """Display links grouped by external file."""
        for external_file, group in grouped_links.items():
            # Create parent node for external file
            file_icon = DataFormatter.get_link_type_icon('Formula')
            parent_text = f"{file_icon} {external_file} ({DataFormatter.format_reference_count(group.reference_count)})"
            parent_id = self.results_tree.insert("", "end", text=parent_text, 
                                                values=("External File", "", ""))
            
            # Add child nodes for each reference
            for link in group.links:
                child_text = f"📄 {link.source_workbook}"
                location = f"{link.source_sheet}!{DataFormatter.format_cell_address(link.source_cell)}" if link.source_sheet else "LinkSource"
                formula = DataFormatter.truncate_formula(link.formula, 80)
                
                self.results_tree.insert(parent_id, "end", text=child_text,
                                       values=(link.link_type, location, formula),
                                       tags=(link.source_workbook, link.source_sheet, link.source_cell))
    
    def _display_flat_links(self, external_links):
        """Display links in flat list format."""
        for link in external_links:
            text = f"{DataFormatter.get_link_type_icon(link.link_type)} {link.source_workbook}"
            location = f"{link.source_sheet}!{DataFormatter.format_cell_address(link.source_cell)}" if link.source_sheet else "LinkSource"
            formula = DataFormatter.truncate_formula(link.formula, 80)
            
            self.results_tree.insert("", "end", text=text,
                                   values=(link.target_file, location, formula),
                                   tags=(link.source_workbook, link.source_sheet, link.source_cell))
    
    def _search_links(self):
        """Search for external links containing the keyword."""
        keyword = self.search_var.get().strip()
        search_field = self.search_field_var.get()
        
        if not keyword:
            self._display_all_links()
            return
        
        # Clear existing items
        for item in self.results_tree.get_children():
            self.results_tree.delete(item)
        
        if self.display_mode_var.get() == "grouped":
            # Get grouped search results
            grouped_results = self.links_manager.get_grouped_search_results(keyword, search_field)
            self._display_grouped_links(grouped_results)
            results_count = sum(group.reference_count for group in grouped_results.values())
        else:
            # Get flat search results
            matching_links = self.links_manager.search_links(keyword, search_field)
            self._display_flat_links(matching_links)
            results_count = len(matching_links)
        
        self.status_label.config(text=f"Search '{keyword}' in {search_field}: found {results_count} matching links")
    
    def _clear_search(self):
        """Clear search and show all links."""
        self.search_var.set("")
        self._display_all_links()
        
        statistics = self.links_manager.get_statistics()
        self.status_label.config(text=f"Showing all links: {statistics['total_links']} total")
    
    def _on_display_mode_change(self, event=None):
        """Handle display mode change."""
        keyword = self.search_var.get().strip()
        if keyword:
            self._search_links()
        else:
            self._display_all_links()
    
    def _on_result_double_click(self, event):
        """Handle double-click on result item."""
        selection = self.results_tree.selection()
        if not selection:
            return
        
        item = self.results_tree.item(selection[0])
        tags = item.get('tags', [])
        
        if len(tags) >= 3:
            workbook_name = tags[0]
            sheet_name = tags[1]
            cell_address = tags[2]
            
            # Try to navigate to the cell
            self._navigate_to_cell(workbook_name, sheet_name, cell_address)
    
    def _navigate_to_cell(self, workbook_name, sheet_name, cell_address):
        """Navigate to specific cell in Excel."""
        success, message = ExcelNavigator.navigate_to_cell(workbook_name, sheet_name, cell_address)
        
        if success:
            messagebox.showinfo("Navigation Success", message)
        else:
            messagebox.showerror("Navigation Error", message)
    
    def _export_results(self):
        """Export search results to file."""
        try:
            file_path = filedialog.asksaveasfilename(
                title="Export External Links Results",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            
            if file_path:
                # Get current search results or all data
                keyword = self.search_var.get().strip()
                if keyword:
                    search_field = self.search_field_var.get()
                    if self.display_mode_var.get() == "grouped":
                        grouped_results = self.links_manager.get_grouped_search_results(keyword, search_field)
                        success, message = ExternalLinksExporter.export_grouped_to_csv(grouped_results, file_path)
                    else:
                        matching_links = self.links_manager.search_links(keyword, search_field)
                        export_data = [
                            {
                                'Source Workbook': link.source_workbook,
                                'Source Sheet': link.source_sheet,
                                'Source Cell': link.source_cell,
                                'External File': link.target_file,
                                'Formula': link.formula,
                                'Link Type': link.link_type
                            }
                            for link in matching_links
                        ]
                        success, message = ExternalLinksExporter.export_to_csv(export_data, file_path)
                else:
                    # Export all data
                    export_data = self.links_manager.export_to_dict()
                    success, message = ExternalLinksExporter.export_to_csv(export_data, file_path)
                
                if success:
                    messagebox.showinfo("Export Success", message)
                else:
                    messagebox.showerror("Export Error", message)
        
        except Exception as e:
            messagebox.showerror("Export Error", f"Export failed:\n{str(e)}")
    
    def _show_batch_update(self):
        """Show batch update dialog."""
        messagebox.showinfo("Batch Update", "Batch update functionality will be available in the next version")
    
    def _show_error(self, message):
        """Show error message."""
        self.progress.stop()
        self.status_label.config(text="Error occurred")
        messagebox.showerror("Error", message)
    
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
    
    def _on_close(self):
        """Handle dialog close event."""
        if self.dialog:
            self.dialog.destroy()
            self.dialog = None