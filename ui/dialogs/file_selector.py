"""
File Selector Dialog for Excel Session Manager

This module contains the file selection dialog used when loading sessions,
allowing users to choose which files from a session to open.
"""
import tkinter as tk
from tkinter import ttk
import os

class DragSelectTreeview(ttk.Treeview):
    """Custom Treeview with drag selection like main interface"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._drag_start = None
        self._drag_mode = None
        self.bind("<Button-1>", self._on_click)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _on_click(self, event):
        iid = self.identify_row(event.y)
        if not iid:
            self.selection_remove(self.selection())
            return
        self.focus(iid)
        if iid in self.selection():
            self.selection_remove(iid)
            self._drag_mode = "unselect"
        else:
            self.selection_add(iid)
            self._drag_mode = "select"
        self._drag_start = iid
        return "break"

    def _on_drag(self, event):
        iid = self.identify_row(event.y)
        if not iid or not self._drag_start:
            return
        iids = self.get_children()
        start_idx = iids.index(self._drag_start)
        cur_idx = iids.index(iid)
        lo = min(start_idx, cur_idx)
        hi = max(start_idx, cur_idx)
        if self._drag_mode == "unselect":
            for i in iids[lo:hi+1]:
                self.selection_remove(i)
        else:
            for i in iids[lo:hi+1]:
                self.selection_add(i)
        return "break"

    def _on_release(self, event):
        self._drag_start = None
        self._drag_mode = None
        self.event_generate("<<TreeviewSelect>>")
        return "break"

class FileSelectionDialog:
    def __init__(self, parent, file_list, title="Select Files to Load"):
        self.parent = parent
        self.file_list = file_list
        self.selected_files = []
        self.result = None
        
        # Create dialog window
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("800x600")
        self.dialog.grab_set()
        self.dialog.resizable(True, True)
        
        # Make dialog modal
        self.dialog.transient(parent)
        self.dialog.focus_set()
        
        self.setup_ui()
        
    def setup_ui(self):
        # Main frame
        main_frame = tk.Frame(self.dialog, padx=10, pady=10)
        main_frame.pack(fill="both", expand=True)
        
        # Title label (use Arial like main interface)
        title_label = tk.Label(main_frame, text=f"Select files to load ({len(self.file_list)} files found):", 
                              font=("Arial", 16, "bold"))
        title_label.pack(anchor="w", pady=(0, 10))
        
        # Select all frame
        select_frame = tk.Frame(main_frame)
        select_frame.pack(fill="x", pady=(0, 10))
        
        select_all_btn = tk.Button(select_frame, text="Select All", 
                                  command=self.select_all,
                                  font=("Arial", 10, "bold"))
        select_all_btn.pack(side="left", padx=(0, 10))
        
        deselect_all_btn = tk.Button(select_frame, text="Deselect All", 
                                    command=self.deselect_all,
                                    font=("Arial", 10))
        deselect_all_btn.pack(side="left")
        
        # File list frame with scrollbar
        list_frame = tk.Frame(main_frame)
        list_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Create treeview container (like main interface)
        treeview_frame = tk.Frame(list_frame)
        treeview_frame.pack(fill="both", expand=True)
        
        # Create treeview with two columns
        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Arial", 12, "bold"))
        style.configure("Treeview", font=("Consolas", 12), rowheight=25)
        
        self.file_treeview = DragSelectTreeview(treeview_frame, 
                                               columns=("col1", "col2"), 
                                               show="headings", 
                                               selectmode="extended")
        
        # Configure columns
        self.file_treeview.column("col1", anchor="w", width=300)
        self.file_treeview.column("col2", anchor="w")
        self.file_treeview.heading("col1", text="File Name")
        self.file_treeview.heading("col2", text="File Path")
        
        # Create scrollbar
        scrollbar = tk.Scrollbar(treeview_frame, orient="vertical", command=self.file_treeview.yview)
        self.file_treeview.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        self.file_treeview.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate treeview with files
        for i, file_path in enumerate(self.file_list):
            filename = os.path.basename(file_path)
            self.file_treeview.insert("", "end", values=(filename, file_path))
        
        # Select all items by default
        for item in self.file_treeview.get_children():
            self.file_treeview.selection_add(item)
        
        # Bind events like main interface
        self.file_treeview.bind("<<TreeviewSelect>>", self.on_selection_change)
        self.file_treeview.bind("<MouseWheel>", lambda event: self.file_treeview.yview_scroll(-1 * int(event.delta / 120), "units"))
        self.file_treeview.bind("<Button-4>", lambda event: self.file_treeview.yview_scroll(-1, "units"))
        self.file_treeview.bind("<Button-5>", lambda event: self.file_treeview.yview_scroll(1, "units"))
        
        # Button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Buttons (use Arial like main interface)
        ok_btn = tk.Button(button_frame, text="Load Selected Files", 
                          command=self.on_ok, font=("Arial", 10, "bold"),
                          bg="#4CAF50", fg="white", width=20)
        ok_btn.pack(side="right", padx=(10, 0))
        
        cancel_btn = tk.Button(button_frame, text="Cancel", 
                              command=self.on_cancel, font=("Arial", 10),
                              width=15)
        cancel_btn.pack(side="right")
        
        # Info label (blue counter, use Arial like main interface)
        self.info_label = tk.Label(button_frame, text="", font=("Arial", 12), fg="blue")
        self.info_label.pack(side="left")
        
        self.update_info_label()
        
    def select_all(self):
        for item in self.file_treeview.get_children():
            self.file_treeview.selection_add(item)
        self.update_info_label()
        
    def deselect_all(self):
        self.file_treeview.selection_remove(self.file_treeview.selection())
        self.update_info_label()
        
    def on_selection_change(self, event):
        self.update_info_label()
        
    def update_info_label(self):
        selected_items = self.file_treeview.selection()
        selected_count = len(selected_items)
        total_count = len(self.file_list)
        self.info_label.config(text=f"Selected: {selected_count}/{total_count} files")
        
    def on_ok(self):
        selected_items = self.file_treeview.selection()
        self.selected_files = []
        for item in selected_items:
            values = self.file_treeview.item(item, "values")
            if values and len(values) >= 2:
                file_path = values[1]  # Get file path from second column
                self.selected_files.append(file_path)
        
        if not self.selected_files:
            import tkinter.messagebox as messagebox
            messagebox.showwarning("Warning", "Please select at least one file to load.")
            return
            
        self.result = "ok"
        self.dialog.destroy()
        
    def on_cancel(self):
        self.result = "cancel"
        self.dialog.destroy()
        
    def show(self):
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
        
        # Wait for dialog to close
        self.dialog.wait_window()
        return self.result, self.selected_files

# Test function
def test_file_selector():
    root = tk.Tk()
    root.withdraw()  # Hide main window
    
    # Sample file list
    test_files = [
        "C:/path/to/file1.xlsx",
        "C:/path/to/file2.xlsx", 
        "C:/path/to/file3.xlsx"
    ]
    
    dialog = FileSelectionDialog(root, test_files)
    result, selected = dialog.show()
    
    print(f"Result: {result}")
    print(f"Selected files: {selected}")

if __name__ == "__main__":
    test_file_selector()