import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font
import pythoncom
import win32com.client
from win32com.client import constants
import win32gui
import win32con
import time
import openpyxl
import os
from datetime import datetime
import threading
from PIL import Image, ImageTk
from excel_session_manager_link_updater import run_excel_link_update
import psutil
import gc

# Import utility functions
from utils.file_utils import get_file_mtime_str, parse_mtime
from utils.ui_utils import calc_row_height, calc_col2_width

# Import UI components
from ui.components.drag_treeview import DragSelectTreeview
from ui.console_popup import ConsolePopup
from ui.dialogs.link_options import LinkUpdateOptionsDialog

# Import core components
from core.session_manager import SessionManager
from core.excel_manager import ExcelManager
from core.process_manager import ProcessManager
from core.performance_monitor import get_performance_monitor

# Import constants and settings
from config.constants import (
    MONO_FONTS, DEFAULT_WINDOW_SIZE, DEFAULT_WINDOW_TITLE, MIN_WINDOW_SIZE,
    NORMAL_GEOMETRY, MINI_WIDGET_SIZE, MINI_WIDGET_POSITION, MINI_WIDGET_ICON,
    DEFAULT_LOG_DIRECTORY, BUTTON_WIDTH, BUTTON_HEIGHT, BUTTON_WRAP_LENGTH,
    TREEVIEW_COLUMN_WIDTH, TREEVIEW_PADDING, DEFAULT_FONT_SIZE, HEADER_FONT_SIZE,
    DEFAULT_CHECK_DAYS
)
from config.settings import settings

def kill_zombie_excel_processes(whitelist_hwnds=None):
    if whitelist_hwnds is None:
        whitelist_hwnds = set()
    zombie_pids = []
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and proc.info['name'].lower() == 'excel.exe':
            try:
                hwnds = []
                def callback(hwnd, hwnds):
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if pid == proc.info['pid']:
                        hwnds.append(hwnd)
                win32gui.EnumWindows(callback, hwnds)
                keep = False
                for hwnd in hwnds:
                    if hwnd in whitelist_hwnds:
                        keep = True
                        break
                if not keep and not hwnds:
                    zombie_pids.append(proc.info['pid'])
            except Exception:
                continue
    for pid in zombie_pids:
        try:
            p = psutil.Process(pid)
            p.kill()
        except Exception:
            continue

# MONO_FONTS moved to config.constants

# calc_row_height and calc_col2_width moved to utils.ui_utils
# DragSelectTreeview moved to ui.components.drag_treeview

# get_file_mtime_str and parse_mtime moved to utils.file_utils
# ConsolePopup moved to ui.console_popup

class ExcelSessionManagerApp:
    def __init__(self, root):
        self.root = root
        self.root.title(settings.window_title)
        self.root.geometry(settings.window_size)
        self.last_log_dir = settings.log_directory
        min_w, min_h = map(int, settings.min_window_size.split('x'))
        self.root.minsize(min_w, min_h)
        self.root.protocol("WM_DELETE_WINDOW", lambda: self.root.destroy())
        self.is_mini = False
        self.normal_geometry = settings.get("ui.window.normal_geometry", NORMAL_GEOMETRY)
        self.mini_side = settings.mini_widget_size
        self.mini_geometry = f"{self.mini_side}x{self.mini_side}+{settings.mini_widget_position}"
        self.floating_icon_btn = None
        
        # Initialize managers
        self.session_manager = SessionManager(self.root)
        self.excel_manager = ExcelManager()
        self.process_manager = ProcessManager()
        
        self.setup_ui()
        self.show_names()

    def setup_ui(self):
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(expand=True, fill="both")

        frame_session = ttk.Frame(self.notebook)
        self.notebook.add(frame_session, text="Session Manager")

        font_frame = tk.Frame(frame_session)
        font_frame.pack(pady=(5, 0), padx=10, anchor="nw", fill='x')

        self.pin_btn = tk.Button(
            font_frame,
            text="Mini Widget (Floating Pin)",
            font=("Microsoft JhengHei", 10),
            command=self.enter_mini,
            relief="raised",
            borderwidth=1,
            padx=5,
            pady=2
        )
        self.pin_btn.pack(side=tk.RIGHT, padx=(10, 0))

        tk.Label(font_frame, text="Font:").pack(side=tk.LEFT)
        self.font_family_var = tk.StringVar(value=MONO_FONTS[0])
        self.font_size_var = tk.IntVar(value=12)
        font_family_menu = ttk.OptionMenu(font_frame, self.font_family_var, MONO_FONTS[0], *MONO_FONTS, command=self.update_treeview_font)
        font_family_menu.pack(side=tk.LEFT, padx=(5, 15))

        tk.Label(font_frame, text="Size:").pack(side=tk.LEFT)
        font_size_spin = ttk.Spinbox(font_frame, from_=8, to=32, textvariable=self.font_size_var, width=4, command=self.update_treeview_font)
        font_size_spin.pack(side=tk.LEFT)
        self.font_size_var.trace_add("write", lambda *a: self.update_treeview_font())

        title_frame = tk.Frame(frame_session)
        title_frame.pack(pady=(10, 0), padx=10, anchor="w", fill='x')

        main_label = tk.Label(title_frame, text="Active Excel Session:", font=("Arial", 16, "bold"))
        main_label.pack(side=tk.LEFT)

        self.count_label = tk.Label(title_frame, text="", font=("Arial", 12))
        self.count_label.pack(side=tk.LEFT, padx=(5, 0), pady=(2,0))

        main_frame = tk.Frame(frame_session)
        main_frame.pack(pady=(5, 10), padx=10, expand=True, fill='both')

        listbox_frame = tk.Frame(main_frame)
        listbox_frame.pack(side=tk.LEFT, fill='both', expand=True, pady=10)

        select_console_frame = tk.Frame(listbox_frame)
        select_console_frame.pack(anchor='w', padx=0, pady=(0, 2), fill='x')

        self.select_all_var = tk.BooleanVar(value=False)
        select_all_cb = tk.Checkbutton(
            select_console_frame,
            text="Select All",
            variable=self.select_all_var,
            command=self.on_select_all_toggle
        )
        select_all_cb.pack(side='left', padx=(8,0))

        self.show_console_progress_var = tk.BooleanVar(value=settings.show_console_by_default)
        show_console_cb = tk.Checkbutton(
            select_console_frame,
            text="Show Progress Console",
            variable=self.show_console_progress_var
        )
        show_console_cb.pack(side='left', padx=(12,0))

        style = ttk.Style()
        style.configure("Treeview.Heading", font=("Consolas", 12, "bold"))
        style.configure("Treeview", font=("Consolas", 12), rowheight=calc_row_height(self.font_size_var.get()))

        self.tree = DragSelectTreeview(listbox_frame, columns=("col1", "col2"), show="headings", selectmode="extended")
        self.tree.column("col1", anchor="w")
        self.tree.column("col2", anchor="e", width=settings.get("ui.treeview.column_width", TREEVIEW_COLUMN_WIDTH), stretch=False)
        tree_yscrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=tree_yscrollbar.set)
        tree_yscrollbar.pack(side="right", fill="y")
        padding = settings.get("ui.treeview.padding", TREEVIEW_PADDING)
        self.tree.pack(expand=True, fill='both', padx=padding, pady=padding)
        self.root.after(200, lambda: self.tree.xview_moveto(1.0))
        self.tree.bind("<MouseWheel>", lambda event: self.tree.yview_scroll(-1 * int(event.delta / 120), "units"))
        self.tree.bind("<Button-4>", lambda event: self.tree.yview_scroll(-1, "units"))
        self.tree.bind("<Button-5>", lambda event: self.tree.yview_scroll(1, "units"))
        self.tree.bind('<<TreeviewSelect>>', self.on_selection_change)
        self.tree.bind("<Button-1>", self.on_treeview_heading_click, add="+")

        side_btn_frame = tk.Frame(main_frame, width=120)
        side_btn_frame.pack(side=tk.RIGHT, padx=(10, 0), pady=(40, 0), fill='y')
        side_btn_frame.pack_propagate(False)
        btn_props = {
            'width': BUTTON_WIDTH, 
            'height': BUTTON_HEIGHT, 
            'wraplength': BUTTON_WRAP_LENGTH, 
            'font': ("Arial", 10, "bold")
        }

        self.refresh_btn = tk.Button(side_btn_frame, text="Refresh\nList", **btn_props, command=self.show_names)
        self.refresh_btn.pack(pady=5, anchor='n')

        self.show_path_btn = tk.Button(side_btn_frame, text="Show\nFile Path", **btn_props, command=self.toggle_path)
        self.show_path_btn.pack(pady=5, anchor='n')

        activate_btn = tk.Button(side_btn_frame, text="Activate\nSelected", **btn_props, command=self.activate_selected_workbooks)
        activate_btn.pack(pady=5, anchor='n')

        minimize_btn = tk.Button(side_btn_frame, text="Minimize All\nExcel", **btn_props, command=self.minimize_all_excel)
        minimize_btn.pack(pady=5, anchor='n')

        save_sess_btn = tk.Button(side_btn_frame, text="Save\nSession", **btn_props, command=self.save_session)
        save_sess_btn.pack(pady=5, anchor='n')

        load_sess_btn = tk.Button(side_btn_frame, text="Load\nSession", **btn_props, command=self.load_session)
        load_sess_btn.pack(pady=5, anchor='n')

        update_links_btn = tk.Button(side_btn_frame, text="Update Recent\nExternal Links", **btn_props, command=self.open_link_update_options)
        update_links_btn.pack(pady=5, anchor='n')
        
        cleanup_btn = tk.Button(side_btn_frame, text="Cleanup Excel\nProcesses", **btn_props, command=self.cleanup_excel_processes)
        cleanup_btn.pack(pady=5, anchor='n')
        
        performance_btn = tk.Button(side_btn_frame, text="Performance\nMonitor", **btn_props, command=self.show_performance_monitor)
        performance_btn.pack(pady=5, anchor='n')
        
        external_links_btn = tk.Button(side_btn_frame, text="External Links\nManager", **btn_props, command=self.show_external_links_manager)
        external_links_btn.pack(pady=5, anchor='n')
        
        btn1 = tk.Button(side_btn_frame, text="Save\nSelected", **btn_props, command=self.save_selected_workbooks)
        btn1.pack(pady=5, anchor='n')

        btn2 = tk.Button(side_btn_frame, text="Close\nWithout Saving", **btn_props, command=lambda: self.close_selected_workbooks(False))
        btn2.pack(pady=5, anchor='n')

        btn3 = tk.Button(side_btn_frame, text="Save and Close\nSelected", **btn_props, command=lambda: self.close_selected_workbooks(True))
        btn3.pack(pady=5, anchor='n')

        self.file_names, self.file_paths, self.sheet_names, self.active_cells = [], [], [], []
        self.showing_path = False
        self.target_captions = []
        self.sort_states = {"col1": "none", "col2": "none"}
        self.original_data = []

    def on_select_all_toggle(self):
        if self.select_all_var.get():
            for iid in self.tree.get_children():
                self.tree.selection_add(iid)
        else:
            self.tree.selection_remove(self.tree.selection())

    def enter_mini(self):
        if self.is_mini:
            return
        self.is_mini = True
        self.normal_geometry = self.root.geometry()
        self.notebook.pack_forget()
        self.root.minsize(self.mini_side, self.mini_side)
        self.root.geometry(self.mini_geometry)
        self.root.resizable(False, False)
        self.root.wm_attributes("-topmost", 1)

        icon_path = settings.get("ui.mini_widget.icon_file", MINI_WIDGET_ICON)
        icon_size = (self.mini_side - 40, self.mini_side - 40)
        if os.path.exists(icon_path):
            try:
                pil_image = Image.open(icon_path).resize(icon_size, Image.LANCZOS)
                icon_image = ImageTk.PhotoImage(pil_image)
                self.floating_icon_btn = tk.Button(
                    self.root, image=icon_image, width=icon_size[0], height=icon_size[1],
                    command=self.exit_mini, borderwidth=0, relief="flat"
                )
                self.floating_icon_btn.image = icon_image
            except Exception:
                self.create_fallback_icon()
        else:
            self.create_fallback_icon()
        self.floating_icon_btn.pack(expand=True, padx=10, pady=10)

    def exit_mini(self):
        if not self.is_mini:
            return
        self.is_mini = False
        if self.floating_icon_btn:
            self.floating_icon_btn.destroy()
            self.floating_icon_btn = None
        self.root.minsize(600, 500)
        self.root.geometry(self.normal_geometry)
        self.root.resizable(True, True)
        self.root.wm_attributes("-topmost", 0)
        self.notebook.pack(expand=True, fill="both")

    def create_fallback_icon(self):
        self.floating_icon_btn = tk.Button(
            self.root, text="🗔", font=("Segoe UI Emoji", 48),
            command=self.exit_mini, borderwidth=0, relief="flat"
        )

    def update_treeview_font(self, *args):
        new_font = (self.font_family_var.get(), self.font_size_var.get())
        style = ttk.Style()
        style.configure("Treeview.Heading", font=(self.font_family_var.get(), self.font_size_var.get(), "bold"))
        self.tree.tag_configure('custom_font', font=new_font)
        for iid in self.tree.get_children():
            self.tree.item(iid, tags=('custom_font',))
        row_h = calc_row_height(self.font_size_var.get())
        style.configure("Treeview", rowheight=row_h)
        test_str = "2025-06-29 11:32:48"
        test_font = font.Font(family=self.font_family_var.get(), size=self.font_size_var.get())
        col2_width = test_font.measure(test_str) + 18
        self.tree.column("col2", width=col2_width, stretch=False)

    def get_open_excel_files(self):
        pythoncom.CoInitialize()
        excel_files, file_paths, sheet_names, active_cells = [], [], [], []
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            for wb in excel.Workbooks:
                excel_files.append(wb.Name)
                file_paths.append(wb.FullName)
                try:
                    sht = wb.ActiveSheet
                    sheet_names.append(sht.Name)
                    cell_addr = sht.Application.ActiveCell.Address
                    active_cells.append(cell_addr)
                except Exception:
                    sheet_names.append("")
                    active_cells.append("")
        except Exception as e:
            print("Error:", e)
        finally:
            pythoncom.CoUninitialize()
        return excel_files, file_paths, sheet_names, active_cells
        
    def open_link_update_options(self):
        """Open the link update options dialog."""
        monitor = get_performance_monitor()
        op_id = monitor.start_operation("update_external_links", {'action': 'dialog_opened'})
        try:
            dialog = LinkUpdateOptionsDialog(self.root, self.last_log_dir, getattr(self, 'last_summary_dir', None))
            self.last_log_dir, self.last_summary_dir = dialog.show()
            monitor.end_operation(op_id, success=True)
        except Exception as e:
            monitor.end_operation(op_id, success=False)
            raise
    
    def cleanup_excel_processes(self):
        """Cleanup Excel processes using ProcessManager."""
        show_console = self.show_console_progress_var.get() if hasattr(self, "show_console_progress_var") else True
        popup = ConsolePopup(self.root, title="Excel Process Cleanup") if show_console else None
        
        def print_to_popup(msg):
            if popup:
                self.root.after(0, lambda: popup.print(msg))
                
        def thread_job():
            # First show health report
            print_to_popup("=== Excel Process Health Check ===")
            health_report = self.process_manager.monitor_excel_health(print_to_popup)
            print_to_popup("")
            
            # Then cleanup if needed
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
            from ui.performance_dialog import PerformanceDialog
            dialog = PerformanceDialog(self.root)
            dialog.show()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open performance monitor:\n{str(e)}")
    
    def show_external_links_manager(self):
        """Show the external links manager dialog."""
        try:
            from ui.external_links_dialog import ExternalLinksDialog
            dialog = ExternalLinksDialog(self.root)
            dialog.show()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open external links manager:\n{str(e)}")
    
    def treeview_sort(self, col):
        items = [self.tree.item(iid, "values") for iid in self.tree.get_children()]
        if not self.original_data or len(self.original_data) != len(items):
            self.original_data = list(items)
        if self.sort_states[col] == "none": self.sort_states[col] = "asc"
        elif self.sort_states[col] == "asc": self.sort_states[col] = "desc"
        else: self.sort_states[col] = "none"
        for other_col in self.sort_states:
            if other_col != col: self.sort_states[other_col] = "none"
        if self.sort_states[col] == "none":
            sorted_items = self.original_data
        else:
            reverse = self.sort_states[col] == "desc"
            if col == "col1": key_func = lambda x: (x[0] or "").lower()
            elif col == "col2": key_func = lambda x: (parse_mtime(x[1]) or datetime.min)
            else: key_func = lambda x: x
            sorted_items = sorted(items, key=key_func, reverse=reverse)
        for iid in self.tree.get_children(): self.tree.delete(iid)
        for v in sorted_items: self.tree.insert("", "end", values=v)
        self.update_treeview_font()

    def get_all_excel_instances(self):
        import pythoncom
        import win32com.client
        instances = []
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
            instances.append(excel)
        except Exception as e:
            print("GetActiveObject error:", e)
        return instances

    def get_open_excel_files(self):
        pythoncom.CoInitialize()
        excel_files, file_paths, sheet_names, active_cells = [], [], [], []
        instances = self.get_all_excel_instances()
        for excel in instances:
            try:
                for wb in excel.Workbooks:
                    if wb.FullName in file_paths:  # avoid duplicates
                        continue
                    excel_files.append(wb.Name)
                    file_paths.append(wb.FullName)
                    try:
                        sht = wb.ActiveSheet
                        sheet_names.append(sht.Name)
                        cell_addr = sht.Application.ActiveCell.Address
                        active_cells.append(cell_addr)
                    except Exception:
                        sheet_names.append("")
                        active_cells.append("")
            except Exception:
                continue
        pythoncom.CoUninitialize()
        return excel_files, file_paths, sheet_names, active_cells

    def show_names(self):
        def get_stable_workbook_list(max_retry=5, wait_sec=0.5):
            last_count = -1
            stable_data = ([], [], [], [])
            for _ in range(max_retry):
                current_data = self.get_open_excel_files()
                current_count = len(current_data[0])
                if current_count == last_count:
                    return current_data
                last_count = current_count
                stable_data = current_data
                time.sleep(wait_sec)
            return stable_data
        def update_gui(data):
            if data is None:
                return
            self.file_names, self.file_paths, self.sheet_names, self.active_cells = data
            self.count_label.config(text=f"({len(self.file_names)} files open)")
            for i in self.tree.get_children(): self.tree.delete(i)
            self.tree.heading("col1", text="File Path" if self.showing_path else "File Name")
            self.tree.heading("col2", text="Last Modified")
            if not self.file_names:
                self.tree.insert("", "end", values=("No Excel files are currently open.", ""))
            elif self.showing_path:
                for path in self.file_paths:
                    mtime = get_file_mtime_str(path)
                    self.tree.insert("", "end", values=(path, mtime))
            else:
                for i, name in enumerate(self.file_names):
                    mtime = get_file_mtime_str(self.file_paths[i])
                    self.tree.insert("", "end", values=(name, mtime))
            self.refresh_btn.config(state=tk.NORMAL)
            self.update_treeview_font()
            self.original_data.clear()
            self.sort_states = {"col1": "none", "col2": "none"}
            self.tree.xview_moveto(1.0)
        def scan_in_thread():
            kill_zombie_excel_processes()
            scan_data = get_stable_workbook_list()
            self.root.after(0, lambda: update_gui(scan_data))
        self.refresh_btn.config(state=tk.DISABLED)
        threading.Thread(target=scan_in_thread, daemon=True).start()


    def toggle_path(self):
        self.showing_path = not self.showing_path
        self.show_names()
        self.show_path_btn.config(text="Hide\nFile Path" if self.showing_path else "Show\nFile Path")

    def get_selected_workbooks(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showinfo("Notice", "Please select one or more Excel files before proceeding.")
            return []
        selected_workbooks = []
        current_shown_files = self.file_paths if self.showing_path else self.file_names
        for item_id in selected_items:
            item_values = self.tree.item(item_id, "values")
            if not item_values or not item_values[0]: continue
            selected_value = item_values[0]
            try:
                idx = current_shown_files.index(selected_value)
                if 0 <= idx < len(self.file_names):
                    selected_workbooks.append((self.file_names[idx], self.file_paths[idx], self.sheet_names[idx], self.active_cells[idx]))
            except ValueError:
                continue
        return selected_workbooks

    def on_selection_change(self, event):
        selected_items = self.tree.selection()
        captions = []
        for item_id in selected_items:
            item_values = self.tree.item(item_id, "values")
            if not item_values or not item_values[0]: continue
            selected_value = item_values[0]
            if self.showing_path:
                captions.append(os.path.basename(selected_value))
            else:
                captions.append(selected_value)
        self.target_captions = captions
        self.select_all_var.set(
            len(self.tree.get_children()) > 0 and
            len(self.tree.selection()) == len(self.tree.get_children())
        )

    def on_treeview_heading_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            col = self.tree.identify_column(event.x)
            if col == "#1": self.treeview_sort("col1")
            elif col == "#2": self.treeview_sort("col2")

    def activate_selected_workbooks(self):
        self.on_selection_change(None)
        if not self.target_captions:
            messagebox.showinfo("Notice", "Please select one or more Excel files to activate.")
            return
        offset_x, offset_y, start_x, start_y = 40, 40, 100, 100
        activated_hwnds, window_index = set(), 0
        def enum_handler(hwnd, ctx):
            nonlocal window_index
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                for caption in ctx["captions"]:
                    if caption in title and " - Excel" in title and hwnd not in activated_hwnds:
                        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                        win32gui.SetForegroundWindow(hwnd)
                        x, y = start_x + window_index * offset_x, start_y + window_index * offset_y
                        rect = win32gui.GetWindowRect(hwnd)
                        w, h = rect[2] - rect[0], rect[3] - rect[1]
                        win32gui.SetWindowPos(hwnd, None, x, y, w, h, win32con.SWP_NOZORDER | win32con.SWP_SHOWWINDOW)
                        activated_hwnds.add(hwnd)
                        window_index += 1
                        time.sleep(0.05)
        win32gui.EnumWindows(enum_handler, {"captions": self.target_captions})

    def save_selected_workbooks(self):
        """Save selected workbooks using ExcelManager."""
        selected = self.get_selected_workbooks()
        if not selected:
            return
            
        show_console = self.show_console_progress_var.get() if hasattr(self, "show_console_progress_var") else True
        popup = ConsolePopup(self.root, title="Save Selected Progress") if show_console else None
        
        def print_to_popup(msg):
            if popup:
                self.root.after(0, lambda: popup.print(msg))
                
        def thread_job():
            self.excel_manager.save_workbooks(selected, print_to_popup)
            self.root.after(0, lambda: messagebox.showinfo("Complete", "Selected Excel files have been saved successfully."))
            self.root.after(0, self.show_names)
            
        threading.Thread(target=thread_job, daemon=True).start()

    def save_selected_workbooks_old(self):
        selected = self.get_selected_workbooks()
        if not selected: return
        show_console = self.show_console_progress_var.get() if hasattr(self, "show_console_progress_var") else True
        popup = ConsolePopup(self.root, title="Save Selected Progress") if show_console else None
        def print_to_popup(msg):
            if popup: self.root.after(0, lambda: popup.print(msg))
        def thread_job():
            pythoncom.CoInitialize()
            excel = None
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                orig_alert = excel.DisplayAlerts
                excel.DisplayAlerts = False
                print_to_popup(f"Saving {len(selected)} file(s)...")
                print_to_popup("-" * 80)
                for idx, (name, path, _, _) in enumerate(selected, 1):
                    print_to_popup(f"({idx}/{len(selected)}) Saving: {path}")
                    t0 = time.time()
                    saved = False
                    for wb in excel.Workbooks:
                        if wb.Name == name and wb.FullName == path:
                            try:
                                # Check if file is readonly
                                if os.path.exists(path):
                                    file_attrs = os.stat(path).st_mode
                                    if not (file_attrs & 0o200):
                                        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                        print_to_popup(f'{ts} |     File "{path}" is read-only. Changes may not be saved.')
                                
                                # Get modification time before save
                                before_save_time = None
                                if os.path.exists(path):
                                    before_save_time = datetime.fromtimestamp(os.path.getmtime(path))
                                    before_save_str = before_save_time.strftime("%Y-%m-%d %H:%M:%S")
                                    print_to_popup(f"Before save modified time: {before_save_str}")
                                
                                # Execute save
                                wb.Save()
                                
                                # Get modification time after save and compare
                                save_success = False
                                retry_count = 0
                                max_retries = 3
                                
                                while not save_success and retry_count < max_retries:
                                    time.sleep(0.1)
                                    
                                    if os.path.exists(path):
                                        after_save_time = datetime.fromtimestamp(os.path.getmtime(path))
                                        after_save_str = after_save_time.strftime("%Y-%m-%d %H:%M:%S")
                                        print_to_popup(f"After save modified time: {after_save_str}")
                                        
                                        # Compare modification times
                                        if before_save_time and before_save_time == after_save_time:
                                            retry_count += 1
                                            print_to_popup(f"Warning: No time change detected, save may have failed! (Retry {retry_count}/{max_retries})")
                                            if retry_count < max_retries:
                                                print_to_popup("Retrying save...")
                                                wb.Save()
                                            else:
                                                print_to_popup(f"Failed after {max_retries} retries")
                                                break
                                        else:
                                            print_to_popup("Success: Modification time changed, save successful!")
                                            save_success = True
                                    else:
                                        print_to_popup("Save completed")
                                        save_success = True
                                
                                print_to_popup(f"({idx}/{len(selected)}) Saved: {path}")
                                saved = True
                            except Exception as e:
                                print_to_popup(f"({idx}/{len(selected)}) Save failed: {path} ({e})")
                            del wb  # <==  workbook handle
                            break
                    if not saved:
                        print_to_popup(f"({idx}/{len(selected)}) Workbook not found in Excel: {path}")
                    t1 = time.time()
                    used_sec = t1 - t0
                    print_to_popup(f"used time: {used_sec:.2f} sec")
                    print_to_popup("-" * 80)
                excel.DisplayAlerts = orig_alert
                print_to_popup("Save selected complete.")
                self.root.after(0, lambda: messagebox.showinfo("Complete", "Selected Excel files have been saved successfully."))
            except Exception as e:
                print_to_popup(f"Error: {str(e)}")
                self.root.after(0, lambda e=e: messagebox.showerror("Error", f"An error occurred while saving the files:\n{str(e)}"))
            finally:
                if excel is not None:
                    del excel  # <== 
                gc.collect()
                pythoncom.CoUninitialize()
                self.root.after(0, self.show_names)
        threading.Thread(target=thread_job, daemon=True).start()


    def close_selected_workbooks(self, save_before_close=False):
        # Add performance monitoring
        monitor = get_performance_monitor()
        operation_name = "close_workbooks_with_save" if save_before_close else "close_workbooks_without_save"
        selected = self.get_selected_workbooks()
        if not selected: return
        op_id = monitor.start_operation(operation_name, {'workbook_count': len(selected)})
        # your patch added between the existing code
        show_console = self.show_console_progress_var.get() if hasattr(self, "show_console_progress_var") else True
        popup = ConsolePopup(self.root, title="Save and Close Progress" if save_before_close else "Close Progress") if show_console else None
        def print_to_popup(msg):
            if popup: self.root.after(0, lambda: popup.print(msg))
        def thread_job():
            pythoncom.CoInitialize()
            excel = None
            orig_alert = None
            try:
                excel = win32com.client.GetActiveObject("Excel.Application")
                try:
                    orig_alert = excel.DisplayAlerts
                    excel.DisplayAlerts = False
                except Exception:
                    pass
                print_to_popup(f"{'Save and close' if save_before_close else 'Close'} {len(selected)} file(s)...")
                print_to_popup("-" * 80)
                for idx, (name, path, _, _) in enumerate(selected, 1):
                    t0 = time.time()
                    closed = False
                    if save_before_close:
                        print_to_popup(f"({idx}/{len(selected)}) Saving and closing: {path}")
                    else:
                        print_to_popup(f"({idx}/{len(selected)}) Closing: {path}")
                    for wb in excel.Workbooks:
                        if wb.Name == name and wb.FullName == path:
                            try:
                                wb.Close(SaveChanges=save_before_close)
                                closed = True
                                if save_before_close:
                                    print_to_popup(f"({idx}/{len(selected)}) Saved and closed: {path}")
                                else:
                                    print_to_popup(f"({idx}/{len(selected)}) Closed: {path}")
                            except Exception as e:
                                print_to_popup(f"({idx}/{len(selected)}) Close failed: {path} ({e})")
                            break
                    if not closed:
                        print_to_popup(f"({idx}/{len(selected)}) Workbook not found in Excel: {path}")
                    t1 = time.time()
                    used_sec = t1 - t0
                    print_to_popup(f"used time: {used_sec:.2f} sec")
                    print_to_popup("-" * 80)
                if excel.Workbooks.Count == 0:
                    excel.Quit()
                print_to_popup(f"{'Save and close' if save_before_close else 'Close'} selected complete.")
                self.root.after(0, lambda: messagebox.showinfo("Complete", f"Selected files have been {'saved and ' if save_before_close else ''}closed."))
            except Exception as e:
                print_to_popup(f"Error: {str(e)}")
                self.root.after(0, lambda e=e: messagebox.showerror("Error", f"An error occurred while closing the files:\n{str(e)}"))
            finally:
                if excel is not None:
                    try:
                        if hasattr(excel, "DisplayAlerts") and orig_alert is not None:
                            excel.DisplayAlerts = orig_alert
                    except Exception:
                        pass
                    del excel
                gc.collect()
                pythoncom.CoUninitialize()
                self.root.after(0, self.show_names)
                # End performance monitoring
                monitor.end_operation(op_id, success=True)
        threading.Thread(target=thread_job, daemon=True).start()

    def minimize_all_excel(self):
        def enum_handler(hwnd, ctx):
            if win32gui.IsWindowVisible(hwnd) and " - Excel" in win32gui.GetWindowText(hwnd):
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
        win32gui.EnumWindows(enum_handler, None)

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
            getattr(self, 'show_console_progress_var', None)
        )
        # Refresh the file list after loading
        self.root.after(200, self.show_names)

def load_session_from_path(self, file_path):
    def thread_job():
        pythoncom.CoInitialize()
        excel = None
        try:
            wb = openpyxl.load_workbook(file_path)
            ws = wb.active
            rows = list(ws.iter_rows(min_row=2, values_only=True))
            valid_rows = [r for r in rows if r and r[0] and os.path.exists(r[0])]
            if not valid_rows:
                self.root.after(0, lambda: messagebox.showwarning("Warning", "No valid file paths found."))
                return
            excel = win32com.client.Dispatch("Excel.Application")
            try:
                excel.Visible = True
            except Exception:
                pass
            excel.AskToUpdateLinks = False
            for idx, r in enumerate(valid_rows, 1):
                path, sheet, cell = (r[0], r[1] if len(r) > 1 else None, r[2] if len(r) > 2 else None)
                # CONSOLE PRINT: Opening
                ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                print_msg = f"{ts} | ({idx}/{len(valid_rows)}) Opening: {path}"
                try:
                    # Print to console (if ConsolePopup used)
                    if hasattr(self, "popup") and self.popup:
                        self.root.after(0, lambda msg=print_msg: self.popup.print(msg))
                    wb_xl = excel.Workbooks.Open(Filename=path, UpdateLinks=0)
                    if wb_xl.ReadOnly:
                        ts2 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        ro_msg = f'{ts2} |     File "{path}" is opened in read-only mode. Changes may not be saved.'
                        if hasattr(self, "popup") and self.popup:
                            self.root.after(0, lambda msg=ro_msg: self.popup.print(msg))
                    #  sheet/cell select
                    if sheet:
                        try:
                            sht = wb_xl.Sheets(sheet)
                            sht.Activate()
                            if cell: sht.Range(cell).Select()
                        except Exception: pass
                    ts3 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    opened_msg = f"{ts3} | ({idx}/{len(valid_rows)}) Opened: {path}"
                    if hasattr(self, "popup") and self.popup:
                        self.root.after(0, lambda msg=opened_msg: self.popup.print(msg))
                except Exception as e:
                    ts4 = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    fail_msg = f"{ts4} | ({idx}/{len(valid_rows)}) Failed to open: {path} ({e})"
                    if hasattr(self, "popup") and self.popup:
                        self.root.after(0, lambda msg=fail_msg: self.popup.print(msg))
            excel.AskToUpdateLinks = True
            self.root.after(0, lambda: messagebox.showinfo("Complete", f"{len(valid_rows)} file(s) opened."))
        except Exception as e:
            self.root.after(0, lambda e=e: messagebox.showerror("Error", f"Error loading session:\n{str(e)}"))
        finally:
            if excel is not None:
                del excel
            gc.collect()
            pythoncom.CoUninitialize()
            self.root.after(200, self.show_names)
    threading.Thread(target=thread_job, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelSessionManagerApp(root)
    root.mainloop()
