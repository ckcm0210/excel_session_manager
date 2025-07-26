"""
External Link Update Options Dialog for Excel Session Manager

This module contains the dialog for configuring external link update options.
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime, timedelta
import threading
from ui.console_popup import ConsolePopup
from excel_session_manager_link_updater import run_excel_link_update


class LinkUpdateOptionsDialog:
    """
    Dialog for configuring external link update options.
    
    Allows users to set parameters for updating external links in Excel files,
    including time filters, display options, and logging preferences.
    """
    
    def __init__(self, parent, last_log_dir=None, last_summary_dir=None):
        """
        Initialize the link update options dialog.
        
        Args:
            parent: Parent window
            last_log_dir: Last used log directory
            last_summary_dir: Last used summary directory
        """
        self.parent = parent
        self.last_log_dir = last_log_dir or r"D:\Pzone\Log"
        self.last_summary_dir = last_summary_dir or r"D:\Pzone\Log"
        
    def show(self):
        """Show the link update options dialog."""
        top = tk.Toplevel(self.parent)
        top.title("Update Recent External Links Options")
        top.geometry("600x380")
        top.grab_set()
    
        frm = tk.Frame(top, padx=14, pady=12)
        frm.pack(fill="both", expand=True)
    
        default_font = ("Arial", 10, "bold")
    
        # Days input
        tk.Label(frm, text="Only update links modified within (days):", anchor="w", font=default_font).grid(row=0, column=0, sticky="w")
        entry_days = tk.Entry(frm, width=6, font=default_font)
        entry_days.insert(0, "14")
        entry_days.grid(row=0, column=1, sticky="w", padx=(6,0))
    
        # Cutoff message
        cutoff_msg_var = tk.StringVar()
        def update_cutoff_msg(*args):
            try:
                days = int(entry_days.get())
                threshold_date = datetime.now() - timedelta(days=days)
                cutoff_msg_var.set(f"Any external link file with last modified time on or after {threshold_date.strftime('%d-%m-%y %H:%M:%S')} will be refreshed and updated.")
            except Exception:
                cutoff_msg_var.set("")
        entry_days.bind("<KeyRelease>", update_cutoff_msg)
        update_cutoff_msg()
        tk.Label(frm, textvariable=cutoff_msg_var, fg="#0055aa", anchor="w", wraplength=440, justify="left", font=default_font).grid(row=1, column=0, columnspan=3, sticky="w", pady=(1, 8))
    
        # Options checkboxes
        show_link_var = tk.BooleanVar(value=True)
        show_full_path_var = tk.BooleanVar(value=True)
        show_last_modified_var = tk.BooleanVar(value=True)
        show_status_var = tk.BooleanVar(value=True)
        save_log_var = tk.BooleanVar(value=True)
        save_scan_summary_var = tk.BooleanVar(value=False)
    
        def on_full_path_check(*args):
            if show_full_path_var.get():
                if not show_link_var.get():
                    show_link_var.set(True)
        show_full_path_var.trace_add("write", on_full_path_check)

        tk.Checkbutton(frm, text="Show external link file name", variable=show_link_var, font=default_font).grid(row=2, column=0, columnspan=3, sticky="w")
        tk.Checkbutton(frm, text="Show external link source full path", variable=show_full_path_var, font=default_font).grid(row=3, column=0, columnspan=3, sticky="w")
        tk.Checkbutton(frm, text="Show last modified date", variable=show_last_modified_var, font=default_font).grid(row=4, column=0, columnspan=3, sticky="w")
        tk.Checkbutton(frm, text="Show update status", variable=show_status_var, font=default_font).grid(row=5, column=0, columnspan=3, sticky="w")
        tk.Checkbutton(frm, text="Save log file", variable=save_log_var, font=default_font).grid(row=6, column=0, columnspan=3, sticky="w")
    
        # Log directory
        tk.Label(frm, text="Log file save to:", anchor="w", font=default_font).grid(row=7, column=0, sticky="w", pady=(6,0))
        entry_logdir = tk.Entry(frm, width=32, font=default_font)
        entry_logdir.insert(0, self.last_log_dir)
        entry_logdir.grid(row=7, column=1, sticky="w", pady=(6,0))
        def browse_logdir():
            d = filedialog.askdirectory(parent=top, initialdir=entry_logdir.get())
            if d:
                entry_logdir.delete(0, tk.END)
                entry_logdir.insert(0, d)
                self.last_log_dir = d
        browse_btn = tk.Button(frm, text="...", width=3, command=browse_logdir, font=default_font)
        browse_btn.grid(row=7, column=2, padx=(6,0), sticky="w", pady=(6,0))
    
        tk.Checkbutton(frm, text="Save external link scan summary", variable=save_scan_summary_var, font=default_font).grid(row=8, column=0, columnspan=3, sticky="w")
    
        # Summary directory
        tk.Label(frm, text="Scan summary save to:", anchor="w", font=default_font).grid(row=9, column=0, sticky="w", pady=(6,0))
        entry_summarydir = tk.Entry(frm, width=32, font=default_font)
        entry_summarydir.insert(0, self.last_summary_dir)
        entry_summarydir.grid(row=9, column=1, sticky="w", pady=(6,0))
        def browse_summarydir():
            d = filedialog.askdirectory(parent=top, initialdir=entry_summarydir.get())
            if d:
                entry_summarydir.delete(0, tk.END)
                entry_summarydir.insert(0, d)
                self.last_summary_dir = d
        browse_summary_btn = tk.Button(frm, text="...", width=3, command=browse_summarydir, font=default_font)
        browse_summary_btn.grid(row=9, column=2, padx=(6,0), sticky="w", pady=(6,0))
    
        # Buttons
        btn_frame = tk.Frame(frm)
        btn_frame.grid(row=99, column=0, columnspan=3, pady=(18,0), sticky="w")
        ok_btn = tk.Button(btn_frame, text="OK", width=12, font=default_font)
        cancel_btn = tk.Button(btn_frame, text="Cancel", width=12, command=top.destroy, font=default_font)
        ok_btn.pack(side=tk.LEFT, padx=(0,14))
        cancel_btn.pack(side=tk.LEFT)
        
        def on_confirm():
            log_dir = entry_logdir.get()
            summary_dir = entry_summarydir.get()
            if save_log_var.get() and not os.path.isdir(log_dir):
                messagebox.showerror("Error", "The selected log file directory does not exist. Please choose another directory.")
                return
            if save_scan_summary_var.get() and not os.path.isdir(summary_dir):
                messagebox.showerror("Error", "The selected scan summary directory does not exist. Please choose another directory.")
                return
            self.last_log_dir = log_dir
            self.last_summary_dir = summary_dir
            options = {
                "CHECK_DAYS": entry_days.get(),
                "SHOW_FULL_PATH": show_full_path_var.get(),
                "SHOW_LINK": show_link_var.get(),
                "SHOW_LAST_MODIFIED": show_last_modified_var.get(),
                "SHOW_STATUS": show_status_var.get(),
                "LOG_DIR": log_dir,
                "SAVE_LOG": save_log_var.get(),
                "SAVE_SCAN_SUMMARY": save_scan_summary_var.get(),
                "SUMMARY_DIR": summary_dir
            }
            popup = ConsolePopup(self.parent, title="Update Recent External Links Console")
    
            def print_to_popup(msg):
                self.parent.after(0, lambda: popup.print(msg))
            threading.Thread(target=lambda: run_excel_link_update(options, print_func=print_to_popup), daemon=True).start()
            top.destroy()
        ok_btn.config(command=on_confirm)
        
        return self.last_log_dir, self.last_summary_dir