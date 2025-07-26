"""
Excel Manager for Excel Session Manager

This module handles all Excel COM operations including getting workbook information,
saving, closing, and activating Excel files.
"""

import pythoncom
import win32com.client
import win32gui
import win32con
import time
import gc
import psutil
from datetime import datetime
from utils.file_utils import get_file_mtime_str
from core.error_handler import handle_error, ErrorCategory, ErrorSeverity, safe_execute
from core.performance_monitor import timed_operation


class ExcelManager:
    """
    Manages Excel COM operations and workbook interactions.
    
    Provides methods for getting workbook information, saving files,
    closing files, and managing Excel application state.
    """
    
    def __init__(self):
        """Initialize the Excel manager."""
        pass
    
    @timed_operation("get_open_workbooks")
    def get_open_workbooks(self):
        """
        Get information about currently open Excel workbooks.
        
        Returns:
            tuple: (file_list, sheet_list, cell_list, path_list) containing
                   workbook names, active sheets, selected cells, and file paths
        """
        def _get_workbooks():
            pythoncom.CoInitialize()
            file_list, sheet_list, cell_list, path_list = [], [], [], []
            
            try:
                excel = win32com.client.Dispatch("Excel.Application")
                for wb in excel.Workbooks:
                    try:
                        file_list.append(wb.Name)
                        path_list.append(wb.FullName)
                        active_sheet = wb.ActiveSheet.Name if wb.ActiveSheet else ""
                        sheet_list.append(active_sheet)
                        try:
                            active_cell = wb.Application.ActiveCell.Address if wb.Application.ActiveCell else ""
                        except Exception:
                            active_cell = ""
                        cell_list.append(active_cell)
                    except Exception as e:
                        handle_error(e, ErrorSeverity.WARNING, ErrorCategory.EXCEL_COM, 
                                   f"Error accessing workbook: {wb.Name if 'wb' in locals() else 'unknown'}", 
                                   show_user=False)
                        file_list.append("Error")
                        path_list.append("Error")
                        sheet_list.append("Error")
                        cell_list.append("Error")
            except Exception as e:
                handle_error(e, ErrorSeverity.ERROR, ErrorCategory.EXCEL_COM, 
                           "Error connecting to Excel application", show_user=False)
            finally:
                pythoncom.CoUninitialize()
                
            return file_list, sheet_list, cell_list, path_list
        
        return safe_execute(_get_workbooks, 
                          context="Getting open Excel workbooks",
                          category=ErrorCategory.EXCEL_COM,
                          default_return=([], [], [], []))
    
    def save_workbooks(self, selected_workbooks, print_func=None):
        """
        Save selected Excel workbooks with before/after timestamp comparison.
        
        Args:
            selected_workbooks: List of tuples (name, path, sheet, cell)
            print_func: Optional function to print progress messages
        """
        if not selected_workbooks:
            return
        
        # Add performance monitoring with workbook count
        from core.performance_monitor import get_performance_monitor
        monitor = get_performance_monitor()
        op_id = monitor.start_operation("save_workbooks", {'workbook_count': len(selected_workbooks)})
        
        try:
            self._save_workbooks_impl(selected_workbooks, print_func)
            monitor.end_operation(op_id, success=True)
        except Exception as e:
            monitor.end_operation(op_id, success=False)
            raise
    
    def _save_workbooks_impl(self, selected_workbooks, print_func=None):
        """Implementation of save workbooks."""
            
        def print_msg(msg):
            if print_func:
                print_func(msg)
            else:
                print(msg)
                
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            print_msg(f"Saving {len(selected_workbooks)} selected file(s)")
            print_msg("-" * 80)
            
            for idx, (name, path, sheet, cell) in enumerate(selected_workbooks, 1):
                print_msg(f"({idx}/{len(selected_workbooks)}) Saving: {name}")
                
                # Get file modification time before save
                mtime_before = get_file_mtime_str(path) if path else "Unknown"
                print_msg(f"  File last modified before save: {mtime_before}")
                
                t0 = time.time()
                
                try:
                    wb = None
                    for workbook in excel.Workbooks:
                        if workbook.Name == name:
                            wb = workbook
                            break
                    
                    if wb:
                        # Save with retry logic
                        max_retries = 3
                        for attempt in range(max_retries):
                            try:
                                wb.Save()
                                break
                            except Exception as e:
                                if attempt < max_retries - 1:
                                    print_msg(f"  Save attempt {attempt + 1} failed, retrying...")
                                    time.sleep(0.1)
                                else:
                                    raise e
                        
                        # Get file modification time after save
                        mtime_after = get_file_mtime_str(path) if path else "Unknown"
                        print_msg(f"  File last modified after save: {mtime_after}")
                        
                        if mtime_before != mtime_after:
                            print_msg(f"  ({idx}/{len(selected_workbooks)}) Saved: {name} [SUCCESS] (File timestamp updated)")
                        else:
                            print_msg(f"  ({idx}/{len(selected_workbooks)}) Saved: {name} [WARNING] (File timestamp unchanged)")
                    else:
                        print_msg(f"  Workbook '{name}' not found in open workbooks")
                        
                except Exception as e:
                    print_msg(f"  ({idx}/{len(selected_workbooks)}) Failed to save: {name} ({e})")
                
                t1 = time.time()
                used_sec = t1 - t0
                print_msg(f"used time: {used_sec:.2f} sec")
                print_msg("-" * 80)
                
            print_msg(f"Save operation completed. Total: {len(selected_workbooks)}")
            
        except Exception as e:
            print_msg(f"Error during save operation: {str(e)}")
        finally:
            gc.collect()
            pythoncom.CoUninitialize()
    
    def save_and_close_workbooks(self, selected_workbooks, print_func=None):
        """
        Save and close selected Excel workbooks with before/after timestamp comparison.
        
        Args:
            selected_workbooks: List of tuples (name, path, sheet, cell)
            print_func: Optional function to print progress messages
        """
        if not selected_workbooks:
            return
        
        # Add performance monitoring with workbook count
        from core.performance_monitor import get_performance_monitor
        monitor = get_performance_monitor()
        op_id = monitor.start_operation("save_and_close_workbooks", {'workbook_count': len(selected_workbooks)})
        
        try:
            self._save_and_close_impl(selected_workbooks, print_func)
            monitor.end_operation(op_id, success=True)
        except Exception as e:
            monitor.end_operation(op_id, success=False)
            raise
    
    def _save_and_close_impl(self, selected_workbooks, print_func=None):
        """Implementation of save and close workbooks."""
        def print_msg(msg):
            if print_func:
                print_func(msg)
            else:
                print(msg)
                
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            print_msg(f"Saving and closing {len(selected_workbooks)} selected file(s)")
            print_msg("-" * 80)
            
            for idx, (name, path, sheet, cell) in enumerate(selected_workbooks, 1):
                print_msg(f"({idx}/{len(selected_workbooks)}) Saving and closing: {name}")
                
                # Get file modification time before save
                mtime_before = get_file_mtime_str(path) if path else "Unknown"
                print_msg(f"  File last modified before save: {mtime_before}")
                
                t0 = time.time()
                
                try:
                    wb = None
                    for workbook in excel.Workbooks:
                        if workbook.Name == name:
                            wb = workbook
                            break
                    
                    if wb:
                        # Save first
                        max_retries = 3
                        for attempt in range(max_retries):
                            try:
                                wb.Save()
                                break
                            except Exception as e:
                                if attempt < max_retries - 1:
                                    print_msg(f"  Save attempt {attempt + 1} failed, retrying...")
                                    time.sleep(0.1)
                                else:
                                    raise e
                        
                        # Get file modification time after save
                        mtime_after = get_file_mtime_str(path) if path else "Unknown"
                        print_msg(f"  File last modified after save: {mtime_after}")
                        
                        # Then close
                        wb.Close(SaveChanges=False)  # Already saved above
                        
                        if mtime_before != mtime_after:
                            print_msg(f"  ({idx}/{len(selected_workbooks)}) Saved and closed: {name} [SUCCESS] (File timestamp updated)")
                        else:
                            print_msg(f"  ({idx}/{len(selected_workbooks)}) Saved and closed: {name} [WARNING] (File timestamp unchanged)")
                    else:
                        print_msg(f"  Workbook '{name}' not found in open workbooks")
                        
                except Exception as e:
                    print_msg(f"  ({idx}/{len(selected_workbooks)}) Failed to save/close: {name} ({e})")
                
                t1 = time.time()
                used_sec = t1 - t0
                print_msg(f"used time: {used_sec:.2f} sec")
                print_msg("-" * 80)
                
            print_msg(f"Save and close operation completed. Total: {len(selected_workbooks)}")
            
        except Exception as e:
            print_msg(f"Error during save and close operation: {str(e)}")
        finally:
            gc.collect()
            pythoncom.CoUninitialize()
    
    def activate_workbooks(self, selected_workbooks):
        """
        Activate (bring to front) selected Excel workbooks.
        
        Args:
            selected_workbooks: List of tuples (name, path, sheet, cell)
        """
        if not selected_workbooks:
            return
        
        # Add performance monitoring with workbook count
        from core.performance_monitor import get_performance_monitor
        monitor = get_performance_monitor()
        op_id = monitor.start_operation("activate_workbooks", {'workbook_count': len(selected_workbooks)})
        
        try:
            self._activate_workbooks_impl(selected_workbooks)
            monitor.end_operation(op_id, success=True)
        except Exception as e:
            monitor.end_operation(op_id, success=False)
            raise
    
    def _activate_workbooks_impl(self, selected_workbooks):
        """Implementation of activate workbooks."""
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            
            for name, path, sheet, cell in selected_workbooks:
                try:
                    wb = None
                    for workbook in excel.Workbooks:
                        if workbook.Name == name:
                            wb = workbook
                            break
                    
                    if wb:
                        wb.Activate()
                        if sheet:
                            try:
                                sht = wb.Sheets(sheet)
                                sht.Activate()
                                if cell:
                                    sht.Range(cell).Select()
                            except Exception as e:
                                print(f"Error selecting sheet/cell: {e}")
                                
                        # Bring Excel window to front
                        try:
                            hwnd = excel.Hwnd
                            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                            win32gui.SetForegroundWindow(hwnd)
                        except Exception as e:
                            print(f"Error bringing window to front: {e}")
                            
                except Exception as e:
                    print(f"Error activating workbook '{name}': {e}")
                    
        except Exception as e:
            print(f"Error during activate operation: {str(e)}")
        finally:
            pythoncom.CoUninitialize()
    
    @timed_operation("minimize_all_excel")
    def minimize_all_excel(self):
        """Minimize all Excel application windows."""
        pythoncom.CoInitialize()
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            try:
                hwnd = excel.Hwnd
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            except Exception as e:
                print(f"Error minimizing Excel: {e}")
        except Exception as e:
            print(f"Error connecting to Excel for minimize: {e}")
        finally:
            pythoncom.CoUninitialize()
    
    def get_workbook_details(self, file_list, sheet_list, cell_list, path_list):
        """
        Get detailed information about workbooks including modification times.
        
        Args:
            file_list: List of workbook names
            sheet_list: List of active sheet names
            cell_list: List of selected cell addresses
            path_list: List of file paths
            
        Returns:
            list: List of tuples (name, path, sheet, cell, mtime_str)
        """
        details = []
        
        for i in range(len(file_list)):
            name = file_list[i] if i < len(file_list) else ""
            path = path_list[i] if i < len(path_list) else ""
            sheet = sheet_list[i] if i < len(sheet_list) else ""
            cell = cell_list[i] if i < len(cell_list) else ""
            
            # Get modification time
            mtime_str = get_file_mtime_str(path) if path else ""
            
            details.append((name, path, sheet, cell, mtime_str))
            
        return details