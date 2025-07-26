"""
External Links Utilities

This module provides utility functions for external links management,
including Excel navigation, file operations, and data formatting.
"""

import pythoncom
import win32com.client
import win32gui
import win32con
from typing import Optional, Tuple
import os
import csv
from datetime import datetime


class ExcelNavigator:
    """Utility class for navigating to specific cells in Excel."""
    
    @staticmethod
    def navigate_to_cell(workbook_name: str, sheet_name: str, cell_address: str) -> Tuple[bool, str]:
        """
        Navigate to a specific cell in Excel.
        
        Args:
            workbook_name: Name of the workbook
            sheet_name: Name of the worksheet
            cell_address: Cell address (e.g., 'A1')
            
        Returns:
            Tuple of (success, message)
        """
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            
            # Find the workbook
            workbook = None
            for wb in excel.Workbooks:
                if wb.Name == workbook_name:
                    workbook = wb
                    break
            
            if not workbook:
                return False, f"Workbook '{workbook_name}' not found"
            
            if not sheet_name or not cell_address:
                # Just activate the workbook
                workbook.Activate()
                excel.Visible = True
                return True, f"Activated workbook '{workbook_name}'"
            
            try:
                # Activate the workbook
                workbook.Activate()
                
                # Activate the worksheet
                worksheet = workbook.Sheets(sheet_name)
                worksheet.Activate()
                
                # Select the cell
                worksheet.Range(cell_address).Select()
                
                # Bring Excel to front
                excel.Visible = True
                excel.WindowState = -4137  # xlNormal
                
                # Try to bring Excel window to foreground
                try:
                    hwnd = excel.Hwnd
                    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                    win32gui.SetForegroundWindow(hwnd)
                except:
                    pass
                
                return True, f"Navigated to {workbook_name} -> {sheet_name} -> {cell_address}"
                
            except Exception as e:
                return False, f"Failed to navigate to sheet/cell: {str(e)}"
                
        except Exception as e:
            return False, f"Excel navigation error: {str(e)}"
        finally:
            pythoncom.CoUninitialize()


class ExternalLinksExporter:
    """Utility class for exporting external links data."""
    
    @staticmethod
    def export_to_csv(data: list, file_path: str, include_timestamp: bool = True) -> Tuple[bool, str]:
        """
        Export external links data to CSV file.
        
        Args:
            data: List of dictionaries containing link data
            file_path: Output file path
            include_timestamp: Whether to include scan timestamp
            
        Returns:
            Tuple of (success, message)
        """
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                if not data:
                    return False, "No data to export"
                
                # Get fieldnames from first item
                fieldnames = list(data[0].keys())
                if include_timestamp:
                    fieldnames.append('Scan Time')
                
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                for row in data:
                    if include_timestamp:
                        row['Scan Time'] = scan_time
                    writer.writerow(row)
            
            return True, f"Data exported successfully to {file_path}"
            
        except Exception as e:
            return False, f"Export failed: {str(e)}"
    
    @staticmethod
    def export_grouped_to_csv(grouped_data: dict, file_path: str) -> Tuple[bool, str]:
        """
        Export grouped external links data to CSV file.
        
        Args:
            grouped_data: Dictionary of ExternalFileGroup objects
            file_path: Output file path
            
        Returns:
            Tuple of (success, message)
        """
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = [
                    'External File', 'Reference Count', 'Source Workbook', 
                    'Source Sheet', 'Source Cell', 'Formula', 'Link Type'
                ]
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                
                scan_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                for external_file, group in grouped_data.items():
                    for i, link in enumerate(group.links):
                        writer.writerow({
                            'External File': external_file if i == 0 else '',  # Only show filename on first row
                            'Reference Count': group.reference_count if i == 0 else '',
                            'Source Workbook': link.source_workbook,
                            'Source Sheet': link.source_sheet,
                            'Source Cell': link.source_cell,
                            'Formula': link.formula,
                            'Link Type': link.link_type
                        })
                    
                    # Add empty row between groups
                    if len(grouped_data) > 1:
                        writer.writerow({field: '' for field in fieldnames})
            
            return True, f"Grouped data exported successfully to {file_path}"
            
        except Exception as e:
            return False, f"Export failed: {str(e)}"


class FormulaAnalyzer:
    """Utility class for analyzing Excel formulas."""
    
    @staticmethod
    def extract_external_references(formula: str) -> list:
        """
        Extract all external references from a formula.
        
        Args:
            formula: Excel formula string
            
        Returns:
            List of external file references
        """
        import re
        
        # Pattern to match external references like [filename.xlsx]SheetName!Range
        pattern = r'\[([^\]]+\.xlsx?m?)\]([^!]*!)?([\w\$:]+)'
        matches = re.findall(pattern, formula, re.IGNORECASE)
        
        references = []
        for match in matches:
            filename = match[0]
            sheet = match[1].rstrip('!') if match[1] else ''
            range_ref = match[2] if match[2] else ''
            
            references.append({
                'filename': filename,
                'sheet': sheet,
                'range': range_ref,
                'full_reference': f"[{filename}]{sheet}!{range_ref}" if sheet else f"[{filename}]{range_ref}"
            })
        
        return references
    
    @staticmethod
    def is_broken_link(formula: str) -> bool:
        """
        Check if a formula contains broken external links.
        
        Args:
            formula: Excel formula string
            
        Returns:
            True if formula appears to contain broken links
        """
        # Common indicators of broken links
        broken_indicators = [
            '#REF!', '#NAME?', '#VALUE!', '#N/A',
            "'[", "']"  # Single quotes around external references often indicate broken links
        ]
        
        return any(indicator in formula for indicator in broken_indicators)
    
    @staticmethod
    def get_formula_complexity_score(formula: str) -> int:
        """
        Calculate a complexity score for a formula.
        
        Args:
            formula: Excel formula string
            
        Returns:
            Complexity score (higher = more complex)
        """
        score = 0
        
        # Count external references
        external_refs = len(re.findall(r'\[[^\]]+\]', formula))
        score += external_refs * 2
        
        # Count nested functions
        nested_functions = formula.count('(') - 1
        score += nested_functions
        
        # Count operators
        operators = ['+', '-', '*', '/', '&', '=', '<', '>']
        for op in operators:
            score += formula.count(op)
        
        # Length factor
        score += len(formula) // 50
        
        return score


class PathUtils:
    """Utility class for file path operations."""
    
    @staticmethod
    def normalize_path(file_path: str) -> str:
        """
        Normalize file path for consistent comparison.
        
        Args:
            file_path: File path to normalize
            
        Returns:
            Normalized file path
        """
        return os.path.normpath(file_path).lower()
    
    @staticmethod
    def extract_filename(file_path: str) -> str:
        """
        Extract filename from full path.
        
        Args:
            file_path: Full file path
            
        Returns:
            Filename only
        """
        return os.path.basename(file_path)
    
    @staticmethod
    def get_file_extension(filename: str) -> str:
        """
        Get file extension from filename.
        
        Args:
            filename: Filename
            
        Returns:
            File extension (including dot)
        """
        return os.path.splitext(filename)[1].lower()
    
    @staticmethod
    def is_excel_file(filename: str) -> bool:
        """
        Check if filename is an Excel file.
        
        Args:
            filename: Filename to check
            
        Returns:
            True if it's an Excel file
        """
        excel_extensions = ['.xlsx', '.xls', '.xlsm', '.xlsb']
        return PathUtils.get_file_extension(filename) in excel_extensions


class DataFormatter:
    """Utility class for formatting data for display."""
    
    @staticmethod
    def truncate_formula(formula: str, max_length: int = 100) -> str:
        """
        Truncate formula for display purposes.
        
        Args:
            formula: Formula string
            max_length: Maximum length
            
        Returns:
            Truncated formula with ellipsis if needed
        """
        if len(formula) <= max_length:
            return formula
        return formula[:max_length-3] + "..."
    
    @staticmethod
    def format_reference_count(count: int) -> str:
        """
        Format reference count for display.
        
        Args:
            count: Number of references
            
        Returns:
            Formatted count string
        """
        if count == 1:
            return "1 reference"
        return f"{count} references"
    
    @staticmethod
    def format_cell_address(address: str) -> str:
        """
        Format cell address for consistent display.
        
        Args:
            address: Cell address
            
        Returns:
            Formatted cell address
        """
        if not address:
            return ""
        
        # Remove $ signs for cleaner display
        return address.replace('$', '')
    
    @staticmethod
    def get_link_type_icon(link_type: str) -> str:
        """
        Get icon/symbol for link type.
        
        Args:
            link_type: Type of link ('LinkSource' or 'Formula')
            
        Returns:
            Icon string
        """
        icons = {
            'LinkSource': '🔗',
            'Formula': '📊'
        }
        return icons.get(link_type, '📄')