"""
External Links Manager Core Module

This module handles the core logic for managing external links in Excel workbooks,
including scanning, searching, and data processing functionality.
"""

import pythoncom
import win32com.client
import re
from typing import Dict, List, Tuple, Set
from dataclasses import dataclass
from collections import defaultdict


@dataclass
class ExternalLink:
    """Data class for external link information."""
    source_workbook: str
    source_sheet: str
    source_cell: str
    target_file: str
    formula: str
    link_type: str  # 'LinkSource' or 'Formula'


@dataclass
class ExternalFileGroup:
    """Data class for grouping links by external file."""
    external_file: str
    links: List[ExternalLink]
    reference_count: int


class ExternalLinksManager:
    """
    Core manager for external links functionality.
    
    Handles scanning Excel workbooks, processing external links,
    and providing search and grouping capabilities.
    """
    
    def __init__(self):
        """Initialize the external links manager."""
        self.external_links: List[ExternalLink] = []
        self.grouped_links: Dict[str, ExternalFileGroup] = {}
        
    def scan_open_workbooks(self) -> Tuple[List[ExternalLink], Dict[str, int]]:
        """
        Scan all open Excel workbooks for external links.
        
        Returns:
            Tuple of (external_links_list, statistics_dict)
        """
        pythoncom.CoInitialize()
        self.external_links = []
        statistics = {
            'total_workbooks': 0,
            'workbooks_with_links': 0,
            'total_links': 0,
            'unique_external_files': 0
        }
        
        try:
            excel = win32com.client.Dispatch("Excel.Application")
            workbooks_with_links = set()
            external_files = set()
            
            for workbook in excel.Workbooks:
                statistics['total_workbooks'] += 1
                workbook_name = workbook.Name
                workbook_links = []
                
                # Method 1: Use Excel's LinkSources
                try:
                    link_sources = workbook.LinkSources(1)  # xlExcelLinks = 1
                    if link_sources:
                        for link_source in link_sources:
                            external_file = self._extract_filename_from_path(link_source)
                            link = ExternalLink(
                                source_workbook=workbook_name,
                                source_sheet='',
                                source_cell='',
                                target_file=external_file,
                                formula=f'LinkSource: {link_source}',
                                link_type='LinkSource'
                            )
                            workbook_links.append(link)
                            external_files.add(external_file)
                except:
                    pass
                
                # Method 2: Scan formulas for external references
                try:
                    for worksheet in workbook.Worksheets:
                        sheet_name = worksheet.Name
                        
                        # Get used range to avoid scanning empty cells
                        used_range = worksheet.UsedRange
                        if used_range:
                            for cell in used_range.Cells:
                                try:
                                    if cell.HasFormula:
                                        formula = cell.Formula
                                        # Check for external references
                                        if self._has_external_reference(formula):
                                            external_files_in_formula = self._extract_external_files_from_formula(formula)
                                            for ext_file in external_files_in_formula:
                                                # Check for duplicates
                                                if not self._is_duplicate_link(workbook_links, sheet_name, cell.Address, ext_file):
                                                    link = ExternalLink(
                                                        source_workbook=workbook_name,
                                                        source_sheet=sheet_name,
                                                        source_cell=cell.Address,
                                                        target_file=ext_file,
                                                        formula=formula,
                                                        link_type='Formula'
                                                    )
                                                    workbook_links.append(link)
                                                    external_files.add(ext_file)
                                except:
                                    continue
                except:
                    continue
                
                # Add workbook links to main list
                if workbook_links:
                    workbooks_with_links.add(workbook_name)
                    self.external_links.extend(workbook_links)
                
            # Update statistics
            statistics['workbooks_with_links'] = len(workbooks_with_links)
            statistics['total_links'] = len(self.external_links)
            statistics['unique_external_files'] = len(external_files)
            
        except Exception as e:
            print(f"Error scanning external links: {e}")
        finally:
            pythoncom.CoUninitialize()
        
        # Group links by external file
        self._group_links_by_external_file()
        
        return self.external_links, statistics
    
    def _extract_filename_from_path(self, file_path: str) -> str:
        """Extract filename from full path."""
        if '\\' in file_path or '/' in file_path:
            return file_path.split('\\')[-1].split('/')[-1]
        return file_path
    
    def _has_external_reference(self, formula: str) -> bool:
        """Check if formula contains external references."""
        return '[' in formula and ']' in formula and any(ext in formula.lower() for ext in ['.xlsx', '.xls', '.xlsm'])
    
    def _extract_external_files_from_formula(self, formula: str) -> List[str]:
        """Extract external file names from a formula."""
        external_files = []
        # Pattern to match [filename.xlsx] or [path\filename.xlsx]
        pattern = r'\[([^\]]+\.xlsx?m?)\]'
        matches = re.findall(pattern, formula, re.IGNORECASE)
        
        for match in matches:
            filename = self._extract_filename_from_path(match)
            external_files.append(filename)
        
        return list(set(external_files))  # Remove duplicates
    
    def _is_duplicate_link(self, existing_links: List[ExternalLink], sheet: str, cell: str, target_file: str) -> bool:
        """Check if a link already exists to avoid duplicates."""
        for link in existing_links:
            if (link.source_sheet == sheet and 
                link.source_cell == cell and 
                link.target_file == target_file):
                return True
        return False
    
    def _group_links_by_external_file(self):
        """Group external links by target external file."""
        self.grouped_links = {}
        file_groups = defaultdict(list)
        
        # Group links by external file
        for link in self.external_links:
            file_groups[link.target_file].append(link)
        
        # Create ExternalFileGroup objects
        for external_file, links in file_groups.items():
            self.grouped_links[external_file] = ExternalFileGroup(
                external_file=external_file,
                links=links,
                reference_count=len(links)
            )
    
    def search_links(self, keyword: str, search_field: str = 'external_file') -> List[ExternalLink]:
        """
        Search external links by keyword in specified field.
        
        Args:
            keyword: Search keyword
            search_field: Field to search in ('external_file', 'formula', 'source_workbook', 'all')
            
        Returns:
            List of matching external links
        """
        if not keyword:
            return self.external_links
        
        keyword_lower = keyword.lower()
        results = []
        
        for link in self.external_links:
            match = False
            
            if search_field == 'external_file':
                match = keyword_lower in link.target_file.lower()
            elif search_field == 'formula':
                match = keyword_lower in link.formula.lower()
            elif search_field == 'source_workbook':
                match = keyword_lower in link.source_workbook.lower()
            elif search_field == 'all':
                match = (keyword_lower in link.target_file.lower() or
                        keyword_lower in link.formula.lower() or
                        keyword_lower in link.source_workbook.lower())
            
            if match:
                results.append(link)
        
        return results
    
    def get_grouped_search_results(self, keyword: str, search_field: str = 'external_file') -> Dict[str, ExternalFileGroup]:
        """
        Get search results grouped by external file.
        
        Args:
            keyword: Search keyword
            search_field: Field to search in
            
        Returns:
            Dictionary of external file groups matching the search
        """
        matching_links = self.search_links(keyword, search_field)
        
        # Group matching links by external file
        file_groups = defaultdict(list)
        for link in matching_links:
            file_groups[link.target_file].append(link)
        
        # Create grouped results
        grouped_results = {}
        for external_file, links in file_groups.items():
            grouped_results[external_file] = ExternalFileGroup(
                external_file=external_file,
                links=links,
                reference_count=len(links)
            )
        
        return grouped_results
    
    def get_statistics(self) -> Dict[str, int]:
        """Get statistics about external links."""
        unique_external_files = len(set(link.target_file for link in self.external_links))
        unique_source_workbooks = len(set(link.source_workbook for link in self.external_links))
        
        return {
            'total_links': len(self.external_links),
            'unique_external_files': unique_external_files,
            'unique_source_workbooks': unique_source_workbooks,
            'link_sources': len([link for link in self.external_links if link.link_type == 'LinkSource']),
            'formula_links': len([link for link in self.external_links if link.link_type == 'Formula'])
        }
    
    def export_to_dict(self) -> List[Dict]:
        """Export external links data to dictionary format for CSV export."""
        return [
            {
                'Source Workbook': link.source_workbook,
                'Source Sheet': link.source_sheet,
                'Source Cell': link.source_cell,
                'External File': link.target_file,
                'Formula': link.formula,
                'Link Type': link.link_type
            }
            for link in self.external_links
        ]