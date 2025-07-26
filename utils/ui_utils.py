"""
UI utility functions for Excel Session Manager

This module contains utility functions for UI calculations,
font management, and display formatting.
"""


def calc_row_height(fsize):
    """
    Calculate TreeView row height based on font size.
    
    Args:
        fsize: Font size in pixels
        
    Returns:
        Calculated row height in pixels
    """
    return int(fsize * 2.1)


def calc_col2_width(fsize):
    """
    Calculate second column width based on font size.
    
    Args:
        fsize: Font size in pixels
        
    Returns:
        Calculated column width in pixels
    """
    return int(fsize * 11.5)