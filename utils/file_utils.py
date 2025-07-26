"""
File utility functions for Excel Session Manager

This module contains utility functions for file operations,
time formatting, and file system interactions.
"""

import os
from datetime import datetime


def get_file_mtime_str(path):
    """
    Get file modification time as formatted string.
    
    Args:
        path: File path to check
        
    Returns:
        Formatted modification time string, or empty string if file doesn't exist
    """
    if os.path.exists(path):
        ts = os.path.getmtime(path)
        dt = datetime.fromtimestamp(ts)
        return f"{dt.year}-{dt.month:02d}-{dt.day:02d} {dt.hour:02d}:{dt.minute:02d}:{dt.second:02d}"
    else:
        return ''


def parse_mtime(mtime_str):
    """
    Parse modification time string back to datetime object.
    
    Args:
        mtime_str: Time string in format "dd/mm/yyyy hh:mm"
        
    Returns:
        datetime object or None if parsing fails
    """
    try:
        parts = mtime_str.split()
        date, time_part = parts[0], parts[1]
        day, month, year = map(int, date.split('/'))
        hour, minute = map(int, time_part.split(':'))
        return datetime(year, month, day, hour, minute)
    except Exception:
        return None