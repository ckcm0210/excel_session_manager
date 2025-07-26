"""
Drag-select TreeView component for Excel Session Manager

This module contains a custom TreeView widget that supports
drag selection functionality for selecting multiple items.
"""

import tkinter as tk
from tkinter import ttk


class DragSelectTreeview(ttk.Treeview):
    """
    Custom TreeView with drag selection support.
    
    This TreeView allows users to click and drag to select multiple
    items, similar to file selection in Windows Explorer.
    """
    
    def __init__(self, *args, **kwargs):
        """Initialize the drag-select TreeView."""
        super().__init__(*args, **kwargs)
        self._drag_start = None
        self._drag_mode = None
        self.bind("<Button-1>", self._on_click)
        self.bind("<B1-Motion>", self._on_drag)
        self.bind("<ButtonRelease-1>", self._on_release)

    def _on_click(self, event):
        """Handle mouse click events for drag selection."""
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
        """Handle mouse drag events for range selection."""
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
        """Handle mouse release events to complete drag selection."""
        self._drag_start = None
        self._drag_mode = None
        self.event_generate("<<TreeviewSelect>>")
        return "break"