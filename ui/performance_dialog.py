"""
Performance Dialog for Excel Session Manager

This module provides a dialog for viewing performance metrics,
system status, and optimization recommendations.
"""

import tkinter as tk
from tkinter import ttk, scrolledtext
from datetime import datetime, timedelta
import threading
from core.performance_monitor import get_performance_monitor


class PerformanceDialog:
    """
    Dialog for displaying performance information and recommendations.
    
    Shows system metrics, operation statistics, and provides
    optimization suggestions to improve application performance.
    """
    
    def __init__(self, parent):
        """
        Initialize the performance dialog.
        
        Args:
            parent: Parent window
        """
        self.parent = parent
        self.monitor = get_performance_monitor()
        self.dialog = None
        self.auto_refresh = True
        self.refresh_interval = 5000  # 5 seconds
        
    def show(self):
        """Show the performance dialog."""
        if self.dialog:
            self.dialog.lift()
            return
            
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Performance Monitor")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        
        # Make dialog non-modal so you can use main window
        self.dialog.transient(self.parent)
        # Remove grab_set() to allow interaction with main window
        
        self._setup_ui()
        self._refresh_data()
        
        # Start auto-refresh
        if self.auto_refresh:
            self._schedule_refresh()
        
        # Handle close event
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_close)
        
        # Center the dialog
        self._center_dialog()
    
    def _setup_ui(self):
        """Setup the user interface."""
        # Main container
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill="both", expand=True)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill="both", expand=True)
        
        # System tab
        self.system_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.system_frame, text="System Status")
        self._setup_system_tab()
        
        # Operations tab
        self.operations_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.operations_frame, text="Operations")
        self._setup_operations_tab()
        
        # Recommendations tab
        self.recommendations_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.recommendations_frame, text="Recommendations")
        self._setup_recommendations_tab()
        
        # Control buttons
        self._setup_control_buttons(main_frame)
    
    def _setup_system_tab(self):
        """Setup the system status tab."""
        # System metrics frame
        metrics_frame = ttk.LabelFrame(self.system_frame, text="System Metrics", padding="10")
        metrics_frame.pack(fill="x", pady=(0, 10))
        
        # CPU usage
        self.cpu_label = ttk.Label(metrics_frame, text="CPU Usage: --", font=("Arial", 10))
        self.cpu_label.grid(row=0, column=0, sticky="w", padx=(0, 20))
        
        self.cpu_progress = ttk.Progressbar(metrics_frame, length=200, mode='determinate')
        self.cpu_progress.grid(row=0, column=1, sticky="w")
        
        # Memory usage
        self.memory_label = ttk.Label(metrics_frame, text="Memory Usage: --", font=("Arial", 10))
        self.memory_label.grid(row=1, column=0, sticky="w", padx=(0, 20), pady=(5, 0))
        
        self.memory_progress = ttk.Progressbar(metrics_frame, length=200, mode='determinate')
        self.memory_progress.grid(row=1, column=1, sticky="w", pady=(5, 0))
        
        # Memory details
        self.memory_details = ttk.Label(metrics_frame, text="", font=("Arial", 9), foreground="gray")
        self.memory_details.grid(row=2, column=0, columnspan=2, sticky="w", pady=(2, 0))
        
        # Excel processes
        self.excel_processes_label = ttk.Label(metrics_frame, text="Excel Processes: --", font=("Arial", 10))
        self.excel_processes_label.grid(row=3, column=0, columnspan=2, sticky="w", pady=(10, 0))
        
        # Performance status
        status_frame = ttk.LabelFrame(self.system_frame, text="Performance Status", padding="10")
        status_frame.pack(fill="both", expand=True)
        
        self.status_text = scrolledtext.ScrolledText(status_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
        self.status_text.pack(fill="both", expand=True)
    
    def _setup_operations_tab(self):
        """Setup the operations statistics tab."""
        # Operations summary
        summary_frame = ttk.LabelFrame(self.operations_frame, text="Operations Summary", padding="10")
        summary_frame.pack(fill="x", pady=(0, 10))
        
        # Create treeview for operations
        columns = ("Operation", "Count", "Success Rate", "Avg Duration", "Max Duration")
        self.operations_tree = ttk.Treeview(summary_frame, columns=columns, show="headings", height=8)
        
        for col in columns:
            self.operations_tree.heading(col, text=col)
            self.operations_tree.column(col, width=120)
        
        # Scrollbar for treeview
        tree_scroll = ttk.Scrollbar(summary_frame, orient="vertical", command=self.operations_tree.yview)
        self.operations_tree.configure(yscrollcommand=tree_scroll.set)
        
        self.operations_tree.pack(side="left", fill="both", expand=True)
        tree_scroll.pack(side="right", fill="y")
        
        # Recent operations
        recent_frame = ttk.LabelFrame(self.operations_frame, text="Recent Operations", padding="10")
        recent_frame.pack(fill="both", expand=True)
        
        self.recent_text = scrolledtext.ScrolledText(recent_frame, height=10, wrap=tk.WORD, font=("Consolas", 9))
        self.recent_text.pack(fill="both", expand=True)
    
    def _setup_recommendations_tab(self):
        """Setup the recommendations tab."""
        # Recommendations text
        self.recommendations_text = scrolledtext.ScrolledText(
            self.recommendations_frame, 
            wrap=tk.WORD, 
            font=("Arial", 10),
            padx=10,
            pady=10
        )
        self.recommendations_text.pack(fill="both", expand=True)
    
    def _setup_control_buttons(self, parent):
        """Setup control buttons."""
        button_frame = ttk.Frame(parent)
        button_frame.pack(fill="x", pady=(10, 0))
        
        # Refresh button
        refresh_btn = ttk.Button(button_frame, text="Refresh Now", command=self._refresh_data)
        refresh_btn.pack(side="left", padx=(0, 10))
        
        # Export button
        export_btn = ttk.Button(button_frame, text="Export Report", command=self._export_report)
        export_btn.pack(side="left", padx=(0, 10))
        
        # Auto-refresh toggle
        self.auto_refresh_var = tk.BooleanVar(value=self.auto_refresh)
        auto_refresh_cb = ttk.Checkbutton(
            button_frame, 
            text="Auto Refresh (5s)", 
            variable=self.auto_refresh_var,
            command=self._toggle_auto_refresh
        )
        auto_refresh_cb.pack(side="left", padx=(0, 10))
        
        # Close button
        close_btn = ttk.Button(button_frame, text="Close", command=self._on_close)
        close_btn.pack(side="right")
    
    def _refresh_data(self):
        """Refresh all performance data."""
        try:
            self._update_system_metrics()
            self._update_operations_data()
            self._update_recommendations()
        except Exception as e:
            print(f"Error refreshing performance data: {e}")
    
    def _update_system_metrics(self):
        """Update system metrics display."""
        summary = self.monitor.get_system_summary()
        
        if 'error' in summary:
            self.cpu_label.config(text=f"CPU Usage: Error - {summary['error']}")
            return
        
        # Update CPU
        cpu_current = summary['cpu']['current']
        cpu_status = summary['cpu']['status']
        self.cpu_label.config(text=f"CPU Usage: {cpu_current:.1f}% ({cpu_status})")
        self.cpu_progress['value'] = cpu_current
        
        # Set progress bar color based on status
        style = ttk.Style()
        if cpu_status == 'critical':
            style.configure("CPU.Horizontal.TProgressbar", background='red')
        elif cpu_status == 'warning':
            style.configure("CPU.Horizontal.TProgressbar", background='orange')
        else:
            style.configure("CPU.Horizontal.TProgressbar", background='green')
        
        # Update Memory
        memory_current = summary['memory']['current']
        memory_status = summary['memory']['status']
        memory_total = summary['memory']['total_gb']
        memory_available = summary['memory']['available_gb']
        
        self.memory_label.config(text=f"Memory Usage: {memory_current:.1f}% ({memory_status})")
        self.memory_progress['value'] = memory_current
        self.memory_details.config(text=f"Available: {memory_available:.1f}GB / Total: {memory_total:.1f}GB")
        
        # Excel processes
        excel_count = self._count_excel_processes()
        self.excel_processes_label.config(text=f"Excel Processes: {excel_count}")
        
        # Update status text
        status_lines = [
            f"Performance Status - {datetime.now().strftime('%H:%M:%S')}",
            "=" * 50,
            f"CPU: {cpu_current:.1f}% ({cpu_status})",
            f"Memory: {memory_current:.1f}% ({memory_status}) - {memory_available:.1f}GB available",
            f"Excel Processes: {excel_count}",
            ""
        ]
        
        if cpu_status != 'normal' or memory_status != 'normal':
            status_lines.append("⚠️ Performance Issues Detected:")
            if cpu_status != 'normal':
                status_lines.append(f"  • High CPU usage ({cpu_current:.1f}%)")
            if memory_status != 'normal':
                status_lines.append(f"  • High memory usage ({memory_current:.1f}%)")
            status_lines.append("")
        
        self.status_text.delete(1.0, tk.END)
        self.status_text.insert(1.0, "\n".join(status_lines))
    
    def _update_operations_data(self):
        """Update operations statistics."""
        # Clear existing data
        for item in self.operations_tree.get_children():
            self.operations_tree.delete(item)
        
        # Get operation statistics
        overall_stats = self.monitor.get_operation_statistics()
        
        # Get unique operations
        operations = set(t.operation for t in self.monitor.operation_timings)
        
        for operation in operations:
            stats = self.monitor.get_operation_statistics(operation)
            if stats['count'] > 0:
                success_rate = f"{stats['success_rate']:.1%}"
                avg_duration = f"{stats.get('avg_duration', 0):.2f}s"
                max_duration = f"{stats.get('max_duration', 0):.2f}s"
                
                self.operations_tree.insert("", "end", values=(
                    operation,
                    stats['count'],
                    success_rate,
                    avg_duration,
                    max_duration
                ))
        
        # Update recent operations
        recent_lines = ["Recent Operations:", "=" * 50]
        
        # Get last 15 operations
        recent_ops = list(self.monitor.operation_timings)[-15:]
        for op in reversed(recent_ops):
            if op.duration:
                status = "✓" if op.success else "✗"
                
                # Calculate timestamp from operation
                if hasattr(op, 'end_time') and op.end_time:
                    timestamp = datetime.fromtimestamp(op.end_time).strftime("%H:%M:%S")
                else:
                    timestamp = "Unknown"
                
                # Get workbook count from context
                workbook_count = ""
                if op.context and isinstance(op.context, dict):
                    if 'workbook_count' in op.context:
                        workbook_count = f" ({op.context['workbook_count']} files)"
                    elif 'selected_count' in op.context:
                        workbook_count = f" ({op.context['selected_count']} files)"
                
                # Format operation name for display
                display_name = op.operation.replace('_', ' ').title()
                
                recent_lines.append(
                    f"[{timestamp}] {status} {display_name}: {op.duration:.2f}s{workbook_count}"
                )
        
        self.recent_text.delete(1.0, tk.END)
        self.recent_text.insert(1.0, "\n".join(recent_lines))
    
    def _update_recommendations(self):
        """Update performance recommendations."""
        recommendations = self.monitor.get_performance_recommendations()
        
        content = [
            "Performance Optimization Recommendations",
            "=" * 50,
            f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
            ""
        ]
        
        for i, rec in enumerate(recommendations, 1):
            content.append(f"{i}. {rec}")
            content.append("")
        
        self.recommendations_text.delete(1.0, tk.END)
        self.recommendations_text.insert(1.0, "\n".join(content))
    
    def _count_excel_processes(self):
        """Count Excel processes."""
        try:
            import psutil
            count = 0
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    count += 1
            return count
        except Exception:
            return 0
    
    def _toggle_auto_refresh(self):
        """Toggle auto-refresh functionality."""
        self.auto_refresh = self.auto_refresh_var.get()
        if self.auto_refresh:
            self._schedule_refresh()
    
    def _schedule_refresh(self):
        """Schedule the next refresh."""
        if self.auto_refresh and self.dialog:
            self.dialog.after(self.refresh_interval, self._auto_refresh)
    
    def _auto_refresh(self):
        """Auto-refresh callback."""
        if self.auto_refresh and self.dialog:
            self._refresh_data()
            self._schedule_refresh()
    
    def _export_report(self):
        """Export performance report."""
        try:
            from tkinter import filedialog
            import json
            
            file_path = filedialog.asksaveasfilename(
                title="Export Performance Report",
                defaultextension=".json",
                filetypes=[("JSON files", "*.json"), ("All files", "*.*")]
            )
            
            if file_path:
                report = self.monitor.export_performance_report()
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(report, f, indent=2, default=str)
                
                tk.messagebox.showinfo("Export Complete", f"Performance report exported to:\n{file_path}")
        
        except Exception as e:
            tk.messagebox.showerror("Export Error", f"Failed to export report:\n{str(e)}")
    
    def _center_dialog(self):
        """Center the dialog on the parent window."""
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (self.dialog.winfo_width() // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (self.dialog.winfo_height() // 2)
        self.dialog.geometry(f"+{x}+{y}")
    
    def _on_close(self):
        """Handle dialog close event."""
        self.auto_refresh = False
        if self.dialog:
            self.dialog.destroy()
            self.dialog = None