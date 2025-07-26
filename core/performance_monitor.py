"""
Performance Monitor for Excel Session Manager

This module provides performance monitoring, metrics collection,
and optimization suggestions for the application.
"""

import time
import psutil
import threading
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Callable
from dataclasses import dataclass
from collections import defaultdict, deque


@dataclass
class PerformanceMetric:
    """Performance metric data structure."""
    name: str
    value: float
    unit: str
    timestamp: datetime
    category: str = "general"


@dataclass
class OperationTiming:
    """Operation timing data structure."""
    operation: str
    start_time: float
    end_time: Optional[float] = None
    duration: Optional[float] = None
    success: bool = True
    context: Dict = None


class PerformanceMonitor:
    """
    Performance monitoring and metrics collection system.
    
    Tracks application performance, resource usage, and provides
    optimization suggestions.
    """
    
    def __init__(self, max_history: int = 1000):
        """
        Initialize the performance monitor.
        
        Args:
            max_history: Maximum number of metrics to keep in history
        """
        self.max_history = max_history
        self.metrics_history: deque = deque(maxlen=max_history)
        self.operation_timings: deque = deque(maxlen=max_history)
        self.active_operations: Dict[str, OperationTiming] = {}
        self.performance_callbacks: List[Callable] = []
        
        # System monitoring
        self.system_metrics = {
            'cpu_percent': deque(maxlen=100),
            'memory_percent': deque(maxlen=100),
            'disk_usage': deque(maxlen=100)
        }
        
        # Performance thresholds
        self.thresholds = {
            'cpu_warning': 80.0,
            'cpu_critical': 95.0,
            'memory_warning': 80.0,
            'memory_critical': 95.0,
            'operation_slow': 5.0,  # seconds
            'operation_very_slow': 15.0  # seconds
        }
        
        # Start background monitoring
        self.monitoring_active = True
        self.monitor_thread = threading.Thread(target=self._background_monitor, daemon=True)
        self.monitor_thread.start()
    
    def _background_monitor(self):
        """Background thread for continuous system monitoring."""
        while self.monitoring_active:
            try:
                # Collect system metrics
                cpu_percent = psutil.cpu_percent(interval=1)
                memory_percent = psutil.virtual_memory().percent
                
                # Fix disk usage for Windows
                try:
                    import os
                    if os.name == 'nt':  # Windows
                        disk_usage = psutil.disk_usage('C:\\').percent
                    else:  # Linux/Mac
                        disk_usage = psutil.disk_usage('/').percent
                except:
                    disk_usage = 0
                
                # Store metrics
                self.system_metrics['cpu_percent'].append(cpu_percent)
                self.system_metrics['memory_percent'].append(memory_percent)
                self.system_metrics['disk_usage'].append(disk_usage)
                
                # Record as performance metrics
                timestamp = datetime.now()
                self.record_metric("cpu_usage", cpu_percent, "percent", "system", timestamp)
                self.record_metric("memory_usage", memory_percent, "percent", "system", timestamp)
                self.record_metric("disk_usage", disk_usage, "percent", "system", timestamp)
                
                # Check for performance issues
                self._check_performance_thresholds(cpu_percent, memory_percent)
                
                time.sleep(5)  # Monitor every 5 seconds
                
            except Exception as e:
                try:
                    print(f"Performance monitoring error: {str(e)}")
                except:
                    print("Performance monitoring error: (unable to display error message)")
                time.sleep(10)  # Wait longer on error
    
    def _check_performance_thresholds(self, cpu_percent: float, memory_percent: float):
        """Check if performance metrics exceed thresholds."""
        issues = []
        
        if cpu_percent >= self.thresholds['cpu_critical']:
            issues.append(f"Critical CPU usage: {cpu_percent:.1f}%")
        elif cpu_percent >= self.thresholds['cpu_warning']:
            issues.append(f"High CPU usage: {cpu_percent:.1f}%")
        
        if memory_percent >= self.thresholds['memory_critical']:
            issues.append(f"Critical memory usage: {memory_percent:.1f}%")
        elif memory_percent >= self.thresholds['memory_warning']:
            issues.append(f"High memory usage: {memory_percent:.1f}%")
        
        if issues:
            self._notify_performance_callbacks({
                'type': 'threshold_exceeded',
                'issues': issues,
                'timestamp': datetime.now(),
                'cpu_percent': cpu_percent,
                'memory_percent': memory_percent
            })
    
    def record_metric(
        self, 
        name: str, 
        value: float, 
        unit: str, 
        category: str = "general",
        timestamp: Optional[datetime] = None
    ):
        """Record a performance metric."""
        if timestamp is None:
            timestamp = datetime.now()
        
        metric = PerformanceMetric(name, value, unit, timestamp, category)
        self.metrics_history.append(metric)
    
    def start_operation(self, operation: str, context: Dict = None) -> str:
        """
        Start timing an operation.
        
        Args:
            operation: Name of the operation
            context: Additional context information
            
        Returns:
            Operation ID for later reference
        """
        operation_id = f"{operation}_{int(time.time() * 1000)}"
        timing = OperationTiming(
            operation=operation,
            start_time=time.time(),
            context=context or {}
        )
        self.active_operations[operation_id] = timing
        return operation_id
    
    def end_operation(self, operation_id: str, success: bool = True):
        """
        End timing an operation.
        
        Args:
            operation_id: ID returned by start_operation
            success: Whether the operation was successful
        """
        if operation_id in self.active_operations:
            timing = self.active_operations.pop(operation_id)
            timing.end_time = time.time()
            timing.duration = timing.end_time - timing.start_time
            timing.success = success
            
            self.operation_timings.append(timing)
            
            # Record as metric
            self.record_metric(
                f"operation_{timing.operation}",
                timing.duration,
                "seconds",
                "operations"
            )
            
            # Check for slow operations
            if timing.duration >= self.thresholds['operation_very_slow']:
                self._notify_performance_callbacks({
                    'type': 'very_slow_operation',
                    'operation': timing.operation,
                    'duration': timing.duration,
                    'context': timing.context
                })
            elif timing.duration >= self.thresholds['operation_slow']:
                self._notify_performance_callbacks({
                    'type': 'slow_operation',
                    'operation': timing.operation,
                    'duration': timing.duration,
                    'context': timing.context
                })
    
    def get_system_summary(self) -> Dict:
        """Get current system performance summary."""
        try:
            # Current values
            cpu_current = psutil.cpu_percent()
            memory = psutil.virtual_memory()
            disk = psutil.disk_usage('/')
            
            # Averages from recent history
            cpu_avg = sum(self.system_metrics['cpu_percent']) / len(self.system_metrics['cpu_percent']) if self.system_metrics['cpu_percent'] else 0
            memory_avg = sum(self.system_metrics['memory_percent']) / len(self.system_metrics['memory_percent']) if self.system_metrics['memory_percent'] else 0
            
            return {
                'cpu': {
                    'current': cpu_current,
                    'average': cpu_avg,
                    'status': self._get_status(cpu_current, 'cpu')
                },
                'memory': {
                    'current': memory.percent,
                    'average': memory_avg,
                    'total_gb': memory.total / (1024**3),
                    'available_gb': memory.available / (1024**3),
                    'status': self._get_status(memory.percent, 'memory')
                },
                'disk': {
                    'usage_percent': (disk.used / disk.total) * 100,
                    'free_gb': disk.free / (1024**3),
                    'total_gb': disk.total / (1024**3)
                }
            }
        except Exception as e:
            return {'error': str(e)}
    
    def _get_status(self, value: float, metric_type: str) -> str:
        """Get status based on threshold values."""
        if metric_type == 'cpu':
            if value >= self.thresholds['cpu_critical']:
                return 'critical'
            elif value >= self.thresholds['cpu_warning']:
                return 'warning'
        elif metric_type == 'memory':
            if value >= self.thresholds['memory_critical']:
                return 'critical'
            elif value >= self.thresholds['memory_warning']:
                return 'warning'
        return 'normal'
    
    def get_operation_statistics(self, operation: str = None) -> Dict:
        """Get statistics for operations."""
        if operation:
            # Filter by specific operation
            timings = [t for t in self.operation_timings if t.operation == operation]
        else:
            timings = list(self.operation_timings)
        
        if not timings:
            return {'count': 0}
        
        durations = [t.duration for t in timings if t.duration is not None]
        successful = [t for t in timings if t.success]
        
        if not durations:
            return {'count': len(timings), 'success_rate': len(successful) / len(timings)}
        
        return {
            'count': len(timings),
            'success_rate': len(successful) / len(timings),
            'avg_duration': sum(durations) / len(durations),
            'min_duration': min(durations),
            'max_duration': max(durations),
            'total_duration': sum(durations)
        }
    
    def get_performance_recommendations(self) -> List[str]:
        """Get performance optimization recommendations."""
        recommendations = []
        
        # Check system metrics
        system_summary = self.get_system_summary()
        
        if 'error' not in system_summary:
            cpu_status = system_summary['cpu']['status']
            memory_status = system_summary['memory']['status']
            
            if cpu_status in ['warning', 'critical']:
                recommendations.append(
                    f"High CPU usage detected ({system_summary['cpu']['current']:.1f}%). "
                    "Consider closing other applications or reducing concurrent operations."
                )
            
            if memory_status in ['warning', 'critical']:
                recommendations.append(
                    f"High memory usage detected ({system_summary['memory']['current']:.1f}%). "
                    "Consider restarting the application or closing unused Excel files."
                )
            
            if system_summary['memory']['available_gb'] < 1.0:
                recommendations.append(
                    "Low available memory (< 1GB). Consider closing other applications."
                )
        
        # Check operation performance
        slow_operations = [
            t for t in self.operation_timings 
            if t.duration and t.duration >= self.thresholds['operation_slow']
        ]
        
        if slow_operations:
            operation_counts = defaultdict(int)
            for op in slow_operations:
                operation_counts[op.operation] += 1
            
            for operation, count in operation_counts.items():
                if count >= 3:  # Multiple slow instances
                    recommendations.append(
                        f"Operation '{operation}' has been slow {count} times. "
                        "Consider optimizing this operation or checking system resources."
                    )
        
        # Excel-specific recommendations
        excel_processes = self._count_excel_processes()
        if excel_processes > 5:
            recommendations.append(
                f"Multiple Excel processes detected ({excel_processes}). "
                "Consider using the process cleanup feature."
            )
        
        if not recommendations:
            recommendations.append("System performance is optimal. No recommendations at this time.")
        
        return recommendations
    
    def _count_excel_processes(self) -> int:
        """Count running Excel processes."""
        try:
            count = 0
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    count += 1
            return count
        except Exception:
            return 0
    
    def register_performance_callback(self, callback: Callable):
        """Register a callback for performance events."""
        self.performance_callbacks.append(callback)
    
    def _notify_performance_callbacks(self, event_data: Dict):
        """Notify registered callbacks of performance events."""
        for callback in self.performance_callbacks:
            try:
                callback(event_data)
            except Exception as e:
                print(f"Error in performance callback: {e}")
    
    def get_metrics_by_category(self, category: str, hours: int = 1) -> List[PerformanceMetric]:
        """Get metrics by category within a time range."""
        cutoff_time = datetime.now() - timedelta(hours=hours)
        return [
            metric for metric in self.metrics_history
            if metric.category == category and metric.timestamp >= cutoff_time
        ]
    
    def export_performance_report(self) -> Dict:
        """Export a comprehensive performance report."""
        return {
            'timestamp': datetime.now().isoformat(),
            'system_summary': self.get_system_summary(),
            'operation_statistics': {
                'overall': self.get_operation_statistics(),
                'by_operation': {
                    op: self.get_operation_statistics(op)
                    for op in set(t.operation for t in self.operation_timings)
                }
            },
            'recommendations': self.get_performance_recommendations(),
            'metrics_count': len(self.metrics_history),
            'operations_count': len(self.operation_timings),
            'active_operations': len(self.active_operations)
        }
    
    def stop_monitoring(self):
        """Stop background monitoring."""
        self.monitoring_active = False
        if self.monitor_thread.is_alive():
            self.monitor_thread.join(timeout=2)


# Global performance monitor instance
_global_performance_monitor: Optional[PerformanceMonitor] = None


def get_performance_monitor() -> PerformanceMonitor:
    """Get the global performance monitor instance."""
    global _global_performance_monitor
    if _global_performance_monitor is None:
        _global_performance_monitor = PerformanceMonitor()
    return _global_performance_monitor


def timed_operation(operation_name: str):
    """Decorator for timing operations."""
    def decorator(func):
        def wrapper(*args, **kwargs):
            monitor = get_performance_monitor()
            op_id = monitor.start_operation(operation_name, {'function': func.__name__})
            try:
                result = func(*args, **kwargs)
                monitor.end_operation(op_id, success=True)
                return result
            except Exception as e:
                monitor.end_operation(op_id, success=False)
                raise
        return wrapper
    return decorator