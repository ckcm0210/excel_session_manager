"""
Process Manager for Excel Session Manager

This module handles Excel process management including cleanup of zombie processes,
process monitoring, and Excel application lifecycle management.
"""

import psutil
import pythoncom
import win32com.client
import time
from datetime import datetime


class ProcessManager:
    """
    Manages Excel processes and application lifecycle.
    
    Provides methods for cleaning up zombie processes, monitoring Excel instances,
    and managing Excel application state.
    """
    
    def __init__(self):
        """Initialize the process manager."""
        pass
    
    def cleanup_zombie_excel_processes(self, print_func=None):
        """
        Clean up zombie Excel processes.
        
        Args:
            print_func: Optional function to print progress messages
        """
        def print_msg(msg):
            if print_func:
                print_func(msg)
            else:
                print(msg)
        
        print_msg("Scanning for zombie Excel processes...")
        cleaned_count = 0
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'status']):
                try:
                    if proc.info['name'] and 'excel' in proc.info['name'].lower():
                        # Check if process is responsive
                        if proc.info['status'] == psutil.STATUS_ZOMBIE:
                            try:
                                proc.terminate()
                                proc.wait(timeout=3)
                                print_msg(f"Terminated zombie Excel process: PID {proc.info['pid']}")
                                cleaned_count += 1
                            except (psutil.NoSuchProcess, psutil.TimeoutExpired):
                                try:
                                    proc.kill()
                                    print_msg(f"Force killed Excel process: PID {proc.info['pid']}")
                                    cleaned_count += 1
                                except Exception:
                                    pass
                        elif not proc.is_running():
                            try:
                                proc.terminate()
                                print_msg(f"Cleaned up non-running Excel process: PID {proc.info['pid']}")
                                cleaned_count += 1
                            except Exception:
                                pass
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                    
        except Exception as e:
            print_msg(f"Error during zombie process cleanup: {e}")
        
        if cleaned_count > 0:
            print_msg(f"Cleaned up {cleaned_count} zombie Excel process(es)")
        else:
            print_msg("No zombie Excel processes found")
    
    def get_excel_process_info(self):
        """
        Get information about running Excel processes.
        
        Returns:
            list: List of dictionaries containing process information
        """
        excel_processes = []
        
        try:
            for proc in psutil.process_iter(['pid', 'name', 'status', 'create_time', 'memory_info']):
                try:
                    if proc.info['name'] and 'excel' in proc.info['name'].lower():
                        create_time = datetime.fromtimestamp(proc.info['create_time'])
                        memory_mb = proc.info['memory_info'].rss / 1024 / 1024
                        
                        excel_processes.append({
                            'pid': proc.info['pid'],
                            'name': proc.info['name'],
                            'status': proc.info['status'],
                            'created': create_time.strftime("%Y-%m-%d %H:%M:%S"),
                            'memory_mb': round(memory_mb, 1)
                        })
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                    
        except Exception as e:
            print(f"Error getting Excel process info: {e}")
        
        return excel_processes
    
    def is_excel_running(self):
        """
        Check if any Excel processes are running.
        
        Returns:
            bool: True if Excel is running, False otherwise
        """
        try:
            for proc in psutil.process_iter(['name']):
                if proc.info['name'] and 'excel' in proc.info['name'].lower():
                    return True
        except Exception:
            pass
        return False
    
    def get_excel_com_connection(self):
        """
        Get Excel COM connection with proper initialization.
        
        Returns:
            Excel Application object or None if connection fails
        """
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            return excel
        except Exception as e:
            print(f"Failed to connect to Excel COM: {e}")
            return None
    
    def release_excel_com_connection(self):
        """Release Excel COM connection and cleanup."""
        try:
            pythoncom.CoUninitialize()
        except Exception as e:
            print(f"Error releasing COM connection: {e}")
    
    def force_close_all_excel(self, print_func=None):
        """
        Force close all Excel processes.
        
        Args:
            print_func: Optional function to print progress messages
        """
        def print_msg(msg):
            if print_func:
                print_func(msg)
            else:
                print(msg)
        
        print_msg("Force closing all Excel processes...")
        closed_count = 0
        
        try:
            # First try to close gracefully through COM
            try:
                pythoncom.CoInitialize()
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Quit()
                print_msg("Sent quit command to Excel application")
                time.sleep(2)  # Give Excel time to close gracefully
            except Exception:
                pass
            finally:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            
            # Then force close any remaining processes
            for proc in psutil.process_iter(['pid', 'name']):
                try:
                    if proc.info['name'] and 'excel' in proc.info['name'].lower():
                        proc.terminate()
                        try:
                            proc.wait(timeout=5)
                            print_msg(f"Terminated Excel process: PID {proc.info['pid']}")
                            closed_count += 1
                        except psutil.TimeoutExpired:
                            proc.kill()
                            print_msg(f"Force killed Excel process: PID {proc.info['pid']}")
                            closed_count += 1
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
                    
        except Exception as e:
            print_msg(f"Error during force close: {e}")
        
        if closed_count > 0:
            print_msg(f"Closed {closed_count} Excel process(es)")
        else:
            print_msg("No Excel processes were running")
    
    def monitor_excel_health(self, print_func=None):
        """
        Monitor Excel process health and report issues.
        
        Args:
            print_func: Optional function to print progress messages
            
        Returns:
            dict: Health report with process status and recommendations
        """
        def print_msg(msg):
            if print_func:
                print_func(msg)
            else:
                print(msg)
        
        health_report = {
            'total_processes': 0,
            'zombie_processes': 0,
            'high_memory_processes': 0,
            'recommendations': []
        }
        
        try:
            excel_processes = self.get_excel_process_info()
            health_report['total_processes'] = len(excel_processes)
            
            for proc_info in excel_processes:
                # Check for zombie processes
                if proc_info['status'] == psutil.STATUS_ZOMBIE:
                    health_report['zombie_processes'] += 1
                
                # Check for high memory usage (>500MB)
                if proc_info['memory_mb'] > 500:
                    health_report['high_memory_processes'] += 1
            
            # Generate recommendations
            if health_report['zombie_processes'] > 0:
                health_report['recommendations'].append(
                    f"Found {health_report['zombie_processes']} zombie process(es). Consider running cleanup."
                )
            
            if health_report['high_memory_processes'] > 0:
                health_report['recommendations'].append(
                    f"Found {health_report['high_memory_processes']} high-memory process(es). Consider restarting Excel."
                )
            
            if health_report['total_processes'] > 3:
                health_report['recommendations'].append(
                    f"Found {health_report['total_processes']} Excel processes. Consider closing unused instances."
                )
            
            if not health_report['recommendations']:
                health_report['recommendations'].append("Excel processes are healthy.")
            
            # Print report
            print_msg(f"Excel Health Report:")
            print_msg(f"  Total processes: {health_report['total_processes']}")
            print_msg(f"  Zombie processes: {health_report['zombie_processes']}")
            print_msg(f"  High memory processes: {health_report['high_memory_processes']}")
            print_msg(f"  Recommendations:")
            for rec in health_report['recommendations']:
                print_msg(f"    - {rec}")
                
        except Exception as e:
            print_msg(f"Error monitoring Excel health: {e}")
            health_report['recommendations'].append(f"Health check failed: {e}")
        
        return health_report