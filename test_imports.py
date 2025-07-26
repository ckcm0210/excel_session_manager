#!/usr/bin/env python3
"""
Test script to verify all imports and basic functionality
"""

def test_imports():
    """Test all module imports"""
    print("Testing module imports...")
    
    try:
        # Test config imports
        from config.constants import MONO_FONTS, DEFAULT_WINDOW_SIZE
        from config.settings import settings
        print("✅ Config modules imported successfully")
        
        # Test core imports
        from core.excel_manager import ExcelManager
        from core.session_manager import SessionManager
        from core.process_manager import ProcessManager
        print("✅ Core modules imported successfully")
        
        # Test UI imports
        from ui.main_window import MainWindow
        from ui.console_popup import ConsolePopup
        from ui.mini_widget import MiniWidget
        from ui.components.drag_treeview import DragSelectTreeview
        from ui.dialogs.file_selector import FileSelectionDialog
        from ui.dialogs.link_options import LinkUpdateOptionsDialog
        print("✅ UI modules imported successfully")
        
        # Test utils imports
        from utils.file_utils import get_file_mtime_str, parse_mtime
        from utils.ui_utils import calc_row_height, calc_col2_width
        from utils.settings_utils import save_user_preferences
        from utils.window_utils import bring_window_to_front
        print("✅ Utils modules imported successfully")
        
        # Test settings
        print(f"✅ Settings loaded - Window size: {settings.window_size}")
        print(f"✅ Available fonts: {len(MONO_FONTS)} fonts")
        
        return True
        
    except ImportError as e:
        print(f"❌ Import error: {e}")
        return False
    except Exception as e:
        print(f"❌ Error: {e}")
        return False

def test_basic_functionality():
    """Test basic functionality without GUI"""
    print("\nTesting basic functionality...")
    
    try:
        from core.excel_manager import ExcelManager
        from core.process_manager import ProcessManager
        
        # Test ExcelManager
        excel_mgr = ExcelManager()
        print("✅ ExcelManager created successfully")
        
        # Test ProcessManager
        process_mgr = ProcessManager()
        excel_processes = process_mgr.get_excel_process_info()
        print(f"✅ ProcessManager working - Found {len(excel_processes)} Excel processes")
        
        # Test settings
        from config.settings import settings
        window_size = settings.window_size
        print(f"✅ Settings working - Window size: {window_size}")
        
        return True
        
    except Exception as e:
        print(f"❌ Functionality test failed: {e}")
        return False

if __name__ == "__main__":
    print("=" * 60)
    print("Excel Session Manager v29 - Module Test")
    print("=" * 60)
    
    imports_ok = test_imports()
    functionality_ok = test_basic_functionality()
    
    print("\n" + "=" * 60)
    if imports_ok and functionality_ok:
        print("🎉 All tests passed! The application is ready to use.")
        print("\nTo start the application:")
        print("  python main_app.py        # New modular entry point")
        print("  python excel_session_manager.py  # Legacy entry point")
    else:
        print("❌ Some tests failed. Please check the error messages above.")
    print("=" * 60)