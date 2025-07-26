# Excel Session Manager v29

A powerful tool for managing Excel workbook sessions with advanced features like process monitoring, external link updates, and session management.

## 🚀 Quick Start

### Installation
No installation required. Just ensure you have Python 3.7+ with the required dependencies.

### Running the Application

**Option 1: New Modular Entry Point (Recommended)**
```bash
python main_app.py
```

**Option 2: Legacy Entry Point (Backward Compatible)**
```bash
python excel_session_manager.py
```

### Testing the Installation
```bash
python test_imports.py
```

## 📋 Features

### Core Functionality
- **Session Management**: Save and load Excel workbook sessions
- **File Operations**: Save, close, and activate selected workbooks
- **External Link Updates**: Update external references in Excel files
- **Process Management**: Monitor and cleanup Excel processes

### Advanced Features
- **Time Stamp Verification**: Verify file saves with before/after timestamps
- **Mini Widget Mode**: Floating compact interface
- **Progress Console**: Real-time operation feedback
- **Drag Selection**: Multi-select files with mouse drag
- **Configurable Settings**: YAML-based configuration system

## 🏗️ Architecture

### Modular Structure
```
excel_session_manager_v29/
├── main_app.py                    # New entry point
├── excel_session_manager.py       # Legacy entry point
├── config/                        # Configuration system
├── core/                          # Core business logic
├── ui/                            # User interface components
└── utils/                         # Utility functions
```

### Key Components
- **MainWindow**: Primary user interface
- **ExcelManager**: Excel COM operations
- **SessionManager**: Session save/load functionality
- **ProcessManager**: Excel process monitoring
- **MiniWidget**: Floating mini interface

## ⚙️ Configuration

### Settings File
Edit `config/settings.yaml` to customize:
- Window size and appearance
- Font preferences
- Default directories
- Console behavior
- External link update options

### Example Configuration
```yaml
ui:
  window:
    default_size: "1200x750"
    title: "Excel Session Manager"
  fonts:
    default_family: "Consolas"
    default_size: 12
```

## 🔧 Usage

### Basic Operations
1. **Refresh List**: Update the list of open Excel files
2. **Save Selected**: Save selected workbooks
3. **Load Session**: Load a previously saved session
4. **Mini Mode**: Switch to floating mini widget

### Advanced Operations
1. **Update External Links**: Scan and update external references
2. **Cleanup Processes**: Monitor and cleanup Excel processes
3. **Session Management**: Save current state for later restoration

### Keyboard Shortcuts
- **Double-click**: Activate selected workbook
- **Drag selection**: Select multiple files
- **Select All**: Toggle all file selection

## 📊 Process Management

### Health Monitoring
The application can monitor Excel process health:
- Detect zombie processes
- Monitor memory usage
- Track process count
- Provide cleanup recommendations

### Automatic Cleanup
- Remove non-responsive processes
- Clean up orphaned Excel instances
- Optimize system performance

## 🔗 External Link Management

### Features
- Scan for external references
- Update links based on modification time
- Configurable time thresholds
- Detailed progress reporting
- Optional summary generation

### Configuration Options
- Days to check for modifications
- Display options (full path, details, status)
- Logging preferences
- Summary file generation

## 📁 Session Management

### Save Sessions
- Capture current Excel workbook state
- Include active sheet and cell information
- Timestamp-based file naming
- Excel-compatible format

### Load Sessions
- File selection dialog
- Validation of file existence
- Progress tracking
- Error handling and recovery

## 🎨 User Interface

### Main Window
- File list with modification times
- Action buttons panel
- Font customization
- Progress console toggle

### Mini Widget
- Floating compact interface
- Always-on-top option
- Icon or text display
- Quick restore functionality

### Console Output
- Real-time progress updates
- Detailed operation logs
- Dark theme interface
- Scrollable history

## 🛠️ Development

### Code Structure
- **Modular Design**: Separated concerns
- **Clean Architecture**: Clear dependencies
- **Configurable**: YAML-based settings
- **Extensible**: Plugin-ready structure

### Adding Features
1. Create new module in appropriate directory
2. Update imports in main components
3. Add configuration options if needed
4. Update documentation

### Testing
```bash
python test_imports.py  # Test module imports
python main_app.py      # Test full application
```

## 📚 Documentation

- `ARCHITECTURE.md`: Detailed code structure
- `REFACTOR_PROGRESS.md`: Development history
- `RESTRUCTURE_SUMMARY.md`: Refactoring summary

## 🔄 Version History

### v29.0 (Current)
- Complete modular architecture
- Mini widget functionality
- Process management features
- Enhanced configuration system
- Comprehensive documentation

### Previous Versions
- v28: Main window extraction
- v27: Process management addition
- v26: Settings system integration
- v25: Core functionality separation

## 🤝 Contributing

### Code Style
- Follow existing naming conventions
- Add docstrings to all functions
- Update documentation for new features
- Test all changes thoroughly

### File Organization
- Place new UI components in `ui/`
- Add business logic to `core/`
- Put utilities in `utils/`
- Update configuration in `config/`

## ⚠️ Requirements

### Python Dependencies
- `tkinter` (usually included with Python)
- `openpyxl` for Excel file handling
- `win32com.client` for Excel COM operations
- `psutil` for process management
- `PIL` (Pillow) for image handling
- `pyyaml` for configuration files

### System Requirements
- Windows OS (for Excel COM integration)
- Microsoft Excel installed
- Python 3.7 or higher

## 🐛 Troubleshooting

### Common Issues
1. **Import Errors**: Run `python test_imports.py` to diagnose
2. **Excel COM Errors**: Ensure Excel is properly installed
3. **Permission Issues**: Run as administrator if needed
4. **Process Cleanup**: Use the cleanup function for stuck processes

### Getting Help
1. Check the console output for detailed error messages
2. Review the log files in the configured log directory
3. Use the process cleanup feature for Excel-related issues
4. Verify configuration settings in `settings.yaml`

---

**Excel Session Manager v29** - A modern, modular approach to Excel session management.