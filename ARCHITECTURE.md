# Excel Session Manager v28 程式碼架構

## 🌳 程式碼結構樹狀圖

```
excel_session_manager_v29/
├── excel_session_manager.py           [~480行] 主程式 (保留向後相容)
├── main_app.py                         [~40行] 新的主程式入口點
├── excel_session_manager_link_updater.py [外部] 外部連結更新器
├── ARCHITECTURE.md                     [文檔] 程式碼架構說明
├── RESTRUCTURE_SUMMARY.md              [文檔] 重構總結
│
├── config/                             [配置系統]
│   ├── __init__.py                     [~1行]
│   ├── constants.py                    [~80行] 應用程式常數
│   ├── settings.py                     [~180行] 設定管理器
│   └── settings.yaml                   [~350行] YAML配置文件 (含註釋)
│
├── core/                               [核心功能]
│   ├── __init__.py                     [~1行]
│   ├── session_manager.py              [~150行] Session 管理
│   ├── excel_manager.py                [~200行] Excel COM 操作
│   ├── process_manager.py              [~180行] 進程管理
│   ├── error_handler.py                [~300行] 錯誤處理系統 ⭐ 新增
│   └── performance_monitor.py          [~400行] 性能監控系統 ⭐ 新增
│
├── ui/                                 [用戶介面]
│   ├── __init__.py                     [~1行]
│   ├── main_window.py                  [~450行] 主視窗類別 ⭐ 新增
│   ├── console_popup.py                [~60行] 控制台彈窗
│   ├── components/
│   │   ├── __init__.py                 [~1行]
│   │   └── drag_treeview.py            [~50行] 拖拽選擇TreeView
│   └── dialogs/
│       ├── __init__.py                 [~1行]
│       ├── link_options.py             [~120行] 外部連結選項對話框
│       └── file_selector.py            [~200行] 檔案選擇對話框
│
└── utils/                              [工具函數]
    ├── __init__.py                     [~1行]
    ├── file_utils.py                   [~30行] 檔案操作工具
    ├── ui_utils.py                     [~15行] UI 計算工具
    ├── settings_utils.py               [~80行] 設定管理工具
    └── window_utils.py                 [~120行] 視窗操作工具
```

## 📊 程式碼行數統計 (不含註釋)

### 主要檔案：
- **excel_session_manager.py**: ~480 行 (原1068行，減少588行)

### 配置系統：
- **config/constants.py**: ~80 行
- **config/settings.py**: ~180 行
- **config/settings.yaml**: ~350 行 (含詳細註釋)

### 核心功能：
- **core/session_manager.py**: ~150 行
- **core/excel_manager.py**: ~200 行
- **core/process_manager.py**: ~180 行

### 用戶介面：
- **ui/main_window.py**: ~450 行 (新增主視窗類別)
- **ui/console_popup.py**: ~60 行
- **ui/components/drag_treeview.py**: ~50 行
- **ui/dialogs/link_options.py**: ~120 行
- **ui/dialogs/file_selector.py**: ~200 行

### 工具函數：
- **utils/file_utils.py**: ~30 行
- **utils/ui_utils.py**: ~15 行
- **utils/settings_utils.py**: ~80 行
- **utils/window_utils.py**: ~120 行

### 總計：
- **原始程式碼**: 1068 行 (單一檔案)
- **重構後總行數**: ~2,800 行 (分散在多個檔案)
- **主檔案減少**: 588 行 (55% 減少)
- **模組化程式碼**: ~2,320 行 (新增的模組化程式碼)
- **新主程式**: ~40 行 (main_app.py)

## 🎯 重構成果

### ✅ 已達成目標：
1. **關注點分離** - 每個模組負責特定功能
2. **可維護性** - 程式碼結構清晰，易於修改
3. **可測試性** - 各模組可獨立測試
4. **可擴展性** - 新功能容易添加
5. **配置化** - 支援 YAML 配置文件自訂
6. **模組重用** - 共用邏輯可被多處使用

### 🔧 新增功能：
1. **進程管理** - Excel 進程健康監控和清理
2. **設定系統** - 完整的 YAML 配置支援
3. **時間戳驗證** - 檔案儲存前後時間比較
4. **詳細日誌** - 操作過程完整記錄

## 📋 模組依賴關係

```
excel_session_manager.py
├── config/settings.py
├── config/constants.py
├── core/session_manager.py
├── core/excel_manager.py
├── core/process_manager.py
├── ui/console_popup.py
├── ui/components/drag_treeview.py
├── ui/dialogs/link_options.py
├── ui/dialogs/file_selector.py
├── utils/file_utils.py
├── utils/ui_utils.py
└── excel_session_manager_link_updater.py
```

## 🚀 未來可能的重構方向：

1. **測試框架** - 添加單元測試
2. **日誌系統** - 統一的日誌管理
3. **插件系統** - 支援第三方擴展
4. **多語言支援** - 國際化功能
5. **主題系統** - 多種 UI 主題
6. **性能優化** - 異步操作和快取

## 📝 更新記錄

- **v28.0** - 完成主要重構，模組化程式碼結構
- **v28.1** - 移動檔案選擇器到 ui/dialogs/ 目錄，重命名為 file_selector.py
- **v28.2** - 調整視窗大小以容納新按鈕
- **v28.3** - 更新所有相關 import 語句

---
*此文件會隨著程式碼結構變更自動更新*