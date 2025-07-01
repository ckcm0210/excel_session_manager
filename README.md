# Excel 會話管理器 (Excel Session Manager)

一個基於 Python `tkinter` 和 `pywin32` 函式庫開發的 Windows GUI 工具，旨在提升處理多個 Microsoft Excel 檔案時的工作效率。此工具提供了一個中央化介面，用以監視、操作並儲存當前開啟的 Excel 工作階段。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

| 主介面                                                                                                            | 迷你懸浮窗模式                                                                                                           |
| ----------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| ![主介面截圖](https://raw.githubusercontent.com/ckcm0210/excel_session_manager/main/screenshots/main_ui.png)           | ![迷你模式截圖](https://raw.githubusercontent.com/ckcm0210/excel_session_manager/main/screenshots/mini_widget_mode.png) |
*(請將此處截圖路徑換成你自己的)*

---

## ✨ 核心功能

*   **實時會話掃描**：非同步偵測並列出所有當前開啟的 Excel 檔案實例。
*   **詳細資訊展示**：提供檔名、完整路徑及最後修改時間，並支援點擊標題進行排序。
*   **直觀列表操作**：透過繼承 `ttk.Treeview` 實現了拖曳多選功能，提升操作流暢度。
*   **批次視窗控制**：可將選定的一個或多個 Excel 視窗批量帶到最前、儲存或關閉。
*   **工作階段持久化**：
    *   **儲存會話**：將當前檔案列表（包含路徑、作用中工作表、選定儲存格）的狀態匯出為一個 `.xlsx` 會話檔。
    *   **載入會話**：從會話檔中讀取數據，批量重新開啟所有檔案並還原其工作狀態。
*   **非侵入式迷你模式**：可將主介面縮小為一個置頂的懸浮圖示，方便在不中斷工作的同時快速存取。
*   **介面個人化**：支援用戶自訂列表的字體和大小。

---

## 🛠️ 架構設計與技術剖析

本專案採用了事件驅動的架構模型，圍繞一個主應用程式類 `ExcelSessionManagerApp` 構建，該類負責管理 UI 狀態、事件處理和與後端邏輯的交互。

### 應用程式架構圖
```
ExcelSessionManagerApp (主類)
│
├── UI 狀態變數 (e.g., self.is_mini, self.showing_path)
│
├── setup_ui()                     # UI 初始化：負責所有元件的創建、佈局和事件綁定
│   ├── 主容器 (ttk.Notebook)
│   ├── 佈局框架 (tk.Frame)
│   └── 元件 (Widgets):
│       ├── DragSelectTreeview (自訂列表)
│       └── 標準按鈕、標籤等
│
├── 功能模組 (作為 Class Methods)
│   │
│   ├── Mini Widget 邏輯 (enter_mini, exit_mini)
│   │   └── 核心: pack_forget(), wm_attributes(), minsize()
│   │
│   ├── COM 互動邏輯 (get_open_excel_files, etc.)
│   │   └── 核心: pywin32, pythoncom, threading, root.after()
│   │
│   ├── 視窗控制邏輯 (activate_selected_workbooks, etc.)
│   │   └── 核心: win32gui, win32con
│   │
│   └── 會話檔案邏輯 (save_session, load_session)
│       └── 核心: openpyxl, filedialog
```

### 技術實現深度解析

1.  **GUI 介面 (`tkinter`)**
    *   **元件選擇**：主要採用 `tkinter.ttk` 的元件，因其外觀比傳統 `tkinter` 元件更貼近現代作業系統的原生風格。
    *   **自訂 `DragSelectTreeview`**：`ttk.Treeview` 本身不支援滑鼠拖曳進行多選。為此，我們創建了一個子類 `DragSelectTreeview`，繼承自 `ttk.Treeview`。透過重寫 `<Button-1>`、`<B1-Motion>` 和 `<ButtonRelease-1>` 的事件處理函數，我們得以攔截並解析滑鼠的拖曳行為，動態地增加或移除 `selection` 中的項目，從而實現了此核心功能。這是一個典型的利用物件導向編程（OOP）擴充現有元件功能的範例。
    *   **佈局管理**：專案混合使用了 `pack()` 和 `place()`。`pack()` 用於大部分的流式佈局，其相對定位非常適合構建可伸縮的介面。而 `place()` 則僅用於需要精確控制位置的場景，但在本專案中，我們最終選擇將所有元件都納入 `pack()` 的管理體系，以獲得更可預測和穩健的佈局行為。

2.  **Windows 自動化 (`pywin32`)**
    *   **`pywin32` 的雙重角色**：此函式庫在本專案中扮演了兩個關鍵但不同的角色：
        1.  **應用程式級互動 (`win32com.client`)**：透過 Component Object Model (COM) 介面與 Excel 應用程式本身進行「對話」。`win32com.client.GetActiveObject("Excel.Application")` 讓我們能獲取到 Excel 的主程式實例，進而可以存取其內部物件模型，如 `Workbooks` 集合、`ActiveSheet` 等，執行儲存、關閉等邏輯操作。
        2.  **作業系統級互動 (`win32gui`)**：當需要操作 Excel 的「視窗」而非其內部數據時，例如移動位置、置頂、最小化，COM 介面便無能為力。此時我們需要使用 `win32gui`，它能直接與 Windows 的視窗管理器溝通。`EnumWindows` 函數遍歷系統中所有頂層視窗，`GetWindowText` 用於標題匹配，而 `SetWindowPos` 則用於修改視窗的幾何屬性。
    *   **線程安全與 `pythoncom`**：COM 有「單元 (Apartment)」的概念，規定了跨線程存取 COM 物件的規則。`tkinter` GUI 運行在主線程，而我們的 Excel 掃描操作被放入一個背景 `threading.Thread` 中以防介面凍結。在這個背景線程中直接呼叫 `win32com` 會違反 COM 的線程規則。因此，必須在線程開始時呼叫 `pythoncom.CoInitialize()` 來為該線程初始化一個新的 COM 單元，並在結束時呼叫 `pythoncom.CoUninitialize()` 來清理。這是 `pywin32` 多線程編程的關鍵所在。

3.  **非同步操作與介面更新**
    *   `tkinter` 本身並非線程安全，任何對 GUI 元件的修改都必須在主線程中執行。當背景線程完成 Excel 掃描後，它不能直接去更新 `Treeview`。
    *   解決方案是使用 `root.after(0, callback)`。背景線程將 `update_gui` 函數作為一個 `callback` 傳遞給 `root.after`，`tkinter` 的主事件循環會在適當的時機（幾乎是立即）安全地在主線程中執行這個 `callback`，從而實現了線程安全的介面更新。

4.  **迷你懸浮窗的實現細節**
    *   **`pack_forget()` vs `destroy()`**：實現介面切換的關鍵在於使用 `pack_forget()` 而非 `destroy()`。前者僅將元件從佈局中移除，但元件本身及其狀態（如 `Treeview` 中的數據）依然存在於記憶體中，方便快速還原。後者則會徹底銷毀元件，需要昂貴的重建成本。
    *   **`minsize` 的陷阱與解決**：開發中最棘手的問題之一。即使 `geometry()` 已被設為小尺寸，視窗依然無法縮小。根本原因在於 `root.minsize(width, height)` 的限制優先級高於 `geometry()`。因此，正確的實現方式是：在進入迷你模式時，同步將 `minsize` 設為目標小尺寸；在退出時，再將其還原為原始的最小尺寸限制，以確保主介面的可用性。

---

## 🚀 如何使用

1.  **環境要求**
    *   Windows 作業系統
    *   Python 3.x
    *   已安裝 Microsoft Excel

2.  **安裝相依套件**
    ```bash
    pip install pywin32 openpyxl pillow
    ```

3.  **執行程式**
    *   如果你下載的是 `.py` 檔：
        ```bash
        python excel_session_manager_with_mini_widget_v4.py
        ```
    *   如果你使用的是 `.ipynb` 檔 (如 `excel_session_manager_v5.ipynb`)，請在 Jupyter 環境中打開並執行所有儲存格。

4.  **自訂懸浮窗圖示**
    *   在程式碼同一個目錄下，放置一張 `.png` 格式的圖片，並將其命名為 `maximize_full_screen.png`。
    *   若程式找不到此檔案，將會自動使用預設的 emoji「🗔」作為替代圖示。

---

## 🔮 未來發展方向

基於當前架構，本專案具有良好的擴展性。以下是一些潛在的發展方向：

1.  **增強的設定管理**
    *   **目標**：允許用戶保存個人化設定，如視窗大小、字體選擇、預設路徑等。
    *   **實現**：引入一個設定檔（如 `config.ini` 或 `settings.json`）。程式啟動時讀取，關閉時寫入。可使用 Python 內建的 `configparser` 或 `json` 函式庫。在 UI 中可增加一個「設定」分頁。

2.  **支援多個 Excel 實例 (Process)**
    *   **挑戰**：目前 `GetActiveObject` 通常只能獲取到用戶最後操作的那個 Excel 實例。如果用戶獨立地打開了多個 Excel 主程式，本工具可能無法全部偵測。
    *   **探索方向**：研究更底層的 COM 或 `psutil` 函式庫，遍歷系統中所有運行的 `EXCEL.EXE` 進程，並嘗試為每個進程單獨建立 COM 連接。這是一個更複雜但能顯著提升實用性的功能。

3.  **打包為獨立執行檔 (.exe)**
    *   **目標**：讓沒有安裝 Python 環境的用戶也能直接使用。
    *   **實現**：使用 `PyInstaller` 或 `cx_Freeze` 等工具。需要處理好圖檔、函式庫依賴等資源的打包路徑問題。

4.  **更健壯的錯誤處理與日誌系統**
    *   **目標**：當 COM 呼叫失敗或檔案路徑無效時，提供更友好的錯誤提示，並將詳細錯誤資訊記錄到日誌檔中，方便排查問題。
    *   **實現**：在所有 `try...except` 區塊中進行更細緻的異常分類。引入 Python 的 `logging` 模組，設定日誌等級、格式和輸出檔案。

5.  **UI/UX 細節優化**
    *   **搜尋/過濾功能**：當檔案列表過長時，在列表上方增加一個搜尋框，可以即時過濾檔名。
    *   **狀態欄**：在視窗底部增加一個狀態欄，用以顯示當前操作的進度（如「正在儲存 5 個檔案...」）或提示資訊。
    *   **滑鼠懸停提示 (Tooltips)**：為所有功能按鈕增加 Tooltips，解釋其具體作用。

---

## 🤝 如何貢獻

歡迎對本專案進行 Fork、提出 Issue 或提交 Pull Request。

1.  Fork 本儲存庫。
2.  創建您的功能分支 (`git checkout -b feature/YourAmazingFeature`)。
3.  Commit 您的變更 (`git commit -m 'Add some AmazingFeature'`)。
4.  將變更推送到分支 (`git push origin feature/YourAmazingFeature`)。
5.  開啟一個 Pull Request。

---

## 📄 授權協議 (License)

本專案採用 [MIT License](https://opensource.org/licenses/MIT) 授權。

---

## 👤 作者

*   **ckcm0210**
