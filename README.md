# Excel 會話管理器 (Excel Session Manager) — 超詳細新手／長者安裝教學 + 技術說明

一個基於 Python `tkinter` 和 `pywin32` 函式庫開發的 Windows GUI 工具，旨在提升處理多個 Microsoft Excel 檔案時的工作效率。此工具提供了一個中央化介面，用以顯示、操作及批量管理所有已開啟的 Excel 檔案，並支援會話保存、還原、批次關閉、迷你浮窗等自動化操作。

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

| 主介面                                                                                                            | 迷你懸浮窗模式                                                          |
| ----------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------- |
| ![主介面截圖](https://raw.githubusercontent.com/ckcm0210/excel_session_manager/main/screenshots/main_ui.png)           | ![迷你模式截圖](https://raw.githubusercontent.com/ckcm0210/excel_session_manager/main/screenshots/mini_ui.png) |

---

## 🚀 極簡超詳細安裝及啟動教學（新手／長者／零經驗都用得）

### 步驟一：打開 Jupyter Notebook

1. **撳住鍵盤上嘅「Windows鍵」再撳一下「S」鍵**（即同時按住「Windows」+「S」），就會自動彈出Windows搜尋欄（適用於大部分Windows版本）。
2. 喺搜尋欄直接輸入 `jupyter notebook`（大細階都得，唔使理），你會見到選單上面出現「Jupyter Notebook」圖示。
3. 撳「Jupyter Notebook」圖示，等幾秒，會自動喺你部機嘅瀏覽器（例如 Chrome、Edge）開一個新分頁，標題寫住「Jupyter」就代表成功。

---

### 步驟二：新建專用資料夾（新Folder）

1. 打開咗Jupyter Notebook嘅網頁之後，**畫面右上角**會見到一個寫住「New」嘅大掣。
2. 用滑鼠撳一下「New」嗰個掣，會彈出一個下拉選單。
3. 喺選單入面搵到「Folder」或者「New Folder」呢個選項，撳一下。
4. 成功嘅話，畫面最上面多咗一行「Untitled Folder」，即係新建咗個資料夾，但未改名。
5. 去到「Untitled Folder」嗰行，最右邊有個小tick box，撳一下揀中佢。
6. 頁面頂部會出現多咗個「Rename」掣，撳「Rename」。
7. 輸入新名：**excel_session_manager**（全部細階英文字母，中間用底線，記得唔好有空格），然後撳「Rename」。
8. 而家你會見到個資料夾名已經變成「excel_session_manager」。

---

### 步驟三：進入新Folder入面

1. 撳一下「excel_session_manager」資料夾名（係藍色字），你就會入咗呢個新資料夾，畫面會變成空白（未有檔案）。

---

### 步驟四：建立Python程式檔案（有兩種簡單方法，二揀一）

#### 方法一：用Jupyter Notebook內置「Text File」手動複製貼上

1. **再撳右上角「New」**，揀「Text File」。
2. 畫面會開一個新分頁，頂部寫住「untitled.txt」。
3. **開定你收到嘅 excel_session_manager_v16.py**（用記事本或任何方法睇到內容）。
4. **將全部內容複製（Ctrl+A 然後 Ctrl+C）**。
5. 返去Jupyter Notebook「untitled.txt」分頁，**將內容全部貼上（Ctrl+V）**。
6. 撳左上角「File」>「Save as...」，輸入檔案名：**excel_session_manager_v16.py**，撳「Save」。
7. 關閉分頁。

8. **重複以上步驟一次**，今次換另一個檔案（excel_session_manager_link_updater.py）：
    - 用同樣方法開新「Text File」
    - 複製檔案內容，貼入新分頁
    - 「Save as...」改名為 **excel_session_manager_link_updater.py**
    - 關閉分頁

#### 方法二：用滑鼠直接拖拉檔案進入Jupyter Notebook

1. **開定你收到嘅py檔案（excel_session_manager_v16.py 同 excel_session_manager_link_updater.py）嘅資料夾視窗**（例如你張桌面、Documents或Download）。
2. **用滑鼠左鍵「拖住」其中一個py檔案**（例如 excel_session_manager_v16.py）。
3. **拖去Jupyter Notebook個網頁**，將滑鼠放喺你啱啱建立好、而家入咗去嘅「excel_session_manager」資料夾畫面上。
4. **一放手（放開滑鼠左鍵）**，檔案就會自動上載並顯示喺Jupyter Notebook資料夾列表。
5. **重複拉多一次另一個py檔案**（excel_session_manager_link_updater.py）入去。

> 兩種方法選一樣就可以，唔需要兩樣一齊做。用拖拉法最快最簡單。

---

### 步驟五：建立Notebook並運行主程式

1. 喺「excel_session_manager」資料夾入面，撳「New」>「Python 3」。
2. 會開一個新Notebook，標題係「Untitled.ipynb」。
3. 第一個格（[ ]:) 入面打：
    ```
    %run excel_session_manager_v16.py
    ```
4. 撳上面「File」>「Save and Rename」，改名做 **啟動Excel工具.ipynb**，撳「Rename」。
5. 撳一下Notebook個格，再撳上面「Run」掣，或者直接撳 **Shift+Enter**。
6. 幾秒之後，Excel Session Manager 的視窗會自動彈出嚟！

---

### 步驟六：完成

- 你可以最小化Jupyter Notebook個網頁，但唔好關閉或者Delete。
- 操作完Excel Session Manager，可以直接關閉視窗。

---

### 常見小貼士

- **如果貼錯檔案內容，Delete舊檔案再重新Create。**
- **改名時要小心，檔案名必須正確（.py 結尾），資料夾名無空格。**
- **任何步驟做錯，可以搵IT同事幫手。**

---

## ✨ 核心功能

*   **實時會話掃描**：非同步偵測並列出所有當前開啟的 Excel 檔案實例。
*   **詳細資訊展示**：提供檔名、完整路徑及最後修改時間，並支援點擊標題進行排序。
*   **直觀列表操作**：透過繼承 `ttk.Treeview` 實現了拖曳多選功能，提升操作流暢度。
*   **批次視窗控制**：可將選定的一個或多個 Excel 視窗批量帶到最前、儲存或關閉。
*   **工作階段持久化**：
    *   **儲存會話**：將當前檔案列表（包含路徑、作用中工作表、選定儲存格）的狀態匯出為一個 `.xlsx` 會話檔。
    *   **載入會話**：從會話檔中讀取數據，批量重新開啟所有檔案並還原其工作狀態。
*   **外部連結自動掃描與更新**：可自訂條件，自動尋找及更新所有有外部連結的 Excel 檔案，並產生操作日誌。
*   **非侵入式迷你模式**：可將主介面縮小為一個置頂的懸浮圖示，方便在不中斷工作的同時快速存取。
*   **介面個人化**：支援用戶自訂列表的字體和大小。

---

## 🛠️ 架構設計與技術剖析（更詳細）

### 程式整體架構

```
ExcelSessionManagerApp (主類)
│
├── UI 狀態變數 (e.g., self.is_mini, self.showing_path)
│
├── setup_ui()                     # UI 初始化：負責所有元件的創建、佈局和事件綁定
│   ├── 主容器 (ttk.Notebook)
│   ├── 佈局框架 (tk.Frame)
│   └── 元件 (Widgets):
│       ├── DragSelectTreeview (自訂多選、行高、動態字體)
│       └── 操作按鈕區、標籤、迷你浮窗icon等
│
├── 功能模組 (Class Methods)
│   │
│   ├── Mini Widget 邏輯 (enter_mini, exit_mini)
│   │   └── pack_forget() 隱藏主UI、使用wm_attributes()設置always on top與minsize限制、動態切換icon顯示
│   │
│   ├── COM 互動邏輯 (get_open_excel_files, save_selected_workbooks, close_selected_workbooks等)
│   │   └── 使用 pythoncom.CoInitialize() 及 win32com.client.GetActiveObject("Excel.Application") 操控Excel
│   │   └── try-except 處理Excel未開啟或多進程情境
│   │   └── threading.Thread + root.after() 解決tkinter非線程安全限制
│   │
│   ├── 視窗控制 (activate_selected_workbooks, minimize_all_excel)
│   │   └── win32gui、win32con API操作Windows視窗，根據檔名/視窗標題配對進行最小化、置前、定位
│   │
│   └── 會話檔案邏輯 (save_session, load_session)
│       └── openpyxl建立/讀取.xlsx會話檔，支援檔案、工作表、儲存格精確還原
│
├── 外部連結掃描與自動更新 (excel_session_manager_link_updater)
│   └── 專用模組，使用 pywin32 掃描已開啟Excel的所有外部連結，根據用戶自訂條件（如修改日期）批量自動更新，並可產生日誌及統計表
```

### 技術細節逐點解構

#### 1. GUI 介面（tkinter + ttk）

- **主要元件**：`ttk.Notebook`分頁、`ttk.Frame`框架、`ttk.Treeview`（自訂支持拖曳多選，動態字體調整）、多粒操作按鈕。
- **自訂 DragSelectTreeview**：繼承自`ttk.Treeview`，重寫滑鼠事件實現連續拖曳多選（解決原生Treeview多選體驗差問題）。
- **動態字體/行高/寬度調整**：用戶可即時調整字體與大小，Treeview行高、欄寬動態計算以對齊內容。
- **迷你模式**：`pack_forget()`隱藏大主介面，僅顯示icon，`wm_attributes("-topmost", 1)`置頂，`minsize`限制縮放。

#### 2. Excel自動化與COM技術（pywin32, pythoncom）

- **Excel Application COM物件**：用`win32com.client.GetActiveObject("Excel.Application")`取得現有Excel應用程式（如未開啟則try-except處理）。
- **文件級操作**：遍歷`excel.Workbooks`，取得所有開啟中的工作簿，獲取檔案名、路徑、目前工作表、選取儲存格。
- **儲存/關閉/還原**：對每個Workbook呼叫`.Save()`、`.Close()`，用openpyxl產生session檔，支援還原至指定sheet/cell。
- **外部連結掃描與更新**：掃描Workbook內所有external links，判斷來源檔案有冇近日期修改，自動決定要唔要刷新連結（UpdateLink方法）。
- **多執行緒與GUI同步**：所有與Excel互動的操作都在threading.Thread背景執行，最後用`root.after()`回到主線程更新UI（避免tkinter thread衝突）。

#### 3. Windows視窗控制（win32gui, win32con）

- **視窗尋找與操作**：用`win32gui.EnumWindows()`遍歷所有Windows，根據Excel視窗標題（通常包含檔名和「 - Excel」）配對，然後調用`ShowWindow`還原/最小化、`SetForegroundWindow`置前、`SetWindowPos`指定位置。
- **Process/Handle安全**：用psutil輔助過濾殭屍Excel進程，避免操作失敗。

#### 4. Session檔案處理（openpyxl）

- **保存Session**：將所有開啟檔案、工作表、儲存格紀錄到一個.xlsx，方便日後還原。
- **還原Session**：逐個打開記錄的檔案，指定Sheet/Cell自動跳轉。
- **檔案路徑驗證**：開啟前自動檢查檔案是否存在，異常時彈窗提示。

#### 5. 非同步及錯誤處理

- **所有耗時操作（如掃描、儲存、關閉、批量更新）均在背景thread執行，確保介面不卡死。**
- **所有與Excel互動的try-except都會將錯誤即時透過popup或console顯示，方便追蹤問題。**
- **外部連結更新可選產生日誌、統計報表（Excel格式），方便後續審計。**

#### 6. 個人化與使用體驗提升

- **字體/大小調整**：支援即時切換多款monospace字體與文字大小。
- **視窗最小化/迷你浮窗**：可一鍵切換簡化模式，主介面暫時收起，只顯示浮動icon，方便多工。
- **進度console**：大批量操作時可選彈出console視窗顯示每一步進度、成功/失敗與用時。

---

## 🚀 如何使用（快速列表）

1.  **環境要求**
    *   Windows 作業系統
    *   Python 3.x
    *   已安裝 Microsoft Excel

2.  **安裝相依套件**
    ```bash
    pip install pywin32 openpyxl pillow psutil
    ```

3.  **按上面新手教學操作，於 Jupyter Notebook 內「%run excel_session_manager_v16.py」啟動 GUI。**

4.  **自訂懸浮窗圖示**
    *   在程式碼同一個目錄下，放置一張 `.png` 格式的圖片，並將其命名為 `maximize_full_screen.png`。
    *   若程式找不到此檔案，將會自動使用預設的 emoji「🗔」作為替代圖示。

---

## 🏆 常見應用場景

- **會計、財務、報表分析每日須同時處理十幾廿個Excel檔案，想一鍵Save/Close/還原。**
- **團隊多同事分批打開同一批Excel，交班時可用Session檔輕鬆交接。**
- **自動檢查大量Excel之間的外部連結，確保數據新鮮。**
- **需要清理殭屍Excel進程或快速重啟工作環境。**

---

## 🔮 未來發展方向（Roadmap）

- 增強設定管理（支援自訂視窗大小、字體、預設路徑等並保存到設定檔）
- 多Excel實例支援（多開Excel主程式時全部偵測/操作）
- 打包為獨立執行檔（無需安裝Python）
- 更健壯錯誤處理與詳細日誌記錄
- 增加搜尋/過濾、狀態列、Tooltips等UI優化

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

---

**全程用滑鼠同Jupyter Notebook，唔使打指令、唔使黑色cmd，全部係圖形介面，長者都用得！有咩困難，請即時搵IT同事協助！**
