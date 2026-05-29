# ODBSee Project Roadmap & TODO

This document tracks planned features, stability fixes, and bugs for ODBSee, directly aligned with the project goals.

---

## 🚀 High Priority (Stability & Performance Bugs)

### 1. Fix Execution Crashes on Cancel/Rerun
*   **Issue:** Double-clicking Execute or starting a query while another is cancelling causes race conditions on global state references (`_currentCts`, `_currentCommand`, `_runningQueryTask`), leading to concurrent thread collisions and crashes.
*   **Plan of Attack:**
    *   [x] **UI State Toggle:** Disable the "Execute" button and show/enable a red "Stop" button next to it when a query is running. This prevents users from initiating multiple concurrent executions.
    *   [x] **Encapsulate Execution Context:** Ensure cancellation tokens are strictly bound to individual tasks.
    *   [x] **Safe Cancellation:** Safeguard `OdbcCommand.Cancel()` calls to catch and ignore driver-specific native exceptions.
*   **Target:** `MainWindow.xaml`, `MainWindow.xaml.cs` -> `ExecuteQuery` (COMPLETED)

### 2. Prevent Memory Crashes on Massive Pulls
*   **Issue:** Pulling large datasets (100k+ rows) causes significant memory overhead because rows are cached in a giant list of string arrays (`_cachedData`). Exporting builds a massive in-memory string using `StringBuilder`, triggering `OutOfMemoryException`.
*   **Plan of Attack:**
    *   [ ] **Memory-Efficient Exporting:** Refactor `BtnExport_Click` to stream rows directly to disk via `StreamWriter` instead of buffering the entire CSV contents in memory using a giant `StringBuilder` string.
    *   [ ] **Virtual Grid Safety Limits:** Warn users when trying to pull more than 100,000 rows into the active UI grid view.
    *   [ ] **Explicit Memory Cleanup:** Ensure references (`_cachedData`, `_columnNames`) are set to null upon closing a tab to prevent memory leaks.
*   **Target:** `MainWindow.xaml.cs` -> `BtnExport_Click`, `GenerateCsv`, `CreateNewTab`

---

## 🛠️ Connection & Compatibility Improvements

### 3. Full 32-bit DSN Support
*   **Issue:** ODBSee compiles as a 64-bit executable, rendering 32-bit ODBC drivers and DSNs completely invisible and inaccessible because a 64-bit process cannot load 32-bit driver DLLs.
*   **Plan of Attack:**
    *   [x] **Compilation Options:** Update the project configuration and documentation to support building a standalone `win-x86` executable configuration for environments running 32-bit legacy drivers (e.g. Access, Excel, or older databases).
    *   [x] **Multi-Hive Registry Probing:** Update `GetOdbcDataSources` to probe both the 64-bit and 32-bit Registry hives using `RegistryView.Registry32` and `RegistryView.Registry64`.
*   **Target:** `ODBSee.csproj`, `MainWindow.xaml.cs` -> `GetOdbcDataSources` (COMPLETED)

### 4. Better Distinguishing Between User/System/Other DSNs
*   **Issue:** The DSN list is a single flat collection, making it difficult to distinguish between User DSNs, 64-bit System DSNs, and 32-bit System DSNs.
*   **Plan of Attack:**
    *   [x] **Define DSN Model:** Create a structured `OdbcDsn` class to store both the DSN `Name` and its `SourceType` (e.g., User, System 64-bit, System 32-bit).
    *   [x] **Grouped UI Menu:** In `BtnDatasource`'s context menu, group DSNs using submenus or category headers to make the list intuitive.
*   **Target:** `MainWindow.xaml.cs` -> `GetOdbcDataSources`, `BtnDatasource_Click` (COMPLETED)

### 5. Custom Connection and Command Timeout Limits
*   **Issue:** Connection and command execution timeouts are hardcoded to 300 seconds (5 minutes).
*   **Plan of Attack:**
    *   [ ] **Add UI Control:** Add a "Timeout (s)" TextBox to the execution panel next to Max Rows (defaulting to 300).
    *   [ ] **Dynamic Connection Strings:** Update `ExecuteQuery` to inject `Connect Timeout={Timeout};` into the connection string and assign `cmd.CommandTimeout` dynamically.
*   **Target:** `MainWindow.xaml`, `MainWindow.xaml.cs` -> `ExecuteQuery`

### 6. Remember ODBC Connection Per Tab
*   **Issue:** DSN selection and credentials are global. Switching tabs does not restore the connection state, leading to accidental queries executed against the wrong database.
*   **Plan of Attack:**
    *   [x] **Extend Tab State:** Add `SelectedDsn`, `User`, and `Password` properties to the `TabState` class.
    *   [x] **Active Switch Sync:** Update `QueryTabs_SelectionChanged` to save active credentials into the previous tab's state, and load the new tab's credentials into the UI input fields.
*   **Target:** `MainWindow.xaml.cs` -> `TabState`, `QueryTabs_SelectionChanged` (COMPLETED)

### 7. Save State of All Tabs and Queries (Session Persistence & Autosave)
*   **Issue:** Closing the application completely wipes all active query texts, tab titles, and configurations.
*   **Plan of Attack:**
    *   [x] **Define Session Model:** Create a JSON-serializable model tracking open tabs, active query texts, selected DSNs, and usernames.
    *   [x] **Save on Exit:** Serialize the active workspace state to a `session_state.json` file in the user's LocalAppData folder on window close.
    *   [x] **Load on Startup:** Deserialize and restore tabs and connection states on application startup.
    *   [x] **Autosave Timer:** Save the active workspace state automatically every 60 seconds.
    *   [x] **Exclude Results:** Do *NOT* serialize the query results cache (`_cachedData`) to prevent application lag, slow IO, and massive file sizes on disk.
*   **Target:** `MainWindow.xaml.cs` -> Window lifecycle event handlers (`Window_Closing`, `Window_Loaded`), and an auto-save timer thread (COMPLETED)

---

## ✨ UI/UX & Data Interaction Features

### 8. Resizable Query Text Box (Grid Splitter)
*   **Issue:** The query text box height is fixed at 150px, limiting visibility when drafting complex SQL scripts.
*   **Plan of Attack:**
    *   [x] **Introduce GridSplitter:** Add a horizontal `GridSplitter` between the query panel row (Row 0) and the results grid row (Row 1) in `MainWindow.xaml`.
    *   [x] **Responsive Row Definition:** Change Row 0's height to a star or auto value with a `MinHeight` constraint to allow dynamic dragging.
*   **Target:** `MainWindow.xaml` (COMPLETED)

### 9. Copy with Headers (UI Checkbox Control)
*   **Issue:** Users want an easy way to switch between copying rows with headers versus raw data values only, without using complex menu setups.
*   **Plan of Attack:**
    *   [x] **Add UI Checkbox:** Add a "Copy with Headers" CheckBox in the bottom status bar or execution control panel.
    *   [x] **Control Highlighted Cell Copying:** When copying selected grid cells (`Ctrl+C` or click copy), check the state of this CheckBox.
    *   [x] **Update Copy Logic:** If checked, prepend the respective column headers before copying the highlighted rows/cells. If unchecked, copy the raw values only.
*   **Target:** `MainWindow.xaml`, `MainWindow.xaml.cs` -> `CopySelectedToCsv` (COMPLETED)

### 10. Clickable Status Bar with Log History
*   **Issue:** Users need a simple way to troubleshoot queries, see execution durations, DSN connection attempts, or inspect past ODBC errors without cluttered modal dialogs.
*   **Plan of Attack:**
    *   [x] **Interactive Label:** Style the status info label (`StatusInfo`) to look clickable (change cursor to Hand, highlight on hover).
    *   [x] **Log Buffer:** Maintain an in-memory buffer of the last 50 status logs, errors, query run times, and cancellation messages with precise timestamps.
    *   [x] **History Viewer Popup:** On clicking the status text, launch a lightweight dialog or popup displaying this log history in chronological order.
*   **Target:** `MainWindow.xaml`, `MainWindow.xaml.cs` (COMPLETED)

### 11. Single Cell Copy Support
*   **Issue:** The grid is locked to `FullRowSelect`, making it impossible to select and copy a single cell's text.
*   **Plan of Attack:**
    *   [x] **Cell Right-Click Context Menu:** Add a ContextMenu to the grid with a "Copy Cell Value" item.
    *   [x] **Hit Testing:** On cell right-click, resolve the specific cell clicked using `WfGrid.HitTest` and copy only its string value to the clipboard.
*   **Target:** `MainWindow.xaml.cs` -> `SetupWinFormsGrid` (COMPLETED via cell-block/single-cell highlighting copying)

### 12. Manual GitHub Release Update Checker
*   **Issue:** Users want a manual, unobtrusive way to verify if they are running the latest version of ODBSee from GitHub, regardless of using standalone EXEs or MSI packages.
*   **Plan of Attack:**
    *   [x] **Status Bar Link:** Add a "Check for Updates" TextBlock link to the status bar (with hand cursor, hover underline, and click hook).
    *   [x] **Async Releases API Probing:** Query `https://api.github.com/repos/daijoubu-dev/ODBSee/releases/latest` using `HttpClient` (with User-Agent headers).
    *   [x] **Strict Tag Comparison:** Clean and parse the latest version tag from GitHub and compare it to the executing assembly's version.
    *   [x] **Dynamic UI Transitions:** If an update is available, color the link red, show the new version number, and bind a mouse listener to open the release URL in the user's default browser. If up-to-date, color the link green and show "Up to Date".
*   **Target:** `MainWindow.xaml`, `MainWindow.xaml.cs` -> `BtnCheckUpdate_MouseLeftButtonUp` (COMPLETED)
