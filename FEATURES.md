# ODBSee Features & Technical Brief

This document defines the core capabilities, architecture, and state management of the ODBSee utility for future reference and AI context.

## 🏗️ Core Architecture
*   **Framework:** .NET 8.0 Windows (WPF).
*   **Data Access:** `System.Data.Odbc` for connecting to any valid DSN.
*   **Hybrid UI:** 
    *   WPF shell for controls and layout.
    *   **Windows Forms `DataGridView`** (via `WindowsFormsHost`) for high-performance tabular results.
    *   **Virtual Mode:** Grid only requests data as needed; actual data is cached in a `List<string[]>`.
    *   **Double Buffering:** Enabled via reflection on the grid for flicker-free scrolling.

## 💾 State Management (Per-Tab Persistence)
*   **Object Model:** `MainWindow.TabState` holds the result data, column names, and sorting state for each tab.
*   **Tagging System:** Each `TabItem.Tag` contains a `TabState` instance.
*   **Active Sync:** The main window maintains private references (`_cachedData`, `_columnNames`, `_lastSortColumnIndex`, `_sortAscending`) that are synchronized whenever `QueryTabs_SelectionChanged` is triggered.
*   **Query Lifecycle:** Results are stored in the `TabState` of the tab that *initiated* the query, even if the user switches tabs during execution.

## ⚡ Key Capabilities

### 1. ODBC Datasource Discovery
*   Automatic scanning of `HKCU` and `HKLM` registry paths for `ODBC Data Sources`.
*   Auto-populates credentials (`LogonID`) from the registry.

### 2. Multi-Tab SQL Workspace
*   Create and manage multiple query tabs independently.
*   Persistent query text and query results per tab.
*   Double-click to rename tab headers.

### 3. Asynchronous Query Execution
*   Queries run on background threads (`Task.Run`).
*   **Robust Cancellation:** Uses `CancellationTokenSource` and `OdbcCommand.Cancel()`. Triggered automatically when "Execute" is hit again.

### 4. DUMMY DSN (Simulation Engine)
*   Virtual "DUMMY" DSN available for offline testing.
*   Generates randomized, realistic data based on the "Max Rows" setting.
*   Useful for stress-testing grid performance without a real database.

### 5. Data Interaction & Export
*   **Virtual Grid Sorting:** Programmatic sorting of the underlying memory cache when column headers are clicked.
*   **CSV Exporting:** Generates UTF-8 encoded CSV files from the current result set.
*   **Manual Clipboard Copy:** Custom `Ctrl+C` logic for copying selected rows as formatted CSV text.

## 🛠️ Internal Implementation Details
*   **Default Max Rows:** 1000 (configurable via UI).
*   **Timeout:** Hardcoded at 300 seconds (5 minutes) for both connection and command execution.
*   **Tab-Awareness:** UI updates (Status labels, Grid) only occur if the initiating tab is still the active tab when a background query finishes.
