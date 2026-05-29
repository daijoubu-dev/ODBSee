# ODBSee Features & Technical Brief

This document defines the core capabilities, architecture, and state management of the ODBSee utility.

## 🏗️ Core Architecture
*   **Framework:** .NET 10.0 Windows (WPF).
*   **Data Access:** `System.Data.Odbc` for connecting to any valid DSN.
*   **Hybrid UI:** 
    *   WPF shell for controls and layout.
    *   **GridSplitter Integration:** Dynamic resizing of the query input area.
    *   **Windows Forms `DataGridView`** (via `WindowsFormsHost`) for high-performance tabular results.
    *   **Virtual Mode:** Grid only requests data as needed; actual data is cached in a `List<string[]>`.
    *   **Double Buffering:** Enabled via reflection on the grid for flicker-free scrolling.

## 💾 State Management & Persistence
*   **Per-Tab State:** `MainWindow.TabState` holds result data, column names, sorting state, and specific connection credentials for each tab.
*   **Active Sync:** UI connection fields (DSN, User, Pass) are synchronized to the active tab state whenever the selection changes.
*   **Session Persistence:** 
    *   **Autosave:** Saves the entire workspace (all tabs, queries, and connection settings) every 60 seconds.
    *   **Restoration:** Automatically restores the last active session on application startup from `LocalAppData/ODBSee/session_state.json`.
    *   **Optimization:** Does not serialize massive query results to disk to ensure fast IO and low memory footprint.

## ⚡ Key Capabilities

### 1. Advanced ODBC Discovery
*   **Multi-Architecture Probing:** Simultaneously scans 64-bit and 32-bit (WOW6432Node) Registry hives to list all available DSNs.
*   **Categorized List:** DSNs are grouped by source type (User, System 64-bit, System 32-bit) in the selection menu.
*   **Credential Auto-Fill:** Attempts to retrieve `LogonID` from the registry for selected DSNs.

### 2. Multi-Tab SQL Workspace
*   Manage multiple independent query environments.
*   Rename tabs via double-click on the header.
*   Close tabs individually; state is managed via a `Tag` system.

### 3. Execution Control & Stability
*   **Non-Blocking UI:** Queries run asynchronously via `Task.Run`.
*   **Execute/Stop Toggle:** The execute button transforms into a "Stop" button during execution, providing thread-safe query cancellation.
*   **Simulation Engine:** "DUMMY" and "DUMMY_10s" datasources for offline performance testing and UI validation.

### 4. Data Interaction & UX
*   **Interactive Status Bar:** Clickable status label launches a log history viewer (last 50 messages/errors with timestamps).
*   **Smart Copying:** 
    *   **Copy with Headers:** Optional toggle to include column names in clipboard exports.
    *   **Flexible Selection:** Supports both full-row and single-cell/block-cell copying to CSV format.
*   **GitHub Update Checker:** Manual hyperlink that probes the GitHub Releases API to verify if a newer version is available.

### 5. Data Export
*   **CSV Exporting:** Generates UTF-8 encoded CSV files from the current result set.
*   **Virtual Grid Sorting:** Programmatic sorting of the underlying memory cache when column headers are clicked.

## 🛠️ Implementation Specs
*   **Default Max Rows:** 1000 (configurable via UI).
*   **Timeout:** Configured at 300 seconds (5 minutes) for both connection and command execution.
*   **Update Agent:** Uses `HttpClient` with a mandatory `User-Agent` header to interact with the GitHub API.
