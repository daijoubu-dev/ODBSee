# ODBSee
A WPF ODBC Client I made because ODBC View has some quirks i would like to fix but can't

Download (Single File .exe ~155MB): https://github.com/daijoubu-dev/ODBSee/releases

# TO DO:
*   [x] 32 bit support (WOW6432Node registry probe & x86 build target)
*   [x] Better distinguishing between User/System/Other DSN
*   [x] "Copy with Headers" UI Checkbox (copies highlighted cells with/without headers based on check state)
*   [x] Single cell copy (right-click or selection-based copying)
*   [x] Resizable query text box (GridSplitter integration)
*   [x] Remember ODBC connection per tab (Selected DSN, username, password saved per tab)
*   [x] Save state of all tabs and queries (Auto-saves every 60 seconds; preserves query texts, tab titles, and connections; excludes massive query results to prevent lag)
*   [x] Clickable status bar showing the last 50 log messages / query execution history
*   [x] Manual "Check for Updates" Hyperlink (Probes GitHub Releases API asynchronously, handles standalone EXEs & MSIs by opening downloads page if an update is found)
*   [x] Custom timeout limit (Settings window: connect & query timeouts, default max rows, copy-with-headers default)
*   [x] Encrypted saved passwords (Windows DPAPI; old plain-text session files are migrated automatically)
*   [x] Type-aware sorting (numbers, dates and text; blanks last; runs in the background)
*   [x] Column header right-click menu (Sort Asc/Desc, Select Column, Copy Column)
*   [x] Ctrl+Enter / F5 run the query from anywhere in the window
*   [x] Reopen closed tabs (Ctrl+Shift+T) and "Clear Results" on the tab right-click menu
*   [x] Report Bug / Feedback dialog (GitHub issue via token, or opens the browser pre-filled)
*   [x] Crash log & activity log in %LocalAppData%\ODBSee
*   [ ] Optimize standalone executable size (Compression & Trimming)

# BUGS
*   [ ] Executions crash when cancelling a running query (hardened: no driver-level cancel during row fetch, cancel runs off the UI thread; monitoring via crash.log / Event Viewer)
*   [ ] Massive pulls can lag and even crash the application (CSV export now streams to disk; still no row safety limit)
*   [ ] App freezes for a few seconds when switching back to it (monitoring; possibly memory from large results)
