# ODBSee
A WPF ODBC Client I made because ODBC View has some quirks i would like to fix but can't

Download (Single File .exe ~155MB): https://github.com/daijoubu-dev/ODBSee/releases

# TO DO:
*   32 bit support (WOW6432Node registry probe & x86 build target)
*   Better distinguishing between User/System/Other DSN
*   [x] "Copy with Headers" UI Checkbox (copies highlighted cells with/without headers based on check state)
*   [x] Single cell copy (right-click or selection-based copying)
*   [x] Resizable query text box (GridSplitter integration)
*   [x] Remember ODBC connection per tab (Selected DSN, username, password saved per tab)
*   [x] Save state of all tabs and queries (Auto-saves every 60 seconds; preserves query texts, tab titles, and connections; excludes massive query results to prevent lag)
*   [x] Clickable status bar showing the last 50 log messages / query execution history
*   [x] Manual "Check for Updates" Hyperlink (Probes GitHub Releases API asynchronously, handles standalone EXEs & MSIs by opening downloads page if an update is found)
*   Custom timeout limit (dynamic Connect Timeout and CommandTimeout)

# BUGS
*   [x] Executions crash when cancelling/rerunning currently running query (Fixed via UI execute/stop toggle and thread-safe cancellation task encapsulation)
*   Massive pulls can lag and even crash the application (needs memory-efficient exporting and safety limits)
