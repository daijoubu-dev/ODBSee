using System;
using System.Globalization;
using Microsoft.Win32;
using System.Data.Odbc;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Threading;
using Button = System.Windows.Controls.Button;
using ContextMenu = System.Windows.Controls.ContextMenu;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using Label = System.Windows.Controls.Label;
using MenuItem = System.Windows.Controls.MenuItem;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Orientation = System.Windows.Controls.Orientation;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using TextBox = System.Windows.Controls.TextBox;
using WinForms = System.Windows.Forms;

namespace ODBSee
{
    public class OdbcDsn
    {
        public string Name { get; set; } = "";
        public string SourceType { get; set; } = "";
    }

    public partial class MainWindow : Window
    {
        private string _selectedDsn = "";

        private readonly AppSettings _settings = AppSettings.Load();

        // These act as references to the current tab's state
        private List<string[]> _cachedData = new List<string[]>();
        private List<string> _columnNames = new List<string>();

        // Trackers for sorting (will be updated from/to tab state)
        private int _lastSortColumnIndex = -1;
        private bool _sortAscending = true;
        private bool _isSorting;

        // Query execution. Only one query runs at a time.
        private CancellationTokenSource? _currentCts;
        private OdbcCommand? _currentCommand;           // guarded by _commandLock
        private bool _inExecuteReader;                  // guarded by _commandLock
        private readonly object _commandLock = new object();

        // Recently closed tabs (most recent first), reopened with Ctrl+Shift+T. Results are kept for these.
        private const int MaxClosedTabs = 5;
        private readonly LinkedList<TabItem> _closedTabs = new LinkedList<TabItem>();

        // Column header right-click menu
        private WinForms.ContextMenuStrip? _headerMenu;
        private int _headerMenuColumn = -1;

        // Log history fields (last 50 messages)
        private readonly List<string> _logHistory = new List<string>();
        private readonly object _logLock = new object();

        public class TabState
        {
            public List<string[]> CachedData { get; set; } = new List<string[]>();
            public List<string> ColumnNames { get; set; } = new List<string>();
            public int LastSortColumnIndex { get; set; } = -1;
            public bool SortAscending { get; set; } = true;
            public string SelectedDsn { get; set; } = "";
            public string User { get; set; } = "";
            public string Password { get; set; } = "";
        }

        public class SessionState
        {
            public List<TabSessionInfo> Tabs { get; set; } = new List<TabSessionInfo>();
            public int ActiveTabIndex { get; set; } = 0;
        }

        public class TabSessionInfo
        {
            public string Header { get; set; } = "";
            public string QueryText { get; set; } = "";
            public string SelectedDsn { get; set; } = "";
            public string User { get; set; } = "";

            /// <summary>Password encrypted with Windows DPAPI (current user). See <see cref="Secret"/>.</summary>
            public string EncryptedPassword { get; set; } = "";

            /// <summary>Legacy plain-text password from older versions. Read once for migration, never written.</summary>
            [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
            public string? Password { get; set; }
        }

        public MainWindow()
        {
            InitializeComponent();
            SetupWinFormsGrid();

            // The WinForms grid doesn't use WPF resources, so it's themed by hand (and re-themed on switch)
            ApplyGridTheme();
            ThemeManager.ThemeChanged += ApplyGridTheme;

            TxtMaxRows.Text = _settings.DefaultMaxRows.ToString();
            ChkCopyWithHeaders.IsChecked = _settings.CopyWithHeadersDefault;

            // Hover underline for the clickable status bar items
            foreach (var link in new[] { StatusInfo, BtnCheckUpdate, BtnSettings, BtnFeedback })
            {
                var l = link;
                l.MouseEnter += (s, e) => l.TextDecorations = TextDecorations.Underline;
                l.MouseLeave += (s, e) => l.TextDecorations = null;
            }

            // Window-wide shortcuts (Ctrl+Enter / F5 to run, Ctrl+Shift+T to reopen a closed tab)
            PreviewKeyDown += Window_PreviewKeyDown;

            // Register closing handler to save session on exit
            this.Closing += (s, e) => SaveSession();

            QueryTabs.Items.Clear(); // Remove the placeholder from XAML

            // Load saved session if exists, otherwise create a new default tab
            if (!LoadSession())
            {
                CreateNewTab("Query 1", "select * from ");
            }

            // Set up auto-save DispatcherTimer (runs every 60 seconds on UI thread)
            var autosaveTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromSeconds(60)
            };
            autosaveTimer.Tick += (s, e) => SaveSession();
            autosaveTimer.Start();

            // Set window title with version
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            if (version != null)
            {
                this.Title = $"ODBSee v{version.ToString(3)}";
            }

            LogStatus("Application started. Ready.");

            _ = RetryPendingFeedbackAsync();
        }

        private void LogStatus(string message)
        {
            var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
            var logLine = $"[{timestamp}] {message}";

            lock (_logLock)
            {
                _logHistory.Add(logLine);
                if (_logHistory.Count > 50)
                {
                    _logHistory.RemoveAt(0);
                }
            }

            // Also keep a rolling log on disk (%LocalAppData%\ODBSee\activity.log) for diagnosing issues later
            AppStorage.AppendLog(AppStorage.ActivityLogFile, logLine);

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => StatusInfo.Text = message);
            }
            else
            {
                StatusInfo.Text = message;
            }
        }

        // ------------------------------------------------------------------
        // Session persistence
        // ------------------------------------------------------------------

        private void SaveSession()
        {
            try
            {
                // Sync the active tab credentials first
                if (QueryTabs.SelectedItem is TabItem activeTab && activeTab.Tag is TabState activeState)
                {
                    activeState.SelectedDsn = _selectedDsn;
                    activeState.User = TxtUser.Text;
                    activeState.Password = TxtPass.Password;
                }

                var session = new SessionState
                {
                    ActiveTabIndex = QueryTabs.SelectedIndex
                };

                foreach (TabItem tab in QueryTabs.Items)
                {
                    if (tab.Tag is TabState state && tab.Content is TextBox tb)
                    {
                        session.Tabs.Add(new TabSessionInfo
                        {
                            Header = GetTabTitle(tab),
                            QueryText = tb.Text,
                            SelectedDsn = state.SelectedDsn,
                            User = state.User,
                            EncryptedPassword = _settings.RememberPasswords ? Secret.Protect(state.Password) : "",
                            Password = null // never write plain text
                        });
                    }
                }

                AppStorage.WriteJson(AppStorage.SessionFile, session);
            }
            catch (Exception ex)
            {
                LogStatus($"Autosave failed: {ex.Message}");
            }
        }

        private bool LoadSession()
        {
            try
            {
                var session = AppStorage.ReadJson<SessionState>(AppStorage.SessionFile);
                if (session == null || session.Tabs.Count == 0) return false;

                bool hadPlainText = false;

                QueryTabs.Items.Clear();
                foreach (var tabInfo in session.Tabs)
                {
                    string password = "";
                    if (!string.IsNullOrEmpty(tabInfo.EncryptedPassword))
                    {
                        password = Secret.Unprotect(tabInfo.EncryptedPassword);
                    }
                    else if (!string.IsNullOrEmpty(tabInfo.Password))
                    {
                        password = tabInfo.Password; // migrate from older plain-text session file
                        hadPlainText = true;
                    }

                    if (!_settings.RememberPasswords) password = "";

                    CreateNewTabInternal(tabInfo.Header, tabInfo.QueryText, new TabState
                    {
                        SelectedDsn = tabInfo.SelectedDsn,
                        User = tabInfo.User,
                        Password = password
                    });
                }

                int indexToSelect = session.ActiveTabIndex;
                if (indexToSelect >= 0 && indexToSelect < QueryTabs.Items.Count)
                {
                    QueryTabs.SelectedIndex = indexToSelect;
                }
                else if (QueryTabs.Items.Count > 0)
                {
                    QueryTabs.SelectedIndex = 0;
                }

                if (hadPlainText)
                {
                    // Rewrite the session file right away so plain-text passwords don't linger on disk
                    SaveSession();
                    LogStatus("Saved passwords are now encrypted.");
                }

                LogStatus($"Restored {session.Tabs.Count} queries from previous session.");
                return true;
            }
            catch (Exception ex)
            {
                LogStatus($"Failed to load session state: {ex.Message}");
                return false;
            }
        }

        // ------------------------------------------------------------------
        // Query execution
        // ------------------------------------------------------------------

        private void SetQueryRunningState(bool isRunning)
        {
            var color = isRunning ? "#D13438" : "#007ACC";
            var text = isRunning ? "■ Stop" : "▶ Execute";

            BtnExecute.Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString(color));
            BtnExecute.IsEnabled = true;

            if (BtnExecute.Content is TextBlock tb)
            {
                tb.Text = text;
            }
            else
            {
                BtnExecute.Content = new TextBlock { Text = text };
            }
        }

        private void CancelQuery()
        {
            var cts = _currentCts;
            if (cts == null || cts.IsCancellationRequested) return;

            LogStatus("Query cancellation requested by user...");
            cts.Cancel();

            if (BtnExecute.Content is TextBlock tb) tb.Text = "Stopping...";
            BtnExecute.IsEnabled = false;

            // While rows are being fetched, the token alone stops the read loop (same path as hitting Max Rows),
            // so we don't poke the driver from another thread mid-fetch. Only a query that is still waiting on the
            // server (inside ExecuteReader) needs an ODBC-level cancel. That runs off the UI thread so a slow
            // driver can't freeze the window, and under a lock so it never touches a command that's been disposed.
            Task.Run(() =>
            {
                lock (_commandLock)
                {
                    if (_currentCommand == null || !_inExecuteReader) return;
                    try
                    {
                        _currentCommand.Cancel();
                    }
                    catch (Exception ex)
                    {
                        LogStatus($"Driver error while cancelling (ignored): {ex.Message}");
                    }
                }
            });
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            ExecuteQuery();
        }

        private async void ExecuteQuery()
        {
            // Button doubles as Stop while a query is running
            if (_currentCts != null)
            {
                CancelQuery();
                return;
            }

            var activeTab = QueryTabs.SelectedItem as TabItem;
            var state = activeTab?.Tag as TabState;
            var activeBox = GetCurrentQueryBox();

            if (activeBox == null || state == null) return;
            if (string.IsNullOrEmpty(_selectedDsn))
            {
                MessageBox.Show("Please select an ODBC Datasource first.", "Datasource Required", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var dsn = _selectedDsn;
            var queryText = activeBox.Text;
            var userId = TxtUser.Text;
            var password = TxtPass.Password;
            int.TryParse(TxtMaxRows.Text, out var maxRows);
            if (maxRows <= 0) maxRows = _settings.DefaultMaxRows > 0 ? _settings.DefaultMaxRows : 1000;
            var connectTimeout = Math.Max(0, _settings.ConnectTimeoutSeconds);
            var commandTimeout = Math.Max(0, _settings.CommandTimeoutSeconds);

            LogStatus($"Executing query on DSN '{dsn}' (Limit: {maxRows} rows)...");
            var connString = $"DSN={dsn};Uid={userId};Pwd={password};";

            var cts = new CancellationTokenSource();
            _currentCts = cts;
            var token = cts.Token;
            var started = DateTime.Now;

            SetQueryRunningState(true);

            try
            {
                Task<(List<string> Columns, List<string[]> Rows)> task;
                if (dsn == "DUMMY" || dsn == "DUMMY_10s")
                {
                    bool isLongDelay = dsn == "DUMMY_10s";
                    task = Task.Run(() => GenerateDummyData(maxRows, isLongDelay, token), token);
                }
                else
                {
                    task = Task.Run(() => FetchData(connString, queryText, maxRows, connectTimeout, commandTimeout, token), token);
                }

                var (columns, rows) = await task;
                var elapsed = DateTime.Now - started;

                state.ColumnNames = columns;
                state.CachedData = rows;
                state.LastSortColumnIndex = -1;
                state.SortAscending = true;

                // Only update the grid if the tab that initiated the query is still the active one
                if (QueryTabs.SelectedItem == activeTab)
                {
                    RefreshGridFromState(state);
                    StatusConnection.Text = "Connected";
                    LogStatus($"Query executed successfully: returned {rows.Count} rows in {elapsed.TotalSeconds:F1}s.");
                }
                else
                {
                    LogStatus($"Background query finished: returned {rows.Count} rows in {elapsed.TotalSeconds:F1}s.");
                }
            }
            catch (OperationCanceledException)
            {
                LogStatus("Query execution cancelled by user.");
            }
            catch (OdbcException ex)
            {
                if (token.IsCancellationRequested)
                {
                    LogStatus("Query cancelled by user (OdbcException handled).");
                }
                else
                {
                    LogStatus($"ODBC Error: {ex.Message}");
                    MessageBox.Show(ex.Message, "ODBC Database Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                if (token.IsCancellationRequested)
                {
                    LogStatus("Query cancelled by user (Exception handled).");
                }
                else
                {
                    LogStatus($"Query Error: {ex.Message}");
                    MessageBox.Show(ex.Message, "Query Execution Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            finally
            {
                if (_currentCts == cts)
                {
                    _currentCts = null;
                }
                SetQueryRunningState(false);
            }
        }

        private static (List<string> Columns, List<string[]> Rows) GenerateDummyData(int maxRows, bool isLongDelay, CancellationToken token)
        {
            var columns = new List<string> { "ID", "Name", "Category", "Value", "Status", "Last Modified" };
            var rows = new List<string[]>();
            var rand = new Random();

            // Loop and sleep in 100ms chunks to ensure rapid responsiveness to stop/cancellation commands
            int totalSleepMs = isLongDelay ? 10000 : 300;
            int intervalMs = 100;
            int elapsedMs = 0;

            while (elapsedMs < totalSleepMs)
            {
                token.ThrowIfCancellationRequested();
                Thread.Sleep(intervalMs);
                elapsedMs += intervalMs;
            }

            for (int i = 1; i <= maxRows; i++)
            {
                token.ThrowIfCancellationRequested();

                rows.Add(new string[] {
                    i.ToString(),
                    $"Dummy Record {i}",
                    rand.Next(0, 2) == 0 ? "Production" : "Staging",
                    (rand.NextDouble() * 1000).ToString("F2"),
                    rand.Next(0, 3) switch { 0 => "Active", 1 => "Pending", _ => "Archived" },
                    DateTime.Now.AddDays(-rand.Next(1, 400)).AddMinutes(-rand.Next(0, 1440)).ToString()
                });
            }

            return (columns, rows);
        }

        private (List<string> Columns, List<string[]> Rows) FetchData(string connString, string sql, int maxRows,
            int connectTimeout, int commandTimeout, CancellationToken token)
        {
            var cols = new List<string>();
            var rows = new List<string[]>();

            using var conn = new OdbcConnection(connString);
            // Set via the property (SQL_ATTR_LOGIN_TIMEOUT): drivers often ignore "Connect Timeout" in the string
            conn.ConnectionTimeout = connectTimeout;
            conn.Open();
            token.ThrowIfCancellationRequested();

            using var cmd = new OdbcCommand(sql, conn);
            cmd.CommandTimeout = commandTimeout;

            lock (_commandLock)
            {
                _currentCommand = cmd;
                _inExecuteReader = true;
            }

            try
            {
                OdbcDataReader reader;
                try
                {
                    reader = cmd.ExecuteReader();
                }
                finally
                {
                    lock (_commandLock) _inExecuteReader = false;
                }

                using (reader)
                {
                    token.ThrowIfCancellationRequested();

                    for (var i = 0; i < reader.FieldCount; i++)
                        cols.Add(reader.GetName(i));

                    int fieldCount = reader.FieldCount;
                    // Check the count first so we don't fetch one row more than we keep
                    while (rows.Count < maxRows && reader.Read())
                    {
                        token.ThrowIfCancellationRequested();

                        string[] row = new string[fieldCount];
                        for (int i = 0; i < fieldCount; i++)
                            row[i] = reader[i]?.ToString() ?? "";
                        rows.Add(row);
                    }
                }
            }
            finally
            {
                // Always clear the shared reference (before cmd is disposed) so Cancel can never hit a dead command
                lock (_commandLock)
                {
                    _currentCommand = null;
                    _inExecuteReader = false;
                }
            }

            return (cols, rows);
        }

        // ------------------------------------------------------------------
        // Results grid
        // ------------------------------------------------------------------

        private void SetupWinFormsGrid()
        {
            WfGrid.VirtualMode = true;
            WfGrid.ReadOnly = true;
            WfGrid.AllowUserToAddRows = false;
            WfGrid.RowHeadersVisible = true; // Make row headers visible so users can select entire rows by clicking headers
            WfGrid.RowHeadersWidth = 40; // Set a nice default width for the row headers
            WfGrid.AllowUserToResizeRows = false;
            WfGrid.AllowUserToResizeColumns = true;
            WfGrid.ClipboardCopyMode = WinForms.DataGridViewClipboardCopyMode.Disable; //disable native copy so we can do our own functional copy
            WfGrid.SelectionMode = WinForms.DataGridViewSelectionMode.RowHeaderSelect; // Allows both cell block and row selections!

            //listener for manual copy method + shortcuts while the grid has focus
            WfGrid.KeyDown += WfGrid_KeyDown;

            typeof(WinForms.DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.SetValue(WfGrid, true, null);

            WfGrid.CellValueNeeded += (s, e) =>
            {
                if (e.RowIndex < _cachedData.Count && e.ColumnIndex < _columnNames.Count)
                {
                    var row = _cachedData[e.RowIndex];
                    if (e.ColumnIndex < row.Length) e.Value = row[e.ColumnIndex];
                }
            };

            // A "Select Column" puts the grid in column-select mode. Any normal left click on a header
            // (row header, column header, corner) switches back to the usual row/cell selection first.
            // Ctrl+click on a column header while in column mode adds/removes columns instead.
            WfGrid.MouseDown += (s, e) =>
            {
                if (e.Button != WinForms.MouseButtons.Left) return;
                if (WfGrid.SelectionMode != WinForms.DataGridViewSelectionMode.ColumnHeaderSelect) return;

                var hit = WfGrid.HitTest(e.X, e.Y);
                bool ctrl = (WinForms.Control.ModifierKeys & WinForms.Keys.Control) == WinForms.Keys.Control;
                bool header = hit.Type == WinForms.DataGridViewHitTestType.ColumnHeader;

                if ((header && !ctrl) ||
                    hit.Type == WinForms.DataGridViewHitTestType.RowHeader ||
                    hit.Type == WinForms.DataGridViewHitTestType.TopLeftHeader)
                {
                    WfGrid.SelectionMode = WinForms.DataGridViewSelectionMode.RowHeaderSelect;
                }
            };

            // Left click on a header sorts, right click opens the column menu
            WfGrid.ColumnHeaderMouseClick += (_, e) =>
            {
                if (e.ColumnIndex < 0) return;

                if (e.Button == WinForms.MouseButtons.Right)
                {
                    ShowHeaderMenu(e.ColumnIndex);
                    return;
                }

                if (e.Button != WinForms.MouseButtons.Left) return;

                // Ctrl+click in column-select mode is a selection gesture, not a sort
                if (WfGrid.SelectionMode == WinForms.DataGridViewSelectionMode.ColumnHeaderSelect) return;

                _ = SortColumnAsync(e.ColumnIndex, null);
            };

            BuildHeaderMenu();
        }

        private void BuildHeaderMenu()
        {
            _headerMenu = new WinForms.ContextMenuStrip();

            var sortAsc = new WinForms.ToolStripMenuItem("Sort Ascending");
            sortAsc.Click += (s, e) => { if (_headerMenuColumn >= 0) _ = SortColumnAsync(_headerMenuColumn, true); };

            var sortDesc = new WinForms.ToolStripMenuItem("Sort Descending");
            sortDesc.Click += (s, e) => { if (_headerMenuColumn >= 0) _ = SortColumnAsync(_headerMenuColumn, false); };

            var select = new WinForms.ToolStripMenuItem("Select Column");
            select.Click += (s, e) => SelectColumn(_headerMenuColumn);

            var copy = new WinForms.ToolStripMenuItem("Copy Column");
            copy.Click += (s, e) => { if (_headerMenuColumn >= 0) CopyColumns(new[] { _headerMenuColumn }); };

            _headerMenu.Items.AddRange(new WinForms.ToolStripItem[]
            {
                sortAsc, sortDesc, new WinForms.ToolStripSeparator(), select, copy
            });
        }

        // Stock WinForms look, captured once so Light mode is exactly what it was before theming existed
        private (WinForms.DataGridViewCellStyle Cells, WinForms.DataGridViewCellStyle ColHeaders, WinForms.DataGridViewCellStyle RowHeaders,
                 System.Drawing.Color Background, System.Drawing.Color GridLines)? _gridLightLook;

        private void ApplyGridTheme()
        {
            _gridLightLook ??= (WfGrid.DefaultCellStyle.Clone(), WfGrid.ColumnHeadersDefaultCellStyle.Clone(),
                                WfGrid.RowHeadersDefaultCellStyle.Clone(), WfGrid.BackgroundColor, WfGrid.GridColor);

            if (ThemeManager.IsDark)
            {
                var panel = System.Drawing.Color.FromArgb(0x25, 0x25, 0x26);
                var header = System.Drawing.Color.FromArgb(0x2D, 0x2D, 0x30);
                var text = System.Drawing.Color.FromArgb(0xDC, 0xDC, 0xDC);
                var selection = System.Drawing.Color.FromArgb(0x26, 0x4F, 0x78);

                WfGrid.EnableHeadersVisualStyles = false; // required for custom header colors
                WfGrid.BackgroundColor = System.Drawing.Color.FromArgb(0x1E, 0x1E, 0x1E);
                WfGrid.GridColor = System.Drawing.Color.FromArgb(0x3F, 0x3F, 0x46);
                WfGrid.DefaultCellStyle.BackColor = panel;
                WfGrid.DefaultCellStyle.ForeColor = text;
                WfGrid.DefaultCellStyle.SelectionBackColor = selection;
                WfGrid.DefaultCellStyle.SelectionForeColor = System.Drawing.Color.White;

                foreach (var style in new[] { WfGrid.ColumnHeadersDefaultCellStyle, WfGrid.RowHeadersDefaultCellStyle })
                {
                    style.BackColor = header;
                    style.ForeColor = text;
                    style.SelectionBackColor = System.Drawing.Color.FromArgb(0x3E, 0x3E, 0x40);
                    style.SelectionForeColor = text;
                }
                WfGrid.ColumnHeadersBorderStyle = WinForms.DataGridViewHeaderBorderStyle.Single;
                WfGrid.RowHeadersBorderStyle = WinForms.DataGridViewHeaderBorderStyle.Single;
            }
            else
            {
                var look = _gridLightLook.Value;
                WfGrid.EnableHeadersVisualStyles = true;
                WfGrid.BackgroundColor = look.Background;
                WfGrid.GridColor = look.GridLines;
                WfGrid.DefaultCellStyle = look.Cells.Clone();
                WfGrid.ColumnHeadersDefaultCellStyle = look.ColHeaders.Clone();
                WfGrid.RowHeadersDefaultCellStyle = look.RowHeaders.Clone();
                WfGrid.ColumnHeadersBorderStyle = WinForms.DataGridViewHeaderBorderStyle.Raised;
                WfGrid.RowHeadersBorderStyle = WinForms.DataGridViewHeaderBorderStyle.Raised;
            }

            // Native scrollbars inside the grid follow the Windows dark/light scrollbar theme
            foreach (WinForms.Control child in WfGrid.Controls)
            {
                if (child is WinForms.ScrollBar) ThemeManager.ApplyNativeTheme(child.Handle);
            }

            if (_headerMenu != null)
            {
                _headerMenu.Renderer = ThemeManager.IsDark
                    ? new WinForms.ToolStripProfessionalRenderer(new DarkMenuColors())
                    : new WinForms.ToolStripProfessionalRenderer();
                foreach (WinForms.ToolStripItem item in _headerMenu.Items)
                {
                    item.ForeColor = ThemeManager.IsDark ? System.Drawing.Color.FromArgb(0xDC, 0xDC, 0xDC) : System.Drawing.SystemColors.ControlText;
                }
            }

            WfGrid.Invalidate(true);
        }

        private sealed class DarkMenuColors : WinForms.ProfessionalColorTable
        {
            private static readonly System.Drawing.Color Bg = System.Drawing.Color.FromArgb(0x2B, 0x2B, 0x2B);
            private static readonly System.Drawing.Color Hover = System.Drawing.Color.FromArgb(0x3E, 0x3E, 0x40);
            private static readonly System.Drawing.Color Line = System.Drawing.Color.FromArgb(0x55, 0x55, 0x59);

            public override System.Drawing.Color ToolStripDropDownBackground => Bg;
            public override System.Drawing.Color ImageMarginGradientBegin => Bg;
            public override System.Drawing.Color ImageMarginGradientMiddle => Bg;
            public override System.Drawing.Color ImageMarginGradientEnd => Bg;
            public override System.Drawing.Color MenuBorder => Line;
            public override System.Drawing.Color MenuItemBorder => Hover;
            public override System.Drawing.Color MenuItemSelected => Hover;
            public override System.Drawing.Color MenuItemSelectedGradientBegin => Hover;
            public override System.Drawing.Color MenuItemSelectedGradientEnd => Hover;
            public override System.Drawing.Color SeparatorDark => Line;
            public override System.Drawing.Color SeparatorLight => Bg;
        }

        private void ShowHeaderMenu(int columnIndex)
        {
            if (_headerMenu == null || columnIndex >= _columnNames.Count) return;
            _headerMenuColumn = columnIndex;
            _headerMenu.Show(WinForms.Cursor.Position);
        }

        private void SelectColumn(int columnIndex)
        {
            if (columnIndex < 0 || columnIndex >= WfGrid.Columns.Count) return;

            // Column selection is native (cheap) in ColumnHeaderSelect mode, even on huge result sets.
            // Our columns use Programmatic sort mode, which this selection mode allows.
            if (WfGrid.SelectionMode != WinForms.DataGridViewSelectionMode.ColumnHeaderSelect)
            {
                WfGrid.SelectionMode = WinForms.DataGridViewSelectionMode.ColumnHeaderSelect;
            }
            WfGrid.ClearSelection();
            WfGrid.Columns[columnIndex].Selected = true;
            WfGrid.Focus();
        }

        private async Task SortColumnAsync(int colIndex, bool? ascending)
        {
            if (_isSorting || _cachedData.Count == 0 || colIndex < 0 || colIndex >= _columnNames.Count) return;
            if (QueryTabs.SelectedItem is not TabItem tab || tab.Tag is not TabState state) return;

            bool asc = ascending ?? (colIndex == _lastSortColumnIndex ? !_sortAscending : true);
            var source = _cachedData;

            _isSorting = true;
            var bigSort = source.Count > 50_000;
            if (bigSort) WfGrid.Cursor = WinForms.Cursors.WaitCursor;

            try
            {
                // Sorting runs in the background so the UI never stalls on big result sets
                GridSort.Kind kind = GridSort.Kind.Text;
                var sorted = await Task.Run(() => GridSort.Sort(source, colIndex, asc, out kind));

                // Data was replaced (new query / cleared) while sorting: drop this result
                if (!ReferenceEquals(state.CachedData, source)) return;

                state.CachedData = sorted;
                state.LastSortColumnIndex = colIndex;
                state.SortAscending = asc;

                if (QueryTabs.SelectedItem == tab)
                {
                    _cachedData = sorted;
                    _lastSortColumnIndex = colIndex;
                    _sortAscending = asc;
                    ApplySortGlyphs();
                    WfGrid.Invalidate();
                }

                if (bigSort)
                {
                    LogStatus($"Sorted {sorted.Count} rows by '{_columnNames.ElementAtOrDefault(colIndex)}' ({kind.ToString().ToLowerInvariant()}, {(asc ? "ascending" : "descending")}).");
                }
            }
            catch (Exception ex)
            {
                LogStatus($"Sort failed: {ex.Message}");
            }
            finally
            {
                _isSorting = false;
                if (bigSort) WfGrid.Cursor = WinForms.Cursors.Default;
            }
        }

        private void ApplySortGlyphs()
        {
            for (int i = 0; i < WfGrid.Columns.Count; i++)
            {
                WfGrid.Columns[i].HeaderCell.SortGlyphDirection =
                    i != _lastSortColumnIndex ? WinForms.SortOrder.None
                    : _sortAscending ? WinForms.SortOrder.Ascending : WinForms.SortOrder.Descending;
            }
        }

        private void QueryTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Only react to the tab control itself (TextBoxes inside tabs raise bubbling SelectionChanged too)
            if (!ReferenceEquals(e.OriginalSource, QueryTabs)) return;

            // 1. Save connection credentials for any tab being deselected
            if (e.RemovedItems.Count > 0 && e.RemovedItems[0] is TabItem prevTab && prevTab.Tag is TabState prevState)
            {
                prevState.SelectedDsn = _selectedDsn;
                prevState.User = TxtUser.Text;
                prevState.Password = TxtPass.Password;
            }

            // 2. Load connection credentials and grid data for the newly selected tab
            if (QueryTabs.SelectedItem is TabItem currentTab && currentTab.Tag is TabState state)
            {
                // Sync tab credentials back to the connection pane
                _selectedDsn = state.SelectedDsn;
                BtnDatasource.Content = string.IsNullOrEmpty(_selectedDsn) ? "Select Datasource..." : NoAccessKey(_selectedDsn);
                TxtUser.Text = state.User;
                TxtPass.Password = state.Password;

                RefreshGridFromState(state);
            }
        }

        private void RefreshGridFromState(TabState state)
        {
            // Update global references used by virtual grid and sorting
            _cachedData = state.CachedData;
            _columnNames = state.ColumnNames;
            _lastSortColumnIndex = state.LastSortColumnIndex;
            _sortAscending = state.SortAscending;

            // Back to the normal selection mode before rebuilding columns
            WfGrid.SelectionMode = WinForms.DataGridViewSelectionMode.RowHeaderSelect;

            WfGrid.Columns.Clear();
            foreach (var colName in _columnNames)
            {
                var col = new WinForms.DataGridViewTextBoxColumn
                {
                    HeaderText = colName,
                    Name = colName,
                    SortMode = WinForms.DataGridViewColumnSortMode.Programmatic
                };
                WfGrid.Columns.Add(col);
            }

            WfGrid.RowCount = _cachedData.Count;
            StatusRowCount.Text = $"{_cachedData.Count} Rows";
            ApplySortGlyphs();

            WfGrid.Invalidate();
        }

        // ------------------------------------------------------------------
        // Datasources
        // ------------------------------------------------------------------

        public List<OdbcDsn> GetOdbcDataSources()
        {
            var sources = new List<OdbcDsn>
            {
                new OdbcDsn { Name = "DUMMY", SourceType = "Simulation" },
                new OdbcDsn { Name = "DUMMY_10s", SourceType = "Simulation" }
            };

            // 1. Scan User DSNs (shared registry view)
            try
            {
                using var userRoot = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default)
                    .OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources");
                if (userRoot != null)
                {
                    foreach (var name in userRoot.GetValueNames())
                    {
                        if (!sources.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                        {
                            sources.Add(new OdbcDsn { Name = name, SourceType = "User DSN" });
                        }
                    }
                }
            }
            catch { /* Ignore registry read errors */ }

            // 2. Scan System DSNs (64-bit registry view)
            try
            {
                using var sysRoot64 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64)
                    .OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources");
                if (sysRoot64 != null)
                {
                    foreach (var name in sysRoot64.GetValueNames())
                    {
                        if (!sources.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                        {
                            sources.Add(new OdbcDsn { Name = name, SourceType = "System DSN (64-bit)" });
                        }
                    }
                }
            }
            catch { /* Ignore registry read errors */ }

            // 3. Scan System DSNs (32-bit registry view under WOW64)
            try
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    using var sysRoot32 = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32)
                        .OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources");
                    if (sysRoot32 != null)
                    {
                        foreach (var name in sysRoot32.GetValueNames())
                        {
                            if (!sources.Any(s => s.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
                            {
                                sources.Add(new OdbcDsn { Name = name, SourceType = "System DSN (32-bit)" });
                            }
                        }
                    }
                }
            }
            catch { /* Ignore registry read errors */ }

            return sources;
        }

        private void BtnDatasource_Click(object sender, RoutedEventArgs e)
        {
            var sources = GetOdbcDataSources();
            DsnMenu.Items.Clear();

            // Group sources by SourceType
            var groups = sources.GroupBy(s => s.SourceType).ToList();

            foreach (var group in groups)
            {
                var groupMenu = new MenuItem
                {
                    Header = group.Key,
                    FontWeight = FontWeights.Bold
                };

                foreach (var source in group)
                {
                    var dsnName = source.Name;
                    var mi = new MenuItem { Header = NoAccessKey(dsnName) };
                    mi.Click += (s2, args) => {
                        _selectedDsn = dsnName;
                        BtnDatasource.Content = NoAccessKey(dsnName);
                        AutoFillDsnDetails(_selectedDsn);
                    };
                    groupMenu.Items.Add(mi);
                }

                DsnMenu.Items.Add(groupMenu);
            }

            DsnMenu.IsOpen = true;
        }

        /// <summary>WPF buttons/menus treat "_" as an access-key marker and hide it ("MY_DSN" shows as "MYDSN"); doubling escapes it.</summary>
        private static string NoAccessKey(string text) => text.Replace("_", "__");

        private void AutoFillDsnDetails(string dsnName)
        {
            if (dsnName == "DUMMY" || dsnName == "DUMMY_10s") return;

            var path = $@"SOFTWARE\ODBC\ODBC.INI\{dsnName}";

            // Check HKEY_CURRENT_USER first
            try
            {
                using var key = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default).OpenSubKey(path);
                if (key != null)
                {
                    var user = key.GetValue("LogonID");
                    if (user != null)
                    {
                        TxtUser.Text = user.ToString();
                        return;
                    }
                }
            }
            catch { }

            // Check HKEY_LOCAL_MACHINE (64-bit view)
            try
            {
                using var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64).OpenSubKey(path);
                if (key != null)
                {
                    var user = key.GetValue("LogonID");
                    if (user != null)
                    {
                        TxtUser.Text = user.ToString();
                        return;
                    }
                }
            }
            catch { }

            // Check HKEY_LOCAL_MACHINE (32-bit view)
            try
            {
                if (Environment.Is64BitOperatingSystem)
                {
                    using var key = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry32).OpenSubKey(path);
                    if (key != null)
                    {
                        var user = key.GetValue("LogonID");
                        if (user != null)
                        {
                            TxtUser.Text = user.ToString();
                            return;
                        }
                    }
                }
            }
            catch { }
        }

        // ------------------------------------------------------------------
        // Keyboard shortcuts
        // ------------------------------------------------------------------

        private enum Shortcut { None, Execute, ReopenTab }

        private static Shortcut MatchShortcut(Key key, ModifierKeys mods)
        {
            if (key == Key.F5 && mods == ModifierKeys.None) return Shortcut.Execute;
            if (key == Key.Enter && mods == ModifierKeys.Control) return Shortcut.Execute;
            if (key == Key.T && mods == (ModifierKeys.Control | ModifierKeys.Shift)) return Shortcut.ReopenTab;
            return Shortcut.None;
        }

        private void RunShortcut(Shortcut shortcut)
        {
            switch (shortcut)
            {
                case Shortcut.Execute:
                    // Shortcuts only start queries; stopping stays a deliberate button click
                    if (_currentCts == null) ExecuteQuery();
                    break;
                case Shortcut.ReopenTab:
                    ReopenClosedTab();
                    break;
            }
        }

        private void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            // The WinForms grid handles its own keys (see WfGrid_KeyDown)
            if (Host.IsKeyboardFocusWithin) return;

            var shortcut = MatchShortcut(e.Key, Keyboard.Modifiers);
            if (shortcut == Shortcut.None) return;

            RunShortcut(shortcut);
            e.Handled = true;
        }

        private void WfGrid_KeyDown(object? sender, WinForms.KeyEventArgs e)
        {
            if (e.Control && !e.Shift && !e.Alt && e.KeyCode == WinForms.Keys.C)
            {
                CopySelectedToCsv();
                e.Handled = true;
                return;
            }

            var mods = ModifierKeys.None;
            if (e.Control) mods |= ModifierKeys.Control;
            if (e.Shift) mods |= ModifierKeys.Shift;
            if (e.Alt) mods |= ModifierKeys.Alt;

            var key = e.KeyCode switch
            {
                WinForms.Keys.F5 => Key.F5,
                WinForms.Keys.Enter => Key.Enter,
                WinForms.Keys.T => Key.T,
                _ => Key.None
            };

            var shortcut = MatchShortcut(key, mods);
            if (shortcut == Shortcut.None) return;

            e.Handled = true;
            e.SuppressKeyPress = true;
            // Let the WinForms key event finish before touching WPF state
            Dispatcher.InvokeAsync(() => RunShortcut(shortcut));
        }

        // ------------------------------------------------------------------
        // Files: open / save / export
        // ------------------------------------------------------------------

        private void BtnOpen_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                Filter = "SQL files (*.sql)|*.sql|Text files (*.txt)|*.txt|All files (*.*)|*.*",
                FilterIndex = 1,
                Title = "Open Query File"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    var fileContent = File.ReadAllText(openFileDialog.FileName);
                    CreateNewTab(Path.GetFileName(openFileDialog.FileName), fileContent);
                    LogStatus($"Loaded query file: {Path.GetFileName(openFileDialog.FileName)}");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error reading file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var activeBox = GetCurrentQueryBox();
            // Don't bother if the box is empty
            if (activeBox == null || string.IsNullOrWhiteSpace(activeBox.Text))
            {
                MessageBox.Show("There is no query text to save.", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var saveFileDialog = new SaveFileDialog
            {
                Filter = "SQL files (*.sql)|*.sql|Text files (*.txt)|*.txt",
                FilterIndex = 1,
                Title = "Save Query As",
                DefaultExt = "sql",
                FileName = "query"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    File.WriteAllText(saveFileDialog.FileName, activeBox.Text);
                    LogStatus($"Saved query file: {Path.GetFileName(saveFileDialog.FileName)}");

                    if (QueryTabs.SelectedItem is TabItem currentTab && currentTab.Header is StackPanel sp)
                    {
                        if (sp.Children[0] is Label lbl)
                            lbl.Content = new TextBlock { Text = Path.GetFileName(saveFileDialog.FileName) };
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private async void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            if (_cachedData.Count == 0) return;

            var sfd = new SaveFileDialog { Filter = "CSV files (*.csv)|*.csv", FileName = "export.csv" };
            if (sfd.ShowDialog() != true) return;

            // Snapshot the references: sorting/clearing replaces these lists rather than modifying them
            var data = _cachedData;
            var headers = _columnNames;
            var path = sfd.FileName;

            BtnExport.IsEnabled = false;
            LogStatus($"Exporting {data.Count} rows to CSV...");
            try
            {
                await Task.Run(() => WriteCsv(path, data, headers));
                LogStatus($"Exported {data.Count} rows to {Path.GetFileName(path)}");
            }
            catch (Exception ex)
            {
                LogStatus($"Export failed: {ex.Message}");
                MessageBox.Show($"Export failed: {ex.Message}", "Export", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                BtnExport.IsEnabled = true;
            }
        }

        /// <summary>Streams rows straight to disk instead of building the whole CSV in memory first.</summary>
        private static void WriteCsv(string path, List<string[]> rows, List<string> headers)
        {
            using var writer = new StreamWriter(path, false, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true), 1 << 16);

            WriteCsvLine(writer, headers);
            foreach (var row in rows)
            {
                WriteCsvLine(writer, row);
            }
        }

        private static void WriteCsvLine(TextWriter writer, IReadOnlyList<string> fields)
        {
            for (int i = 0; i < fields.Count; i++)
            {
                if (i > 0) writer.Write(',');
                writer.Write(EscapeCsvField(fields[i]));
            }
            writer.Write("\r\n");
        }

        // ------------------------------------------------------------------
        // Clipboard
        // ------------------------------------------------------------------

        private void CopySelectedToCsv()
        {
            // Whole columns selected (via "Select Column"): copy straight from the cache, which is fast
            // even on huge result sets (enumerating SelectedCells would build an object per cell).
            if (WfGrid.SelectionMode == WinForms.DataGridViewSelectionMode.ColumnHeaderSelect && WfGrid.SelectedColumns.Count > 0)
            {
                var cols = WfGrid.SelectedColumns.Cast<WinForms.DataGridViewColumn>().Select(c => c.Index).OrderBy(i => i).ToArray();
                CopyColumns(cols);
                return;
            }

            if (WfGrid.SelectedCells.Count == 0) return;

            // Group selected cells by RowIndex to format as rows
            var cellsByRow = WfGrid.SelectedCells.Cast<WinForms.DataGridViewCell>()
                .GroupBy(c => c.RowIndex)
                .OrderBy(g => g.Key)
                .ToList();

            var csvRows = new List<string>();

            // Check if "Copy with Headers" is checked in WPF CheckBox
            bool includeHeaders = ChkCopyWithHeaders.IsChecked == true;

            if (includeHeaders)
            {
                // Find all unique selected column indices, sorted
                var selectedColIndices = WfGrid.SelectedCells.Cast<WinForms.DataGridViewCell>()
                    .Select(c => c.ColumnIndex)
                    .Distinct()
                    .OrderBy(c => c)
                    .ToList();

                var headerRow = selectedColIndices.Select(colIndex => colIndex < _columnNames.Count ? _columnNames[colIndex] : $"Col{colIndex}");
                csvRows.Add(string.Join(",", headerRow.Select(EscapeCsvField)));
            }

            foreach (var rowGroup in cellsByRow)
            {
                // Order cells in this row by ColumnIndex
                var orderedCells = rowGroup.OrderBy(c => c.ColumnIndex).ToList();
                var rowData = orderedCells.Select(c => c.Value?.ToString() ?? "");
                csvRows.Add(string.Join(",", rowData.Select(EscapeCsvField)));
            }

            var csv = string.Join(Environment.NewLine, csvRows);

            if (!string.IsNullOrEmpty(csv))
            {
                System.Windows.Clipboard.SetText(csv);
                LogStatus($"Copied {WfGrid.SelectedCells.Count} selected cells to clipboard.");
            }
        }

        private void CopyColumns(int[] columnIndices)
        {
            var cols = columnIndices.Where(i => i >= 0 && i < _columnNames.Count).ToArray();
            if (cols.Length == 0) return;

            var sb = new StringBuilder();
            if (ChkCopyWithHeaders.IsChecked == true)
            {
                sb.Append(string.Join(",", cols.Select(i => EscapeCsvField(_columnNames[i]))));
                sb.Append(Environment.NewLine);
            }

            foreach (var row in _cachedData)
            {
                for (int c = 0; c < cols.Length; c++)
                {
                    if (c > 0) sb.Append(',');
                    sb.Append(EscapeCsvField(cols[c] < row.Length ? row[cols[c]] : ""));
                }
                sb.Append(Environment.NewLine);
            }

            try
            {
                System.Windows.Clipboard.SetText(sb.ToString());
                var name = cols.Length == 1 ? $"column '{_columnNames[cols[0]]}'" : $"{cols.Length} columns";
                LogStatus($"Copied {name} ({_cachedData.Count} rows) to clipboard.");
            }
            catch (Exception ex)
            {
                LogStatus($"Copy failed: {ex.Message}");
            }
        }

        private static string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains(',') || field.Contains('"') || field.Contains('\n') || field.Contains('\r'))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        // ------------------------------------------------------------------
        // Tabs
        // ------------------------------------------------------------------

        private TextBox? GetCurrentQueryBox()
        {
            if (QueryTabs.SelectedItem is TabItem currentTab)
            {
                return currentTab.Content as TextBox;
            }
            return null;
        }

        private static string GetTabTitle(TabItem tab)
        {
            if (tab.Header is StackPanel sp && sp.Children.Count > 0 && sp.Children[0] is Label lbl)
            {
                return (lbl.Content as TextBlock)?.Text ?? lbl.Content?.ToString() ?? "Query";
            }
            return "Query";
        }

        private void AddNewTab_Click(object sender, RoutedEventArgs e)
        {
            CreateNewTab("New Query", "select * from ");
        }

        private void CreateNewTab(string header, string content)
        {
            var initialTabState = new TabState
            {
                SelectedDsn = _selectedDsn,
                User = TxtUser.Text,
                Password = TxtPass.Password
            };
            CreateNewTabInternal(header, content, initialTabState);
        }

        private void CreateNewTabInternal(string header, string content, TabState state)
        {
            var newBox = new TextBox
            {
                Text = content,
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                FontSize = 13,
                AcceptsReturn = true,
                Padding = new Thickness(8),
                BorderThickness = new Thickness(0),
                TextWrapping = TextWrapping.Wrap,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto
            };

            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };

            // FIX: Using Label because TextBlock does NOT have MouseDoubleClick
            var titleLabel = new Label
            {
                Content = new TextBlock { Text = header }, // TextBlock, not a string: keeps "_" in names visible
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(0),
                MinWidth = 40
            };

            void StartRename()
            {
                if (!headerStack.Children.Contains(titleLabel)) return;

                var editBox = new TextBox { Text = (titleLabel.Content as TextBlock)?.Text ?? "", MinWidth = 60 };
                headerStack.Children.Insert(0, editBox);
                headerStack.Children.Remove(titleLabel);
                editBox.Focus();
                editBox.SelectAll();

                void FinishRename()
                {
                    if (headerStack.Children.Contains(editBox))
                    {
                        titleLabel.Content = new TextBlock { Text = string.IsNullOrWhiteSpace(editBox.Text) ? "Untitled" : editBox.Text };
                        headerStack.Children.Insert(0, titleLabel);
                        headerStack.Children.Remove(editBox);
                    }
                }

                editBox.KeyDown += (s2, e2) => { if (e2.Key == Key.Enter) FinishRename(); };
                editBox.LostFocus += (s2, e2) => FinishRename();
            }

            titleLabel.MouseDoubleClick += (s, e) => StartRename();

            var closeBtn = new Button
            {
                Content = "×",
                Margin = new Thickness(10, 0, 0, 0),
                Padding = new Thickness(2, 0, 2, 0),
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center,
                ToolTip = "Close tab (Ctrl+Shift+T reopens it)"
            };

            var newTab = new TabItem { Content = newBox, Tag = state };

            closeBtn.Click += (s, e) => CloseTab(newTab);

            // Right-click menu on the tab header
            var renameItem = new MenuItem { Header = "Rename" };
            renameItem.Click += (s, e) => StartRename();
            var clearItem = new MenuItem { Header = "Clear Results", ToolTip = "Free this tab's result data (query text is kept)" };
            clearItem.Click += (s, e) => ClearResults(newTab);
            var closeItem = new MenuItem { Header = "Close Tab" };
            closeItem.Click += (s, e) => CloseTab(newTab);
            var reopenItem = new MenuItem { Header = "Reopen Closed Tab", InputGestureText = "Ctrl+Shift+T" };
            reopenItem.Click += (s, e) => ReopenClosedTab();

            var tabMenu = new ContextMenu();
            tabMenu.Items.Add(renameItem);
            tabMenu.Items.Add(clearItem);
            tabMenu.Items.Add(new Separator());
            tabMenu.Items.Add(closeItem);
            tabMenu.Items.Add(reopenItem);
            tabMenu.Opened += (s, e) =>
            {
                closeItem.IsEnabled = QueryTabs.Items.Count > 1;
                reopenItem.IsEnabled = _closedTabs.Count > 0;
                clearItem.IsEnabled = state.CachedData.Count > 0 || state.ColumnNames.Count > 0;
            };
            headerStack.ContextMenu = tabMenu;

            headerStack.Children.Add(titleLabel);
            headerStack.Children.Add(closeBtn);
            newTab.Header = headerStack;

            QueryTabs.Items.Add(newTab);
            QueryTabs.SelectedItem = newTab;
        }

        private void CloseTab(TabItem tab)
        {
            if (QueryTabs.Items.Count <= 1 || !QueryTabs.Items.Contains(tab)) return;

            QueryTabs.Items.Remove(tab);

            // Keep it (with its results) so Ctrl+Shift+T can bring it back; older ones are dropped to free memory
            _closedTabs.AddFirst(tab);
            while (_closedTabs.Count > MaxClosedTabs)
            {
                var dropped = _closedTabs.Last!.Value;
                _closedTabs.RemoveLast();
                if (dropped.Tag is TabState droppedState)
                {
                    droppedState.CachedData = new List<string[]>();
                    droppedState.ColumnNames = new List<string>();
                }
            }

            LogStatus($"Closed '{GetTabTitle(tab)}'. Press Ctrl+Shift+T to reopen.");
        }

        private void ReopenClosedTab()
        {
            if (_closedTabs.Count == 0) return;

            var tab = _closedTabs.First!.Value;
            _closedTabs.RemoveFirst();

            QueryTabs.Items.Add(tab);
            QueryTabs.SelectedItem = tab;
            LogStatus($"Reopened '{GetTabTitle(tab)}'.");
        }

        private void ClearResults(TabItem tab)
        {
            if (tab.Tag is not TabState state) return;

            var rowCount = state.CachedData.Count;
            state.CachedData = new List<string[]>();
            state.ColumnNames = new List<string>();
            state.LastSortColumnIndex = -1;
            state.SortAscending = true;

            if (QueryTabs.SelectedItem == tab)
            {
                RefreshGridFromState(state);
            }

            // Big result sets live on the large object heap; collect now so the memory is actually returned
            GC.Collect();
            LogStatus($"Cleared {rowCount} rows from '{GetTabTitle(tab)}'.");
        }

        // ------------------------------------------------------------------
        // Status bar: logs, settings, feedback, updates
        // ------------------------------------------------------------------

        private void StatusInfo_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var logWindow = new Window
            {
                Title = "ODBSee Log History",
                Width = 650,
                Height = 450,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                ShowInTaskbar = false
            };
            logWindow.SetResourceReference(BackgroundProperty, "Brush.Panel");
            logWindow.SetResourceReference(ForegroundProperty, "Brush.Text");

            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var header = new TextBlock
            {
                Text = "ODBSee Activity Logs (Last 50 messages)",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 10)
            };
            Grid.SetRow(header, 0);
            grid.Children.Add(header);

            string logsText;
            lock (_logLock)
            {
                logsText = string.Join(Environment.NewLine, _logHistory);
            }

            var logBox = new TextBox
            {
                Text = string.IsNullOrEmpty(logsText) ? "No log history recorded yet." : logsText,
                IsReadOnly = true,
                FontFamily = new System.Windows.Media.FontFamily("Consolas"),
                FontSize = 11,
                AcceptsReturn = true,
                VerticalScrollBarVisibility = ScrollBarVisibility.Auto,
                HorizontalScrollBarVisibility = ScrollBarVisibility.Auto,
                Padding = new Thickness(8),
                BorderThickness = new Thickness(1)
            };
            logBox.SetResourceReference(BackgroundProperty, "Brush.Code");
            logBox.SetResourceReference(ForegroundProperty, "Brush.Text");
            logBox.SetResourceReference(BorderBrushProperty, "Brush.BorderStrong");
            Grid.SetRow(logBox, 1);
            grid.Children.Add(logBox);

            var btnPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new Thickness(0, 12, 0, 0)
            };
            Grid.SetRow(btnPanel, 2);
            grid.Children.Add(btnPanel);

            var copyBtn = new Button
            {
                Content = "Copy to Clipboard",
                Padding = new Thickness(12, 6, 12, 6),
                Margin = new Thickness(0, 0, 10, 0)
            };
            copyBtn.Click += (s2, e2) =>
            {
                try
                {
                    System.Windows.Clipboard.SetText(logsText);
                    MessageBox.Show("Logs copied to clipboard!", "Copied", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed to copy logs: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            };
            btnPanel.Children.Add(copyBtn);

            var closeBtn = new Button
            {
                Content = "Close",
                Padding = new Thickness(18, 6, 18, 6),
                IsDefault = true
            };
            closeBtn.Click += (s2, e2) => logWindow.Close();
            btnPanel.Children.Add(closeBtn);

            logWindow.Content = grid;
            logWindow.ShowDialog();
        }

        private void BtnSettings_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var window = new SettingsWindow(_settings) { Owner = this };
            if (window.ShowDialog() == true)
            {
                // Re-save the session so a change to "Remember passwords" takes effect on disk right away
                SaveSession();
                if (_settings.Theme != ThemeManager.CurrentSetting) ThemeManager.Apply(_settings.Theme);
                LogStatus($"Settings saved (connect timeout {_settings.ConnectTimeoutSeconds}s, query timeout {_settings.CommandTimeoutSeconds}s).");
            }
        }

        private void BtnFeedback_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            var window = new FeedbackWindow(_settings.GitHubToken) { Owner = this };
            if (window.ShowDialog() == true && !string.IsNullOrEmpty(window.ResultMessage))
            {
                LogStatus(window.ResultMessage);
            }
        }

        private async Task RetryPendingFeedbackAsync()
        {
            try
            {
                var token = _settings.GitHubToken;
                if (string.IsNullOrWhiteSpace(token) || FeedbackService.PendingCount == 0) return;

                var (sent, remaining) = await FeedbackService.RetryPendingAsync(token);
                if (sent > 0) LogStatus($"Sent {sent} saved feedback item(s) to GitHub.");
                if (remaining > 0) LogStatus($"{remaining} feedback item(s) still waiting to be sent.");
            }
            catch (Exception ex)
            {
                LogStatus($"Retrying saved feedback failed: {ex.Message}");
            }
        }

        private async void BtnCheckUpdate_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            LogStatus("Checking for updates on GitHub...");
            BtnCheckUpdate.Text = "Checking...";
            BtnCheckUpdate.Foreground = System.Windows.Media.Brushes.Gray;

            try
            {
                using var client = new HttpClient();
                // GitHub API requires a User-Agent header
                client.DefaultRequestHeaders.UserAgent.ParseAdd("ODBSee-Updater");

                var url = "https://api.github.com/repos/daijoubu-dev/ODBSee/releases/latest";
                var response = await client.GetStringAsync(url);

                using var doc = JsonDocument.Parse(response);
                var root = doc.RootElement;

                if (root.TryGetProperty("tag_name", out var tagProperty) &&
                    root.TryGetProperty("html_url", out var urlProperty))
                {
                    var latestTag = tagProperty.GetString() ?? "";
                    var htmlUrl = urlProperty.GetString() ?? "https://github.com/daijoubu-dev/ODBSee/releases";

                    var cleanTag = latestTag.TrimStart('v', 'V');
                    if (Version.TryParse(cleanTag, out var latestVersion))
                    {
                        var currentVersion = Assembly.GetExecutingAssembly().GetName().Version ?? new Version(1, 0, 0, 0);

                        if (latestVersion > currentVersion)
                        {
                            LogStatus($"New version available: {latestTag}!");
                            BtnCheckUpdate.Text = $"Update Available! ({latestTag})";
                            BtnCheckUpdate.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#D13438"));
                            BtnCheckUpdate.ToolTip = "Click to open GitHub download page";

                            // Re-bind click event to open URL
                            BtnCheckUpdate.MouseLeftButtonUp -= BtnCheckUpdate_MouseLeftButtonUp;
                            BtnCheckUpdate.MouseLeftButtonUp += (s2, e2) =>
                            {
                                try
                                {
                                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                                    {
                                        FileName = htmlUrl,
                                        UseShellExecute = true
                                    });
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"Could not open release URL: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                                }
                            };
                        }
                        else
                        {
                            LogStatus($"ODBSee is up to date (Version {currentVersion.ToString(3)}).");
                            BtnCheckUpdate.Text = "Up to Date";
                            BtnCheckUpdate.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#228B22"));
                            BtnCheckUpdate.ToolTip = "You are running the latest version";
                        }
                    }
                    else
                    {
                        throw new Exception("Could not parse latest version tag.");
                    }
                }
                else
                {
                    throw new Exception("Latest release tag not found in response.");
                }
            }
            catch (Exception ex)
            {
                LogStatus($"Update check failed: {ex.Message}");
                BtnCheckUpdate.Text = "Check for Updates";
                BtnCheckUpdate.SetResourceReference(ForegroundProperty, "Brush.Link");
                BtnCheckUpdate.ToolTip = "Click to check for updates on GitHub";
                MessageBox.Show($"Update check failed: {ex.Message}{Environment.NewLine}Please check your internet connection.", "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}
