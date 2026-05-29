using System;
using System.Globalization;
using Microsoft.Win32;
using System.Data.Odbc;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;
using Button = System.Windows.Controls.Button;
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
        
        // These now act as references to the current tab's state
        private List<string[]> _cachedData = new List<string[]>();
        private List<string> _columnNames = new List<string>();

        // Trackers for sorting (will be updated from/to tab state)
        private int _lastSortColumnIndex = -1;
        private bool _sortAscending = true;

        private OdbcCommand? _currentCommand;
        private CancellationTokenSource? _currentCts;
        private Task? _runningQueryTask;

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
            public string Password { get; set; } = "";
        }

        public MainWindow()
        {
            InitializeComponent();
            SetupWinFormsGrid();

            // Set up hover underline dynamic micro-animation for status bar
            StatusInfo.MouseEnter += (s, e) => StatusInfo.TextDecorations = TextDecorations.Underline;
            StatusInfo.MouseLeave += (s, e) => StatusInfo.TextDecorations = null;

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

            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => StatusInfo.Text = message);
            }
            else
            {
                StatusInfo.Text = message;
            }
        }

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
                        string headerText = "Query";
                        if (tab.Header is StackPanel sp && sp.Children[0] is Label lbl)
                        {
                            headerText = lbl.Content?.ToString() ?? "Query";
                        }

                        session.Tabs.Add(new TabSessionInfo
                        {
                            Header = headerText,
                            QueryText = tb.Text,
                            SelectedDsn = state.SelectedDsn,
                            User = state.User,
                            Password = state.Password
                        });
                    }
                }

                var appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ODBSee");
                if (!Directory.Exists(appDataPath))
                {
                    Directory.CreateDirectory(appDataPath);
                }

                var filePath = Path.Combine(appDataPath, "session_state.json");
                var json = JsonSerializer.Serialize(session, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json, Encoding.UTF8);
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
                var appDataPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ODBSee");
                var filePath = Path.Combine(appDataPath, "session_state.json");
                if (!File.Exists(filePath)) return false;

                var json = File.ReadAllText(filePath, Encoding.UTF8);
                var session = JsonSerializer.Deserialize<SessionState>(json);
                if (session == null || session.Tabs.Count == 0) return false;

                QueryTabs.Items.Clear();
                foreach (var tabInfo in session.Tabs)
                {
                    CreateNewTabInternal(tabInfo.Header, tabInfo.QueryText, new TabState
                    {
                        SelectedDsn = tabInfo.SelectedDsn,
                        User = tabInfo.User,
                        Password = tabInfo.Password
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

                LogStatus($"Restored {session.Tabs.Count} queries from previous session.");
                return true;
            }
            catch (Exception ex)
            {
                LogStatus($"Failed to load session state: {ex.Message}");
                return false;
            }
        }

        private void SetQueryRunningState(bool isRunning)
        {
            if (isRunning)
            {
                BtnExecute.Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#D13438"));
                
                if (BtnExecute.Content is TextBlock tb)
                {
                    tb.Text = "■ Stop";
                }
                else
                {
                    BtnExecute.Content = new TextBlock { Text = "■ Stop" };
                }
            }
            else
            {
                BtnExecute.Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#007ACC"));
                
                if (BtnExecute.Content is TextBlock tb)
                {
                    tb.Text = "▶ Execute";
                }
                else
                {
                    BtnExecute.Content = new TextBlock { Text = "▶ Execute" };
                }
            }
        }

        private void CancelQuery()
        {
            if (_currentCts != null)
            {
                LogStatus("Query cancellation requested by user...");
                try
                {
                    _currentCts.Cancel();
                    _currentCommand?.Cancel();
                }
                catch (Exception ex)
                {
                    LogStatus($"Error while calling Cancel on command: {ex.Message}");
                }
            }
        }

        private void StatusInfo_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            var logWindow = new Window
            {
                Title = "ODBSee Log History",
                Width = 650,
                Height = 450,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                Owner = this,
                Background = System.Windows.Media.Brushes.White,
                ShowInTaskbar = false
            };

            var grid = new Grid { Margin = new Thickness(12) };
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            grid.RowDefinitions.Add(new RowDefinition { Height = new GridLength(1, GridUnitType.Star) });
            grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var header = new TextBlock
            {
                Text = "ODBSee Activity Logs (Last 50 messages)",
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                Margin = new Thickness(0, 0, 0, 10),
                Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#333333"))
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
                Background = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#F8F9FA")),
                Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#212529")),
                Padding = new Thickness(8),
                BorderBrush = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#CED4DA")),
                BorderThickness = new Thickness(1)
            };
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

            //listener for manual copy method
            WfGrid.KeyDown += WfGrid_KeyDown;

            typeof(DataGridView).GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
                ?.SetValue(WfGrid, true, null);

            WfGrid.CellValueNeeded += (s, e) =>
            {
                if (e.RowIndex < _cachedData.Count && e.ColumnIndex < _columnNames.Count)
                {
                    e.Value = _cachedData[e.RowIndex][e.ColumnIndex];
                }
            };

            
            WfGrid.ColumnHeaderMouseClick += (_, e) =>
            {
                if (_cachedData.Count == 0) return;

                var colIndex = e.ColumnIndex;

                
                if (colIndex == _lastSortColumnIndex)
                {
                    _sortAscending = !_sortAscending;
                }
                else
                {
                    _sortAscending = true;
                    _lastSortColumnIndex = colIndex;
                }

                // Perform the sort and sync back to tab state
                if (_sortAscending)
                {
                    _cachedData = _cachedData.OrderBy(x => x[colIndex]).ToList();
                }
                else
                {
                    _cachedData = _cachedData.OrderByDescending(x => x[colIndex]).ToList();
                }

                // Sync back to the active tab state so it persists
                if (QueryTabs.SelectedItem is TabItem currentTab && currentTab.Tag is TabState state)
                {
                    state.CachedData = _cachedData;
                    state.LastSortColumnIndex = _lastSortColumnIndex;
                    state.SortAscending = _sortAscending;
                }

                WfGrid.Invalidate();
            };
        }

        private void QueryTabs_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
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
                BtnDatasource.Content = string.IsNullOrEmpty(_selectedDsn) ? "Select Datasource..." : _selectedDsn;
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
            
            WfGrid.Invalidate();
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            ExecuteQuery();
        }

        private async void ExecuteQuery()
        {
            // Cancel any existing execution
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

            var queryText = activeBox.Text;
            var userId = TxtUser.Text;
            var password = TxtPass.Password;
            int.TryParse(TxtMaxRows.Text, out var maxRows);
            if (maxRows <= 0) maxRows = 1000;

            LogStatus($"Executing query on DSN '{_selectedDsn}' (Limit: {maxRows} rows)...");
            var connString = $"DSN={_selectedDsn};Uid={userId};Pwd={password};Connect Timeout=300;";

            _currentCts = new CancellationTokenSource();
            var token = _currentCts.Token;

            // Set UI executing state
            SetQueryRunningState(true);

            try
            {
                if (_selectedDsn == "DUMMY" || _selectedDsn == "DUMMY_10s")
                {
                    bool isLongDelay = _selectedDsn == "DUMMY_10s";
                    _runningQueryTask = Task.Run(() => GenerateDummyData(maxRows, isLongDelay), token);
                }
                else
                {
                    _runningQueryTask = Task.Run(() => FetchData(connString, queryText, maxRows), token);
                }

                var (columns, rows) = await (Task<(List<string>, List<string[]>)>)_runningQueryTask;

                state.ColumnNames = columns;
                state.CachedData = rows;
                state.LastSortColumnIndex = -1;
                state.SortAscending = true;

                // Only update the grid if the tab that initiated the query is still the active one
                if (QueryTabs.SelectedItem == activeTab)
                {
                    RefreshGridFromState(state);
                    StatusConnection.Text = "Connected";
                    LogStatus($"Query executed successfully: returned {rows.Count} rows.");
                }
                else
                {
                    LogStatus($"Background query finished: returned {rows.Count} rows.");
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
                if (_currentCts?.Token == token)
                {
                    _currentCts = null;
                    _runningQueryTask = null;
                }
                SetQueryRunningState(false);
            }
        }

        private (List<string> Columns, List<string[]> Rows) GenerateDummyData(int maxRows, bool isLongDelay)
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
                if (_currentCts?.IsCancellationRequested == true)
                    throw new OperationCanceledException();
                
                Thread.Sleep(intervalMs);
                elapsedMs += intervalMs;
            }

            for (int i = 1; i <= maxRows; i++)
            {
                if (_currentCts?.IsCancellationRequested == true)
                    throw new OperationCanceledException();

                rows.Add(new string[] { 
                    i.ToString(), 
                    $"Dummy Record {i}", 
                    rand.Next(0, 2) == 0 ? "Production" : "Staging",
                    (rand.NextDouble() * 1000).ToString("F2"), 
                    rand.Next(0, 3) switch { 0 => "Active", 1 => "Pending", _ => "Archived" },
                    DateTime.Now.AddDays(-rand.Next(1, 30)).ToString("yyyy-MM-dd HH:mm:ss")
                });
            }

            return (columns, rows);
        }

        private (List<string> Columns, List<string[]> Rows) FetchData(string connString, string sql, int maxRows)
        {
            var cols = new List<string>();
            var rows = new List<string[]>();
            using var conn = new OdbcConnection(connString);
            conn.Open();
            using var cmd = new OdbcCommand(sql, conn);
            _currentCommand = cmd;
            cmd.CommandTimeout = 300; //5 mins
            using var reader = cmd.ExecuteReader();
            for (var i = 0; i < reader.FieldCount; i++)
                cols.Add(reader.GetName(i));
            int count = 0;
            while (reader.Read() && count < maxRows)
            {
                if (_currentCts?.IsCancellationRequested == true)
                    throw new OperationCanceledException();

                string[] row = new string[reader.FieldCount];
                for (int i = 0; i < reader.FieldCount; i++)
                    row[i] = reader[i]?.ToString() ?? "";
                rows.Add(row);
                count++;
            }
            _currentCommand = null;
            return (cols, rows);
        }

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
                    var mi = new MenuItem { Header = source.Name };
                    mi.Click += (s2, args) => {
                        _selectedDsn = mi.Header.ToString() ?? "";
                        BtnDatasource.Content = _selectedDsn;
                        AutoFillDsnDetails(_selectedDsn);
                    };
                    groupMenu.Items.Add(mi);
                }

                DsnMenu.Items.Add(groupMenu);
            }

            DsnMenu.IsOpen = true;
        }

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

        private void TxtQuery_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.Enter &&
                (System.Windows.Input.Keyboard.Modifiers & System.Windows.Input.ModifierKeys.Control) == System.Windows.Input.ModifierKeys.Control)
            {
                ExecuteQuery();
                e.Handled = true;
            }
        }

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
                            lbl.Content = Path.GetFileName(saveFileDialog.FileName);
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
            if (sfd.ShowDialog() == true)
            {
                LogStatus("Exporting data to CSV...");
                
                var csv = await Task.Run(() => GenerateCsv(_cachedData, _columnNames));

                File.WriteAllText(sfd.FileName, csv, Encoding.UTF8);
                LogStatus($"Exported result set successfully to {Path.GetFileName(sfd.FileName)}");
            }
        }

        //Manual copy method
        private void WfGrid_KeyDown(object sender, WinForms.KeyEventArgs e)
        {
            switch (e.Control)
            {

                case (true):

                    switch (e.KeyCode)
                    {
                        case (Keys.C):
                            CopySelectedToCsv();
                            e.Handled = true;
                            break;
                    }

                    break;
            }
        }



        private void CopySelectedToCsv()
        {
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


        //CSV generation helpers
        private string GenerateCsv(IEnumerable<string[]> rows, IEnumerable<string> headers = null)
        {
            var sb = new StringBuilder();

            // Add headers
            if (headers != null)
            {
                sb.AppendLine(string.Join(",", headers.Select(EscapeCsvField)));
            }

            // Add rows
            foreach (var row in rows)
            {
                sb.AppendLine(string.Join(",", row.Select(EscapeCsvField)));
            }

            return sb.ToString();
        }

        private string EscapeCsvField(string field)
        {
            if (string.IsNullOrEmpty(field)) return "";
            if (field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r"))
            {
                return $"\"{field.Replace("\"", "\"\"")}\"";
            }
            return field;
        }

        private TextBox GetCurrentQueryBox()
        {
            if (QueryTabs.SelectedItem is TabItem currentTab)
            {
                // If you set the TextBox as the Content of the TabItem
                return currentTab.Content as TextBox;
            }
            return null;
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
            newBox.KeyDown += TxtQuery_KeyDown;

            var headerStack = new StackPanel { Orientation = Orientation.Horizontal };

            // FIX: Using Label because TextBlock does NOT have MouseDoubleClick
            var titleLabel = new Label
            {
                Content = header,
                VerticalAlignment = VerticalAlignment.Center,
                Padding = new Thickness(0),
                MinWidth = 40
            };

            titleLabel.MouseDoubleClick += (s, e) =>
            {
                var editBox = new TextBox { Text = titleLabel.Content.ToString(), MinWidth = 60 };
                headerStack.Children.Insert(0, editBox);
                headerStack.Children.Remove(titleLabel);
                editBox.Focus();
                editBox.SelectAll();

                void FinishRename()
                {
                    if (headerStack.Children.Contains(editBox))
                    {
                        titleLabel.Content = string.IsNullOrWhiteSpace(editBox.Text) ? "Untitled" : editBox.Text;
                        headerStack.Children.Insert(0, titleLabel);
                        headerStack.Children.Remove(editBox);
                    }
                }

                editBox.KeyDown += (s2, e2) => { if (e2.Key == System.Windows.Input.Key.Enter) FinishRename(); };
                editBox.LostFocus += (s2, e2) => FinishRename();
            };

            var closeBtn = new Button
            {
                Content = "×",
                Margin = new Thickness(10, 0, 0, 0),
                Padding = new Thickness(2, 0, 2, 0),
                Background = System.Windows.Media.Brushes.Transparent,
                BorderThickness = new Thickness(0),
                FontWeight = FontWeights.Bold,
                VerticalAlignment = VerticalAlignment.Center
            };

            var newTab = new TabItem { Content = newBox, Tag = state };

            closeBtn.Click += (s, e) =>
            {
                if (QueryTabs.Items.Count > 1)
                    QueryTabs.Items.Remove(newTab);
            };

            headerStack.Children.Add(titleLabel);
            headerStack.Children.Add(closeBtn);
            newTab.Header = headerStack;

            QueryTabs.Items.Add(newTab);
            QueryTabs.SelectedItem = newTab;
        }

        private async void BtnCheckUpdate_MouseLeftButtonUp(object sender, System.Windows.Input.MouseButtonEventArgs e)
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
                BtnCheckUpdate.Foreground = new System.Windows.Media.SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#007ACC"));
                BtnCheckUpdate.ToolTip = "Click to check for updates on GitHub";
                MessageBox.Show($"Update check failed: {ex.Message}{Environment.NewLine}Please check your internet connection.", "Update Check Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
    }
}