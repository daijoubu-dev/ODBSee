using Microsoft.Win32;
using System.Data.Odbc;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using Button = System.Windows.Controls.Button;
using Label = System.Windows.Controls.Label;
using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Orientation = System.Windows.Controls.Orientation;
using SaveFileDialog = Microsoft.Win32.SaveFileDialog;
using TextBox = System.Windows.Controls.TextBox;
using WinForms = System.Windows.Forms;

namespace ODBSee
{
    public partial class MainWindow : Window
    {
        private string _selectedDsn = "";
        private List<string[]> _cachedData = new List<string[]>();
        private List<string> _columnNames = new List<string>();

        // Trackers for sorting
        private int _lastSortColumnIndex = -1;
        private bool _sortAscending = true;

        private OdbcCommand _currentCommand;

        public MainWindow()
        {
            InitializeComponent();
            SetupWinFormsGrid();

            QueryTabs.Items.Clear(); // Remove the placeholder from XAML
            CreateNewTab("Query 1", "select * from ");
        }

        private void SetupWinFormsGrid()
        {
            WfGrid.VirtualMode = true;
            WfGrid.ReadOnly = true;
            WfGrid.AllowUserToAddRows = false;
            WfGrid.RowHeadersVisible = false;
            WfGrid.AllowUserToResizeRows = false;
            WfGrid.AllowUserToResizeColumns = true;
            WfGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.Disable; //disable native copy so we can do our own functional copy
            WfGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            WfGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

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

                // Perform the sort
                if (_sortAscending)
                {
                    _cachedData = _cachedData.OrderBy(x => x[colIndex]).ToList();
                }
                else
                {
                    _cachedData = _cachedData.OrderByDescending(x => x[colIndex]).ToList();
                }

                WfGrid.Invalidate();
            };
        }

        private void BtnExecute_Click(object sender, RoutedEventArgs e)
        {
            ExecuteQuery();
        }

        private async void ExecuteQuery()
        {
            if (_currentCommand != null)
            {
                try
                {
                    _currentCommand.Cancel();
                    StatusInfo.Text = "Previous query cancelled. Starting new...";
                }
                catch { /* Ignore cancellation errors */ }
            }

            var activeBox = GetCurrentQueryBox();
            if (activeBox == null || string.IsNullOrEmpty(_selectedDsn)) return;

            var queryText = activeBox.Text;
            var userId = TxtUser.Text;
            var password = TxtPass.Password;
            int.TryParse(TxtMaxRows.Text, out var maxRows);
            if (maxRows <= 0) maxRows = 1000;

            StatusInfo.Text = "Executing Query...";
            var connString = $"DSN={_selectedDsn};Uid={userId};Pwd={password};Connect Timeout=300;";

            try
            {
                var result = await Task.Run(() => FetchData(connString, queryText, maxRows));

                _columnNames = result.Columns;
                _cachedData = result.Rows;

                // Reset sort trackers for the new result set
                _lastSortColumnIndex = -1;
                _sortAscending = true;

                WfGrid.Columns.Clear();
                foreach (var colName in _columnNames)
                {
                    // Set SortMode to Programmatic for performance (so the grid doesn't try to handle it internally)
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
                StatusConnection.Text = "Connected";
                StatusInfo.Text = "Ready";
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                StatusInfo.Text = "Error";
            }
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
                string[] row = new string[reader.FieldCount];
                for (int i = 0; i < reader.FieldCount; i++)
                    row[i] = reader[i]?.ToString() ?? "";
                rows.Add(row);
                count++;
            }
            _currentCommand = null;
            return (cols, rows);
        }

        public List<string> GetOdbcDataSources()
        {
            var sources = new List<string>();
            using (var root = Registry.CurrentUser.OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources"))
                if (root != null) foreach (var name in root.GetValueNames()) sources.Add(name);
            using (var root = Registry.LocalMachine.OpenSubKey(@"Software\ODBC\ODBC.INI\ODBC Data Sources"))
                if (root != null) foreach (var name in root.GetValueNames()) sources.Add(name);
            return sources;
        }

        private void BtnDatasource_Click(object sender, RoutedEventArgs e)
        {
            var sources = GetOdbcDataSources();
            DsnMenu.Items.Clear();
            foreach (var source in sources)
            {
                var mi = new MenuItem { Header = source };
                mi.Click += (s, args) => {
                    _selectedDsn = mi.Header.ToString();
                    BtnDatasource.Content = _selectedDsn;
                    AutoFillDsnDetails(_selectedDsn);
                };
                DsnMenu.Items.Add(mi);
            }
            DsnMenu.IsOpen = true;
        }

        private void AutoFillDsnDetails(string dsnName)
        {
            var path = $@"SOFTWARE\ODBC\ODBC.INI\{dsnName}";
            using var key = Registry.CurrentUser.OpenSubKey(path) ?? Registry.LocalMachine.OpenSubKey(path);
            if (key != null)
            {
                var user = key.GetValue("LogonID");
                if (user != null) TxtUser.Text = user.ToString();
            }
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
                    StatusInfo.Text = $"Loaded: {Path.GetFileName(openFileDialog.FileName)}";
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
                    StatusInfo.Text = $"Saved: {Path.GetFileName(saveFileDialog.FileName)}";

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
                StatusInfo.Text = "Exporting...";
                
                var csv = await Task.Run(() => GenerateCsv(_cachedData, _columnNames));

                File.WriteAllText(sfd.FileName, csv, Encoding.UTF8);
                StatusInfo.Text = $"Exported {Path.GetFileName(sfd.FileName)}";
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
            if (WfGrid.SelectedRows.Count == 0) return;

            var selectedData = WfGrid.SelectedRows.Cast<WinForms.DataGridViewRow>()
                .OrderBy(r => r.Index)
                .Select(r => _cachedData[r.Index]);

            var csv = GenerateCsv(selectedData);

            if (!string.IsNullOrEmpty(csv))
            {
                System.Windows.Clipboard.SetText(csv);
                StatusInfo.Text = "Copied selected rows to clipboard";
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

        // Logic for the [+] Button
        private void AddNewTab_Click(object sender, RoutedEventArgs e)
        {
            CreateNewTab("New Query", "select * from ");
        }

        private void CreateNewTab(string header, string content)
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

            // DOUBLE CLICK TO RENAME LOGIC
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

            // Create the TabItem first so the button can reference it
            var newTab = new TabItem { Content = newBox };

            // FIX: Remove 'newTab' specifically, not just the 'SelectedItem'
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
    }
}