using System.Windows;

namespace ODBSee
{
    public partial class SettingsWindow : Window
    {
        private readonly AppSettings _settings;

        public SettingsWindow(AppSettings settings)
        {
            InitializeComponent();
            _settings = settings;

            TxtConnectTimeout.Text = settings.ConnectTimeoutSeconds.ToString();
            TxtCommandTimeout.Text = settings.CommandTimeoutSeconds.ToString();
            TxtDefaultMaxRows.Text = settings.DefaultMaxRows.ToString();
            ChkCopyWithHeaders.IsChecked = settings.CopyWithHeadersDefault;
            ChkRememberPasswords.IsChecked = settings.RememberPasswords;
            TxtGitHubToken.Password = settings.GitHubToken;

            RbThemeLight.IsChecked = settings.Theme == ThemeManager.Light;
            RbThemeDark.IsChecked = settings.Theme == ThemeManager.Dark;
            RbThemeSystem.IsChecked = settings.Theme == ThemeManager.System;
            if (RbThemeLight.IsChecked != true && RbThemeDark.IsChecked != true && RbThemeSystem.IsChecked != true)
                RbThemeLight.IsChecked = true;
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (!TryReadNonNegative(TxtConnectTimeout.Text, "Connect timeout", out var connect) ||
                !TryReadNonNegative(TxtCommandTimeout.Text, "Query timeout", out var command) ||
                !TryReadNonNegative(TxtDefaultMaxRows.Text, "Default max rows", out var maxRows))
            {
                return;
            }

            _settings.ConnectTimeoutSeconds = connect;
            _settings.CommandTimeoutSeconds = command;
            _settings.DefaultMaxRows = maxRows == 0 ? 1000 : maxRows;
            _settings.CopyWithHeadersDefault = ChkCopyWithHeaders.IsChecked == true;
            _settings.RememberPasswords = ChkRememberPasswords.IsChecked == true;
            _settings.GitHubToken = TxtGitHubToken.Password;
            _settings.Theme = RbThemeDark.IsChecked == true ? ThemeManager.Dark
                            : RbThemeSystem.IsChecked == true ? ThemeManager.System
                            : ThemeManager.Light;

            try
            {
                _settings.Save();
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Could not save settings: {ex.Message}", "Settings", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            DialogResult = true;
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = AppStorage.Folder,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message, "Settings", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private bool TryReadNonNegative(string text, string label, out int value)
        {
            if (int.TryParse(text.Trim(), out value) && value >= 0) return true;
            System.Windows.MessageBox.Show($"{label} must be a whole number (0 or more).", "Settings", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }
    }
}
