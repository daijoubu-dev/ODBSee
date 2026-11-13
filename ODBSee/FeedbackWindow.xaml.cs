using System.Windows;

namespace ODBSee
{
    public partial class FeedbackWindow : Window
    {
        private readonly string _token;

        /// <summary>Set after closing: a short line for the status bar / log.</summary>
        public string ResultMessage { get; private set; } = "";

        public FeedbackWindow(string token)
        {
            InitializeComponent();
            _token = token;
            TxtMode.Text = string.IsNullOrWhiteSpace(_token)
                ? "No GitHub token set: Send opens the issue page in your browser."
                : "Sends directly to GitHub.";
            Loaded += (s, e) => TxtTitle.Focus();
        }

        private async void BtnSend_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtTitle.Text))
            {
                System.Windows.MessageBox.Show("Please enter a title.", "Feedback", MessageBoxButton.OK, MessageBoxImage.Warning);
                TxtTitle.Focus();
                return;
            }

            var item = new FeedbackItem
            {
                Kind = RbBug.IsChecked == true ? "bug" : "enhancement",
                Title = TxtTitle.Text.Trim(),
                Body = TxtBody.Text,
                CreatedUtc = DateTime.UtcNow
            };

            if (string.IsNullOrWhiteSpace(_token))
            {
                try
                {
                    FeedbackService.OpenInBrowser(item);
                    ResultMessage = "Opened the GitHub issue page in your browser.";
                    DialogResult = true;
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show($"Could not open the browser: {ex.Message}", "Feedback", MessageBoxButton.OK, MessageBoxImage.Error);
                }
                return;
            }

            BtnSend.IsEnabled = false;
            BtnSend.Content = "Sending...";
            try
            {
                var url = await FeedbackService.SendAsync(item, _token);
                ResultMessage = $"Feedback sent: {url}";
                DialogResult = true;
            }
            catch (Exception ex)
            {
                try
                {
                    FeedbackService.Enqueue(item);
                    ResultMessage = $"Feedback couldn't be sent ({ex.Message}). Saved; will retry next launch.";
                    System.Windows.MessageBox.Show($"Couldn't send right now:\n{ex.Message}\n\nIt's been saved and will be sent automatically next time ODBSee starts.",
                        "Feedback", MessageBoxButton.OK, MessageBoxImage.Information);
                    DialogResult = true;
                }
                catch (Exception saveEx)
                {
                    BtnSend.IsEnabled = true;
                    BtnSend.Content = "Send";
                    System.Windows.MessageBox.Show($"Couldn't send ({ex.Message}) or save ({saveEx.Message}).", "Feedback", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }
}
