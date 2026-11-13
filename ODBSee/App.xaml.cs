using System.Text;
using System.Windows;

namespace ODBSee
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        protected override void OnStartup(StartupEventArgs e)
        {
            // Write anything unexpected to %LocalAppData%\ODBSee\crash.log so bugs hit at work can be diagnosed later.
            // Note: native crashes inside an ODBC driver (access violations) can't be caught by .NET at all;
            // those show up in Windows Event Viewer > Windows Logs > Application as "Application Error".
            DispatcherUnhandledException += (s, args) =>
            {
                WriteCrash("UI thread", args.Exception);
                System.Windows.MessageBox.Show(
                    $"Something went wrong:\n\n{args.Exception.Message}\n\nDetails were saved to:\n{AppStorage.CrashLogFile}",
                    "ODBSee Error", MessageBoxButton.OK, MessageBoxImage.Error);
                args.Handled = true; // keep the app alive
            };

            AppDomain.CurrentDomain.UnhandledException += (s, args) =>
                WriteCrash($"AppDomain (terminating: {args.IsTerminating})", args.ExceptionObject as Exception);

            TaskScheduler.UnobservedTaskException += (s, args) =>
            {
                WriteCrash("Unobserved task", args.Exception);
                args.SetObserved();
            };

            // Apply the saved theme before the main window is created so it opens in the right colors
            ThemeManager.Apply(AppSettings.Load().Theme);

            base.OnStartup(e);
        }

        public static void WriteCrash(string source, Exception? ex)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"===== {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff} | {source} | v{System.Reflection.Assembly.GetExecutingAssembly().GetName().Version?.ToString(3)} =====");
            sb.AppendLine(ex?.ToString() ?? "(no exception object)");
            AppStorage.AppendLog(AppStorage.CrashLogFile, sb.ToString());
        }
    }
}
