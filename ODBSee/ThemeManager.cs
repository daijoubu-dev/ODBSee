using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Media;
using Microsoft.Win32;
using Color = System.Windows.Media.Color;
using ColorConverter = System.Windows.Media.ColorConverter;

namespace ODBSee
{
    /// <summary>
    /// Light / Dark / Match-Windows theming.
    /// Colors live in Application resources as brushes (referenced with DynamicResource, so switching is live).
    /// Light mode uses the stock WPF control look exactly as before; Dark mode additionally merges
    /// Themes/DarkControls.xaml, which restyles the handful of controls the app uses.
    /// The WinForms results grid is themed separately by MainWindow (listens to <see cref="ThemeChanged"/>).
    /// </summary>
    public static class ThemeManager
    {
        public const string Light = "Light";
        public const string Dark = "Dark";
        public const string System = "System";

        public static bool IsDark { get; private set; }
        public static string CurrentSetting { get; private set; } = Light;

        public static event Action? ThemeChanged;

        private static ResourceDictionary? _darkControls;
        private static bool _hooked;

        private static readonly Dictionary<string, (string Light, string Dark)> Palette = new()
        {
            ["Brush.Window"]         = ("#F5F5F5", "#1E1E1E"),
            ["Brush.Panel"]          = ("#FFFFFF", "#252526"),
            ["Brush.Input"]          = ("#FFFFFF", "#1E1E1E"),
            ["Brush.Code"]           = ("#F8F9FA", "#1E1E1E"),
            ["Brush.Text"]           = ("#000000", "#DCDCDC"),
            ["Brush.Muted"]          = ("#666666", "#A0A0A0"),
            ["Brush.Subtle"]         = ("#888888", "#8A8A8A"),
            ["Brush.Border"]         = ("#DDDDDD", "#3F3F46"),
            ["Brush.BorderStrong"]   = ("#CCCCCC", "#555559"),
            ["Brush.Button"]         = ("#EEEEEE", "#333337"),
            ["Brush.HoverOverlay"]   = ("#14000000", "#22FFFFFF"),
            ["Brush.PressedOverlay"] = ("#28000000", "#3AFFFFFF"),
            ["Brush.StatusBar"]      = ("#EEEEEE", "#2D2D30"),
            ["Brush.Link"]           = ("#007ACC", "#4FA8E8"),
            ["Brush.Warning"]        = ("#A0522D", "#E0A060"),
            ["Brush.Selection"]      = ("#3399FF", "#264F78"),
            ["Brush.Menu"]           = ("#F5F5F5", "#2B2B2B"),
            ["Brush.MenuHover"]      = ("#E0E0E0", "#3E3E40"),
            ["Brush.Tab"]            = ("#EDEDED", "#2D2D30"),
            ["Brush.ScrollTrack"]    = ("#F0F0F0", "#252526"),
            ["Brush.ScrollThumb"]    = ("#C2C2C2", "#4A4A4F"),
            ["Brush.ScrollThumbHover"] = ("#A0A0A0", "#6A6A70"),
        };

        public static void Apply(string? setting)
        {
            CurrentSetting = setting is Dark or System ? setting : Light;
            IsDark = CurrentSetting == Dark || (CurrentSetting == System && WindowsUsesDarkApps());

            var resources = global::System.Windows.Application.Current.Resources;

            foreach (var (key, colors) in Palette)
            {
                var brush = new SolidColorBrush((Color)ColorConverter.ConvertFromString(IsDark ? colors.Dark : colors.Light));
                brush.Freeze();
                resources[key] = brush;
            }

            _darkControls ??= new ResourceDictionary
            {
                Source = new Uri("pack://application:,,,/ODBSee;component/Themes/DarkControls.xaml", UriKind.Absolute)
            };

            var merged = resources.MergedDictionaries;
            if (IsDark && !merged.Contains(_darkControls)) merged.Add(_darkControls);
            if (!IsDark && merged.Contains(_darkControls)) merged.Remove(_darkControls);

            if (!_hooked)
            {
                _hooked = true;
                // Dark title bar for every window as it loads
                EventManager.RegisterClassHandler(typeof(Window), FrameworkElement.LoadedEvent,
                    new RoutedEventHandler((s, e) => { if (s is Window w) ApplyTitleBar(w); }));
                // Follow Windows' setting live when "Match Windows" is selected
                SystemEvents.UserPreferenceChanged += (s, e) =>
                {
                    if (CurrentSetting == System && e.Category == UserPreferenceCategory.General)
                    {
                        global::System.Windows.Application.Current?.Dispatcher.InvokeAsync(() =>
                        {
                            if (WindowsUsesDarkApps() != IsDark) Apply(System);
                        });
                    }
                };
            }

            foreach (Window w in global::System.Windows.Application.Current.Windows)
            {
                ApplyTitleBar(w);
            }

            ThemeChanged?.Invoke();
        }

        public static bool WindowsUsesDarkApps()
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
                return key?.GetValue("AppsUseLightTheme") is int v && v == 0;
            }
            catch
            {
                return false;
            }
        }

        // ---- Title bar (Windows 10 20H1+ / Windows 11) ----

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        private const int DWMWA_USE_IMMERSIVE_DARK_MODE = 20;

        public static void ApplyTitleBar(Window window)
        {
            try
            {
                var hwnd = new WindowInteropHelper(window).Handle;
                if (hwnd == IntPtr.Zero) return;
                int value = IsDark ? 1 : 0;
                DwmSetWindowAttribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, ref value, sizeof(int));
            }
            catch { /* older Windows: title bar just stays light */ }
        }

        // ---- WinForms helpers (results grid scrollbars) ----

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hwnd, string? appName, string? idList);

        public static void ApplyNativeTheme(IntPtr hwnd)
        {
            try { SetWindowTheme(hwnd, IsDark ? "DarkMode_Explorer" : "Explorer", null); } catch { }
        }
    }
}
