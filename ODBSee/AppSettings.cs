using System.Text.Json.Serialization;

namespace ODBSee
{
    /// <summary>User preferences, stored in %LocalAppData%\ODBSee\settings.json.</summary>
    public class AppSettings
    {
        /// <summary>"Light", "Dark" or "System" (match Windows). See <see cref="ThemeManager"/>.</summary>
        public string Theme { get; set; } = ThemeManager.Light;

        /// <summary>Seconds to wait when opening a connection. 0 = driver default / no limit.</summary>
        public int ConnectTimeoutSeconds { get; set; } = 300;

        /// <summary>Seconds a query may run before timing out. 0 = no limit.</summary>
        public int CommandTimeoutSeconds { get; set; } = 300;

        public int DefaultMaxRows { get; set; } = 1000;

        public bool CopyWithHeadersDefault { get; set; } = true;

        /// <summary>Remember passwords between sessions (stored encrypted with Windows DPAPI).</summary>
        public bool RememberPasswords { get; set; } = true;

        /// <summary>DPAPI-encrypted GitHub token used by Send Feedback. Never stored in plain text.</summary>
        public string GitHubTokenEncrypted { get; set; } = "";

        [JsonIgnore]
        public string GitHubToken
        {
            get => Secret.Unprotect(GitHubTokenEncrypted);
            set => GitHubTokenEncrypted = Secret.Protect(value?.Trim() ?? "");
        }

        public static AppSettings Load()
        {
            try
            {
                return AppStorage.ReadJson<AppSettings>(AppStorage.SettingsFile) ?? new AppSettings();
            }
            catch
            {
                return new AppSettings();
            }
        }

        public void Save() => AppStorage.WriteJson(AppStorage.SettingsFile, this);
    }
}
