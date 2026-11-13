using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace ODBSee
{
    /// <summary>
    /// Paths and small helpers for everything ODBSee keeps in %LocalAppData%\ODBSee.
    /// </summary>
    public static class AppStorage
    {
        public static string Folder
        {
            get
            {
                var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ODBSee");
                Directory.CreateDirectory(path);
                return path;
            }
        }

        public static string SessionFile => Path.Combine(Folder, "session_state.json");
        public static string SettingsFile => Path.Combine(Folder, "settings.json");
        public static string CrashLogFile => Path.Combine(Folder, "crash.log");
        public static string ActivityLogFile => Path.Combine(Folder, "activity.log");
        public static string PendingFeedbackFile => Path.Combine(Folder, "pending_feedback.json");

        private static readonly object _fileLock = new object();
        private const long MaxLogBytes = 1024 * 1024; // keep each log file under ~1 MB

        /// <summary>Appends a line to a log file, rolling it over to .old when it gets too big. Never throws.</summary>
        public static void AppendLog(string file, string text)
        {
            try
            {
                lock (_fileLock)
                {
                    var info = new FileInfo(file);
                    if (info.Exists && info.Length > MaxLogBytes)
                    {
                        File.Copy(file, file + ".old", overwrite: true);
                        File.Delete(file);
                    }
                    File.AppendAllText(file, text + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch { /* logging must never take the app down */ }
        }

        /// <summary>Writes a file atomically (temp file + replace) so a crash mid-write can't corrupt it.</summary>
        public static void WriteAllTextAtomic(string file, string contents)
        {
            var tmp = file + ".tmp";
            File.WriteAllText(tmp, contents, Encoding.UTF8);
            File.Move(tmp, file, overwrite: true);
        }

        public static T? ReadJson<T>(string file) where T : class
        {
            if (!File.Exists(file)) return null;
            return JsonSerializer.Deserialize<T>(File.ReadAllText(file, Encoding.UTF8));
        }

        public static void WriteJson<T>(string file, T value)
        {
            WriteAllTextAtomic(file, JsonSerializer.Serialize(value, new JsonSerializerOptions { WriteIndented = true }));
        }
    }

    /// <summary>
    /// Encrypts secrets with Windows DPAPI (CurrentUser scope): only the same Windows account
    /// on the same machine can decrypt them. Nothing to manage, no key stored by us.
    /// </summary>
    public static class Secret
    {
        private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("ODBSee.v1");

        public static string Protect(string plain)
        {
            if (string.IsNullOrEmpty(plain)) return "";
            var bytes = ProtectedData.Protect(Encoding.UTF8.GetBytes(plain), Entropy, DataProtectionScope.CurrentUser);
            return Convert.ToBase64String(bytes);
        }

        public static string Unprotect(string? cipher)
        {
            if (string.IsNullOrEmpty(cipher)) return "";
            try
            {
                var bytes = ProtectedData.Unprotect(Convert.FromBase64String(cipher), Entropy, DataProtectionScope.CurrentUser);
                return Encoding.UTF8.GetString(bytes);
            }
            catch
            {
                // Different user/machine or corrupted value: behave as "not remembered".
                return "";
            }
        }
    }
}
