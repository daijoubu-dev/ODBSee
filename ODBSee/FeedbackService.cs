using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace ODBSee
{
    public class FeedbackItem
    {
        /// <summary>"bug" or "enhancement" (matches the GitHub label).</summary>
        public string Kind { get; set; } = "bug";
        public string Title { get; set; } = "";
        public string Body { get; set; } = "";
        public DateTime CreatedUtc { get; set; } = DateTime.UtcNow;
    }

    /// <summary>
    /// Files feedback as GitHub issues. Only sends what the user typed plus the app and Windows version:
    /// no DSNs, usernames, passwords, query text or error messages (issues on the repo are public).
    /// </summary>
    public static class FeedbackService
    {
        private const string Repo = "daijoubu-dev/ODBSee";
        private const string FromAppLabel = "from-app";

        public static string AppVersion =>
            Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "unknown";

        public static string BuildBody(FeedbackItem item)
        {
            var sb = new StringBuilder();
            sb.AppendLine(string.IsNullOrWhiteSpace(item.Body) ? "_(no description)_" : item.Body.Trim());
            sb.AppendLine();
            sb.AppendLine("---");
            sb.AppendLine($"_Sent from ODBSee v{AppVersion} ({RuntimeInformation.ProcessArchitecture}) on {RuntimeInformation.OSDescription}, {item.CreatedUtc:yyyy-MM-dd HH:mm} UTC_");
            return sb.ToString();
        }

        /// <summary>Creates the issue via the GitHub API. Returns the issue's URL. Throws on failure.</summary>
        public static async Task<string> SendAsync(FeedbackItem item, string token)
        {
            using var client = new HttpClient { Timeout = TimeSpan.FromSeconds(20) };
            client.DefaultRequestHeaders.UserAgent.ParseAdd("ODBSee-Feedback");
            client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
            client.DefaultRequestHeaders.Add("X-GitHub-Api-Version", "2022-11-28");
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

            var url = $"https://api.github.com/repos/{Repo}/issues";
            var labels = new[] { FromAppLabel, item.Kind };

            var response = await PostIssue(client, url, item, labels);

            // A token that can't create labels may get 422 for labels the repo doesn't have yet: retry without them.
            if (response.StatusCode == HttpStatusCode.UnprocessableEntity)
            {
                response.Dispose();
                response = await PostIssue(client, url, item, Array.Empty<string>());
            }

            using (response)
            {
                var text = await response.Content.ReadAsStringAsync();
                if (!response.IsSuccessStatusCode)
                {
                    var hint = response.StatusCode switch
                    {
                        HttpStatusCode.Unauthorized => " (token is invalid or expired)",
                        HttpStatusCode.Forbidden or HttpStatusCode.NotFound => " (token doesn't have Issues write access to the ODBSee repo)",
                        _ => ""
                    };
                    throw new HttpRequestException($"GitHub returned {(int)response.StatusCode} {response.ReasonPhrase}{hint}");
                }

                using var doc = JsonDocument.Parse(text);
                return doc.RootElement.TryGetProperty("html_url", out var u) ? u.GetString() ?? "" : "";
            }
        }

        private static Task<HttpResponseMessage> PostIssue(HttpClient client, string url, FeedbackItem item, string[] labels)
        {
            var payload = JsonSerializer.Serialize(new
            {
                title = item.Title.Trim(),
                body = BuildBody(item),
                labels
            });
            return client.PostAsync(url, new StringContent(payload, Encoding.UTF8, "application/json"));
        }

        /// <summary>No token: open GitHub's "new issue" page with everything pre-filled; the user submits it there.</summary>
        public static void OpenInBrowser(FeedbackItem item)
        {
            var body = BuildBody(item);
            if (body.Length > 6000) body = body.Substring(0, 6000) + "\n…(truncated)";

            var url = $"https://github.com/{Repo}/issues/new" +
                      $"?title={Uri.EscapeDataString(item.Title.Trim())}" +
                      $"&body={Uri.EscapeDataString(body)}" +
                      $"&labels={Uri.EscapeDataString(FromAppLabel + "," + item.Kind)}";

            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = url, UseShellExecute = true });
        }

        // ---- Offline queue: feedback that failed to send is kept and retried on next launch ----

        private static readonly object _queueLock = new object();

        public static void Enqueue(FeedbackItem item)
        {
            lock (_queueLock)
            {
                var list = LoadQueue();
                list.Add(item);
                AppStorage.WriteJson(AppStorage.PendingFeedbackFile, list);
            }
        }

        public static int PendingCount
        {
            get { lock (_queueLock) return LoadQueue().Count; }
        }

        /// <summary>Tries to send every queued item. Returns (sent, stillPending).</summary>
        public static async Task<(int Sent, int Remaining)> RetryPendingAsync(string token)
        {
            List<FeedbackItem> pending;
            lock (_queueLock) pending = LoadQueue();
            if (pending.Count == 0 || string.IsNullOrWhiteSpace(token)) return (0, pending.Count);

            var failed = new List<FeedbackItem>();
            int sent = 0;
            foreach (var item in pending)
            {
                try
                {
                    await SendAsync(item, token);
                    sent++;
                }
                catch
                {
                    failed.Add(item);
                }
            }

            lock (_queueLock)
            {
                // Keep anything that failed, plus anything queued while we were sending.
                var current = LoadQueue();
                var added = current.Where(c => !pending.Any(p => p.CreatedUtc == c.CreatedUtc && p.Title == c.Title)).ToList();
                var remaining = failed.Concat(added).ToList();
                if (remaining.Count == 0)
                {
                    try { File.Delete(AppStorage.PendingFeedbackFile); } catch { }
                }
                else
                {
                    AppStorage.WriteJson(AppStorage.PendingFeedbackFile, remaining);
                }
                return (sent, remaining.Count);
            }
        }

        private static List<FeedbackItem> LoadQueue()
        {
            try
            {
                return AppStorage.ReadJson<List<FeedbackItem>>(AppStorage.PendingFeedbackFile) ?? new List<FeedbackItem>();
            }
            catch
            {
                return new List<FeedbackItem>();
            }
        }
    }
}
