using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookMailSorter
{
    public partial class ThisAddIn
    {
        private System.Windows.Forms.Timer _statusClearTimer;
        private string _statusMessageToken;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Logger.Init("OutlookMailSorter");

            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                var buildTime = GetBuildTimestamp(assembly).ToLocalTime();
                var status = $"OutlookMailSorter ready — Built: {buildTime:yyyy-MM-dd HH:mm zzz}";

                Logger.Info(status);
                SetStatusBar(status, TimeSpan.FromSeconds(12));
            }
            catch (Exception ex)
            {
                // Safe startup: never crash Outlook
                Logger.Error(ex, "Startup failed.");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            try
            {
                if (_statusClearTimer != null)
                {
                    _statusClearTimer.Tick -= StatusClearTimer_Tick;
                    _statusClearTimer.Stop();
                    _statusClearTimer.Dispose();
                    _statusClearTimer = null;
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Shutdown encountered an error.");
            }
        }

        // Non-blocking status bar update with optional auto-clear
        private void SetStatusBar(string message, TimeSpan? clearAfter = null)
        {
            try
            {
                var explorer = GetAnyExplorer();
                if (explorer != null)
                {
                    explorer.StatusBar = message;

                    // Generate a message token so we clear only if message hasn't changed
                    _statusMessageToken = Guid.NewGuid().ToString("N");
                    var tokenAtSetTime = _statusMessageToken;

                    if (clearAfter.HasValue)
                    {
                        EnsureStatusTimer();
                        _statusClearTimer.Tag = tokenAtSetTime;
                        _statusClearTimer.Interval = Math.Max(250, (int)clearAfter.Value.TotalMilliseconds);
                        _statusClearTimer.Stop();
                        _statusClearTimer.Start();
                    }
                }
            }
            catch (Exception ex)
            {
                // Do not propagate COM issues back to Outlook
                Logger.Warn($"Failed to set status bar: {ex.Message}");
            }
        }

        private void EnsureStatusTimer()
        {
            if (_statusClearTimer != null) return;

            _statusClearTimer = new System.Windows.Forms.Timer();
            _statusClearTimer.Tick += StatusClearTimer_Tick;
        }

        private void StatusClearTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                _statusClearTimer.Stop();

                // Clear only if the token hasn't changed (message was not updated since timer started)
                var tokenFromTimer = _statusClearTimer.Tag as string;
                if (!string.Equals(tokenFromTimer, _statusMessageToken, StringComparison.Ordinal))
                {
                    // Message changed, don't clear
                    return;
                }

                var explorer = GetAnyExplorer();
                if (explorer != null)
                {
                    // Set to empty string clears the status bar message
                    explorer.StatusBar = string.Empty;
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"Failed to clear status bar: {ex.Message}");
            }
        }

        private Outlook.Explorer GetAnyExplorer()
        {
            try
            {
                // ActiveExplorer is preferred
                var active = this.Application?.ActiveExplorer();
                if (active != null) return active;

                // Fall back to the first available explorer if any
                var explorers = this.Application?.Explorers;
                if (explorers != null && explorers.Count > 0)
                {
                    return explorers[1];
                }
            }
            catch (Exception ex)
            {
                Logger.Warn($"GetAnyExplorer failed: {ex.Message}");
            }

            return null;
        }

        // Read commit/build timestamp using multiple strategies, fallback to PE header linker time
        private static DateTimeOffset GetBuildTimestamp(Assembly assembly)
        {
            // 1) AssemblyMetadataAttribute("CommitTimestamp") or ("CommitDate")
            if (TryGetCommitTimeFromMetadata(assembly, out var dto)) return dto;

            // 2) AssemblyInformationalVersionAttribute parsing (look for CommitDate or date suffix)
            if (TryGetCommitTimeFromInformationalVersion(assembly, out dto)) return dto;

            // 3) Embedded resource "CommitTimestamp" / "CommitTimestamp.txt"
            if (TryGetCommitTimeFromResource(assembly, out dto)) return dto;

            // 4) Sidecar file next to assembly
            if (TryGetCommitTimeFromSidecarFile(assembly, out dto)) return dto;

            // 5) Fallback to PE header link time (UTC)
            return GetLinkerTimeUtc(assembly);
        }

        private static bool TryGetCommitTimeFromMetadata(Assembly assembly, out DateTimeOffset commitTime)
        {
            commitTime = default;

            try
            {
                var metadataAttributes = assembly.GetCustomAttributes<AssemblyMetadataAttribute>()?.ToArray() ?? Array.Empty<AssemblyMetadataAttribute>();
                foreach (var attr in metadataAttributes)
                {
                    if (attr == null) continue;
                    if (string.Equals(attr.Key, "CommitTimestamp", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attr.Key, "CommitDate", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attr.Key, "BuildTimestamp", StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(attr.Key, "BuildDateUtc", StringComparison.OrdinalIgnoreCase))
                    {
                        if (TryParseAnyTimestamp(attr.Value, out commitTime))
                        {
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug($"TryGetCommitTimeFromMetadata failed: {ex.Message}");
            }

            return false;
        }

        private static bool TryGetCommitTimeFromInformationalVersion(Assembly assembly, out DateTimeOffset commitTime)
        {
            commitTime = default;

            try
            {
                var infoAttr = assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>();
                if (infoAttr == null || string.IsNullOrWhiteSpace(infoAttr.InformationalVersion))
                    return false;

                var s = infoAttr.InformationalVersion;

                // Look for patterns like "CommitDate: 2025-01-02T03:04:05Z" or "+Branch.main.Sha.<hash>.Date.2025-01-02T03:04:05Z"
                // Extract potential ISO8601 tokens
                string[] candidates = s.Split(' ', '+', ';', ',', '|')
                    .Concat(s.Split('/','\\','-'))
                    .Distinct()
                    .ToArray();

                foreach (var c in candidates)
                {
                    if (TryParseAnyTimestamp(c, out commitTime))
                    {
                        return true;
                    }
                }

                // Also handle "Date=..." or "CommitDate=..."
                foreach (var segment in s.Split(' ', ';', ',', '|'))
                {
                    var kvp = segment.Split(new[] { '=', ':' }, 2);
                    if (kvp.Length == 2)
                    {
                        var key = kvp[0].Trim();
                        var val = kvp[1].Trim();
                        if (key.Equals("CommitDate", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("CommitTimestamp", StringComparison.OrdinalIgnoreCase) ||
                            key.Equals("BuildDateUtc", StringComparison.OrdinalIgnoreCase))
                        {
                            if (TryParseAnyTimestamp(val, out commitTime))
                                return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug($"TryGetCommitTimeFromInformationalVersion failed: {ex.Message}");
            }

            return false;
        }

        private static bool TryGetCommitTimeFromResource(Assembly assembly, out DateTimeOffset commitTime)
        {
            commitTime = default;

            try
            {
                var names = assembly.GetManifestResourceNames() ?? Array.Empty<string>();
                var resourceName = names.FirstOrDefault(n =>
                    n.EndsWith(".CommitTimestamp", StringComparison.OrdinalIgnoreCase) ||
                    n.EndsWith(".CommitTimestamp.txt", StringComparison.OrdinalIgnoreCase) ||
                    n.EndsWith(".BuildTimestamp", StringComparison.OrdinalIgnoreCase) ||
                    n.EndsWith(".BuildTimestamp.txt", StringComparison.OrdinalIgnoreCase));

                if (resourceName == null) return false;

                using var stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null) return false;

                using var reader = new StreamReader(stream, Encoding.UTF8, true);
                var content = reader.ReadToEnd()?.Trim();
                if (string.IsNullOrWhiteSpace(content)) return false;

                return TryParseAnyTimestamp(content, out commitTime);
            }
            catch (Exception ex)
            {
                Logger.Debug($"TryGetCommitTimeFromResource failed: {ex.Message}");
            }

            return false;
        }

        private static bool TryGetCommitTimeFromSidecarFile(Assembly assembly, out DateTimeOffset commitTime)
        {
            commitTime = default;

            try
            {
                var asmPath = assembly.Location;
                var dir = Path.GetDirectoryName(asmPath);
                if (string.IsNullOrEmpty(dir)) return false;

                string[] candidates =
                {
                    Path.Combine(dir, "CommitTimestamp"),
                    Path.Combine(dir, "CommitTimestamp.txt"),
                    Path.Combine(dir, "BuildTimestamp"),
                    Path.Combine(dir, "BuildTimestamp.txt")
                };

                foreach (var path in candidates)
                {
                    if (File.Exists(path))
                    {
                        var content = File.ReadAllText(path).Trim();
                        if (TryParseAnyTimestamp(content, out commitTime))
                            return true;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Debug($"TryGetCommitTimeFromSidecarFile failed: {ex.Message}");
            }

            return false;
        }

        private static bool TryParseAnyTimestamp(string value, out DateTimeOffset dto)
        {
            dto = default;
            if (string.IsNullOrWhiteSpace(value)) return false;

            // Try ISO8601 or general date/time formats
            if (DateTimeOffset.TryParse(value, out dto)) return true;

            // Try Unix epoch seconds
            if (long.TryParse(value, out var seconds) &&
                seconds > 0 && seconds < 32503680000) // < year 3000
            {
                dto = DateTimeOffset.FromUnixTimeSeconds(seconds);
                return true;
            }

            // Try Unix epoch milliseconds
            if (long.TryParse(value, out var ms) &&
                ms > 1000000000 && ms < 32503680000000) // heuristic
            {
                dto = DateTimeOffset.FromUnixTimeMilliseconds(ms);
                return true;
            }

            return false;
        }

        // Returns PE header linker time in UTC
        private static DateTimeOffset GetLinkerTimeUtc(Assembly assembly)
        {
            const int peHeaderOffset = 60;
            const int linkerTimestampOffset = 8;

            try
            {
                var filePath = assembly.Location;
                var buffer = new byte[2048];

                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    _ = stream.Read(buffer, 0, buffer.Length);
                }

                int peHeader = BitConverter.ToInt32(buffer, peHeaderOffset);
                int secondsSince1970 = BitConverter.ToInt32(buffer, peHeader + linkerTimestampOffset);
                var epoch = new DateTimeOffset(1970, 1, 1, 0, 0, 0, TimeSpan.Zero);
                return epoch.AddSeconds(secondsSince1970);
            }
            catch (Exception ex)
            {
                Logger.Debug($"GetLinkerTimeUtc failed: {ex.Message}");
                return DateTimeOffset.UtcNow;
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        internal static class Logger
        {
            private static readonly object _sync = new object();
            private static string _appName = "App";
            private static string _logFile;
            private static bool _initialized;

            public static void Init(string appName)
            {
                if (_initialized) return;

                lock (_sync)
                {
                    if (_initialized) return;

                    _appName = string.IsNullOrWhiteSpace(appName) ? "App" : appName;
                    try
                    {
                        var dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), _appName);
                        Directory.CreateDirectory(dir);
                        _logFile = Path.Combine(dir, "log.txt");
                    }
                    catch
                    {
                        _logFile = null;
                    }

                    _initialized = true;
                    Info($"{_appName} logger initialized.");
                }
            }

            public static void Debug(string message) => Write("DEBUG", message);
            public static void Info(string message) => Write("INFO", message);
            public static void Warn(string message) => Write("WARN", message);
            public static void Error(string message) => Write("ERROR", message);
            public static void Error(Exception ex, string message = null)
            {
                var combined = message == null ? ex.ToString() : $"{message}{Environment.NewLine}{ex}";
                Write("ERROR", combined);
            }

            private static void Write(string level, string message)
            {
                try
                {
                    var line = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss.fff zzz} [{level}] {message}";
                    System.Diagnostics.Debug.WriteLine(line);

                    if (string.IsNullOrEmpty(_logFile)) return;

                    lock (_sync)
                    {
                        RotateIfNeeded();
                        File.AppendAllText(_logFile, line + Environment.NewLine, Encoding.UTF8);
                    }
                }
                catch
                {
                    // Swallow all logger exceptions
                }
            }

            private static void RotateIfNeeded()
            {
                try
                {
                    const long maxBytes = 512 * 1024; // 512 KB
                    if (!File.Exists(_logFile)) return;

                    var fi = new FileInfo(_logFile);
                    if (fi.Length <= maxBytes) return;

                    var archive = Path.Combine(Path.GetDirectoryName(_logFile) ?? "", $"log-{DateTime.Now:yyyyMMdd-HHmmss}.txt");
                    File.Move(_logFile, archive);
                }
                catch
                {
                    // Ignore rotation errors
                }
            }
        }
    }
}