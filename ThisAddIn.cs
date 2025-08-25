using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookMailSorter
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // Initialize logger quickly on startup (non-blocking).
                Logger.Initialize();
                Logger.Log("ThisAddIn_Startup: starting initialization.");

                // Subscribe to events on the Outlook (UI) thread. Do minimal work in handlers.
                this.Application.NewMailEx += Application_NewMailEx;

                // Subscribe to the Inbox ItemAdd event to detect newly delivered items.
                var inbox = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inbox != null)
                {
                    _inboxItems = inbox.Items;
                    _inboxItems.ItemAdd += InboxItems_ItemAdd;

                    // Release the local reference to the folder (we keep Items only).
                    try
                    {
                        Marshal.ReleaseComObject(inbox);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ThisAddIn_Startup: failed to release inbox COM object: {ex}");
                    }
                    inbox = null;
                }

                Logger.Log("ThisAddIn_Startup: event subscriptions completed.");

                // Read commit info from possible locations (created at commit time). Fallback to current UTC if missing.
                string commitInfo = ReadCommitInfo();

                // Obtain local build (link) time safely.
                string buildTimeLocal = "unavailable";
                try
                {
                    var localDt = GetLinkerTime(Assembly.GetExecutingAssembly(), TimeZoneInfo.Local);
                    buildTimeLocal = localDt.ToString("o");
                }
                catch (Exception ex)
                {
                    Logger.Log($"ThisAddIn_Startup: failed to obtain build time: {ex}");
                }

                Logger.Log($"ThisAddIn_Startup: commit info obtained: {Truncate(commitInfo, 500)}; build time (local): {buildTimeLocal}");

#if DEBUG
                // Only show a blocking MessageBox in DEBUG builds to avoid blocking Outlook for end users.
                try
                {
                    var message = $"OutlookMailSorter initialized.{Environment.NewLine}Commit info: {commitInfo}{Environment.NewLine}Build time (local): {buildTimeLocal}";
                    Logger.Log("ThisAddIn_Startup: showing startup MessageBox with commit info and build time (DEBUG).");
                    MessageBox.Show(message, "OutlookMailSorter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    // Protect startup from MessageBox exceptions (rare) and log instead of throwing.
                    Logger.Log($"ThisAddIn_Startup: MessageBox.Show threw an exception: {ex}");
                }
#else
                Logger.Log("ThisAddIn_Startup: skipping MessageBox display in non-DEBUG build.");
#endif
                var buildTime = GetLinkerTime(Assembly.GetExecutingAssembly());
                MessageBox.Show($"Good night - Built: {buildTime:yyyy-MM-dd HH:mm zzz}", "OutlookMailSorter");
                // TODO: Place any initialization code here
            }
            catch (Exception ex)
            {
                // Safe-startup: do not crash Outlook on startup
                Debug.WriteLine($"ThisAddIn_Startup failed: {ex}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        // Returns the assembly linker timestamp converted to local time (or provided timezone)
        private static DateTime GetLinkerTime(Assembly assembly, TimeZoneInfo targetTimeZone = null)
        {
            const int peHeaderOffset = 60;
            const int linkerTimestampOffset = 8;

            string filePath = assembly.Location;
            byte[] buffer = new byte[2048];

            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                stream.Read(buffer, 0, buffer.Length);
            }

            int peHeader = BitConverter.ToInt32(buffer, peHeaderOffset);
            int secondsSince1970 = BitConverter.ToInt32(buffer, peHeader + linkerTimestampOffset);
            DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime linkTimeUtc = epoch.AddSeconds(secondsSince1970);

            var tz = targetTimeZone ?? TimeZoneInfo.Local;
            return TimeZoneInfo.ConvertTimeFromUtc(linkTimeUtc, tz);
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion

        // Helper to truncate long strings for log safety
        private static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return (value.Length <= maxLength) ? value : value.Substring(0, maxLength) + "...";
        }

        // Read commit info from several known locations in a safe, non-throwing way.
        private static string ReadCommitInfo()
        {
            try
            {
                string content = null;

                // 1) Primary: next to the add-in assembly
                try
                {
                    var assemblyLocation = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                    if (!string.IsNullOrEmpty(assemblyLocation))
                    {
                        var primaryPath = Path.Combine(assemblyLocation, "COMMIT_INFO.txt");
                        if (File.Exists(primaryPath))
                        {
                            content = File.ReadAllText(primaryPath).Trim();
                            Logger.Log($"ReadCommitInfo: found COMMIT_INFO at assembly folder: {primaryPath}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"ReadCommitInfo: error reading from assembly folder: {ex}");
                }

                // 2) Fallback: LocalApplicationData\OutlookMailSorter\COMMIT_INFO.txt
                if (string.IsNullOrEmpty(content))
                {
                    try
                    {
                        var fallbackPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "OutlookMailSorter",
                            "COMMIT_INFO.txt");
                        if (File.Exists(fallbackPath))
                        {
                            content = File.ReadAllText(fallbackPath).Trim();
                            Logger.Log($"ReadCommitInfo: found COMMIT_INFO at LocalApplicationData: {fallbackPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ReadCommitInfo: error reading from LocalApplicationData: {ex}");
                    }
                }

                // 3) Fallback: search upward from the base directory (covers build/CI layouts)
                if (string.IsNullOrEmpty(content))
                {
                    try
                    {
                        var repoRoot = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;
                        var dir = new DirectoryInfo(repoRoot);
                        for (int i = 0; i < 6 && dir != null; i++)
                        {
                            var path = Path.Combine(dir.FullName, "COMMIT_INFO.txt");
                            if (File.Exists(path))
                            {
                                content = File.ReadAllText(path).Trim();
                                Logger.Log($"ReadCommitInfo: found COMMIT_INFO by searching up from base: {path}");
                                break;
                            }
                            dir = dir.Parent;
                        }
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ReadCommitInfo: error during upward search: {ex}");
                    }
                }

                if (!string.IsNullOrEmpty(content))
                {
                    // Normalize whitespace for display and truncate to a reasonable length.
                    content = content.Replace("\r\n", " ").Replace("\n", " ").Trim();
                    content = Truncate(content, 1000);
                    return content;
                }

                // Final fallback: timestamp indicating absence.
                return $"{DateTime.UtcNow.ToString("o")} (no commit info available)";
            }
            catch (Exception ex)
            {
                Logger.Log($"ReadCommitInfo: unexpected exception: {ex}");
                return $"Failed to read commit info: {ex.Message}";
            }
        }

        // Returns the build (link) time for the specified assembly, converted to the target time zone (Local by default).
        private static DateTime GetLinkerTime(Assembly assembly, TimeZoneInfo targetTimeZone = null)
        {
            if (assembly == null) throw new ArgumentNullException(nameof(assembly));

            var filePath = assembly.Location;
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
            {
                // Fallback: if the path is not available, return "now" to avoid throwing during startup.
                var nowUtc = DateTime.UtcNow;
                if (targetTimeZone == null) targetTimeZone = TimeZoneInfo.Local;
                return TimeZoneInfo.ConvertTimeFromUtc(nowUtc, targetTimeZone);
            }

            const int peHeaderOffset = 60;
            const int linkerTimestampOffset = 8;

            byte[] buffer = new byte[2048];
            using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                stream.Read(buffer, 0, 2048);
            }

            int headerPos = BitConverter.ToInt32(buffer, peHeaderOffset);
            int secondsSince1970 = BitConverter.ToInt32(buffer, headerPos + linkerTimestampOffset);
            var epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            var linkTimeUtc = epoch.AddSeconds(secondsSince1970);

            if (targetTimeZone == null) targetTimeZone = TimeZoneInfo.Local;
            return TimeZoneInfo.ConvertTimeFromUtc(linkTimeUtc, targetTimeZone);
        }

        // Simple file logger for diagnostics (safe to use on startup).
        private static class Logger
        {
            private static string _logFilePath;
            private static readonly object _sync = new object();
            private static bool _initialized;

            public static void Initialize()
            {
                if (_initialized) return;
                lock (_sync)
                {
                    if (_initialized) return;
                    try
                    {
                        var folder = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                            "OutlookMailSorter");
                        Directory.CreateDirectory(folder);
                        _logFilePath = Path.Combine(folder, "logs.txt");

                        // Create or append a header so the file exists and it's clear when the process started.
                        try
                        {
                            var header = $"=== OutlookMailSorter log started at {DateTime.UtcNow.ToString("o")} (UTC) ==={Environment.NewLine}";
                            File.AppendAllText(_logFilePath, header);
                        }
                        catch (Exception ex)
                        {
                            // If header write fails, still continue without throwing.
                            _logFilePath = _logFilePath; // no-op to avoid warnings
                            // We will rely on Log's internal try/catch.
                        }

                        _initialized = true;
                        Log("Logger initialized.");
                    }
                    catch
                    {
                        _logFilePath = null;
                        _initialized = true;
                    }
                }
            }

            public static void Log(string message)
            {
                try
                {
                    var ts = DateTime.UtcNow.ToString("o");
                    var line = $"{ts} - {message}{Environment.NewLine}";

                    lock (_sync)
                    {
                        if (!string.IsNullOrEmpty(_logFilePath))
                        {
                            File.AppendAllText(_logFilePath, line);
                        }
                    }
                }
                catch
                {
                    // Do not propagate exceptions from logging.
                }
            }
        }
    }
}