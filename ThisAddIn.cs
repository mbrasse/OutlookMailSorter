using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookMailSorter
{
    public partial class ThisAddIn
    {
        // Keep a reference to the Inbox Items collection so we can unsubscribe and release.
        private Outlook.Items _inboxItems;

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

                // Read commit info from COMMIT_INFO.txt (created at commit time). Fallback to current UTC if missing.
                string commitTimestamp = null;
                try
                {
                    // Try repository root locations relative to the assembly base directory.
                    var repoRoot = AppDomain.CurrentDomain.BaseDirectory ?? Environment.CurrentDirectory;

                    // Search upward from the base directory for COMMIT_INFO.txt (covering bin/Debug/bin/Release cases).
                    string candidate = null;
                    var dir = new DirectoryInfo(repoRoot);
                    for (int i = 0; i < 4 && dir != null; i++)
                    {
                        var path = Path.Combine(dir.FullName, "COMMIT_INFO.txt");
                        if (File.Exists(path))
                        {
                            candidate = path;
                            break;
                        }
                        dir = dir.Parent;
                    }

                    if (!string.IsNullOrEmpty(candidate))
                    {
                        commitTimestamp = File.ReadAllText(candidate).Trim();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"ThisAddIn_Startup: error reading COMMIT_INFO.txt: {ex}");
                    commitTimestamp = null;
                }

                if (string.IsNullOrEmpty(commitTimestamp))
                {
                    commitTimestamp = DateTime.UtcNow.ToString("o") + " (UTC)"; // fallback
                }

                // PER USER REQUEST: Keep the MessageBox for now — include commit timestamp in the text.
                // WARNING: This is blocking on the Outlook UI thread. Consider removing or replacing later.
                try
                {
                    var message = $"OutlookMailSorter initialized.\nLatest commit timestamp: {commitTimestamp}";
                    MessageBox.Show(message, "OutlookMailSorter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    // Protect startup from MessageBox exceptions (rare) and log instead of throwing.
                    Logger.Log($"ThisAddIn_Startup: MessageBox.Show threw an exception: {ex}");
                }
            }
            catch (Exception ex)
            {
                // Log startup exceptions; do NOT show blocking UI.
                Logger.Log($"ThisAddIn_Startup: exception during startup: {ex}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Logger.Log("ThisAddIn_Shutdown: shutting down.");

            try
            {
                // Unsubscribe from NewMailEx
                try
                {
                    this.Application.NewMailEx -= Application_NewMailEx;
                }
                catch (Exception ex)
                {
                    Logger.Log($"ThisAddIn_Shutdown: error unsubscribing NewMailEx: {ex}");
                }

                // Unsubscribe and release the Inbox Items collection
                if (_inboxItems != null)
                {
                    try
                    {
                        _inboxItems.ItemAdd -= InboxItems_ItemAdd;
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ThisAddIn_Shutdown: error unsubscribing ItemAdd: {ex}");
                    }

                    try
                    {
                        Marshal.ReleaseComObject(_inboxItems);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ThisAddIn_Shutdown: error releasing _inboxItems COM object: {ex}");
                    }
                    finally
                    {
                        _inboxItems = null;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"ThisAddIn_Shutdown: unexpected exception: {ex}");
            }
            finally
            {
                Logger.Log("ThisAddIn_Shutdown: complete.");
            }
        }

        // NewMailEx is called on the Outlook thread. Keep processing minimal here.
        private void Application_NewMailEx(string entryIDCollection)
        {
            try
            {
                Logger.Log($"Application_NewMailEx: received entries: {entryIDCollection}");

                // Offload heavier/CPU-bound work to a background task. Do NOT access Outlook COM objects from background threads.
                Task.Run(() =>
                {
                    Logger.Log($"Background task: new mail reference processing for IDs: {entryIDCollection}");
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"Application_NewMailEx: exception: {ex}");
            }
        }

        // ItemAdd handler runs on the Outlook (UI) thread. Keep it short.
        private void InboxItems_ItemAdd(object item)
        {
            Outlook.MailItem mail = null;
            try
            {
                mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    // Only read a small number of properties while on the UI thread.
                    string subject = mail.Subject;
                    string sender = mail.SenderName;

                    Logger.Log($"InboxItems_ItemAdd: new mail - Subject: \"{Truncate(subject, 250)}\", Sender: \"{Truncate(sender, 200)}\"");

                    // Offload any heavy processing to a background task. IMPORTANT:
                    // Access to Outlook COM objects from background threads is unsafe.
                    Task.Run(() => {
                        // Placeholder for CPU-bound or non-COM work.
                        Logger.Log("Background processing for the newly arrived mail (non-COM work).");
                    });
                }
                else
                {
                    Logger.Log("InboxItems_ItemAdd: item is not a MailItem.");
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"InboxItems_ItemAdd: exception: {ex}");
            }
            finally
            {
                if (mail != null)
                {
                    try
                    {
                        Marshal.ReleaseComObject(mail);
                    }
                    catch
                    {
                        // swallowing errors from ReleaseComObject
                    }
                    mail = null;
                }
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify the contents of this method with the code editor.
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

                        // Optional: limit file size, rotate if too large - omitted for brevity.
                        Log("Logger initialized.");
                        _initialized = true;
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
                    var line = $"{ts}{message}{Environment.NewLine}";

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