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
        // Keep a reference to the Inbox Items collection so we can unsubscribe and release properly.
        private Outlook.Items _inboxItems;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                // Initialize logger first (quick, non-blocking).
                Logger.Initialize();
                Logger.Log("ThisAddIn_Startup: initialization started.");

                // Subscribe to Outlook events (short handlers only).
                this.Application.NewMailEx += Application_NewMailEx;

                // Subscribe to ItemAdd on the Inbox folder (to detect delivered items).
                var inboxFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox) as Outlook.Folder;
                if (inboxFolder != null)
                {
                    _inboxItems = inboxFolder.Items;
                    _inboxItems.ItemAdd += InboxItems_ItemAdd;

                    // Release folder COM reference (we keep Items reference).
                    try
                    {
                        Marshal.ReleaseComObject(inboxFolder);
                    }
                    catch (Exception ex)
                    {
                        Logger.Log($"ThisAddIn_Startup: failed to release inboxFolder COM object: {ex}");
                    }
                    inboxFolder = null;
                }
                else
                {
                    Logger.Log("ThisAddIn_Startup: inbox folder was null.");
                }

                Logger.Log("ThisAddIn_Startup: event subscriptions completed.");

                // Read commit info from COMMIT_INFO.txt (created at commit time). Fallback to current UTC if missing.
                string commitTimestamp = null;
                try
                {
                    string candidate = null;
                    // Try common locations relative to the assembly base directory.
                    var repoProbe = AppDomain.CurrentDomain.BaseDirectory;

                    // Candidate at base dir
                    candidate = Path.Combine(repoProbe, "COMMIT_INFO.txt");

                    // Also try parent directories up to two levels (bin/Debug or bin/Release scenarios)
                    var parent = Directory.GetParent(repoProbe);
                    if ((candidate == null || !File.Exists(candidate)) && parent != null)
                    {
                        var alt = Path.Combine(parent.FullName, "COMMIT_INFO.txt");
                        if (File.Exists(alt)) candidate = alt;
                    }

                    if ((candidate == null || !File.Exists(candidate)) && parent != null && parent.Parent != null)
                    {
                        var alt2 = Path.Combine(parent.Parent.FullName, "COMMIT_INFO.txt");
                        if (File.Exists(alt2)) candidate = alt2;
                    }

                    if (!string.IsNullOrEmpty(candidate) && File.Exists(candidate))
                    {
                        commitTimestamp = File.ReadAllText(candidate).Trim();
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log($"ThisAddIn_Startup: failed reading COMMIT_INFO.txt: {ex}");
                    commitTimestamp = null;
                }

                if (string.IsNullOrEmpty(commitTimestamp))
                {
                    commitTimestamp = DateTime.UtcNow.ToString("o") + " (UTC)"; // fallback
                }

                // PER YOUR REQUEST: Keep the MessageBox but update its text to include commit timestamp.
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
                // Log any startup exception (do not show blocking UI).
                Logger.Log($"ThisAddIn_Startup: exception during startup: {ex}");
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Logger.Log("ThisAddIn_Shutdown: shutting down.");

            // Unsubscribe events and release COM references defensively.
            try
            {
                try
                {
                    this.Application.NewMailEx -= Application_NewMailEx;
                }
                catch (Exception ex)
                {
                    Logger.Log($"ThisAddIn_Shutdown: error unsubscribing NewMailEx: {ex}");
                }

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

        // Called on Outlook UI thread. Keep brief and offload heavier work.
        private void Application_NewMailEx(string entryIDCollection)
        {
            try
            {
                Logger.Log($"Application_NewMailEx: entryIDs: {entryIDCollection}");

                // Do minimal work on the Outlook thread; offload CPU-bound or non-COM work to background tasks.
                Task.Run(() =>
                {
                    // Placeholder for background processing that does not touch Outlook COM objects.
                    Logger.Log($"Background worker: handling NewMailEx IDs: {entryIDCollection}");
                });
            }
            catch (Exception ex)
            {
                Logger.Log($"Application_NewMailEx: exception: {ex}");
            }
        }

        // Runs on Outlook UI thread whenever an item is added to Inbox.
        private void InboxItems_ItemAdd(object item)
        {
            Outlook.MailItem mail = null;
            try
            {
                mail = item as Outlook.MailItem;
                if (mail != null)
                {
                    // Read small set of properties while on UI thread.
                    string subject = Truncate(mail.Subject, 250);
                    string sender = Truncate(mail.SenderName, 200);

                    Logger.Log($"InboxItems_ItemAdd: Subject=\"{subject}\", Sender=\"{sender}\"");

                    // If you need to perform operations that access Outlook COM objects,
                    // do them on the Outlook thread in short bursts or use a marshalling strategy.
                    Task.Run(() =>
                    {
                        // Placeholder for non-COM work (e.g., update DB, call web APIs).
                        Logger.Log("Background processing (non-COM) for new MailItem.");
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
                        // Swallow release exceptions to avoid interfering with Outlook.
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

        // Helper to truncate long strings for logging
        private static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return (value.Length <= maxLength) ? value : value.Substring(0, maxLength) + "...";
        }

        // Lightweight file logger (resilient; does not throw).
        private static class Logger
        {
            private static readonly object _sync = new object();
            private static string _logFilePath;
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

                        // Note: for production you may want to rotate or limit log size.
                        Log("Logger initialized.");
                    }
                    catch
                    {
                        // If logging cannot be initialized, mark as initialized to avoid repeated attempts.
                        _logFilePath = null;
                    }
                    finally
                    {
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
                        // If there's no file path, drop the log (don't throw).
                    }
                }
                catch
                {
                    // Swallow all exceptions from logging to avoid destabilizing Outlook.
                }
            }
        }
    }
}