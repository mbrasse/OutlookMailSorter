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
        private const string AppName = "OutlookMailSorter";
        private const string CommitInfoFileName = "COMMIT_INFO.txt";

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var buildTime = GetLinkerTime(asm);
                var commitInfo = GetCommitInfoSafe(asm);

                var message = $"Good night - Built: {buildTime:yyyy-MM-dd HH:mm zzz}";
                if (!string.IsNullOrWhiteSpace(commitInfo))
                {
                    message += $" | {commitInfo}";
                }

                // Keep any UI very lightweight; never let it crash startup.
                try
                {
                    MessageBox.Show(message, AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception uiEx)
                {
                    Debug.WriteLine($"Startup UI message failed: {uiEx}");
                }

                // TODO: Place any initialization code here (keep it safe and guarded).
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
            // must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        // Safely read optional commit/build info from COMMIT_INFO.txt co-located with the assembly.
        private static string GetCommitInfoSafe(Assembly assembly)
        {
            try
            {
                var location = assembly?.Location;
                if (string.IsNullOrWhiteSpace(location))
                {
                    return null;
                }

                var baseDir = Path.GetDirectoryName(location);
                if (string.IsNullOrWhiteSpace(baseDir))
                {
                    return null;
                }

                var filePath = Path.Combine(baseDir, CommitInfoFileName);
                if (!File.Exists(filePath))
                {
                    return null;
                }

                // Read all text safely; normalize whitespace and limit size.
                var text = File.ReadAllText(filePath)?.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    return null;
                }

                // Normalize newlines and prevent excessively long content from blocking UI.
                text = text.Replace("\r", " ").Replace("\n", " ").Trim();
                const int maxLength = 200;
                if (text.Length > maxLength)
                {
                    text = text.Substring(0, maxLength) + "...";
                }

                return text;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetCommitInfoSafe failed: {ex}");
                return null;
            }
        }

        // Returns the assembly linker timestamp converted to local time (or provided timezone)
        private static DateTime GetLinkerTime(Assembly assembly, TimeZoneInfo targetTimeZone = null)
        {
            try
            {
                const int peHeaderOffset = 60;
                const int linkerTimestampOffset = 8;

                var filePath = assembly?.Location;
                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    return DateTime.Now;
                }

                byte[] buffer = new byte[2048];
                using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    int bytesRead = stream.Read(buffer, 0, buffer.Length);
                    if (bytesRead < 256)
                    {
                        return DateTime.Now;
                    }
                }

                int peHeader = BitConverter.ToInt32(buffer, peHeaderOffset);
                int secondsSince1970 = BitConverter.ToInt32(buffer, peHeader + linkerTimestampOffset);
                DateTime epoch = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
                DateTime linkTimeUtc = epoch.AddSeconds(secondsSince1970);

                var tz = targetTimeZone ?? TimeZoneInfo.Local;
                return TimeZoneInfo.ConvertTimeFromUtc(linkTimeUtc, tz);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"GetLinkerTime failed: {ex}");
                return DateTime.Now;
            }
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
    }
}