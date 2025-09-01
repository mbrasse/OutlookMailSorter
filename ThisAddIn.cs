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
    }
}