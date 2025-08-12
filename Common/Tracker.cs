using System;
using System.Diagnostics;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HDR_EMEA.Common
{
    internal static class Tracker
    {
        /// <summary>
        /// Returns (and creates if needed) the network folder for tracking.
        /// Falls back to AppData if the UNC cannot be created.
        /// </summary>
        private static string GetTrackerDirectory()
        {
            // 1) Your network tracking location:
            string networkPath = @"\\omesrv3\ASMECServiceCenter\BIMdev\EMEA_RevitToolbar\RevitToolTracking";

            try
            {
                if (!Directory.Exists(networkPath))
                    Directory.CreateDirectory(networkPath);
                return networkPath;
            }
            catch
            {
                // 2) Fallback to AppData
                string fallback = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "HDR_EMEA",
                    "RevitToolTracking");
                Directory.CreateDirectory(fallback);
                return fallback;
            }
        }

        // Overload for commands that don’t pass UIApplication
        public static void LogCommandUsage(string commandName, DateTime startTime, bool success)
            => LogCommandUsage(commandName, startTime, success, null);

        // Full version with UIApplication for extra context
        public static void LogCommandUsage(string commandName, DateTime startTime, bool success, UIApplication uiapp)
        {
            try
            {
                string trackerDir = GetTrackerDirectory();
                string filePath = Path.Combine(trackerDir, $"{commandName.Replace(" ", "_")}_UsageLog.txt");

                DateTime endTime = DateTime.Now;
                double durationSec = (endTime - startTime).TotalSeconds;
                string user = Environment.UserName;
                string machine = Environment.MachineName;
                string revitVer = uiapp?.Application?.VersionNumber ?? "Unknown";
                string revitBuild = uiapp?.Application?.VersionBuild ?? "Unknown";
                string docTitle = uiapp?.ActiveUIDocument?.Document?.Title ?? "NoDocument";

                // Tab-delimited record:  
                // Timestamp    User    Machine    Command    Duration    Success/Fail    RevitVer    DocTitle
                string logLine = string.Join("\t",
                    endTime.ToString("yyyy-MM-dd HH:mm:ss"),
                    user,
                    machine,
                    commandName,
                    $"{durationSec:F2}s",
                    success ? "Success" : "Failure",
                    $"Revit {revitVer} ({revitBuild})",
                    docTitle
                );

                File.AppendAllText(filePath, logLine + Environment.NewLine);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[Tracker] Could not log usage: {ex.Message}");
            }
        }
    }
}
