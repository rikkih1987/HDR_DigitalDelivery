using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace HDR_EMEA.Common
{ 
    internal static class TagMiConfig
    {
        private static string GetSettingsPath()
        {
            const string companyAddin = "HDR_EMEA";
            const string appName = "TagMi";
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                companyAddin, appName);
            Directory.CreateDirectory(dir);
            return Path.Combine(dir, "TagCategories.json");
        }

        public static HashSet<int> LoadSavedBics()
        {
            try
            {
                string path = GetSettingsPath();
                if (File.Exists(path))
                {
                    var json = File.ReadAllText(path);
                    var ids = JsonSerializer.Deserialize<List<int>>(json);
                    return new HashSet<int>(ids ?? new List<int>());
                }
            }
            catch
            {
                // swallow and return empty
            }

            return new HashSet<int>();
        }

        public static void SaveBics(IEnumerable<int> values)
        {
            try
            {
                string path = GetSettingsPath();
                var json = JsonSerializer.Serialize(values, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(path, json);
            }
            catch
            {
                Autodesk.Revit.UI.TaskDialog.Show("HDR", "Could not save tag category preferences.");
            }
        }
    }
}
