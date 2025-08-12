using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using CheckBox = System.Windows.Controls.CheckBox;

namespace HDR_EMEA.Forms
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class TagSelection : Window
    {
        private readonly Document _doc;
        private readonly string _settingsPath;

        public TagSelection(Document doc)
        {
            InitializeComponent();
            _doc = doc;

            string companyAddin = "HDR_EMEA";
            string appName = "TagMi";
            string dir = System.IO.Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                companyAddin, appName);
            Directory.CreateDirectory(dir);
            _settingsPath = System.IO.Path.Combine(dir, "TagCategories.json");

            Loaded += TagSelection_Loaded;
        }

        private void TagSelection_Loaded(object sender, RoutedEventArgs e)
        {
            // Load any saved selections first (per-user)
            var saved = LoadSaved();

            var tagBics = Enum.GetValues(typeof(BuiltInCategory))
                              .Cast<BuiltInCategory>()
                              .Where(bic => bic.ToString().EndsWith("Tags"));

            var items = new List<(BuiltInCategory bic, string label)>();

            foreach (var bic in tagBics)
            {
                // Only include categories that exist in this document/session
                Category cat = Category.GetCategory(_doc, new ElementId((int)bic));
                if (cat == null) continue;

                string label;
                try { label = LabelUtils.GetLabelFor(bic); }
                catch { label = bic.ToString().Replace("OST_", string.Empty); }

                items.Add((bic, label));
            }

            // Sort nicely by label
            foreach (var item in items.OrderBy(i => i.label, StringComparer.CurrentCultureIgnoreCase))
            {
                var cb = new CheckBox
                {
                    Content = item.label,
                    Tag = item.bic,
                    Margin = new Thickness(0, 2, 0, 2),
                    IsChecked = saved.Contains((int)item.bic)
                };
                stackSchedules.Children.Add(cb);
            }
        }

        private HashSet<int> LoadSaved()
        {
            try
            {
                if (File.Exists(_settingsPath))
                {
                    var json = File.ReadAllText(_settingsPath);
                    var ids = System.Text.Json.JsonSerializer.Deserialize<List<int>>(json);
                    return new HashSet<int>(ids ?? new List<int>());
                }
            }
            catch { /* ignore and use defaults */ }
            return new HashSet<int>();
        }

        private void SaveSelections()
        {
            var selected = stackSchedules.Children
                .OfType<CheckBox>()
                .Where(cb => cb.IsChecked == true)
                .Select(cb => (int)(BuiltInCategory)cb.Tag)
                .ToList();

            try
            {
                var json = System.Text.Json.JsonSerializer.Serialize(selected, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(_settingsPath, json);
            }
            catch
            {
                TaskDialog.Show("HDR", "Could not save tag category preferences.");
            }
        }

        private void btnSave_Click(object sender, RoutedEventArgs e)
        {
            SaveSelections();
            DialogResult = true;
            Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
