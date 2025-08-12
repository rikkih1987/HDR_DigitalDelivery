using System;
using System.Collections.Generic;
using System.ComponentModel;
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
using Button = System.Windows.Controls.Button;

namespace HDR_EMEA.Forms
{
    /// <summary>
    /// Interaction logic for Window1.xaml
    /// </summary>
    public partial class TagButtons : Window
    {
        private readonly Document _doc;
        private readonly List<TagItem> _allItems = new();

        public BuiltInCategory? SelectedBic { get; private set; }

        public class TagItem
        {
            public BuiltInCategory Bic { get; set; }
            public string Label { get; set; }       // nice short label
            public string FullLabel { get; set; }   // tooltip text / debug
        }

        public TagButtons(Document doc, IEnumerable<BuiltInCategory> tagBics)
        {
            InitializeComponent();
            _doc = doc;

            // Build button list (only categories that exist in this doc)
            foreach (var tagBic in tagBics)
            {
                var cat = Category.GetCategory(_doc, new ElementId((int)tagBic));
                if (cat == null) continue;

                string label;
                try { label = LabelUtils.GetLabelFor(tagBic); }
                catch { label = tagBic.ToString().Replace("OST_", ""); }

                _allItems.Add(new TagItem
                {
                    Bic = tagBic,
                    Label = label,
                    FullLabel = $"{label} ({tagBic})"
                });
            }

            // Sort and bind
            var sorted = _allItems.OrderBy(i => i.Label, StringComparer.CurrentCultureIgnoreCase).ToList();

            icButtons.ItemsSource = sorted;
        }

        public event Action<BuiltInCategory> TagButtonClicked;
        private void TagButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is BuiltInCategory bic)
            {
                SelectedBic = bic;  
                TagButtonClicked?.Invoke(bic);
                DialogResult = true; 
                Close();
            }
        }

        private void btnClose_Click(object sender, RoutedEventArgs e) => Close();
    }
}