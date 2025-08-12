using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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

namespace HDR_EMEA.Forms
{
    /// <summary>
    /// Interaction logic for SheetManagerForm.xaml
    /// </summary>
    public partial class SheetManagerForm : Window
    {
        public class RevisionDisplay
        {
            public string DisplayText { get; set; }
            public int RevisionId { get; set; } 
        }
        public int? SelectedRevisionId { get; private set; }
        public List<RevisionDisplay> AvailableRevisions { get; set; } = new();
        public List<string> AvailableBannerStatuses { get; set; } = new List<string>();
        public List<string> AvailableSuitabilityStatuses { get; set; } = new List<string>();
        public class SheetInfo : INotifyPropertyChanged
        {
            public bool IsSelected { get; set; }
            public string SheetName { get; set; }
            public string SheetNumber { get; set; }
            public string CurrentRev { get; set; }
            public string CurrentBanner { get; set; }
            public string CurrentStatus { get; set; }

            public event PropertyChangedEventHandler PropertyChanged;
            protected void Raise(string name)
                => PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
        public ObservableCollection<SheetInfo> Sheets { get; set; } = new ObservableCollection<SheetInfo>();
        public string SelectedBanner { get; private set; }
        public string SelectedStatus { get; private set; }
        public SheetManagerForm(
            List<RevisionDisplay> availableRevisions,
            List<string> availableBannerStatuses,
            List<string> availableSuitabilityStatuses,
            List<SheetInfo> sheets)
        {
            AvailableRevisions = availableRevisions;
            AvailableBannerStatuses = availableBannerStatuses;
            AvailableSuitabilityStatuses = availableSuitabilityStatuses;

            foreach (var sheet in sheets)
            {
                Sheets.Add(sheet);
            } 

            InitializeComponent();
            this.DataContext = this;
        }

        private void HeaderCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (dgFamilyGrid.ItemsSource is IEnumerable<FamilyTypeSelection> items)
            {
                bool isChecked = (sender as System.Windows.Controls.CheckBox)?.IsChecked == true;

                foreach (var item in items)
                {
                    item.IsSelected = isChecked;
                }

                dgFamilyGrid.Items.Refresh();
            }
        }

        private void btnSet_Click(object sender, RoutedEventArgs e)
        {
            var selectedSheets = Sheets.Where(s => s.IsSelected).ToList();
            if (!selectedSheets.Any())
            {
                this.Close();
                TaskDialog.Show("Sheet Manager", "No sheets selected.");
                return;
            }

            var revDisplay = chkRevision.IsChecked == true ? cbRevision.SelectedItem as RevisionDisplay : null;
            int? selectedRevisionId = revDisplay?.RevisionId;

            string selectedBanner = chkBanner.IsChecked == true ? cbBanner.SelectedItem as string : null;
            string selectedStatus = chkStatus.IsChecked == true ? cbStatus.SelectedItem as string : null;

            if (selectedRevisionId == null && selectedBanner == null && selectedStatus == null)
            {
                this.Close();
                TaskDialog.Show("Sheet Manager", "No change selected.");
                return;
            }

            this.DialogResult = true;

            SelectedRevisionId = selectedRevisionId;
            SelectedBanner = selectedBanner;
            SelectedStatus = selectedStatus;

            this.Close();
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
