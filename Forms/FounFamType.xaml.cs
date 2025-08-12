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
using System.ComponentModel;
using CheckBox = System.Windows.Controls.CheckBox;

namespace HDR_EMEA.Forms
{
    /// <summary>
    /// Interaction logic for FounFamType.xaml
    /// </summary>
    public class FamilyTypeSelection : INotifyPropertyChanged
    {
        public string FamilyName { get; set; }
        public string FamilyType { get; set; }
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                if (_isSelected != value)
                {
                    _isSelected = value;
                    OnPropertyChanged(nameof(IsSelected));
                }
            }
        }
        public FamilySymbol Symbol { get; set; }

        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged(string name) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public partial class FounFamType : Window
    {
        // Backing lists
        private readonly List<FamilySymbol> _filteredSymbols;
        private readonly List<FamilySymbol> _allSymbols;

        // This is the field your handlers and GetSelectedSymbols() use
        private List<FamilyTypeSelection> familyTypeSelections;

        /// <summary>
        /// Original constructor (kept so other callers don't break). 
        /// Shows exactly the symbols passed in.
        /// </summary>
        public FounFamType(List<FamilySymbol> symbols)
        {
            InitializeComponent();

            _allSymbols = symbols ?? new List<FamilySymbol>();
            _filteredSymbols = _allSymbols; // No filter if using original ctor

            ApplySymbols(showAll: true);
        }

        /// <summary>
        /// New constructor: pass filtered + all used symbols.
        /// By default (toggle off) shows filtered list.
        /// </summary>
        public FounFamType(List<FamilySymbol> filteredSymbols, List<FamilySymbol> allSymbols)
        {
            InitializeComponent();

            _filteredSymbols = filteredSymbols ?? new List<FamilySymbol>();
            _allSymbols = allSymbols ?? new List<FamilySymbol>();

            // Default view: filtered (toggle off)
            ApplySymbols(showAll: false);
        }

        /// <summary>
        /// Builds the grid rows from either the filtered or full set.
        /// </summary>
        private void ApplySymbols(bool showAll)
        {
            var symbols = showAll ? _allSymbols : _filteredSymbols;

            familyTypeSelections = symbols.Select(s => new FamilyTypeSelection
            {
                FamilyName = s.Family?.Name,
                FamilyType = s.Name,
                Symbol = s,
                IsSelected = false
            }).ToList();

            dgFamilyGrid.ItemsSource = familyTypeSelections;
            dgFamilyGrid.Items.Refresh();
        }

        // Toggle handlers wired from XAML
        private void tglShowAll_Checked(object sender, RoutedEventArgs e) => ApplySymbols(showAll: true);
        private void tglShowAll_Unchecked(object sender, RoutedEventArgs e) => ApplySymbols(showAll: false);

        public List<FamilySymbol> GetSelectedSymbols() =>
            familyTypeSelections?
                .Where(s => s.IsSelected)
                .Select(s => s.Symbol)
                .ToList() ?? new List<FamilySymbol>();

        private void HeaderCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (dgFamilyGrid.ItemsSource is IEnumerable<FamilyTypeSelection> rows)
            {
                bool isChecked = (sender as CheckBox)?.IsChecked == true;
                foreach (var r in rows) r.IsSelected = isChecked;
                dgFamilyGrid.Items.Refresh();
            }
        }

        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
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