using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Linq;
using System.Windows;
using WpfCheckBox = System.Windows.Controls.CheckBox;

namespace HDR_EMEA.Forms
{
    public partial class SchedulePickerWindow : Window
    {
        private readonly UIDocument _uiDoc;
        private readonly ViewSchedule[] _schedules;

        /// <summary>The schedules the user picked (can be 1…N).</summary>
        public ViewSchedule[] SelectedSchedules { get; private set; }

        public SchedulePickerWindow(UIDocument uiDoc)
        {
            InitializeComponent();
            _uiDoc = uiDoc;

            // load schedules
            _schedules = new FilteredElementCollector(_uiDoc.Document)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(vs => !vs.IsTitleblockRevisionSchedule)
                .OrderBy(vs => vs.Name)
                .ToArray();

            // create a CheckBox for each schedule
            foreach (var vs in _schedules)
            {
                var cb = new WpfCheckBox
                {
                    Content = vs.Name,
                    Tag = vs,
                    Margin = new Thickness(2, 4, 2, 4)
                };
                stackSchedules.Children.Add(cb);
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            SelectedSchedules = stackSchedules.Children
                .OfType<WpfCheckBox>()
                .Where(cb => cb.IsChecked == true)
                .Select(cb => (ViewSchedule)cb.Tag)
                .ToArray();

            if (SelectedSchedules.Length == 0)
            {
                System.Windows.MessageBox.Show(
                    this,
                    "Please check at least one schedule.",
                    "Selection Required",
                    MessageBoxButton.OK,
                    MessageBoxImage.Warning
                );
                return;
            }
            DialogResult = true;
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            SelectedSchedules = new ViewSchedule[0];
            DialogResult = false;
        }
    }
}
