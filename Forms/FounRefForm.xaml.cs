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
using Autodesk.Revit.UI;

namespace HDR_EMEA.Forms
{
    /// <summary>
    /// Interaction logic for FounRef.xaml
    /// </summary>
    public partial class FounRefForm : Window
    {
        private FoundationCommandHandler _handler;
        private ExternalEvent _exEvent;
        public FounRefForm(FoundationCommandHandler handler, ExternalEvent exEvent)
        {
            InitializeComponent();
            _handler = handler;
            _exEvent = exEvent;
        }

        private void btnConfirm_Click(object sender, RoutedEventArgs e)
        {
            if (rbFullMarking.IsChecked == true)
            {
                _handler.CurrentCommand = FoundationCommandHandler.CommandType.FullMarking; // Full pilecap/pile marking
            }
            else if (rbUpdateMarks.IsChecked == true)
            {
                _handler.CurrentCommand = FoundationCommandHandler.CommandType.UpdateMarks; // Only unmarked foundations
            }
            else if (rbMarkBySelection.IsChecked == true)
            {
                _handler.CurrentCommand = FoundationCommandHandler.CommandType.MarkBySelection; // Open FounFamType, apply marks
            }
            else if (rbMarkByGrids.IsChecked == true)
            {
                _handler.CurrentCommand = FoundationCommandHandler.CommandType.MarkByGrids; // Open FounFamType, grid-based numbering
            }
            else
            {
                TaskDialog.Show("HDR", "Please select an option.");
            }

            _exEvent.Raise();
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
