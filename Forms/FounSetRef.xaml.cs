using Autodesk.Revit.DB.Structure;
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

namespace HDR_EMEA.Forms
{
    /// <summary>
    /// Interaction logic for FounSetRef.xaml
    /// </summary>
    public partial class FounSetRef : Window
    {
        public string Prefix => prefix.Text;
        public string StartNumber => startnumber.Text;
        public string Suffix => suffix.Text;

        public FounSetRef()
        {
            InitializeComponent();
        }

        private void btnSET_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnFinish_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }
    }
}
