using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
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
    /*
    public class HealthToCanvasTopConverter : IValueConverter
    {
        public double MaxHeight { get; set; } = 660; // Total height in pixels (depends on your layout)

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is double health)
            {
                health = Math.Max(0, Math.Min(10, health));
                double ratio = 1.0 - (health / 10.0);
                double offset = 0;
                if (parameter != null && double.TryParse(parameter.ToString(), out double parsed))
                {
                    offset = parsed; // Use the parameter value as an offset
                }

                return (ratio * MaxHeight) + offset;
            }
            return 0;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture) =>
            throw new NotImplementedException();
    }
    */
    
    public partial class ModelHealthWPF : Window
    {
        public ModelHealthViewModel ViewModel { get; set; }
        public ModelHealthWPF(ModelHealthViewModel viewModel)
        {
            InitializeComponent();
            ViewModel = viewModel;
            this.DataContext = ViewModel;

            //if (this.Resources["HealthToCanvasTopConverter"] is HealthToCanvasTopConverter converter)
            //{
            //    converter.MaxHeight = 660; // Set the maximum height for the converter
            //}
        }

        private void btnEXPORT_Click(object sender, RoutedEventArgs e)
        {
            var modelName = ViewModel.ModelName?.Replace(" ", "_") ?? " ";
            var timestamp = DateTime.Now.ToString("dd-MM-yyyy_HH'h'mm'm'");
            string userName = Environment.UserName;

            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                FileName = $"{modelName}_{timestamp}_ModelHealthReport",
                DefaultExt = ".xlsx",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                Title = "Export Model Health Report"
            };

            if (saveFileDialog.ShowDialog() == true)
                try
                {
                    using (var workbook = new ClosedXML.Excel.XLWorkbook())
                    {
                        var ws = workbook.Worksheets.Add("Model Health");

                        int col = 1;
                        void Write(string label, object value)
                        {
                            ws.Cell(col, 1).Value = label;
                            ws.Cell(col, 2).Value = value?.ToString() ?? string.Empty;
                            col++;
                        }

                        var vm = ViewModel;

                        Write("Model Size",             vm.HC_ModelSize);
                        Write("Group Instances",        vm.HC_GroupInstances);
                        Write("Group Types",            vm.HC_GroupTypes);
                        Write("Unused Groups",          vm.HC_UnusedGroups);
                        Write("Model Elements",         vm.HC_ModelElements);
                        Write("Design Options",         vm.HC_DesignOptions);
                        Write("Worksets",               vm.HC_Worksets);
                        Write("CAD Imports",            vm.HC_CADImports);
                        Write("CAD Links",              vm.HC_CADLinks);
                        Write("Images",                 vm.HC_Images);
                        Write("Area of Rooms",          vm.HC_AreaofRooms);
                        Write("Rooms",                  vm.HC_Rooms);
                        Write("Unplaced Rooms",         vm.HC_UnplacedRooms);
                        Write("Unenclosed Rooms",       vm.HC_UnenclosedRooms);
                        Write("Loaded Families",        vm.HC_LoadedFamilies);
                        Write("Unused Families (%)",    vm.HC_UnusedFamilies_Percentage);
                        Write("Inplace Families",       vm.HC_InplaceFamilies);
                        Write("Warning Count",          vm.HC_ErrorCount);
                        Write("Approved Families (%)",  vm.HC_ApprovedFamiliesPercentage);
                        Write("Total Views",            vm.HC_TotalViews);
                        Write("Views With No VT",       vm.HC_ViewsWithNoVT);
                        Write("Total Sheets",           vm.HC_TotalSheets);
                        Write("Views Not On Sheet",     vm.HC_ViewsNotOnSheet);
                        Write("Clipping Disabled",      vm.HC_ClippingDisabled);
                        Write("Materials",              vm.HC_Materials);
                        Write("Line Styles",            vm.HC_LineStyles);
                        Write("Line Patterns",          vm.HC_LinePatterns);
                        Write("Fill Patterns",          vm.HC_FillPatterns);
                        Write("Text Styles",            vm.HC_TextStyles);
                        //Write("Duplicate Elements", vm.HC_DuplicateElements);

                        workbook.SaveAs(saveFileDialog.FileName);

                        // Silent export to network folder using Project Name
                        string networkRootPath = @"\\omesrv3\ASMECServiceCenter\BIMdev\EMEA_RevitToolbar\ModelHealthExports"; // Replace accordingly
                        if (Directory.Exists(networkRootPath))
                        {
                            string projectName = ViewModel?.ProjectName ?? "UnknownProject";
                            string sanitizedProjectName = string.Join("_", projectName.Split(System.IO.Path.GetInvalidFileNameChars()));
                            string projectFolderPath = System.IO.Path.Combine(networkRootPath, sanitizedProjectName);

                            if (!Directory.Exists(projectFolderPath))
                                Directory.CreateDirectory(projectFolderPath);

                            string networkFileName = $"{modelName}_{timestamp}_ModelHealthReport_{userName}.xlsx";
                            string fullNetworkPath = System.IO.Path.Combine(projectFolderPath, networkFileName);


                            workbook.SaveAs(fullNetworkPath);
                        }

                        this.Close();
                        TaskDialog.Show("Export", "Export successful!");
                    }
                }
                catch (Exception ex)
                {
                    this.Close();
                    TaskDialog.Show("Error", $"Failed to export: {ex.Message}");
                }
        }
    }

    public class ModelHealthViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        private void Notify([CallerMemberName] string name = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));

        private void NotifyStatus(params string[] properties)
        {
            foreach (var prop in properties)
                Notify(prop);
        }

        // Utility Colours
        private SolidColorBrush Green => new SolidColorBrush(System.Windows.Media.Color.FromRgb(0, 192, 48));
        private SolidColorBrush Orange => new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 160, 66));
        private SolidColorBrush Red => new SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 81, 81));

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        /*
        private const double ProjectHealthBase = 1.0; // Base value for project health calculations

        public double CadImportScore => HC_CADImports < 6 ? 0 : ProjectHealthBase * 1.25;
        public double CadLinkScore => HC_CADLinks < 14 ? 0 : ProjectHealthBase * 0.75;
        public double ClipScore => HC_ClippingDisabled < HC_TotalViews * 0.3 ? 0 : ProjectHealthBase * 1;
        public double DesignOptionScore => HC_DesignOptions < HC_ModelSize * 0.05 ? 0 : ProjectHealthBase * 1;
        public double GroupScore => HC_UnusedGroups < UnusedGroups_Target ? 0 : ProjectHealthBase * 1;
        public double ImageScore => HC_Images < 24 ? 0 : ProjectHealthBase * 1;
        public double InplaceFamiliesScore => InplaceFamilies_HealthCheck < 1 ? 0 : ProjectHealthBase * 1.5;
        public double ModelSizeScore => HC_ModelSize < 600 ? 0 : ProjectHealthBase * 1.5;
        public double SheetScore => HC_ViewsNotOnSheet < ViewsNotOnSheets_Target + (ViewsNotOnSheets_Target * 0.5) ? 0 : ProjectHealthBase * 1;
        public double TemplateScore => HC_ViewsWithNoVT < HC_TotalViews * 0.15 ? 0 : ProjectHealthBase * 0.5;
        public double UnenclosedRoomScore => UnenclosedRoom_HealthCheck < 4 ? 0 : ProjectHealthBase * 1;
        public double UnplacedRoomScore => UnplacedRoom_HealthCheck < 4 ? 0 : ProjectHealthBase * 1;
        public double UnusedFamiliesScore => HC_UnusedFamilies_Percentage < 35 ? 0 : ProjectHealthBase * 0.5;
        public double WarningCountScore => HC_ErrorCount < HC_ModelSize ? 0 : ProjectHealthBase * 1.5;
        public double WorksetScore => HC_Worksets < 27 ? 0 : ProjectHealthBase * 0.5;

        public double TotalScore =>
            (ModelSizeScore + GroupScore + DesignOptionScore + WorksetScore + CadImportScore + CadLinkScore +
            ImageScore + UnplacedRoomScore + UnenclosedRoomScore + UnusedFamiliesScore + InplaceFamiliesScore +
            WarningCountScore + TemplateScore + SheetScore + ClipScore) / 100;

        public double ProjectHealth => Math.Round(10 - (TotalScore * 10), 1);
        */
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Model Sizes Colour
        private double _hC_ModelSize;
        public double HC_ModelSize
        {
            get => _hC_ModelSize;
            set { _hC_ModelSize = value; Notify(); NotifyStatus(nameof(HC_ModelSizeStatus), nameof(HC_DesignOptionsStatus), nameof(HC_ErrorCountStatus)); }
        }
        public SolidColorBrush HC_ModelSizeStatus =>
            HC_ModelSize > 600 ? Red :
            (HC_ModelSize > 500 ? Orange : Green);
        //
        private int _hC_GroupInstances;
        public int HC_GroupInstances { get => _hC_GroupInstances; set { _hC_GroupInstances = value; Notify(); } }
        //
        private int _hC_GroupTypes;
        public int HC_GroupTypes { get => _hC_GroupTypes; set { _hC_GroupTypes = value; Notify(); } }

        // UnusedGroups Colours & Target
        private int _hC_UnusedGroups;
        public int HC_UnusedGroups
        {
            get => _hC_UnusedGroups;
            set { _hC_UnusedGroups = value; Notify(); Notify(nameof(HC_UnusedGroupsStatus)); }
        }
        public double UnusedGroups_Target => HC_GroupInstances * 0.2;
        public SolidColorBrush HC_UnusedGroupsStatus =>
            HC_UnusedGroups > UnusedGroups_Target * 2 ? Red :
            (HC_UnusedGroups > UnusedGroups_Target * 1.25 ? Orange : Green);

        //Model Elements
        private int _hC_ModelElements;
        public int HC_ModelElements
        {
            get => _hC_ModelElements;
            set { _hC_ModelElements = value; Notify(); }
        }

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // Design Options
        private int _hC_DesignOptions;
        public int HC_DesignOptions
        {
            get => _hC_DesignOptions;
            set { _hC_DesignOptions = value; Notify(); Notify(nameof(HC_DesignOptionsStatus)); }
        }
        public SolidColorBrush HC_DesignOptionsStatus =>
            HC_DesignOptions > HC_ModelSize * 0.05 ? Red :
            (HC_DesignOptions > HC_ModelSize * 0.025 ? Orange : Green);

        // Worksets
        private int _hC_Worksets;
        public int HC_Worksets
        {
            get => _hC_Worksets;
            set { _hC_Worksets = value; Notify(); Notify(nameof(HC_WorksetsStatus)); }
        }
        public SolidColorBrush HC_WorksetsStatus =>
            HC_Worksets > 26 ? Red :
            (HC_Worksets > 22 ? Orange : Green);

        // CADImports
        private int _hC_CADImports;
        public int HC_CADImports
        {
            get => _hC_CADImports;
            set { _hC_CADImports = value; Notify(); Notify(nameof(HC_CADImportsStatus)); }
        }
        public SolidColorBrush HC_CADImportsStatus =>
            HC_CADImports > 5 ? Red :
            (HC_CADImports > 0 ? Orange : Green);

        // CADLinks
        private int _hC_CADLinks;
        public int HC_CADLinks
        {
            get => _hC_CADLinks;
            set { _hC_CADLinks = value; Notify(); Notify(nameof(HC_CADLinksStatus)); }
        }
        public SolidColorBrush HC_CADLinksStatus =>
            HC_CADLinks > 13 ? Red :
            (HC_CADLinks > 9 ? Orange : Green);

        // Images
        private int _hC_Images;
        public int HC_Images
        {
            get => _hC_Images;
            set { _hC_Images = value; Notify(); Notify(nameof(HC_ImagesStatus)); }
        }
        public SolidColorBrush HC_ImagesStatus =>
            HC_Images > 24 ? Red :
            (HC_Images > 9 ? Orange : Green);
        
        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        private double _hC_AreaofRooms;
        public double HC_AreaofRooms { get => _hC_AreaofRooms; set { _hC_AreaofRooms = value; Notify(); } }
        //
        private int _hC_Rooms;
        public int HC_Rooms { get => _hC_Rooms; set { _hC_Rooms = value; Notify(); } }
        //
        private int _hC_UnplacedRooms;
        public int HC_UnplacedRooms { get => _hC_UnplacedRooms; set { _hC_UnplacedRooms = value; Notify(); } }

        // UnplacedRoom_HealthCheck
        public double UnplacedRoom_HealthCheck =>
            (HC_Rooms + HC_UnplacedRooms) > 0 ? (double)HC_UnplacedRooms / (HC_Rooms + HC_UnplacedRooms) * 100 : 0;

        public SolidColorBrush UnplacedRoom_HealthCheckStatus =>
            UnplacedRoom_HealthCheck > 4 ? Red :
            (UnplacedRoom_HealthCheck > 2 ? Orange : Green);

        //
        private int _hC_UnenclosedRooms;
        public int HC_UnenclosedRooms { get => _hC_UnenclosedRooms; set { _hC_UnenclosedRooms = value; Notify(); } }
        // UnenclosedRoom_HealthCheck
        public double UnenclosedRoom_HealthCheck =>
            (HC_Rooms + HC_UnenclosedRooms) > 0 ? (double)HC_UnenclosedRooms / (HC_Rooms + HC_UnenclosedRooms) * 100 : 0;

        public SolidColorBrush UnenclosedRoom_HealthCheckStatus =>
            UnenclosedRoom_HealthCheck > 4 ? Red :
            (UnenclosedRoom_HealthCheck > 2 ? Orange : Green);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        private int _hC_LoadedFamilies;
        public int HC_LoadedFamilies { get => _hC_LoadedFamilies; set { _hC_LoadedFamilies = value; Notify(); } }

        // UnusedFamilies_Percentage 
        private double _hC_UnusedFamilies_Percentage;
        public double HC_UnusedFamilies_Percentage 
        { 
            get => _hC_UnusedFamilies_Percentage;
            set { _hC_UnusedFamilies_Percentage = value; Notify(); Notify(nameof(HC_UnusedFamiliesStatus)); }
        }
        public SolidColorBrush HC_UnusedFamiliesStatus =>
            HC_UnusedFamilies_Percentage > 35.00 ? Red :
            (HC_UnusedFamilies_Percentage > 25.00 ? Orange : Green);

        //
        private int _hC_InplaceFamilies;
        public int HC_InplaceFamilies { get => _hC_InplaceFamilies; set { _hC_InplaceFamilies = value; Notify(); } }

        // InplaceFamilies_HealthCheck
        private double _inplaceFamilies_HealthCheck;
        public double InplaceFamilies_HealthCheck
        {
            get => _inplaceFamilies_HealthCheck;
            set { _inplaceFamilies_HealthCheck = value; Notify(); Notify(nameof(InplaceFamilies_HealthCheckStatus)); }
        }
        public SolidColorBrush InplaceFamilies_HealthCheckStatus =>
            InplaceFamilies_HealthCheck > 1 ? Red :
            (InplaceFamilies_HealthCheck > 0 ? Orange : Green);

        // Warning Count
        private int _hC_ErrorCount;
        public int HC_ErrorCount
        {
            get => _hC_ErrorCount;
            set { _hC_ErrorCount = value; Notify(); Notify(nameof(HC_ErrorCountStatus)); }
        }
        public SolidColorBrush HC_ErrorCountStatus =>
            HC_ErrorCount > HC_ModelSize ? Red :
            (HC_ErrorCount > HC_ModelSize * 0.6 ? Orange : Green);

        private double _hC_ApprovedFamiliesPercentage;
        public double HC_ApprovedFamiliesPercentage
        {
            get => _hC_ApprovedFamiliesPercentage;
            set { _hC_ApprovedFamiliesPercentage = value; Notify(); Notify(nameof(HC_ApprovedFamiliesStatus)); }
        }
        public SolidColorBrush HC_ApprovedFamiliesStatus =>
            HC_ApprovedFamiliesPercentage < 40 ? Red :
            (HC_ApprovedFamiliesPercentage < 70 ? Orange : Green);


        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        // TotalViews and ViewsWithNoVT 
        public int HC_TotalViews
        {
            get => _hC_TotalViews;
            set { _hC_TotalViews = value; Notify(); Notify(nameof(HC_ViewsWithNoVTStatus)); }
        }
        private int _hC_ViewsWithNoVT;
        public int HC_ViewsWithNoVT
        {
            get => _hC_ViewsWithNoVT;
            set { _hC_ViewsWithNoVT = value; Notify(); Notify(nameof(HC_ViewsWithNoVTStatus)); }
        }
        private int _hC_TotalViews;
        public SolidColorBrush HC_ViewsWithNoVTStatus =>
            HC_ViewsWithNoVT >  HC_TotalViews * 0.15 ? Red :
            (HC_ViewsWithNoVT > HC_TotalViews * 0.1 ? Orange : Green);

        //
        private int _hC_TotalSheets;
        public int HC_TotalSheets { get => _hC_TotalSheets; set { _hC_TotalSheets = value; Notify(); } }

        // ViewsNotOnSheet & Target
        private int _hC_ViewsNotOnSheet;
        public int HC_ViewsNotOnSheet
        {
            get => _hC_ViewsNotOnSheet;
            set { _hC_ViewsNotOnSheet = value; Notify(); Notify(nameof(HC_ViewsNotOnSheetStatus)); }
        }
        public double ViewsNotOnSheets_Target => HC_TotalViews * 0.2;
        public SolidColorBrush HC_ViewsNotOnSheetStatus =>
            HC_ViewsNotOnSheet > ViewsNotOnSheets_Target + ViewsNotOnSheets_Target * 0.5 ? Red :
            (HC_ViewsNotOnSheet > ViewsNotOnSheets_Target + ViewsNotOnSheets_Target * 0.1 ? Orange : Green);

        //
        private int _hC_ClippingDisabled;
        public int HC_ClippingDisabled
        {
            get => _hC_ClippingDisabled;
            set { _hC_ClippingDisabled = value; Notify(); Notify(nameof(HC_ClippingDisabledStatus)); }
        }
        public SolidColorBrush HC_ClippingDisabledStatus =>
            HC_ClippingDisabled > HC_TotalViews * 0.3 ? Red :
            (HC_ClippingDisabled > HC_TotalViews * 0.1 ? Orange : Green);

        ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //
        private int _hC_Materials;
        public int HC_Materials { get => _hC_Materials; set { _hC_Materials = value; Notify(); } }
        //
        private int _hC_LineStyles;
        public int HC_LineStyles { get => _hC_LineStyles; set { _hC_LineStyles = value; Notify(); } }
        //
        private int _hC_LinePatterns;
        public int HC_LinePatterns { get => _hC_LinePatterns; set { _hC_LinePatterns = value; Notify(); } }
        //
        private int _hC_FillPatterns;
        public int HC_FillPatterns { get => _hC_FillPatterns; set { _hC_FillPatterns = value; Notify(); } }
        //
        private int _hC_TextStyles;
        public int HC_TextStyles { get => _hC_TextStyles; set { _hC_TextStyles = value; Notify(); } }



        // Duplicate Elements
        private int _hC_DuplicateElements;
        public int HC_DuplicateElements
        {
            get => _hC_DuplicateElements;
            set { _hC_DuplicateElements = value; Notify(); Notify(nameof(HC_DuplicateElementsStatus)); }
        }
        public SolidColorBrush HC_DuplicateElementsStatus =>
            HC_DuplicateElements > 99 ? Red :
            (HC_DuplicateElements > 49 ? Orange : Green);

        public string ModelName { get; set; }
        public string ProjectName { get; set; }

        // Joining Elements ?? potencial 

        // Piles not hosted 
    }
}
