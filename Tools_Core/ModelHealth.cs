using Autodesk.Revit.UI;
using HDR_EMEA.Forms;
using System.Security.Cryptography.X509Certificates;
using System.Windows.Media;
using static HDR_EMEA.Common.ButtonDataClass;

namespace HDR_EMEA
{
    [Transaction(TransactionMode.Manual)]
    public class ModelHealth : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            bool success = false;

            try
            {
                // Revit application and document variables
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                // Your code goes here
                ModelHealthViewModel viewModel = new ModelHealthViewModel()
                {
                    ModelName = Utils.GetParameterValue(doc.ProjectInformation, "Model Name"),
                    ProjectName = Utils.GetParameterValue(doc.ProjectInformation, "Project Name")
                };
                ModelHealthAssesment modelHealth = new ModelHealthAssesment(viewModel);
                ExternalEvent externalEvent = ExternalEvent.Create(modelHealth);

                HDR_EMEA.Handlers.AppContext_.ExternalEventModelHealth = externalEvent;

                // Open form
                Forms.ModelHealthWPF formModelHealth = new Forms.ModelHealthWPF(viewModel)
                {
                    Width = 780,
                    Height = 830,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    Topmost = true,
                };

                formModelHealth.Show();
                externalEvent.Raise();
                success = true;
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Model Health Error", ex.Message);
                return Result.Failed;
            }
            finally
            {
                Tracker.LogCommandUsage("ModelHealth", startTime, success);
            }
        }


        internal class ModelHealthAssesment : IExternalEventHandler
        {
            public ModelHealthViewModel ViewModel { get;}
            public string ModelName { get; set; }
            public string ProjectName { get; set; }
            public ModelHealthAssesment (ModelHealthViewModel viewModel)
            {
                ViewModel = viewModel;
            }
            public void Execute(UIApplication app)
            {
                Document doc = app.ActiveUIDocument.Document;
                //
                ViewModel.HC_ModelSize      = GetModelSize(doc);
                ViewModel.HC_GroupInstances = ElementCollectorUtils.GetGroupInstanceCount(doc);
                ViewModel.HC_GroupTypes     = ElementCollectorUtils.GetGroupTypeCount(doc);
                ViewModel.HC_UnusedGroups   = ElementCollectorUtils.GetUnusedGroupTypeCount(doc);
                ViewModel.HC_ModelElements  = ElementCollectorUtils.Get3DModelElementCount(doc);
                //
                ViewModel.HC_DesignOptions  = ElementCollectorUtils.GetDesignOptionCount(doc);
                ViewModel.HC_Worksets       = ElementCollectorUtils.GetWorksetCount(doc);
                ViewModel.HC_CADImports     = ElementCollectorUtils.GetCADImportCount(doc);
                ViewModel.HC_CADLinks       = ElementCollectorUtils.GetCADLinkCount(doc);
                ViewModel.HC_Images         = ElementCollectorUtils.GetImageCount(doc);
                //
                ViewModel.HC_AreaofRooms     = ElementCollectorUtils.GetTotalRoomArea(doc);
                ViewModel.HC_Rooms           = ElementCollectorUtils.GetRoomCount(doc);
                ViewModel.HC_UnplacedRooms   = ElementCollectorUtils.GetUnplacedRoomCount(doc);
                ViewModel.HC_UnenclosedRooms = ElementCollectorUtils.GetUnenclosedRoomCount(doc);
                //
                ViewModel.HC_LoadedFamilies             = ElementCollectorUtils.GetLoadedFamilyCount(doc);
                ViewModel.HC_UnusedFamilies_Percentage  = ElementCollectorUtils.GetUnusedFamilyPercentage(doc);
                ViewModel.HC_InplaceFamilies            = ElementCollectorUtils.GetInplaceFamilyCount(doc);
                ViewModel.HC_ErrorCount                 = ElementCollectorUtils.GetWarningCount(doc);
                ViewModel.HC_ApprovedFamiliesPercentage = ElementCollectorUtils.GetApprovedFamilyPercentage(doc);
                //
                ViewModel.HC_TotalViews      = ElementCollectorUtils.GetViewCount(doc);
                ViewModel.HC_TotalSheets     = ElementCollectorUtils.GetSheetCount(doc);
                ViewModel.HC_ViewsWithNoVT   = ElementCollectorUtils.GetViewsWithNoTemplate(doc);
                ViewModel.HC_ViewsNotOnSheet = ElementCollectorUtils.GetViewsNotOnSheetCount(doc);
                ViewModel.HC_ClippingDisabled = ElementCollectorUtils.GetViewsWithClippingDisabled(doc);
                //
                ViewModel.HC_Materials      = ElementCollectorUtils.GetMaterialCount(doc);
                ViewModel.HC_LineStyles     = ElementCollectorUtils.GetLineStyleCount(doc);
                ViewModel.HC_LinePatterns   = ElementCollectorUtils.GetLinePatternCount(doc);
                ViewModel.HC_FillPatterns   = ElementCollectorUtils.GetFillPatternCount(doc);
                ViewModel.HC_TextStyles     = ElementCollectorUtils.GetTextStyleCount(doc);
            }
            public string GetName() => "Model Health Assesment";


        }
        private static double GetModelSize(Document doc)
        {
            double? sizeMB = null;

            var path = doc.PathName;
            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    var fileInfo = new FileInfo(path);
                    sizeMB = Math.Round(fileInfo.Length / 1_000_000.0, 2);
                }
                catch { sizeMB = null; }
            }

            if (!sizeMB.HasValue)
            {
                ModelPath centralPath = doc.GetWorksharingCentralModelPath();
                string centralPathString = ModelPathUtils.ConvertModelPathToUserVisiblePath(centralPath);

                if (!string.IsNullOrEmpty(centralPathString) && File.Exists(centralPathString))
                {
                    try
                    {
                        var fileInfo = new FileInfo(centralPathString);
                        sizeMB = Math.Round(fileInfo.Length / 1_000_000.0, 2);
                    }
                    catch { sizeMB = null; }
                }
            }

            return sizeMB ?? 0;
        }

        internal static Common.ButtonDataClass GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnModelHealth";
            string buttonTitle = "Model Health";
            string buttonTooltip = @"
Run a model health check to assess the current state of the project and export key metrics for future reference.

Version 1.0

Instructions:
- Click the button to generate the model health report.
- Review the displayed metrics in the UI.
- Click 'EXPORT' to save the report as an Excel file.
- The export file will be named using the project identifier and current date.";

            ButtonStatus status = ButtonStatus.Standard;

            // These would be toggled as needed:
            bool isNew = false;
            bool isUpdate = false;

            if (isNew)
                status = ButtonDataClass.ButtonStatus.New;
            else if (isUpdate)
                status = ButtonDataClass.ButtonStatus.Update;

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.ModelHealth_32,
                Properties.Resources.ModelHealth_16,
                status,
                buttonTooltip);

            return myButtonData;

        }
    }

}
