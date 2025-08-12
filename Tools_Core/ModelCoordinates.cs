using System;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class ModelCoordinates : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // read project base point and survey point from the model
            var (pbpEW_m, pbpNS_m, pbpElev_m, pbpAngle_rad) = ElementCollectorUtils.GetProjectBasePoint(doc);
            var (spEW_m, spNS_m, spElev_m) = ElementCollectorUtils.GetSurveyPoint(doc);

            // find the one FamilyInstance that holds our parameters
            var modelCoordInstance = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                .Cast<FamilyInstance>()
                .FirstOrDefault(fi => fi.Symbol?.Family?.Name == "HDR_QC_ModelCoordinates");

            if (modelCoordInstance == null)
            {
                TaskDialog.Show("Model Coordinates", "No instance of 'HDR_QC_ModelCoordinates' found.");
                return Result.Failed;
            }

            // the values we want to push into the parameters
            var paramMap = new Dictionary<string, double>
            {
                { "HC_ProjectBasePointEW",               pbpEW_m },
                { "HC_ProjectBasePointNS",               pbpNS_m },
                { "HC_ProjectBasePointElev",             pbpElev_m },
                { "HC_SurveyPointEW",                    spEW_m },
                { "HC_SurveyPointNS",                    spNS_m },
                { "HC_SurveyPointElev",                  spElev_m },
                { "HC_ProjectBasePointAngleToTrueNorth", pbpAngle_rad }
            };

            // this one is already in the right units (radians), so don't convert it
            HashSet<string> unitlessKeys = new HashSet<string>
            {
                "HC_ProjectBasePointAngleToTrueNorth"
            };

            using (var tx = new Transaction(doc, "Set Model Coordinates"))
            {
                tx.Start();

                foreach (var kvp in paramMap)
                {
                    Parameter param = modelCoordInstance.Symbol.LookupParameter(kvp.Key);
                    if (param == null
                        || param.IsReadOnly
                        || param.StorageType != StorageType.Double)
                    {
                        continue;
                    }

                    double valueToSet = kvp.Value;

                    // convert metres -> internal feet for everything except the angle
                    if (!unitlessKeys.Contains(kvp.Key))
                    {
                        valueToSet = UnitUtils.ConvertToInternalUnits(valueToSet, UnitTypeId.Meters);
                    }

                    param.Set(valueToSet);
                }

                tx.Commit();
            }

            TaskDialog.Show(
                "Model Coordinates",
                "Parameters updated with Project Base Point and Survey Point NS values."
            );
            return Result.Succeeded;
        }

        internal static Common.ButtonDataClass GetButtonData()
        {
            string buttonInternalName = "btnModelCoordinates";
            string buttonTitle = "Model Coordinates";
            string buttonTooltip = @"
Populate or update project parameters with the current Project Base Point and Survey Point North/South coordinates.

Version 1.0

Instructions:
- Click the button to update the parameters in this model.
- Ensure parameters 'ProjectBasePointNS' and 'SurveyPointNS' exist in the project.
- Values are updated in metres.";

            Common.ButtonDataClass.ButtonStatus status = Common.ButtonDataClass.ButtonStatus.Standard;
            bool isNew = false, isUpdate = false;

            if (isNew)
                status = Common.ButtonDataClass.ButtonStatus.New;
            else if (isUpdate)
                status = Common.ButtonDataClass.ButtonStatus.Update;

            return new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.ModelCoordinates_32,
                Properties.Resources.ModelCoordinates_16,
                status,
                buttonTooltip
            );
        }
    }
}
