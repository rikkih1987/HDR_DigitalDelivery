using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HDR_EMEA.Forms;
using static HDR_EMEA.Common.ButtonDataClass;

namespace HDR_EMEA
{
    [Transaction(TransactionMode.Manual)]
    public class PileCoordinates : IExternalCommand
    {
        private readonly List<string> parameterNames = new()
        {
            "Coordinates_Survey_X",
            "Coordinates_Survey_Y",
            "Coordinates_Survey_Z",
            "Coordinates_Project_X",
            "Coordinates_Project_Y",
            "Coordinates_Project_Z",
            "Coordinates_ScriptLastExecuted"
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            bool success = false;

            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // Save embedded shared parameter file to a temp location
                string tempPath = Path.Combine(Path.GetTempPath(), "HDR_SharedParameters.txt");
                using (Stream resource = Assembly.GetExecutingAssembly().GetManifestResourceStream("HDR_EMEA.Resources.HDR_SharedParameters.txt"))
                using (FileStream file = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                {
                    resource.CopyTo(file);
                }

                uiapp.Application.SharedParametersFilename = tempPath;
                DefinitionFile sharedParameterFile = uiapp.Application.OpenSharedParameterFile();
                if (sharedParameterFile == null)
                {
                    message = "Shared parameter file could not be loaded.";
                    return Result.Failed;
                }

                DefinitionGroup group = sharedParameterFile.Groups.get_Item("Foundation Coordinates") ??
                                        sharedParameterFile.Groups.Create("Foundation Coordinates");

                CategorySet categorySet = uiapp.Application.Create.NewCategorySet();
                categorySet.Insert(doc.Settings.Categories.get_Item(BuiltInCategory.OST_StructuralFoundation));
                BindingMap bindingMap = doc.ParameterBindings;

                using (Transaction t = new Transaction(doc, "Bind Shared Parameters"))
                {
                    t.Start();
                    foreach (string paramName in parameterNames)
                    {
                        Definition def = group.Definitions.get_Item(paramName);
                        if (def == null) continue;

                        if (!bindingMap.Contains(def))
                        {
                            InstanceBinding binding = uiapp.Application.Create.NewInstanceBinding(categorySet);

                            #if REVIT2025
                            // Revit 2025+ has only the 2-parameter Insert
                             bindingMap.Insert(def, binding);
                            #else
                            // Revit 2023/2024 needs the 3-parameter Insert
                            bindingMap.Insert(def, binding, BuiltInParameterGroup.PG_DATA);
                            #endif
                        }
                    }
                    t.Commit();
                }

                // Get all placed FamilyInstances of Structural Foundations
                var placedInstances = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .OfClass(typeof(FamilyInstance))
                    .Cast<FamilyInstance>()
                    .ToList();

                // Extract distinct FamilySymbols used in placed instances
                List<FamilySymbol> allSymbols = placedInstances
                    .Select(f => f.Symbol)
                    .Where(sym => sym != null)
                    .Distinct(new SymbolEqualityComparer())
                    .ToList();

                // filteredSymbols = only Circular/Square pile types
                List<FamilySymbol> filteredSymbols = allSymbols
                    .Where(sym =>
                    {
                        Parameter p = sym.LookupParameter("Foundation_Type");
                        if (p == null || !p.HasValue) return false;
                        string val = p.AsString()?.Trim().ToLower();
                        return val == "circular pile" || val == "square pile";
                    })
                    .ToList();

                Forms.FounFamType currentFormSFF = new Forms.FounFamType(filteredSymbols,allSymbols)
                {
                    Width = 800,
                    Height = 450,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    Topmost = true,
                };

                bool? dialogResult = currentFormSFF.ShowDialog();
                List<FamilySymbol> selectedSymbols = currentFormSFF.GetSelectedSymbols();

                if (dialogResult != true || !selectedSymbols.Any())
                    return Result.Cancelled;

                int updatedCount = 0;
                using (Transaction t = new Transaction(doc, "Set Pile Coordinates"))
                {
                    t.Start();

                    Transform sharedToProject = doc.ActiveProjectLocation.GetTotalTransform();
                    Transform projectToShared = sharedToProject.Inverse;


                    var symbolIds = selectedSymbols.Select(symbol => symbol.Id).ToHashSet();

                    var instances = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                        .OfClass(typeof(FamilyInstance))
                        .Cast<FamilyInstance>()
                        .Where(fi => symbolIds.Contains(fi.Symbol.Id))
                        .ToList();

                    foreach (FamilyInstance inst in instances)
                    {
                        if (inst.Location is not LocationPoint location) continue;

                        XYZ projectPoint = location.Point;
                        XYZ surveyPoint = projectToShared.OfPoint(projectPoint);


                        double sx = Math.Round(UnitUtils.ConvertFromInternalUnits(surveyPoint.X, UnitTypeId.Meters), 3);
                        double sy = Math.Round(UnitUtils.ConvertFromInternalUnits(surveyPoint.Y, UnitTypeId.Meters), 3);
                        double sz = Math.Round(UnitUtils.ConvertFromInternalUnits(surveyPoint.Z, UnitTypeId.Meters), 3);
                        double px = Math.Round(UnitUtils.ConvertFromInternalUnits(projectPoint.X, UnitTypeId.Millimeters), 3);
                        double py = Math.Round(UnitUtils.ConvertFromInternalUnits(projectPoint.Y, UnitTypeId.Millimeters), 3);
                        double pz = Math.Round(UnitUtils.ConvertFromInternalUnits(projectPoint.Z, UnitTypeId.Millimeters), 3);

                        bool updated = false;
                        updated |= Utils.SetParameter(inst, "Coordinates_Survey_X", sx.ToString("F3"));
                        updated |= Utils.SetParameter(inst, "Coordinates_Survey_Y", sy.ToString("F3"));
                        updated |= Utils.SetParameter(inst, "Coordinates_Survey_Z", sz.ToString("F3"));
                        updated |= Utils.SetParameter(inst, "Coordinates_Project_X", px.ToString("F3"));
                        updated |= Utils.SetParameter(inst, "Coordinates_Project_Y", py.ToString("F3"));
                        updated |= Utils.SetParameter(inst, "Coordinates_Project_Z", pz.ToString("F3"));

                        Utils.SetParameter(inst, "Coordinates_ScriptLastExecuted", DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));

                        if (updated) updatedCount++;
                    }

                    Element projectInfo = new FilteredElementCollector(doc)
                        .OfCategory(BuiltInCategory.OST_ProjectInformation)
                        .FirstElement();

                    if (projectInfo != null)
                    {
                        Parameter piParam = projectInfo.LookupParameter("Coordinates_ScriptLastExecuted");
                        if (piParam != null && !piParam.IsReadOnly)
                        {
                            piParam.Set(DateTime.Now.ToString("dd/MM/yyyy HH:mm:ss"));
                        }
                    }

                    t.Commit();
                }

                TaskDialog.Show("Results", $"Coordinates updated for {updatedCount} structural foundation elements.");
                success = true;
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("PileCoordinates Error", ex.Message);
                return Result.Failed;
            }
            finally
            {
                Tracker.LogCommandUsage("PileCoordinates", startTime, success, uiapp);
            }
        }

        class SymbolEqualityComparer : IEqualityComparer<FamilySymbol>
        {
            public bool Equals(FamilySymbol x, FamilySymbol y) => x?.Id == y?.Id;
            public int GetHashCode(FamilySymbol obj) => obj.Id.GetHashCode();
        }

        internal static Common.ButtonDataClass GetButtonData()
        {
            string buttonInternalName = "btnPileCoordinates";
            string buttonTitle = "Pile Coordinates";

            ButtonStatus status = ButtonStatus.Standard;
            string buttonTooltip = @"
Extracts and updates survey and project coordinates for selected foundation types placed in the model.

- Shared parameters are automatically bound if missing
- Includes X, Y, Z in both Survey and Project coordinates
- Timestamp is written to each element and project info

Version 1.2";

            Common.ButtonDataClass myButtonData = new(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.PileCoordinates_16,
                Properties.Resources.PileCoordinates_16,

                status,
                buttonTooltip);

            return myButtonData;
        }
    }
}
