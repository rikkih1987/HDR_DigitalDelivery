using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using System.Linq;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class WorksetRefLink : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get current selection
            var selIds = uidoc.Selection.GetElementIds();
            if (!selIds.Any())
            {
                TaskDialog.Show("Workset Ref Link", "At least one linked element must be selected.");
                return Result.Failed;
            }

            try
            {
                // Enable worksharing if needed
                if (!doc.IsWorkshared && doc.CanEnableWorksharing())
                {
                    using (Transaction tx = new Transaction(doc, "Enable Worksharing"))
                    {
                        tx.Start();
                        doc.EnableWorksharing("Shared Levels and Grids", "Workset1");
                        tx.Commit();
                    }
                }

                // Collect existing user worksets
                var existingWS = new FilteredWorksetCollector(doc)
                                     .OfKind(WorksetKind.UserWorkset)
                                     .ToWorksets();
                var existingNames = existingWS.Select(ws => ws.Name).ToList();

                // Create and assign worksets
                using (Transaction tx = new Transaction(doc, "Create Workset for Linked Models"))
                {
                    tx.Start();

                    foreach (var id in selIds)
                    {
                        Element el = doc.GetElement(id);
                        string linkedModelName = null;

                        // Handle Revit Link Instances
                        if (el is RevitLinkInstance linkInst)
                        {
                            linkedModelName = "ZLink_RVT_" + linkInst.Name.Split(':')[0];
                        }
                        // Handle Import Instances (CAD)
                        else if (el is ImportInstance importInst)
                        {
                            var param = importInst.get_Parameter(BuiltInParameter.IMPORT_SYMBOL_NAME);
                            if (param != null)
                            {
                                string symName = param.AsString();
                                if (!string.IsNullOrEmpty(symName))
                                    linkedModelName = "ZLink_CAD" + symName;
                            }
                        }

                        if (!string.IsNullOrEmpty(linkedModelName) && !existingNames.Contains(linkedModelName))
                        {
                            Workset newWS = Workset.Create(doc, linkedModelName);
                            existingNames.Add(linkedModelName);
                        }

                        if (!string.IsNullOrEmpty(linkedModelName))
                        {
                            Workset workset = new FilteredWorksetCollector(doc)
                                                  .OfKind(WorksetKind.UserWorkset)
                                                  .FirstOrDefault(ws => ws.Name == linkedModelName);
                            if (workset != null)
                            {
                                var wsParam = el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM);
                                if (wsParam != null && !wsParam.IsReadOnly)
                                {
                                    wsParam.Set(workset.Id.IntegerValue);
                                }
                            }
                        }
                    }

                    tx.Commit();
                }
            }
            catch (System.Exception ex)
            {
                TaskDialog.Show("Error", "An error occurred: " + ex.Message);
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "WorksetRefLink",
                "Workset Ref Link",
                "HDR_EMEA.Tools_Core.WorksetRefLink",
                Properties.Resources.WorksetRefLink_32,
                Properties.Resources.WorksetRefLink_16,
                "Create and assign worksets for linked models"
            );
        }
    }
}