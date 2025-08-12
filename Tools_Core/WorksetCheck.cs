using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class WorksetCheck : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Ensure the model is workshared
            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Workset Check", "Model is not workshared.");
                return Result.Failed;
            }

            // Find the 3D ViewFamilyType
            var vft = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewFamilyType))
                .Cast<ViewFamilyType>()
                .FirstOrDefault(x => x.ViewFamily == ViewFamily.ThreeDimensional);

            if (vft == null)
            {
                TaskDialog.Show("Workset Check", "No 3D ViewFamilyType found.");
                return Result.Failed;
            }

            // Remove the default view template from the 3D view type so new views won't inherit it
            if (vft.DefaultTemplateId != ElementId.InvalidElementId)
            {
                using (var tx = new Transaction(doc, "Remove 3D View Type Template"))
                {
                    tx.Start();
                    vft.DefaultTemplateId = ElementId.InvalidElementId;
                    tx.Commit();
                }
            }

            // Collect user worksets
            IList<Workset> worksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .ToWorksets();

            if (worksets == null || worksets.Count == 0)
            {
                TaskDialog.Show("Workset Check", "No user worksets found.");
                return Result.Failed;
            }

            // Delete any existing views named "Workset_View - <Name>"
            var namesToDelete = new HashSet<string>(
                worksets.Select(ws => "Workset_View - " + ws.Name),
                StringComparer.OrdinalIgnoreCase);

            var existingViews = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => namesToDelete.Contains(v.Name))
                .Select(v => v.Id)
                .ToList();

            if (existingViews.Any())
            {
                using (var tx = new Transaction(doc, "Delete Existing Workset Views"))
                {
                    tx.Start();
                    doc.Delete(existingViews);
                    tx.Commit();
                }
            }

            // Create a new 3D view for each workset
            int createdCount = 0;
            using (var tx = new Transaction(doc, "Create Workset Views"))
            {
                tx.Start();
                foreach (var ws in worksets)
                {
                    var view3d = View3D.CreateIsometric(doc, vft.Id);
                    view3d.Name = "Workset_View - " + ws.Name;

                    // Remove any view template from the new view itself
                    view3d.ViewTemplateId = ElementId.InvalidElementId;

                    // Set visibility: show only the target workset
                    foreach (var otherWs in worksets)
                    {
                        var visibility = otherWs.Id == ws.Id
                            ? WorksetVisibility.Visible
                            : WorksetVisibility.Hidden;
                        view3d.SetWorksetVisibility(otherWs.Id, visibility);
                    }

                    createdCount++;
                }
                tx.Commit();
            }

            TaskDialog.Show(
                "Workset Check",
                createdCount > 0
                    ? $"Successfully created {createdCount} workset view{(createdCount > 1 ? "s" : string.Empty)}."
                    : "No new workset views were created."
            );

            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "WorksetCheck",
                "Workset Check",
                "HDR_EMEA.Tools_Core.WorksetCheck",
                Properties.Resources.WorksetCheck_32,
                Properties.Resources.WorksetCheck_16,
                "Generate an independent 3D view per workset, named 'Workset_View – <WorksetName>'"
            );
        }
    }
}
