using System;
using System.Linq;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HDR_EMEA.Tools_Structural
{
    [Transaction(TransactionMode.Manual)]
    public class PileProximityDelete : IExternalCommand
    {
        // Match all year variants, e.g. HDR_General_PileChecker23/24/25
        private const string CheckerFamilyPrefix = "HDR_General_PileChecker";

        public Result Execute(ExternalCommandData data,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uiDoc = data.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Find all matching families by prefix
            var families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(f => f.Name.StartsWith(CheckerFamilyPrefix, StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (!families.Any())
            {
                TaskDialog.Show("Delete Pile Checkers",
                    $"No families found starting with '{CheckerFamilyPrefix}'.");
                return Result.Cancelled;
            }

            // Collect all instances of those families
            var famIds = new HashSet<ElementId>(families.Select(f => f.Id));

            var instanceIds = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family != null && famIds.Contains(fi.Symbol.Family.Id))
                .Select(fi => fi.Id)
                .ToList();

            int deletedInstances = 0;
            int deletedFamilies = 0;

            using (var tx = new Transaction(doc, "Delete Pile Proximity Checkers"))
            {
                tx.Start();

                if (instanceIds.Any())
                {
                    var result = doc.Delete(instanceIds);
                    deletedInstances = result?.Count ?? 0;
                }

                // Attempt to delete family definitions now that instances are gone
                foreach (var fam in families)
                {
                    try
                    {
                        // Skip if any type still in use for some reason
                        if (fam.GetFamilySymbolIds().Any(sid =>
                        {
                            var sym = doc.GetElement(sid) as FamilySymbol;
                            // If symbol has any placed instances it is still in use
                            return new FilteredElementCollector(doc)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .Any(fi => fi.Symbol?.Id == sid);
                        }))
                            continue;

                        var res = doc.Delete(fam.Id);
                        if (res != null && res.Count > 0)
                            deletedFamilies++;
                    }
                    catch
                    {
                        // Ignore failures (e.g. still in use or protected)
                    }
                }

                tx.Commit();
            }

            TaskDialog.Show("Delete Pile Checkers",
                $"Deleted {deletedInstances} checker instance(s).\n" +
                $"Removed {deletedFamilies} checker family definition(s) matching '{CheckerFamilyPrefix}'.");
            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "PileProximityDelete",
                "Delete\nCheckers",
                "HDR_EMEA.Tools_Structural.PileProximityDelete",
                Properties.Resources.PileProximityDelete_32,
                Properties.Resources.PileProximityDelete_16,
                $"Remove all instances (and unused families) whose name starts with '{CheckerFamilyPrefix}'."
            );
        }
    }
}
