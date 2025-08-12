using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Structure;

namespace HDR_EMEA.Tools_Structural
{
    [Transaction(TransactionMode.Manual)]
    public class PileProximityAdd : IExternalCommand
    {
        // Remember the most recently loaded family (by internal name) for quicker lookups in-session
        static string _loadedFamilyInternalName = null;

        public Result Execute(ExternalCommandData data,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uiDoc = data.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            RemoveOldCheckers(doc);

            FamilySymbol checkerSym = FindCheckerSymbol(doc);
            if (checkerSym == null)
            {
                var fam = LoadEmbeddedFamilyForVersion(doc, data);
                if (fam == null)
                {
                    TaskDialog.Show("Pile Proximity",
                        "Failed to load embedded PileChecker family for this Revit version.");
                    return Result.Failed;
                }
                _loadedFamilyInternalName = fam.Name;
                checkerSym = doc.GetElement(
                    fam.GetFamilySymbolIds().FirstOrDefault()) as FamilySymbol;
            }

            if (checkerSym == null)
            {
                TaskDialog.Show("Pile Proximity", "Could not find a valid type in the loaded family.");
                return Result.Failed;
            }

            if (!checkerSym.IsActive)
            {
                using (var tx = new Transaction(doc, "Activate Checker"))
                {
                    tx.Start();
                    checkerSym.Activate();
                    tx.Commit();
                }
            }

            return PlaceAndColorCheckers(uiDoc, doc, checkerSym);
        }

        void RemoveOldCheckers(Document doc)
        {
            var sym = FindCheckerSymbol(doc);
            if (sym == null) return;

            ElementId famId = sym.Family.Id;
            var oldIds = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi => fi.Symbol?.Family?.Id == famId)
                .Select(fi => fi.Id)
                .ToList();

            if (!oldIds.Any()) return;
            using (var tx = new Transaction(doc, "Remove Old Pile Checkers"))
            {
                tx.Start();
                doc.Delete(oldIds);
                tx.Commit();
            }
        }

        FamilySymbol FindCheckerSymbol(Document doc)
        {
            // Fast path if we know what we just loaded this session
            if (!string.IsNullOrWhiteSpace(_loadedFamilyInternalName))
            {
                var famFast = new FilteredElementCollector(doc)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .FirstOrDefault(f =>
                        f.Name.Equals(_loadedFamilyInternalName, StringComparison.OrdinalIgnoreCase));

                if (famFast != null)
                {
                    var sid = famFast.GetFamilySymbolIds().FirstOrDefault();
                    if (sid != ElementId.InvalidElementId)
                        return doc.GetElement(sid) as FamilySymbol;
                }
            }

            // Generic search by family name prefix, covers 23/24/25 variants
            const string prefix = "HDR_General_PileChecker";
            var fam = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .FirstOrDefault(f =>
                    f.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));

            if (fam != null)
            {
                var sid = fam.GetFamilySymbolIds().FirstOrDefault();
                if (sid != ElementId.InvalidElementId)
                    return doc.GetElement(sid) as FamilySymbol;
            }

            return null;
        }

        Family LoadEmbeddedFamilyForVersion(Document doc, ExternalCommandData data)
        {
            // Determine which embedded resource to look for based on version
            string ver = data.Application.Application.VersionNumber; // "2023", "2024", "2025", etc.
            string suffix = ver switch
            {
                "2023" => "23",
                "2024" => "24",
                "2025" => "25",
                _ => "25" // default to latest asset if running on a newer version
            };

            // Expected full name
            string expected = $"HDR_EMEA.Resources.HDR_General_PileChecker{suffix}.rfa";

            // Resolve to an actual manifest resource (handles namespace drift)
            var asm = Assembly.GetExecutingAssembly();
            string resourceName =
                asm.GetManifestResourceNames()
                   .FirstOrDefault(n => n.Equals(expected, StringComparison.OrdinalIgnoreCase))
                ?? asm.GetManifestResourceNames()
                      .FirstOrDefault(n => n.EndsWith($".HDR_General_PileChecker{suffix}.rfa", StringComparison.OrdinalIgnoreCase));

            if (string.IsNullOrEmpty(resourceName))
                return null;

            // Write to temp
            string tempFile = Path.Combine(
                Path.GetTempPath(),
                $"HDR_General_PileChecker{suffix}.rfa");
            try { if (File.Exists(tempFile)) File.Delete(tempFile); } catch { /* ignore */ }

            using (var rs = asm.GetManifestResourceStream(resourceName))
            {
                if (rs == null) return null;
                using (var fs = new FileStream(tempFile, FileMode.Create, FileAccess.Write))
                    rs.CopyTo(fs);
            }

            // Load with overwrite options
            using (var tx = new Transaction(doc, "Load Embedded PileChecker"))
            {
                tx.Start();
                var options = new AlwaysOverwriteFamilyLoadOptions();
                bool ok = doc.LoadFamily(tempFile, options, out Family fam);
                tx.Commit();
                return ok ? fam : null;
            }
        }

        Result PlaceAndColorCheckers(
            UIDocument uiDoc, Document doc, FamilySymbol checkerSym)
        {
            // collect only foundations whose Foundation_Type = "Circular Pile"
            var piles = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .Where(fi =>
                {
                    // check Foundation_Type
                    Parameter p = fi.Symbol.LookupParameter("Foundation_Type")
                                  ?? fi.LookupParameter("Foundation_Type");
                    return p != null
                           && string.Equals(p.AsString(),
                                            "Circular Pile",
                                            StringComparison.OrdinalIgnoreCase);
                })
                .ToList();

            if (!piles.Any())
            {
                TaskDialog.Show("Pile Proximity",
                                "No Circular Pile foundations found.");
                return Result.Cancelled;
            }

            var placed = new List<(ElementId id, XYZ loc, double radius)>();
            using (var tx = new Transaction(doc, "Add Pile Proximity Checkers"))
            {
                tx.Start();
                foreach (var fi in piles)
                {
                    if (!(fi.Location is LocationPoint lp)) continue;

                    // try instance first, then type
                    Parameter diaParam = fi.LookupParameter("Pile_Size")
                                         ?? fi.Symbol.LookupParameter("Pile_Size");
                    double dia = diaParam?.AsDouble() ?? 0;
                    if (dia <= 0) continue;

                    double radius = dia * 1.5;
                    var lvl = doc.GetElement(fi.LevelId) as Level;

                    var inst = doc.Create.NewFamilyInstance(
                        lp.Point, checkerSym, lvl,
                        StructuralType.NonStructural);
                    inst.LookupParameter("Pile_Size")?.Set(dia);

                    placed.Add((inst.Id, lp.Point, radius));
                }
                tx.Commit();
            }

            if (!placed.Any())
            {
                TaskDialog.Show("Pile Proximity", "No checkers placed.");
                return Result.Failed;
            }

            // prepare overrides
            ElementId solidFillId = new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .First(fp => fp.GetFillPattern().IsSolidFill)
                .Id;

            var clashOV = new OverrideGraphicSettings()
                .SetSurfaceForegroundPatternId(solidFillId)
                .SetSurfaceBackgroundPatternId(solidFillId)
                .SetSurfaceForegroundPatternColor(new Color(255, 0, 0))
                .SetSurfaceBackgroundPatternColor(new Color(255, 0, 0))
                .SetSurfaceTransparency(50);

            var okOV = new OverrideGraphicSettings()
                .SetSurfaceForegroundPatternId(solidFillId)
                .SetSurfaceBackgroundPatternId(solidFillId)
                .SetSurfaceForegroundPatternColor(new Color(0, 128, 0))
                .SetSurfaceBackgroundPatternColor(new Color(0, 128, 0))
                .SetSurfaceTransparency(50);

            double tol = UnitUtils.ConvertToInternalUnits(
                             1.0, UnitTypeId.Millimeters);
            var view = uiDoc.ActiveView;
            using (var tx = new Transaction(doc, "Highlight Pile Proximity"))
            {
                tx.Start();
                foreach (var (id, loc, rad) in placed)
                {
                    bool overlaps = placed
                        .Where(x => x.id != id)
                        .Any(x => x.loc.DistanceTo(loc)
                                  < (rad + x.radius) - tol);

                    view.SetElementOverrides(
                        id, overlaps ? clashOV : okOV);
                }
                tx.Commit();
            }

            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "PileProximityAdd",
                "Pile Proximity",
                "HDR_EMEA.Tools_Structural.PileProximityAdd",
                Properties.Resources.PileProximityAdd_32,
                Properties.Resources.PileProximityAdd_16,
                "Place pile checker families around Circular Pile foundations"
            );
        }

        // Overwrite family + parameters when reloading
        private class AlwaysOverwriteFamilyLoadOptions : IFamilyLoadOptions
        {
            public bool OnFamilyFound(bool familyInUse, out bool overwriteParameterValues)
            {
                overwriteParameterValues = true;
                return true; // overwrite family
            }

            public bool OnSharedFamilyFound(Family sharedFamily, bool familyInUse,
                out FamilySource source, out bool overwriteParameterValues)
            {
                source = FamilySource.Family; // prefer the incoming family
                overwriteParameterValues = true;
                return true;
            }
        }
    }
}
