using HDR_EMEA;
using System.Windows.Forms;

public class FoundationCommandHandler : IExternalEventHandler
{
    public enum CommandType
    {
        None,
        FullMarking, // Full pilecap/pile marking
        UpdateMarks, // Only unmarked or duplicate mark foundations
        MarkBySelection, // Open FounFamType, apply marks
        MarkByGrids  // Open FounFamType, grid-based numbering
    }

    public CommandType CurrentCommand { get; set; } = CommandType.None;

    public void Execute(UIApplication app)
    {
        UIDocument uidoc = app.ActiveUIDocument;
        Document doc = uidoc.Document;

        DateTime startTime = DateTime.Now;
        bool success = false;
        string commandLabel = CurrentCommand.ToString();

        try
        {
            switch (CurrentCommand)
            {
                case CommandType.FullMarking:
                    FoundationCommands.RunFullMarking(app);
                    break;
                case CommandType.UpdateMarks:
                    FoundationCommands.RunUpdateMarks(app);
                    break;
                case CommandType.MarkBySelection:
                    FoundationCommands.RunMarkBySelection(app);
                    break;
                case CommandType.MarkByGrids:
                    FoundationCommands.RunMarkByGrids(app);
                    break;
                default:
                    TaskDialog.Show("Error", "No command selected.");
                    return;
            }
            success = true;
        }
        catch (Exception ex)
        {
            TaskDialog.Show("Foundation Command Error", ex.Message);
        }
        finally
        {
            Tracker.LogCommandUsage($"FounRef_{commandLabel}", startTime, success);
        }
    }
    public string GetName() => "FoundationCommandHandler";

    public static class FoundationCommands
    {
        public static void RunFullMarking(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Collect all structural foundation instances
            var foundations = ElementCollectorUtils.GetAllStructuralFoundations(doc);

            List<FamilyInstance> allFoundations = foundations.OfType<FamilyInstance>().ToList();

            // Filter pilecaps (top-level only)
            var pileCaps = allFoundations
                .Where(f => f.SuperComponent == null &&
                            f.Symbol.FamilyName.ToLower().Contains("pilecap"))
                .ToList();

            // Sort pilecaps by X then Y (left to right, top down)
            pileCaps.Sort((a, b) =>
            {
                XYZ aLoc = ElementCollectorUtils.GetLocation(a);
                XYZ bLoc = ElementCollectorUtils.GetLocation(b);
                int cmpY = bLoc.Y.CompareTo(aLoc.Y); 
                return cmpY != 0 ? cmpY : aLoc.X.CompareTo(bLoc.X);
            });

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Set PileCap and Pile Marks");

                int capIndex = 1;

                foreach (var pileCap in pileCaps)
                {
                    string capMark = $"F{capIndex.ToString("D2")}";
                    Utils.SetParameter(pileCap, "Mark", capMark);

                    // Find nested piles under this pilecap
                    var nestedPiles = allFoundations
                        .Where(f => f.SuperComponent?.Id == pileCap.Id &&
                                    f.Symbol.FamilyName.ToLower().Contains("pile"))
                        .ToList();

                    // Sort nested piles spatially
                    nestedPiles.Sort((a, b) =>
                    {
                        XYZ aLoc = ElementCollectorUtils.GetLocation(a);
                        XYZ bLoc = ElementCollectorUtils.GetLocation(b);
                        int cmpY = bLoc.Y.CompareTo(aLoc.Y); 
                        return cmpY != 0 ? cmpY : aLoc.X.CompareTo(bLoc.X);
                    });

                    // Assign pile marks 
                    for (int i = 0; i < nestedPiles.Count; i++)
                    {
                        string pileMark = $"{capMark}-P{(i + 1).ToString("D2")}";
                        Utils.SetParameter(nestedPiles[i], "Mark", pileMark);
                    }

                    capIndex++;
                }


                // Get all structural foundation elements
                var allSlabs = ElementCollectorUtils.GetAllStructuralFoundations(doc)
                    .Where(f =>
                    {
                        var type = doc.GetElement(f.GetTypeId()) as ElementType;
                        return type != null && type.Name.ToLower().Contains("corebase");
                    })
                    .OrderByDescending(f => ElementCollectorUtils.GetLocation(f).Y) 
                    .ThenBy(f => ElementCollectorUtils.GetLocation(f).X)
                    .ToList();

                int coreIndex = 1;
                foreach (var core in allSlabs)
                {
                    string coreMark = $"C{coreIndex.ToString("D2")}";
                    Utils.SetParameter(core, "Mark", coreMark);
                    coreIndex++;
                }

                // Build slab-pile map based on spatial overlap
                var slabPileMap = new Dictionary<ElementId, List<FamilyInstance>>();
                var view = doc.ActiveView;

                foreach (var slab in allSlabs)
                {
                    BoundingBoxXYZ bbox = slab.get_BoundingBox(view);
                    if (bbox == null) continue;

                    List<FamilyInstance> matchedPiles = new();

                    foreach (var pile in allFoundations.Where(f => f.Symbol.FamilyName.ToLower().Contains("pile")))
                    {
                        if (pile.Location is not LocationPoint pileLoc) continue;

                        XYZ pt = pileLoc.Point;

                        bool withinXY = (bbox.Min.X <= pt.X && pt.X <= bbox.Max.X) &&
                                        (bbox.Min.Y <= pt.Y && pt.Y <= bbox.Max.Y);
                        bool belowOrTouchingZ = pt.Z <= bbox.Max.Z + 0.01;

                        if (withinXY && belowOrTouchingZ)
                        {
                            matchedPiles.Add(pile);
                        }
                    }

                    if (matchedPiles.Count > 0)
                    {
                        slabPileMap[slab.Id] = matchedPiles;
                    }
                }

                // Assign marks to matched piles based on slab mark
                foreach (var kvp in slabPileMap)
                {
                    Element slab = doc.GetElement(kvp.Key);
                    string coreMark = slab.LookupParameter("Mark")?.AsString();
                    if (string.IsNullOrWhiteSpace(coreMark)) continue;

                    var piles = kvp.Value;
                    piles.Sort((a, b) =>
                    {
                        XYZ aLoc = ElementCollectorUtils.GetLocation(a);
                        XYZ bLoc = ElementCollectorUtils.GetLocation(b);
                        int cmpY = bLoc.Y.CompareTo(aLoc.Y);
                        return cmpY != 0 ? cmpY : aLoc.X.CompareTo(bLoc.X);
                    });

                    for (int i = 0; i < piles.Count; i++)
                    {
                        string pileMark = $"{coreMark}-P{(i + 1).ToString("D2")}";
                        Utils.SetParameter(piles[i], "Mark", pileMark);
                    }
                }


                t.Commit();
            }
        }

        public static void RunUpdateMarks(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Collect foundations
            var allFoundationElems = ElementCollectorUtils.GetAllStructuralFoundations(doc);
            var allFoundationsFI = allFoundationElems.OfType<FamilyInstance>().ToList();

            // Identify piles and pilecaps (FamilyInstances)
            var pileCaps = allFoundationsFI
                .Where(f => f.SuperComponent == null &&
                            f.Symbol.FamilyName.ToLower().Contains("pilecap"))
                .OrderByDescending(f => ElementCollectorUtils.GetLocation(f).Y)
                .ThenBy(f => ElementCollectorUtils.GetLocation(f).X)
                .ToList();

            var piles = allFoundationsFI
                .Where(f => f.Symbol.FamilyName.ToLower().Contains("pile"))
                .ToList();

            // Identify CoreBase slabs (can be system or component)
            var coreSlabs = allFoundationElems
                .Where(e =>
                {
                    var type = doc.GetElement(e.GetTypeId()) as ElementType;
                    return type != null && type.Name.ToLower().Contains("corebase");
                })
                .OrderByDescending(f => ElementCollectorUtils.GetLocation(f).Y)
                .ThenBy(f => ElementCollectorUtils.GetLocation(f).X)
                .ToList();

            // Build mark counts across ALL structural foundation elements (for uniqueness check)
            Dictionary<string, int> markCounts = new(StringComparer.InvariantCultureIgnoreCase);
            foreach (var e in allFoundationElems)
            {
                string m = e.LookupParameter("Mark")?.AsString();
                if (!string.IsNullOrWhiteSpace(m))
                {
                    if (!markCounts.ContainsKey(m)) markCounts[m] = 0;
                    markCounts[m]++;
                }
            }

            // Helper: predicate => needs assignment (empty or duplicate)
            bool NeedsMark(Element e)
            {
                string m = e.LookupParameter("Mark")?.AsString();
                if (string.IsNullOrWhiteSpace(m)) return true;
                return markCounts.TryGetValue(m, out int count) && count > 1;
            }

            // Track used marks and fast uniqueness checks
            HashSet<string> usedMarks = new HashSet<string>(
                markCounts.Keys, StringComparer.InvariantCultureIgnoreCase);

            // Helpers to generate next unique codes
            string NextUniqueF(ref int i)
            {
                string code;
                do { code = $"F{(i++).ToString("D2")}"; } while (usedMarks.Contains(code));
                usedMarks.Add(code);
                return code;
            }
            string NextUniqueC(ref int i)
            {
                string code;
                do { code = $"C{(i++).ToString("D2")}"; } while (usedMarks.Contains(code));
                usedMarks.Add(code);
                return code;
            }
            string NextUniquePile(string baseMark, ref int i)
            {
                string code;
                do { code = $"{baseMark}-P{(i++).ToString("D2")}"; } while (usedMarks.Contains(code));
                usedMarks.Add(code);
                return code;
            }

            using (Transaction t = new Transaction(doc, "Fix Unmarked/Duplicate Foundation Marks"))
            {
                t.Start();

                // ========= 1) PILECAPS (only empty or duplicates) =========
                int capIndex = 1;

                foreach (var cap in pileCaps)
                {
                    // If cap has a unique existing mark, keep it
                    if (!NeedsMark(cap))
                        continue;

                    // Generate next unique Fxx
                    string capMark = NextUniqueF(ref capIndex);
                    Utils.SetParameter(cap, "Mark", capMark);

                    // Nested piles under this cap: only fix empty/duplicates
                    var nested = allFoundationsFI
                        .Where(f => f.SuperComponent?.Id == cap.Id &&
                                    f.Symbol.FamilyName.ToLower().Contains("pile"))
                        .ToList();

                    // Spatial sort for consistent numbering
                    nested.Sort((a, b) =>
                    {
                        XYZ aLoc = ElementCollectorUtils.GetLocation(a);
                        XYZ bLoc = ElementCollectorUtils.GetLocation(b);
                        int cmpY = bLoc.Y.CompareTo(aLoc.Y);
                        return cmpY != 0 ? cmpY : aLoc.X.CompareTo(bLoc.X);
                    });

                    int pileIdx = 1;
                    foreach (var p in nested)
                    {
                        if (!NeedsMark(p)) continue; // keep unique existing

                        string pMark = NextUniquePile(capMark, ref pileIdx);
                        Utils.SetParameter(p, "Mark", pMark);
                    }
                }

                // ========= 2) COREBASE SLABS (only empty or duplicates) =========
                int coreIndex = 1;

                // First, ensure CoreBase slabs themselves have unique Cxx where needed
                foreach (var slab in coreSlabs)
                {
                    if (!NeedsMark(slab)) continue;

                    string cMark = NextUniqueC(ref coreIndex);
                    Utils.SetParameter(slab, "Mark", cMark);
                    // also patch our counts/hash to reflect the new assignment
                }

                // Rebuild markCounts/usedMarks after changes so far (keeps things consistent)
                markCounts.Clear();
                usedMarks.Clear();
                foreach (var e in allFoundationElems)
                {
                    string m = e.LookupParameter("Mark")?.AsString();
                    if (!string.IsNullOrWhiteSpace(m))
                    {
                        if (!markCounts.ContainsKey(m)) markCounts[m] = 0;
                        markCounts[m]++;
                    }
                }
                foreach (var k in markCounts.Keys) usedMarks.Add(k);

                // ========= 3) PILES UNDER COREBASE (spatial), only empty/duplicates =========
                var view = doc.ActiveView;

                foreach (var slab in coreSlabs)
                {
                    string coreMark = slab.LookupParameter("Mark")?.AsString();
                    if (string.IsNullOrWhiteSpace(coreMark)) continue; // should be set above

                    BoundingBoxXYZ bbox = slab.get_BoundingBox(view);
                    if (bbox == null) continue;

                    // Collect piles under this slab
                    var pilesUnder = new List<FamilyInstance>();
                    foreach (var pile in piles)
                    {
                        // skip piles that already have a unique mark
                        if (!NeedsMark(pile)) continue;

                        if (pile.Location is not LocationPoint lp) continue;
                        XYZ pt = lp.Point;

                        bool withinXY = (bbox.Min.X <= pt.X && pt.X <= bbox.Max.X) &&
                                        (bbox.Min.Y <= pt.Y && pt.Y <= bbox.Max.Y);
                        bool belowOrTouchingZ = pt.Z <= bbox.Max.Z + 0.01;

                        if (withinXY && belowOrTouchingZ)
                            pilesUnder.Add(pile);
                    }

                    if (!pilesUnder.Any()) continue;

                    // Sort spatially for consistent numbering
                    pilesUnder.Sort((a, b) =>
                    {
                        XYZ aLoc = ElementCollectorUtils.GetLocation(a);
                        XYZ bLoc = ElementCollectorUtils.GetLocation(b);
                        int cmpY = bLoc.Y.CompareTo(aLoc.Y);
                        return cmpY != 0 ? cmpY : aLoc.X.CompareTo(bLoc.X);
                    });

                    int pileIdx = 1;
                    foreach (var p in pilesUnder)
                    {
                        // Only assign if still empty/duplicate (in case something changed above)
                        if (!NeedsMark(p)) continue;

                        string pMark = NextUniquePile(coreMark, ref pileIdx);
                        Utils.SetParameter(p, "Mark", pMark);
                    }
                }

                t.Commit();
            }
        }


        public static void RunMarkBySelection(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get all Structural Foundations FamilySymbols in the document
            List<FamilySymbol> foundationSymbols = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .Cast<FamilySymbol>()
                .ToList();

            // Show the family selection form
            HDR_EMEA.Forms.FounFamType formSFF = new HDR_EMEA.Forms.FounFamType(foundationSymbols)
            {
                Width = 800,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            bool? dialogResult = formSFF.ShowDialog();
            if (dialogResult != true || !formSFF.GetSelectedSymbols().Any())
                return;

            var selectedSymbols = formSFF.GetSelectedSymbols();
            var selectionFilter = new FoundationTypeSelectionFilter(selectedSymbols);

            // Rectangle selection first
            IList<Element> rectSelectedElements = uidoc.Selection.PickElementsByRectangle(selectionFilter);
            HashSet<ElementId> selectedIds = rectSelectedElements.Select(e => e.Id).ToHashSet();
            uidoc.Selection.SetElementIds(selectedIds.ToList());

            // Prompt user
            TaskDialog.Show("Selection", "Initial rectangle selection complete.\nNow use window selection or individual clicks to modify selection.\nClick 'Finish' when done.");

            // One-time multiple selection with Finish/Cancel
            try
            {
                IList<Reference> refs = uidoc.Selection.PickObjects(
                    ObjectType.Element,
                    selectionFilter,
                    "Select/deselect foundations. Click Finish to complete."
                );

                foreach (Reference r in refs)
                {
                    if (selectedIds.Contains(r.ElementId))
                        selectedIds.Remove(r.ElementId);
                    else
                        selectedIds.Add(r.ElementId);
                }

                uidoc.Selection.SetElementIds(selectedIds.ToList());
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                TaskDialog.Show("Cancelled", "Selection was cancelled. No changes applied.");
                return;
            }

            List<Element> selectedElements = selectedIds
                .Select(id => doc.GetElement(id))
                .Where(e => e != null)
                .ToList();

            // Show form to set mark options
            HDR_EMEA.Forms.FounSetRef formPSR = new HDR_EMEA.Forms.FounSetRef()
            {
                Width = 500,
                Height = 280,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            if (formPSR.ShowDialog() != true)
                return;

            string prefix = formPSR.Prefix;
            string suffix = formPSR.Suffix;
            int startNum = 1;
            int.TryParse(formPSR.StartNumber, out startNum);

            HashSet<string> existingMarks = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .WhereElementIsNotElementType()
                .Select(e => e.LookupParameter("Mark")?.AsString())
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .ToHashSet();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Set Foundation References for Selected");

                foreach (Element el in selectedElements)
                {
                    string proposedMark;
                    do
                    {
                        proposedMark = $"{prefix}{startNum:D2}{suffix}";
                        startNum++;
                    } while (existingMarks.Contains(proposedMark));
                    existingMarks.Add(proposedMark);

                    Parameter markParam = el.LookupParameter("Mark");
                    if (markParam != null && !markParam.HasValue)
                    {
                        markParam.Set(proposedMark);
                    }
                }

                t.Commit();
            }
        }


        public static void RunMarkByGrids(UIApplication app)
        {
            UIDocument uidoc = app.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Your code goes here
            // Get all Grids in the Document
            FilteredElementCollector gridCollector = new FilteredElementCollector(doc);
            ICollection<Element> grids = gridCollector.OfClass(typeof(Grid)).ToElements();

            List<Grid> gridList = grids.Cast<Grid>().ToList();

            // Hashset to store unique intersections
            HashSet<string> uniqueIntersections = new HashSet<string>();
            List<Tuple<XYZ, string, string>> intersectionPoints = new List<Tuple<XYZ, string, string>>();

            // Loop through each grid and find intersections
            for (int i = 0; i < gridList.Count; i++)
            {
                for (int j = 0; j < gridList.Count; j++)
                {
                    XYZ gridIntersection = GetIntersection(gridList[i], gridList[j]);
                    if (gridIntersection != null)
                    {
                        string grid1Name = gridList[i].Name;
                        string grid2Name = gridList[j].Name;
                        string intersectionKey = $"{gridIntersection.X:F6},{gridIntersection.Y:F6}";

                        //Add intersection if it is Unique
                        if (uniqueIntersections.Add(intersectionKey))
                        {
                            intersectionPoints.Add(new Tuple<XYZ, string, string>(gridIntersection, grid1Name, grid2Name));
                            //TaskDialog.Show("Intersection", $"Intersection between Grid {grid1Name} and Grid {grid2Name} at ({gridIntersection.X}, {gridIntersection.Y}");
                        }

                    }

                }
            }

            // Get all placed FamilyInstances of Structural Foundations
            var placedInstances = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            // Distinct symbols by Id (avoids needing a comparer type)
            List<FamilySymbol> allSymbols = placedInstances
                .Select(f => f.Symbol)
                .Where(sym => sym != null)
                .GroupBy(sym => sym.Id.IntegerValue)
                .Select(g => g.First())
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

            // Fallback: if nothing matches, just show all to avoid an empty dialog
            if (!filteredSymbols.Any())
                filteredSymbols = allSymbols;

            // open the form Pass the family names
            HDR_EMEA.Forms.FounFamType formSFF = new HDR_EMEA.Forms.FounFamType(filteredSymbols, allSymbols)
            {
                Width = 800,
                Height = 450,
                WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                Topmost = true,
            };

            bool? dialogResult = formSFF.ShowDialog();
            if (dialogResult != true || !formSFF.GetSelectedSymbols().Any())
                return;

            List<FamilySymbol> selectedSymbols = formSFF.GetSelectedSymbols();

            // Create a lookup for selected families and types
            var selectedFamilyKeys = new HashSet<(string FamilyName, string TypeName)>(
                selectedSymbols.Select(s => (s.Family.Name, s.Name))
            );

            // Filter foundations using Family Name and Type Name
            FilteredElementCollector foundationCollector = new FilteredElementCollector(doc);
            ICollection<Element> foundations = foundationCollector
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .WhereElementIsNotElementType()
                .Where(f =>
                {
                    if (f is FamilyInstance instance)
                    {
                        string famName = instance.Symbol.Family.Name;
                        string typeName = instance.Symbol.Name;
                        return selectedFamilyKeys.Contains((famName, typeName));
                    }
                    return false;
                })
                .OrderByDescending(f => GetFoundationLocation(f).Y)
                .ThenBy(f => GetFoundationLocation(f).X)
                .ToList();

            //FilteredElementCollector foundationCollector = new FilteredElementCollector(doc);
            //ICollection<Element> foundations = foundationCollector.OfCategory(BuiltInCategory.OST_StructuralFoundation)
            //    .WhereElementIsNotElementType()
            //    .ToElements()
            //    .Where(f => f.Name.Equals(selectedFounFamily, StringComparison.InvariantCultureIgnoreCase))
            //    .OrderByDescending(f => GetFoundationLocation(f).Y)
            //    .ThenBy(f => GetFoundationLocation(f).X)
            //    .ToList();


            // use PickObjects to select foundations
            //Selection sel = uiapp.ActiveUIDocument.Selection;
            //IList<Reference> selectedReference;
            //try
            //{
            //    selectedReference = sel.PickObjects(ObjectType.Element, new FoundationSelectionFilter(foundationFilter), "Select Foundations");
            //}
            //catch (OperationCanceledException)
            //{
            //    return Result.Cancelled;
            //}

            //List<Element> selectedFoundations = selectedReference.Select(r => doc.GetElement(r)).ToList();
            //

            // Dictionary to store groups of foundations by main foundation Id
            Dictionary<ElementId, List<Element>> foundationGroups = new Dictionary<ElementId, List<Element>>();

            // Group foundations based on their main foundation Id
            foreach (Element f in foundations)
            {
                FamilyInstance familyInstance = f as FamilyInstance;
                ElementId mainFoundationId = familyInstance?.SuperComponent?.Id ?? f.Id;

                if (!foundationGroups.ContainsKey(mainFoundationId))
                {
                    foundationGroups[mainFoundationId] = new List<Element>();
                }
                foundationGroups[mainFoundationId].Add(f);
            }

            // Set Mark Parameter for each foundation based on nearest grid intersection
            HashSet<string> usedMarks = new HashSet<string>();
            Dictionary<string, int> gridPairCounters = new Dictionary<string, int>();

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Set Foundation Marks to nearest Grid Intersection");

                // Collect all existing Mark parameter values in the document
                HashSet<string> existingMarks = new HashSet<string>();
                FilteredElementCollector founMarkCollector = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .WhereElementIsNotElementType();

                foreach (Element elem in founMarkCollector)
                {
                    Parameter markParam = elem.LookupParameter("Mark");
                    if (markParam != null && !string.IsNullOrEmpty(markParam.AsString()))
                    {
                        existingMarks.Add(markParam.AsString());
                    }
                }

                foreach (var group in foundationGroups)
                {
                    List<Element> foundationGroup = group.Value;

                    foreach (Element foundation in foundationGroup)
                    {
                        //Check if this foundation has a "mark" parameter that starts with "P"
                        Parameter markParm = foundation.LookupParameter("Mark");
                        if (markParm != null && !string.IsNullOrEmpty(markParm.AsString()) && markParm.AsString().StartsWith("P"))
                        {
                            continue;
                        }

                        XYZ foundationLocation = GetMainFoundationLocation(doc, foundation);
                        if (foundationLocation != null)
                        {
                            double minDistance = double.MaxValue;
                            Tuple<XYZ, string, string> nearestIntersection = null;

                            foreach (var intersection in intersectionPoints)
                            {
                                double distance = foundationLocation.DistanceTo(intersection.Item1);
                                if (distance < minDistance)
                                {
                                    minDistance = distance;
                                    nearestIntersection = intersection;
                                }
                            }
                            if (nearestIntersection != null)
                            {
                                string grid1Name = nearestIntersection.Item2;
                                string grid2Name = nearestIntersection.Item3;
                                string gridPair = $"{grid1Name}/{grid2Name}";

                                // Initialize counter for new grid pairs
                                if (!gridPairCounters.ContainsKey(gridPair))
                                {
                                    gridPairCounters[gridPair] = 1;
                                }

                                // Generate the Mark and check for uniqueness
                                string mark;
                                do
                                {
                                    mark = $"P-[{grid1Name}/{grid2Name}] {gridPairCounters[gridPair]:D2}";
                                    gridPairCounters[gridPair]++;
                                }
                                while (existingMarks.Contains(mark) || !usedMarks.Add(mark)); // Avoid Duplicates

                                // Set the Mark Parameter
                                if (markParm != null)
                                {
                                    markParm.Set(mark);
                                }
                            }
                        }
                    }
                }

                t.Commit();
            }
        }
    }


    private static XYZ GetIntersection(Grid grid1, Grid grid2)
    {
        Curve curve1 = grid1.Curve;
        Curve curve2 = grid2.Curve;

        IntersectionResultArray resultArray;
        SetComparisonResult result = curve1.Intersect(curve2, out resultArray);

        if (result == SetComparisonResult.Overlap)
        {
            if (resultArray != null && resultArray.Size > 0)
            {
                return resultArray.get_Item(0).XYZPoint;
            }
        }
        return null;
    }

    private static XYZ GetFoundationLocation(Element foundation)
    {
        LocationPoint locationPoint = foundation.Location as LocationPoint;
        return locationPoint?.Point;
    }

    private static XYZ GetMainFoundationLocation(Document doc, Element foundation)
    {
        FamilyInstance instance = foundation as FamilyInstance;
        if (instance != null && instance.SuperComponent != null)
        {
            Element superComponent = doc.GetElement(instance.SuperComponent.Id);
            if (superComponent != null)
            {
                return GetMainFoundationLocation(doc, superComponent);
            }
        }

        return GetFoundationLocation(foundation);
    }

    internal class FoundationSelectionFilter : ISelectionFilter
    {
        private string _foundationFilter;

        public FoundationSelectionFilter(string foundationFilter)
        {
            _foundationFilter = foundationFilter;
        }

        public bool AllowElement(Element elem)
        {
            return elem.Category != null &&
                elem.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralFoundation &&
                elem.Name.Equals(_foundationFilter, StringComparison.InvariantCultureIgnoreCase);
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            return false;
        }
    }
    internal class FoundationTypeSelectionFilter : ISelectionFilter
    {
        private readonly HashSet<(string FamilyName, string TypeName)> _selectedPairs;

        public FoundationTypeSelectionFilter(IEnumerable<FamilySymbol> symbols)
        {
            _selectedPairs = symbols
                .Select(s => (s.Family.Name, s.Name))
                .ToHashSet();
        }

        public bool AllowElement(Element elem)
        {
            if (elem is FamilyInstance instance)
            {
                string famName = instance.Symbol.Family.Name;
                string typeName = instance.Symbol.Name;
                return _selectedPairs.Contains((famName, typeName));
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position) => true;
    }


    public static string GetMethod()
    {
        var method = MethodBase.GetCurrentMethod().DeclaringType?.FullName;
        return method;
    }
}


