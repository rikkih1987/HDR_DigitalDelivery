using System.Windows.Controls;

namespace HDR_EMEA.Common
{
    internal static class ElementCollectorUtils
    {
        // Generic method to collect elements by category
        internal static List<Element> GetElementsByCategory(Document doc, BuiltInCategory category)
        {
            return new FilteredElementCollector(doc)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .ToList();
        }

        // Generic method to collect elements by class
        internal static List<T> GetElementsByClass<T>(Document doc) where T : Element
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(T))
                .WhereElementIsNotElementType()
                .Cast<T>()
                .ToList();
        }
    
        internal static List<Element> GetAllStructuralFoundations(Document doc) =>
            GetElementsByCategory(doc, BuiltInCategory.OST_StructuralFoundation);

        internal static List<Element> GetAllStructuralColumns(Document doc) =>
            GetElementsByCategory(doc, BuiltInCategory.OST_StructuralColumns);

        internal static List<Element> GetAllStructuralFraming(Document doc) =>
            GetElementsByCategory(doc, BuiltInCategory.OST_StructuralFraming);

        internal static List<Element> GetAllFloors(Document doc) =>
            GetElementsByCategory(doc, BuiltInCategory.OST_Floors);

        internal static List<Element> GetAllWalls(Document doc) =>
            GetElementsByCategory(doc, BuiltInCategory.OST_Walls);

        // Utility: Get all used BuiltInCategories
        internal static List<BuiltInCategory> GetUsedBuiltInCategories(Document doc)
        {
            return new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Select(e => e.Category?.Id)
                .Where(id => id != null && id != ElementId.InvalidElementId)
                .Distinct()
                .Select(id => (BuiltInCategory)id.IntegerValue)
                .Where(bic => Enum.IsDefined(typeof(BuiltInCategory), bic))
                .OrderBy(bic => bic.ToString())
                .ToList();
        }


        internal static ElementId GetSolidFillPatternId(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FillPatternElement))
                .Cast<FillPatternElement>()
                .First(fp => fp.GetFillPattern().IsSolidFill)
                .Id;
        }

        internal static FillPatternElement GetFillPatternByName(Document doc, string name)
        {
            FillPatternElement curFPE = null;

            curFPE = FillPatternElement.GetFillPatternElementByName(doc, FillPatternTarget.Drafting, name);

            return curFPE;
        }

        internal static XYZ GetLocation(Element element)
        {
            if (element.Location is LocationPoint locPoint)
                return locPoint.Point;
            return XYZ.Zero;
        }


        internal static int GetGroupInstanceCount(Document doc) =>
        new FilteredElementCollector(doc).OfClass(typeof(Group)).Count();
        internal static int GetGroupTypeCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(GroupType)).Count();
        internal static int GetUnusedGroupTypeCount(Document doc)
        {
            try
            {
                var groupTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(GroupType))
                    .Cast<GroupType>()
                    .ToList();

                var groupInstance = new FilteredElementCollector(doc)
                    .OfClass(typeof(Group))
                    .Cast<Group>()
                    .ToList();

                int unusedCount = groupTypes.Count(gt => !groupInstance.Any(g => g.GroupType.Id == gt.Id));

                return unusedCount;
            }
            catch
            {
                return 0;
            }
        }

        internal static int GetDesignOptionCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(DesignOption)).Count();
        internal static int GetWorksetCount(Document doc) =>
            new FilteredWorksetCollector(doc).Where(w => w.Kind == WorksetKind.UserWorkset).Count();
        internal static int GetCADImportCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(ImportInstance))
            .Count();
        internal static int GetCADLinkCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Count();
        internal static int GetImageCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(ImageType)).Count();


        internal static double GetTotalRoomArea(Document doc)
        {
            const double SqFtToSqMeters = 0.092903;

            double totalSqFt = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
                .WhereElementIsNotElementType()
                .Cast<Room>()
                .Where(r => r.Area > 0)
                .Sum(r => r.Area);

            double totalSqM = totalSqFt * SqFtToSqMeters;
            return Math.Round(totalSqM, 2);
        }
        internal static int GetRoomCount(Document doc) =>
            new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .Count();
        internal static int GetUnplacedRoomCount(Document doc) =>
            new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .Cast<Room>().Count(r => r.Area == 0 || r.Location == null);
        internal static int GetUnenclosedRoomCount(Document doc) =>
            new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms)
            .WhereElementIsNotElementType()
            .Cast<Room>()
            .Count(r => r.Area <= 0);


        internal static int GetLoadedFamilyCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(Family)).Cast<Family>().Count();
        internal static int GetInplaceFamilyCount(Document doc)
        {
            var usedFamilyIds = new HashSet<ElementId>();

            var instances = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>();

            foreach ( var instance in instances)
            {
                var family = instance.Symbol?.Family;
                if (family != null)
                {
                    usedFamilyIds.Add(family.Id);
                }
            }
            return usedFamilyIds.Count();
        }
        internal static double GetUnusedFamilyPercentage(Document doc)
        {
            int total = GetLoadedFamilyCount(doc);
            int used = GetInplaceFamilyCount (doc);

            if (total == 0) return 0;

            double unused = total - used;
            double percent = (unused / (double)total) * 100;

            return Math.Round(percent, 2);
        }

        internal static int GetWarningCount(Document doc) =>
        doc.GetWarnings().Count;


        internal static List<View> GetFilteredWorkingViews(Document doc)
        {
            ViewType[] targetTypes = new[]
            {
                ViewType.FloorPlan,
                ViewType.CeilingPlan,
                ViewType.Section,
                ViewType.Elevation,
                ViewType.ThreeD,
                ViewType.DraftingView,
                ViewType.EngineeringPlan,
                ViewType.AreaPlan,
                ViewType.Detail
            };

            return new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => !v.IsTemplate && targetTypes.Contains(v.ViewType))
                .ToList();
        }
        internal static int GetViewCount(Document doc) =>
            GetFilteredWorkingViews (doc).Count;
        internal static int GetSheetCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(ViewSheet))
            .Count();
        internal static int GetViewsWithNoTemplate(Document doc) =>
            GetFilteredWorkingViews(doc)
                .Count(v => v.ViewTemplateId == null || v.ViewTemplateId == ElementId.InvalidElementId);

        internal static int GetViewsNotOnSheetCount(Document doc)
        {
            var allViews = GetFilteredWorkingViews(doc);
            var viewsOnSheets = new HashSet<ElementId>();

            var sheets = new FilteredElementCollector(doc)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>();

            foreach (var sheet in sheets)
            {
                var viewIds = sheet.GetAllPlacedViews();
                foreach (var id in viewIds)
                    viewsOnSheets.Add(id);
            }

            return allViews.Count(v => !viewsOnSheets.Contains(v.Id));
        }

        internal static int GetViewsWithClippingDisabled(Document doc)
        {
            return GetFilteredWorkingViews(doc)
                .Where(v => !v.CropBoxActive)
                .Count();
        }

        internal static int GetMaterialCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(Material)).Count();

        internal static int GetLineStyleCount(Document doc)
        {
            Category lineCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
            if (lineCat == null)
                return 0;

            return lineCat.SubCategories.Size;
        }
        internal static int GetLinePatternCount(Document doc) =>
        new FilteredElementCollector(doc).OfClass(typeof(LinePatternElement))
            .Count();
        internal static int GetFillPatternCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(FillPatternElement))
            .Count();
        internal static int GetTextStyleCount(Document doc) =>
            new FilteredElementCollector(doc).OfClass(typeof(TextNoteType)).Count();


        internal static (double eastWestMM, double northSouthMM, double elevMM, double angleDeg) GetProjectBasePoint(Document doc)
        {
            var pbp = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .FirstElement() as BasePoint;

            double ToMeters(double feet) => Math.Round(feet * 0.3048, 3);

            if (pbp != null)
            {
                double ew =    pbp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)?.AsDouble() ?? 0;
                double ns =    pbp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)?.AsDouble() ?? 0;
                double elev =  pbp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)?.AsDouble() ?? 0;
                double angleRad = pbp.get_Parameter(BuiltInParameter.BASEPOINT_ANGLETON_PARAM)?.AsDouble() ?? 0;

                return (
                    ToMeters(ew),
                    ToMeters(ns),
                    ToMeters(elev),
                    angleRad
                    );
            }

            return (0, 0, 0, 0);
        }
        internal static (double eastWestMM, double northSouthMM, double elevMM) GetSurveyPoint(Document doc)
        {
            var sp = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_SharedBasePoint)
                .FirstElement() as BasePoint;

            double ToMeters(double feet) => Math.Round(feet * 0.3048, 3);

            if (sp != null)
            {
                double ew =   sp.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM)?.AsDouble() ?? 0;
                double ns =   sp.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM)?.AsDouble() ?? 0;
                double elev = sp.get_Parameter(BuiltInParameter.BASEPOINT_ELEVATION_PARAM)?.AsDouble() ?? 0;

                return (
                    ToMeters(ew),
                    ToMeters(ns), 
                    ToMeters(elev)
                    );
            }

            return (0, 0, 0);
        }
        public static double GetApprovedFamilyPercentage(Document doc)
        {
            var families = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            int totalTypes = 0;
            int approvedTypes = 0;

            foreach (var family in families)
            {
                foreach (ElementId symbolId in family.GetFamilySymbolIds())
                {
                    var symbol = doc.GetElement(symbolId) as FamilySymbol;
                    if (symbol == null) continue;

                    totalTypes++;
                    Parameter param = symbol.LookupParameter("Copyright");
                    if (param != null && param.HasValue && param.StorageType == StorageType.String)
                    {
                        string value = param.AsString();
                        if (!string.IsNullOrEmpty(value) && value.ToUpperInvariant().Contains("HDR"))
                            approvedTypes++;
                    }
                }
            }

            return totalTypes > 0 ? Math.Round((double)approvedTypes / totalTypes * 100, 2) : 0;
        }

        internal static int Get3DModelElementCount(Document doc)
        {
            // 1) Pick a non-template 3D view (prefer {3D})
            var view3D = new FilteredElementCollector(doc)
                .OfClass(typeof(View3D)).Cast<View3D>()
                .FirstOrDefault(v => !v.IsTemplate && v.Name == "{3D}")
                ?? new FilteredElementCollector(doc)
                    .OfClass(typeof(View3D)).Cast<View3D>()
                    .FirstOrDefault(v => !v.IsTemplate);

            // Fallback if the model truly has no 3D view
            bool useViewFilter = view3D != null;
            ElementId viewId = useViewFilter ? view3D.Id : ElementId.InvalidElementId;

            // 2) Types to exclude even though they may be "Model" category
            var excludedTypes = new HashSet<Type>
            {
                typeof(CurveElement),      
                typeof(Sketch),
                typeof(SketchPlane),
                typeof(SpatialElement),     
                typeof(RevitLinkInstance),  
                typeof(ImportInstance),     
                typeof(Group),              
                typeof(AssemblyInstance)
            };

            // 3) Collector limited to what’s visible in the chosen 3D view
            var collector = useViewFilter
                ? new FilteredElementCollector(doc, viewId)
                : new FilteredElementCollector(doc);

            return collector
                .WhereElementIsNotElementType()
                .Where(e =>
                {
                    // Must be a "Model" category
                    if (e.Category == null || e.Category.CategoryType != CategoryType.Model)
                        return false;

                    // Exclude types above
                    foreach (var t in excludedTypes)
                        if (t.IsAssignableFrom(e.GetType()))
                            return false;

                    // Must have a bbox in the 3D view (extra guard when we have a view)
                    if (useViewFilter && e.get_BoundingBox(view3D) == null)
                        return false;

                    return true;
                })
                .Count();
        }
    }
}
