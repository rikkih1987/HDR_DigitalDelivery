using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HDR_EMEA.Common
{
    internal static class TagUtils
    {
        private static readonly Dictionary<BuiltInCategory, BuiltInCategory> ExplicitMap =
            new Dictionary<BuiltInCategory, BuiltInCategory>
        {
            { BuiltInCategory.OST_WallTags,                 BuiltInCategory.OST_Walls },
            { BuiltInCategory.OST_DoorTags,                 BuiltInCategory.OST_Doors },
            { BuiltInCategory.OST_WindowTags,               BuiltInCategory.OST_Windows },
            { BuiltInCategory.OST_RoomTags,                 BuiltInCategory.OST_Rooms },
            { BuiltInCategory.OST_AreaTags,                 BuiltInCategory.OST_Areas },
            { BuiltInCategory.OST_CeilingTags,              BuiltInCategory.OST_Ceilings },
            { BuiltInCategory.OST_FloorTags,                BuiltInCategory.OST_Floors },
            { BuiltInCategory.OST_RoofTags,                 BuiltInCategory.OST_Roofs },
            { BuiltInCategory.OST_ColumnTags,               BuiltInCategory.OST_Columns },
            { BuiltInCategory.OST_StructuralColumnTags,     BuiltInCategory.OST_StructuralColumns },
            { BuiltInCategory.OST_StructuralFoundationTags, BuiltInCategory.OST_StructuralFoundation },
            { BuiltInCategory.OST_StructuralFramingTags,    BuiltInCategory.OST_StructuralFraming },
            { BuiltInCategory.OST_PipeTags,                 BuiltInCategory.OST_PipeCurves },
            { BuiltInCategory.OST_DuctTags,                 BuiltInCategory.OST_DuctCurves },
            { BuiltInCategory.OST_CableTrayTags,            BuiltInCategory.OST_CableTray },
            { BuiltInCategory.OST_ConduitTags,              BuiltInCategory.OST_Conduit },
            { BuiltInCategory.OST_GenericModelTags,         BuiltInCategory.OST_GenericModel },
            { BuiltInCategory.OST_ParkingTags,              BuiltInCategory.OST_Parking },
        };

        public static bool TryMapTagToElement(BuiltInCategory tagBic, out BuiltInCategory elementBic)
        {
            if (ExplicitMap.TryGetValue(tagBic, out elementBic))
                return true;

            string name = tagBic.ToString(); // e.g. "OST_WallTags"
            if (name.EndsWith("Tags", StringComparison.Ordinal))
            {
                var candidate = name.Substring(0, name.Length - "Tags".Length); // "OST_Wall"
                if (Enum.TryParse(candidate, out elementBic))
                    return true;

                // plural heuristic: "OST_StructuralColumnTags" -> "OST_StructuralColumns"
                if (!candidate.EndsWith("s", StringComparison.Ordinal))
                {
                    var plural = candidate + "s";
                    if (Enum.TryParse(plural, out elementBic))
                        return true;
                }
            }

            elementBic = 0;
            return false;
        }

        public static List<(BuiltInCategory TagBic, BuiltInCategory ElementBic, int Count)>
            CheckForMissingTaggableElements(Document doc, IEnumerable<BuiltInCategory> tagBics)
        {
            var results = new List<(BuiltInCategory TagBic, BuiltInCategory ElementBic, int Count)>();

            foreach (var tagBic in tagBics)
            {
                if (!TryMapTagToElement(tagBic, out var elementBic))
                    continue;

                int count = new FilteredElementCollector(doc)
                    .OfCategory(elementBic)
                    .WhereElementIsNotElementType()
                    .Count();

                results.Add((tagBic, elementBic, count));
            }

            // Only keep those with zero elements present
            return results.Where(r => r.Count == 0).ToList();
        }
    }
}