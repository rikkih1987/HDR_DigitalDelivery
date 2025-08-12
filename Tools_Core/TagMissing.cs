using Autodesk.Revit.DB.Mechanical;
using HDR_EMEA.Common;
using HDR_EMEA.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    internal class TagMissing : IExternalCommand
    {
        public Result Execute(ExternalCommandData data, ref string message, ElementSet elements)
        {
            var uiapp = data.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;
            var view = doc.ActiveView;

            if (view is ViewSheet)
            {
                TaskDialog.Show("Tag Missing", "You're on a Sheet. Please activate a model view.");
                return Result.Cancelled;
            }

            bool shiftHeld = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
            if (shiftHeld)
            {
                var dlg1 = new HDR_EMEA.Forms.TagSelection(doc)
                {
                    Width = 600,
                    Height = 600,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    Topmost = true
                };
                dlg1.ShowDialog();
                return Result.Succeeded;
            }

            // Load saved tag categories
            var saved = Common.TagMiConfig.LoadSavedBics();
            var tagBics = saved
                .Select(i => (BuiltInCategory)i)
                .Where(bic => Enum.IsDefined(typeof(BuiltInCategory), bic))
                .ToList();

            if (!tagBics.Any())
            {
                TaskDialog.Show("Tag Missing", "No tag categories selected. Hold Shift when launching to configure required tags.");
                return Result.Cancelled;
            }

            // Let the user pick one category
            var chooser = new HDR_EMEA.Forms.TagButtons(doc, tagBics)
            {
                Owner = System.Windows.Application.Current?.MainWindow,
                Topmost = true
            };

            BuiltInCategory? chosenTagBic = null;
            chooser.TagButtonClicked += (bic) => chosenTagBic = bic;
            var dlgResult = chooser.ShowDialog();
            if (dlgResult != true || chosenTagBic == null)
                return Result.Cancelled;

            var tagBicValue = chosenTagBic.Value;

            if (!Common.TagUtils.TryMapTagToElement(tagBicValue, out var elementBic))
            {
                TaskDialog.Show("Tag Missing", $"Could not map {tagBicValue} to an element category.");
                return Result.Cancelled;
            }

            // Find untagged elements
            var untagged = CollectUntaggedInView(doc, view, tagBicValue, elementBic);
            if (!untagged.Any())
            {
                TaskDialog.Show("Tag Missing", $"All {elementBic} in this view appear to be tagged.");
                return Result.Succeeded;
            }

            uidoc.Selection.SetElementIds(untagged);
            TaskDialog.Show("Missing Tags",
                $"Category: {tagBicValue} → Elements: {elementBic}\n" +
                $"Selected {untagged.Count} untagged element(s) in the active view.");
            return Result.Succeeded;
        }

        private static List<ElementId> CollectUntaggedInView(Document doc, View view,
            BuiltInCategory tagBic, BuiltInCategory elementBic)
        {
            var tagElems = new FilteredElementCollector(doc, view.Id)
                .OfCategory(tagBic)
                .WhereElementIsNotElementType()
                .ToElements();

            var taggedIds = new HashSet<ElementId>(new ElementIdEqualityComparer());
            foreach (var tag in tagElems)
            {
                foreach (var id in GetTaggedLocalElementIds(tag))
                {
                    if (id != ElementId.InvalidElementId)
                        taggedIds.Add(id);
                }
            }

            var candidateIds = new FilteredElementCollector(doc, view.Id)
                .OfCategory(elementBic)
                .WhereElementIsNotElementType()
                .ToElementIds();

            return candidateIds.Where(id => !taggedIds.Contains(id)).ToList();
        }

        private static ISet<ElementId> GetTaggedLocalElementIds(Element tagElement)
        {
            return tagElement switch
            {
                IndependentTag it => it.GetTaggedLocalElementIds() ?? new HashSet<ElementId>(),
                RoomTag rt => rt.Room != null ? new HashSet<ElementId> { rt.Room.Id } : new HashSet<ElementId>(),
                AreaTag at => at.Area != null ? new HashSet<ElementId> { at.Area.Id } : new HashSet<ElementId>(),
                SpaceTag st => st.Space != null ? new HashSet<ElementId> { st.Space.Id } : new HashSet<ElementId>(),
                _ => new HashSet<ElementId>()
            };
        }

        private sealed class ElementIdEqualityComparer : IEqualityComparer<ElementId>
        {
            public bool Equals(ElementId x, ElementId y) => x?.IntegerValue == y?.IntegerValue;
            public int GetHashCode(ElementId obj) => obj?.IntegerValue.GetHashCode() ?? 0;
        }

        public static Common.ButtonDataClass GetButtonData() =>
            new Common.ButtonDataClass(
                "TagMissing",
                "Missing Tags",
                "HDR_EMEA.Tools_Core.TagMissing",
                Properties.Resources.TagMissing_32,
                Properties.Resources.TagMissing_16,
                @"
Find and select elements in the active view that do not have tags.

How to Use:
- Click the button to run the Missing Tag check.
- Choose the tag category button to check for untagged elements.
- Untagged elements found will be selected.

Shift Key Feature:
- Hold SHIFT when launching to configure preferred tag categories.
- Selections are saved for future sessions.

Notes:
- Only elements visible in the active view are checked.
- Categories with no elements in the active view will return no results."
            );
    }
}
