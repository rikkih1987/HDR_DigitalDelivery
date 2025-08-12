using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class TagDuplicates : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;
            var curView = doc.ActiveView;

            if (curView is ViewSheet)
            {
                TaskDialog.Show("Duplicate Tags", "You're on a Sheet. Please activate a model view.");
                return Result.Cancelled;
            }

            bool shiftHeld = (Keyboard.Modifiers & ModifierKeys.Shift) == ModifierKeys.Shift;
            if (shiftHeld)
            {
                var dlg1 = new Forms.TagSelection(doc)
                {
                    Width = 600,
                    Height = 600,
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                    Topmost = true
                };
                dlg1.ShowDialog();
                return Result.Succeeded;
            }

            var saved = Common.TagMiConfig.LoadSavedBics();
            var tagBics = saved
                .Select(i => (BuiltInCategory)i)
                .Where(bic => Enum.IsDefined(typeof(BuiltInCategory), bic))
                .ToList();

            if (!tagBics.Any())
            {
                TaskDialog.Show("Duplicate Tags", "No tag categories selected. Hold Shift when launching to configure.");
                return Result.Cancelled;
            }

            var chooser = new Forms.TagButtons(doc, tagBics)
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
                TaskDialog.Show("Duplicate Tags", $"Could not map {tagBicValue} to an element category.");
                return Result.Cancelled;
            }

            // Collect tags of chosen category
            var tagCollector = new FilteredElementCollector(doc, curView.Id)
                .OfCategory(tagBicValue)
                .WhereElementIsNotElementType();

            var elementToTags = new Dictionary<ElementId, List<ElementId>>();

            foreach (var tag in tagCollector)
            {
                if (tag is IndependentTag independentTag)
                {
                    var taggedIds = independentTag.GetTaggedLocalElementIds();
                    foreach (var elemId in taggedIds)
                    {
                        if (elemId == ElementId.InvalidElementId) continue;

                        if (!elementToTags.ContainsKey(elemId))
                            elementToTags[elemId] = new List<ElementId>();

                        elementToTags[elemId].Add(independentTag.Id);
                    }
                }
            }

            // Get first tag for each duplicated element
            var duplicateTagIds = elementToTags
                .Where(kvp => kvp.Value.Count > 1)
                .Select(kvp => kvp.Value.First())
                .ToList();

            if (duplicateTagIds.Any())
            {
                uidoc.Selection.SetElementIds(duplicateTagIds);
                TaskDialog.Show("Duplicate Tags",
                    $"{duplicateTagIds.Count} elements in '{elementBic}' have duplicate tags.\n" +
                    "The first tag of each set has been selected.");
            }
            else
            {
                TaskDialog.Show("Duplicate Tags", $"No duplicate tags found for {elementBic} in the active view.");
            }

            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData() =>
            new Common.ButtonDataClass(
                "TagDuplicates",
                "Duplicate Tags",
                "HDR_EMEA.Tools_Core.TagDuplicates",
                Properties.Resources.TagDuplicates_32,
                Properties.Resources.TagDuplicates_16,
                @"
Finds elements in the active view that have more than one tag of the same category.

How to Use:
- Click the button to open the tag category selection window.
- Choose the tag category to check for duplicates.
- If duplicates are found, the first tag in each duplicate set will be selected.

Shift Key Feature:
- Hold SHIFT when launching to configure preferred tag categories.
- Selections are saved for future sessions.

Notes:
- Only elements visible in the active view are checked.
- If no duplicates are found, a confirmation message appears."
            );
    }
}