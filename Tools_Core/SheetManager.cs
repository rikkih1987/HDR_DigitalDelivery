using DocumentFormat.OpenXml.Spreadsheet;
using HDR_EMEA.Forms;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static HDR_EMEA.Common.ButtonDataClass;
using static HDR_EMEA.Forms.SheetManagerForm;

namespace HDR_EMEA
{
    [Transaction(TransactionMode.Manual)]
    public class SheetManager : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            DateTime startTime = DateTime.Now;
            bool success = false;

            try
            {
                // Revit application and document variables
                UIApplication uiapp = commandData.Application;
                UIDocument uidoc = uiapp.ActiveUIDocument;
                Document doc = uidoc.Document;

                while (true)
                {
                    // 1) Build the latest view-model each loop
                    var collector = new FilteredElementCollector(doc)
                                        .OfClass(typeof(ViewSheet))
                                        .Cast<ViewSheet>();

                    var sheetList = new List<SheetManagerForm.SheetInfo>();

                    var allRevisions = new FilteredElementCollector(doc)
                                            .OfClass(typeof(Revision))
                                            .Cast<Revision>()
                                            .ToList();

                    var dummySheet = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewSheet))
                        .Cast<ViewSheet>()
                        .FirstOrDefault();

                    // Build revisions: "Issued To - Description", select by Id
                    var revisions = new List<SheetManagerForm.RevisionDisplay>();
                    foreach (var rev in allRevisions)
                    {
                        string issuedTo = rev.IssuedTo ?? string.Empty;
                        string description = rev.Description ?? string.Empty;
                        string label =
                            string.IsNullOrWhiteSpace(issuedTo) && string.IsNullOrWhiteSpace(description) ? "<unnamed revision>" :
                            string.IsNullOrWhiteSpace(issuedTo) ? description :
                            string.IsNullOrWhiteSpace(description) ? issuedTo :
                            $"{issuedTo} - {description}";

                        revisions.Add(new SheetManagerForm.RevisionDisplay
                        {
                            DisplayText = label,
                            RevisionId = rev.Id.IntegerValue
                        });
                    }

                    foreach (var vs in collector)
                    {
                        string name = vs.Name;
                        string number = string.Join("-",
                            vs.LookupParameter("Titleblock_ISO_01Project")?.AsString(),
                            vs.LookupParameter("Titleblock_ISO_02Originator")?.AsString(),
                            vs.LookupParameter("Titleblock_ISO_03FuctionalBreakdown")?.AsString(),
                            vs.LookupParameter("Titleblock_ISO_04SpatialBreakdown")?.AsString(),
                            vs.LookupParameter("Titleblock_ISO_05Type")?.AsString(),
                            vs.LookupParameter("Titleblock_ISO_06Discipline")?.AsString(),
                            vs.LookupParameter("Sheet Number")?.AsString()
                        );

                        string revOnSheet = null;
                        var revisionIds = vs.GetAllRevisionIds();
                        if (revisionIds.Count > 0)
                        {
                            var lastRevId = revisionIds.Last();
                            revOnSheet = vs.GetRevisionNumberOnSheet(lastRevId);
                        }

                        FamilyInstance titleBlock = new FilteredElementCollector(doc, vs.Id)
                            .OfCategory(BuiltInCategory.OST_TitleBlocks)
                            .OfClass(typeof(FamilyInstance))
                            .Cast<FamilyInstance>()
                            .FirstOrDefault();

                        string banner = string.Empty;
                        string status = string.Empty;

                        if (titleBlock != null)
                        {
                            var bannerParam = titleBlock.LookupParameter("Titleblock_BannerStatus");
                            var statusParam = titleBlock.LookupParameter("Titleblock_Status");

                            if (bannerParam == null || statusParam == null)
                                continue;

                            if (bannerParam.HasValue)
                                banner = bannerParam.AsValueString();

                            if (statusParam.HasValue)
                                status = statusParam.AsValueString();

                            var bannerVisibility = titleBlock.LookupParameter("Titleblock_Banner");
                            if (bannerVisibility != null && bannerVisibility.StorageType == StorageType.Integer && bannerVisibility.AsInteger() == 0)
                            {
                                banner = "Off";
                            }
                        }
                        else
                        {
                            continue;
                        }

                        sheetList.Add(new SheetManagerForm.SheetInfo
                        {
                            SheetName = name,
                            SheetNumber = number,
                            CurrentRev = revOnSheet,
                            CurrentBanner = banner,
                            CurrentStatus = status
                        });
                    }

                    if (!sheetList.Any())
                    {
                        TaskDialog.Show("Sheet Manager",
                            "No sheets were found with both 'Titleblock_BannerStatus' and 'Titleblock_Status' parameters.");
                        break; // nothing to do
                    }

                    // Banner types (family name filter)
                    var bannerTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                        .Cast<FamilySymbol>()
                        .Where(fs => fs.Family != null &&
                                     fs.Family.Name.StartsWith("HDR_Annotation_Banner", StringComparison.OrdinalIgnoreCase))
                        .Select(fs => fs.Name)
                        .Distinct()
                        .OrderBy(name => name)
                        .ToList();
                    bannerTypes.Insert(0, "Off");

                    // Suitability types (family name filter)
                    var statusTypes = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_GenericAnnotation)
                        .Cast<FamilySymbol>()
                        .Where(fs => fs.Family != null &&
                                     fs.Family.Name.StartsWith("HDR_Annotation_Suitability", StringComparison.OrdinalIgnoreCase))
                        .Select(fs => fs.Name)
                        .Distinct()
                        .OrderBy(name => name)
                        .ToList();

                    // 2) Open the form
                    var currentFormRBS = new Forms.SheetManagerForm(revisions, bannerTypes, statusTypes, sheetList)
                    {
                        Width = 1500,
                        Height = 1000,
                        WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen,
                        Topmost = true
                    };

                    bool? dialogResult = currentFormRBS.ShowDialog();

                    // 3) If user cancelled, exit the loop (and command)
                    if (dialogResult != true)
                        break;

                    // 4) Apply changes, then show success and loop to refresh
                    using (Transaction t = new Transaction(doc, "Update Sheet Rev,Banner,Status Parameters"))
                    {
                        t.Start();

                        foreach (var sheet in currentFormRBS.Sheets.Where(s => s.IsSelected))
                        {
                            var vs = new FilteredElementCollector(doc)
                                .OfClass(typeof(ViewSheet))
                                .Cast<ViewSheet>()
                                .FirstOrDefault(x => x.Name == sheet.SheetName);

                            if (vs == null) continue;

                            FamilyInstance titleBlock = new FilteredElementCollector(doc, vs.Id)
                                .OfCategory(BuiltInCategory.OST_TitleBlocks)
                                .OfClass(typeof(FamilyInstance))
                                .Cast<FamilyInstance>()
                                .FirstOrDefault();

                            if (titleBlock != null)
                            {
                                if (!string.IsNullOrWhiteSpace(currentFormRBS.SelectedBanner))
                                {
                                    var p = titleBlock.LookupParameter("Titleblock_BannerStatus");
                                    if (currentFormRBS.SelectedBanner != "Off")
                                    {
                                        SetFamilyTypeByName(doc, p, currentFormRBS.SelectedBanner);
                                    }
                                }

                                var bannerBool = titleBlock.LookupParameter("Titleblock_Banner");
                                if (bannerBool != null && bannerBool.StorageType == StorageType.Integer)
                                {
                                    bannerBool.Set(currentFormRBS.SelectedBanner != "Off" ? 1 : 0);
                                }

                                if (!string.IsNullOrWhiteSpace(currentFormRBS.SelectedStatus))
                                {
                                    var p = titleBlock.LookupParameter("Titleblock_Status");
                                    if (p != null && p.StorageType == StorageType.ElementId)
                                    {
                                        SetFamilyTypeByName(doc, p, currentFormRBS.SelectedStatus);
                                    }
                                }
                            }

                            // Add selected revision by Id (per-sheet numbering handled by Revit)
                            if (currentFormRBS.SelectedRevisionId.HasValue)
                            {
                                var revId = new ElementId(currentFormRBS.SelectedRevisionId.Value);
                                var revision = doc.GetElement(revId) as Revision;

                                if (revision != null)
                                {
                                    var revisionsOnSheet = vs.GetAdditionalRevisionIds().ToList();
                                    if (!revisionsOnSheet.Contains(revision.Id))
                                    {
                                        revisionsOnSheet.Add(revision.Id);
                                        vs.SetAdditionalRevisionIds(revisionsOnSheet);
                                    }
                                }
                            }
                        }

                        t.Commit();
                    }

                    TaskDialog.Show("Sheet Manager", "Process completed successfully.");
                    // loop continues; the form will reopen with refreshed data
                }

                // mark success once the user exits the loop
                success = true;
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Sheet Manager Error", ex.Message);
                return Result.Failed;
            }
            finally
            {
                Tracker.LogCommandUsage("SheetManager", startTime, success);
            }
        }

        private void SetFamilyTypeByName(Document doc, Autodesk.Revit.DB.Parameter param, string typeName)
        {
            var type = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(x => x.Name == typeName);

            if (type != null)
            {
                param.Set(type.Id);
            }
        }

        /// <summary>
        /// Constructs the ribbon button data for the Revision/Banner/Status command.  
        /// Returning our custom <see cref="Common.ButtonDataClass"/> wrapper keeps the API
        /// consistent with other commands so the calling code in <see cref="App"/> can
        /// uniformly access the underlying <see cref="PushButtonData"/> via the <c>Data</c> property.
        /// </summary>
        /// <returns>A new instance of <see cref="Common.ButtonDataClass"/> configured for this tool.</returns>
        internal static Common.ButtonDataClass GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnSheetManager";
            string buttonTitle = "Sheet Manager";

            ButtonStatus status = ButtonStatus.Standard;

            // These would be toggled as needed:
            bool isNew = false;
            bool isUpdate = false;
            string buttonTooltip = @"
Displays and updates Revision, Banner Status, and Suitability Status for selected sheets in the project.

Version 1.0

How to Use:
- Click the button to open the update window.
- Use the checkboxes to indicate which categories (Revision, Banner Status, Suitability Status) you want to apply.
- Select a value from each checked dropdown.
- Select one or more sheets from the table using the checkboxes.
- Click 'Set' to apply the selected values to the title blocks of the checked sheets.

Notes:
- Revision selection will add the chosen revision to each selected sheet's revision schedule.
- Banner and Suitability values are applied via their corresponding Family Type parameters on the title block instance.
- Sheets or title blocks without the required parameters will not be shown in the list of sheets.
- A confirmation message will appear once the process completes successfully.";

            if (isNew)
                status = ButtonStatus.New;
            else if (isUpdate)
                status = ButtonStatus.Update;

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.SheetManager_32,
                Properties.Resources.SheetManager_16,
                status,
                buttonTooltip);

            // Return the wrapper itself instead of the inner PushButtonData.  
            // This allows the calling code to access the Data property uniformly.
            return myButtonData;
        }
    }

}
