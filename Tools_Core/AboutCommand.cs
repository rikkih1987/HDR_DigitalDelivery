using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using HDR_EMEA.Common;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class AboutCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string title = "HDR EMEA Tools";
            string version = "1.0.0";
            string description = "This Revit add-in contains tools developed for HDR’s EMEA region to improve automation and delivery workflows.";

            string contact = "Dagan.Barnard@hdrinc.com | Rikki.Hartigan@hdrinc.com | Chris.Adams@hdrinc.com";

            TaskDialog dialog = new TaskDialog("About " + title)
            {
                MainInstruction = title,
                MainContent = $"Version: {version}\n\n{description}\n\nContact:\n{contact}",
                CommonButtons = TaskDialogCommonButtons.Close,
                TitleAutoPrefix = false
            };

            dialog.Show();
            return Result.Succeeded;
        }

        public static ButtonDataClass GetButtonData()
        {
            return new ButtonDataClass(
                "cmdAbout",
                "About",
                typeof(AboutCommand).FullName,
                Properties.Resources.About_32,
                Properties.Resources.About_16,
                "About these tools"
            );
        }
    }
}
