using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Diagnostics;
using HDR_EMEA.Common;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class VisitDDCCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Process.Start("https://hdrinc.sharepoint.com/teams/DigitalDesignCommunity/SitePages/BES_Digital_Delivery.aspx");
            return Result.Succeeded;
        }

        public static ButtonDataClass GetButtonData()
        {
            return new ButtonDataClass(
                "cmdVisitDDC",
                "DDC",
                typeof(VisitDDCCommand).FullName,
                Properties.Resources.SharePoint_32,
                Properties.Resources.SharePoint_16,
                "Visit the BES Digital Design Community homepage"
            );
        }
    }
}
