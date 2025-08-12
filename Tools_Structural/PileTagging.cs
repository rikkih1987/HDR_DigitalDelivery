using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using WinForms = System.Windows.Forms;

namespace HDR_EMEA.Tools_Structural
{
    [Transaction(TransactionMode.Manual)]
    public class PileTagging : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, Autodesk.Revit.DB.ElementSet elements)
        {
            WinForms.MessageBox.Show("To be developed", "Pile Tagging");
            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "PileTagging",
                "Pile Tagging",
                "HDR_EMEA.Tools_Structural.PileTagging",
                Properties.Resources.PileTagging_32,   // largeImage (32x32)
                Properties.Resources.PileTagging_16,   // smallImage (16x16)
                "Placeholder for Pile Tagging tool"  // tooltip
            );
        }
    }
}
