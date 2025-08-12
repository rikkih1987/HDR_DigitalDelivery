using FilterTreeControlWPF;
using System.Windows.Media.Media3D;
using static HDR_EMEA.Common.ButtonDataClass;

namespace HDR_EMEA
{
    [Transaction(TransactionMode.Manual)]
    public class FounRef : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Revit application and document variables
            //UIApplication uiapp = commandData.Application;
            //UIDocument uidoc = uiapp.ActiveUIDocument;
            //Document doc = uidoc.Document;
            FoundationCommandHandler handler = new FoundationCommandHandler();
            ExternalEvent exEvent = ExternalEvent.Create(handler);

            // Your code goes here
            Forms.FounRefForm form = new Forms.FounRefForm(handler, exEvent);
            form.Width = 800;
            form.Height = 450;
            form.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;
            form.Topmost = true;


            form.ShowDialog();

            return Result.Succeeded;
        }
        /// <summary>
        /// Creates the ribbon button data for the Foundation Reference tool.  
        /// Returning a <see cref="Common.ButtonDataClass"/> aligns with the rest of the add‑in,
        /// allowing consumers to uniformly access the underlying <see cref="PushButtonData"/> via
        /// the <c>Data</c> property.
        /// </summary>
        internal static Common.ButtonDataClass GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnFoundationRef";
            string buttonTitle = "Foundation Ref.";
            string buttonTooltip = @"
This is a tooltip for FounRef WIP.

Version 1.0
";

            ButtonStatus status = ButtonStatus.Standard;

            // These would be toggled as needed:
            bool isNew = false;
            bool isUpdate = false;

            if (isNew)
                status = ButtonStatus.New;
            else if (isUpdate)
                status = ButtonStatus.Update;

            Common.ButtonDataClass myButtonData = new Common.ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.FoundationRef_32,
                Properties.Resources.FoundationRef_16,
                status,
                buttonTooltip);

            // Return the wrapper instead of the PushButtonData so callers can access .Data consistently.
            return myButtonData;
        }
    }

}
