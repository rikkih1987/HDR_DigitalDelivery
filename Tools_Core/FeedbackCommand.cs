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
    public class FeedbackCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Process.Start("https://forms.office.com/Pages/ResponsePage.aspx?id=AeJnNtzLs0ibQl0tPxbiqdIYaRnLm-ZEkAWOmZHnybZUQlgzTlBNUVFCUkJLOEk1WjBERFgxOVUyTy4u");
            return Result.Succeeded;
        }

        public static ButtonDataClass GetButtonData()
        {
            return new ButtonDataClass(
                "cmdFeedback",
                "Feedback",
                typeof(FeedbackCommand).FullName,
                Properties.Resources.SubmitFeedback_32,
                Properties.Resources.SubmitFeedback_16,

                "Submit feedback for the HDR EMEA tools"
            );
        }
    }
}
