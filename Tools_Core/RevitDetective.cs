using System;
using System.Linq;
using WinForms = System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class RevitDetective : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Build the options dialog
            var options = new[]
            {
                "Who Created Active View?",
                "Who Created Selected Element?",
                "Who Reloaded Keynotes Last?"
            };

            using (var form = new WinForms.Form())
            {
                form.Text = "Revit Detective";
                form.Width = 300;
                form.Height = 180;
                form.StartPosition = WinForms.FormStartPosition.CenterParent;

                var listBox = new WinForms.ListBox
                {
                    DataSource = options,
                    Dock = WinForms.DockStyle.Top,
                    Height = 90
                };
                form.Controls.Add(listBox);

                var btnOk = new WinForms.Button
                {
                    Text = "OK",
                    Dock = WinForms.DockStyle.Bottom,
                    DialogResult = WinForms.DialogResult.OK
                };
                form.Controls.Add(btnOk);
                form.AcceptButton = btnOk;

                if (form.ShowDialog() != WinForms.DialogResult.OK)
                    return Result.Cancelled;

                string choice = listBox.SelectedItem as string;
                var uiApp = commandData.Application;
                var uiDoc = uiApp.ActiveUIDocument;
                var doc = uiDoc.Document;

                try
                {
                    switch (choice)
                    {
                        case "Who Created Active View?":
                            WhoCreatedActiveView(doc);
                            break;
                        case "Who Created Selected Element?":
                            WhoCreatedSelection(uiDoc);
                            break;
                        case "Who Reloaded Keynotes Last?":
                            WhoReloadedKeynotes(doc);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    TaskDialog.Show("Revit Detective", "Error: " + ex.Message);
                    return Result.Failed;
                }
            }

            return Result.Succeeded;
        }

        private void WhoCreatedActiveView(Document doc)
        {
            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Revit Detective", "Model is not workshared.");
                return;
            }

            View activeView = doc.ActiveView;
            var info = WorksharingUtils.GetWorksharingTooltipInfo(doc, activeView.Id);

            string msg =
                $"Creator of view \"{activeView.Name}\" (Id: {activeView.Id}):\n" +
                $"{info.Creator}";
            TaskDialog.Show("Revit Detective", msg);
        }

        private void WhoCreatedSelection(UIDocument uiDoc)
        {
            Document doc = uiDoc.Document;
            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Revit Detective", "Model is not workshared.");
                return;
            }

            var selIds = uiDoc.Selection.GetElementIds();
            if (selIds.Count != 1)
            {
                TaskDialog.Show("Revit Detective", "Exactly one element must be selected.");
                return;
            }

            ElementId id = selIds.First();
            var info = WorksharingUtils.GetWorksharingTooltipInfo(doc, id);

            string msg =
                $"Creator:       {info.Creator}\n" +
                $"Current Owner: {info.Owner}";
            TaskDialog.Show("Revit Detective", msg);
        }

        private void WhoReloadedKeynotes(Document doc)
        {
            if (!doc.IsWorkshared)
            {
                TaskDialog.Show("Revit Detective", "Model is not workshared.");
                return;
            }

            string path = doc.PathName;
            if (String.IsNullOrEmpty(path))
            {
                TaskDialog.Show("Revit Detective", "Model is not saved yet.");
                return;
            }

            ModelPath mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(path);
            var tx = TransmissionData.ReadTransmissionData(mp);
            var refs = tx.GetAllExternalFileReferenceIds();

            foreach (var rId in refs)
            {
                var last = tx.GetLastSavedReferenceData(rId);
                if (last.ExternalFileReferenceType == ExternalFileReferenceType.KeynoteTable)
                {
                    string userPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(last.GetPath());
                    if (!String.IsNullOrEmpty(userPath))
                    {
                        Element keynote = doc.GetElement(last.GetReferencingId());
                        Parameter editedBy = keynote.get_Parameter(BuiltInParameter.EDITED_BY);

                        string who = (editedBy != null && !String.IsNullOrEmpty(editedBy.AsString()))
                            ? editedBy.AsString()
                            : "<nobody>";
                        string msg = $"Keynote table (Id: {keynote.Id}) was last reloaded by:\n{who}";
                        TaskDialog.Show("Revit Detective", msg);
                        return;
                    }
                }
            }

            TaskDialog.Show("Revit Detective", "No keynote table reference found.");
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "RevitDetective",
                "Revit Detective",
                "HDR_EMEA.Tools_Core.RevitDetective",
                Properties.Resources.RevitDetective_32,   // largeImage (32x32)
                Properties.Resources.RevitDetective_16,   // smallImage (16x16)
                "Inspect who made specific changes in the model"
            );
        }
    }
}
