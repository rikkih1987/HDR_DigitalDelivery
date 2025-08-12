using System;
using System.Linq;
using WinForms = System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace HDR_EMEA.Tools_Structural
{
    [Transaction(TransactionMode.Manual)]
    public class PileLengths : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            var uiDoc = commandData.Application.ActiveUIDocument;
            var doc = uiDoc.Document;

            // Prompt user for pile length in mm
            string input = ShowInputDialog("Pile Length", "Enter Pile Length (mm):");
            if (string.IsNullOrWhiteSpace(input))
                return Result.Cancelled;

            if (!double.TryParse(input, out double mm))
            {
                WinForms.MessageBox.Show(
                    $"Invalid number: '{input}'",
                    "Pile Length",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Error
                );
                return Result.Failed;
            }

            // Convert mm -> internal feet
            double lengthInternal = UnitUtils.ConvertToInternalUnits(
                mm,
                UnitTypeId.Millimeters   // <— correct enum
            );

            // Collect all foundation family instances
            var foundations = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                .OfClass(typeof(FamilyInstance))
                .Cast<FamilyInstance>()
                .ToList();

            if (foundations.Count == 0)
            {
                WinForms.MessageBox.Show(
                    "No structural foundation elements found.",
                    "Pile Length",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Information
                );
                return Result.Cancelled;
            }

            int updated = 0;
            using (var tx = new Transaction(doc, "Set Pile Length"))
            {
                tx.Start();
                foreach (var fi in foundations)
                {
                    var p = fi.LookupParameter("Pile_Length");
                    if (p != null && !p.IsReadOnly)
                    {
                        p.Set(lengthInternal);
                        updated++;
                    }
                }
                tx.Commit();
            }

            WinForms.MessageBox.Show(
                $"Updated Pile_Length for {updated} elements.",
                "Pile Length",
                WinForms.MessageBoxButtons.OK,
                WinForms.MessageBoxIcon.Information
            );
            return Result.Succeeded;
        }

        private string ShowInputDialog(string title, string prompt)
        {
            var form = new WinForms.Form
            {
                Width = 400,
                Height = 140,
                Text = title,
                FormBorderStyle = WinForms.FormBorderStyle.FixedDialog,
                StartPosition = WinForms.FormStartPosition.CenterScreen,
                MinimizeBox = false,
                MaximizeBox = false
            };
            var lbl = new WinForms.Label { Left = 10, Top = 10, Width = 360, Text = prompt };
            var txt = new WinForms.TextBox { Left = 10, Top = 35, Width = 360 };
            var btnOk = new WinForms.Button
            {
                Text = "OK",
                Left = 200,
                Width = 80,
                Top = 70,
                DialogResult = WinForms.DialogResult.OK
            };
            var btnCancel = new WinForms.Button
            {
                Text = "Cancel",
                Left = 290,
                Width = 80,
                Top = 70,
                DialogResult = WinForms.DialogResult.Cancel
            };
            form.Controls.AddRange(new System.Windows.Forms.Control[] { lbl, txt, btnOk, btnCancel });
            form.AcceptButton = btnOk;
            form.CancelButton = btnCancel;
            return form.ShowDialog() == WinForms.DialogResult.OK ? txt.Text : null;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "PileLengths",
                "Pile Lengths",
                "HDR_EMEA.Tools_Structural.PileLengths",
                Properties.Resources.PileLengths_32,
                Properties.Resources.PileLengths_16,
                "Prompt for pile length and apply to all foundations"
            );
        }
    }
}
