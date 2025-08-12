using System;
using System.Collections.Generic;
using System.Linq;
using WinForms = System.Windows.Forms;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class WorksetCreate : IExternalCommand
    {
        // Map default worksets to your company naming
        private static readonly Dictionary<string, string> DefaultRenames = new Dictionary<string, string>
        {
            { "Workset 1",               "QA3_ScopeBoxes"          },
            { "Workset1",                "QA3_ScopeBoxes"          }, // some versions omit the space
            { "Worksets",                "QA3_ScopeBoxes"          }, // some versions omit the space
            { "Shared Levels and Grids", "QA1_LevelsGrids"         },
            { "Shared Levels & Grids",   "QA1_LevelsGrids"         }
        };

        // Define worksets per discipline—note that none of these use QA1/QA2/QA3 names
        private static readonly string[] MEPWorksets = new[]
        {
            "QA2_SpacesSeparationLines",
            "QA3_ReferencePlanes",
            "E1_Containment",
            "E1_ElectricalEquipment",
            "E1_ITTelecoms",
            "E1_FireAlarm",
            "E1_Lighting",
            "E1_Security",
            "E1_SmallPower",
            "F1_SprinklerPipework",
            "M1_Pipework",
            "M1_Ventilation",
            "P1_Condensate",
            "P1_Drainage",
            "P1_WaterServices",
            "Z1_Mass",
            "Z1_MEPEquipment",
            "Z1_BuildersWork",
            "Z1_PlantReplacementZones",
            "ZLink_CAD_Description",
            "ZLink_RVT_ArchModel",
            "ZLink_RVT_MEPModel",
            "ZLink_RVT_CivilsModel"
        };

        private static readonly string[] StructureWorksets = new[]
        {
            "QA3_ReferencePlanes",
            "S1_SubStructure",
            "S1_SuperStructure",
            "S2_Existing",
            "ZLink_CAD_Description",
            "ZLink_RVT_ArchModel",
            "ZLink_RVT_MEPModel",
            "ZLink_RVT_CivilsModel"
        };

        private static readonly string[] ArchitectureWorksets = new[]
        {
            "QA2_SpacesSeparationLines",
            "QA3_ReferencePlanes",
            "A1_CoreShell",
            "A1_Equipment",
            "A1_Existing",
            "A1_Furnishings",
            "A1_InteriorConstruction",
            "A2_Lighting",
            "A2_Structure",
            "A3_LifeSafety",
            "A3_WallProtection",
            "ZLink_CAD_Description",
            "ZLink_RVT_MEPModel",
            "ZLink_RVT_StructModel",
            "ZLink_RVT_CivilsModel"
        };

        private static readonly string[] CivilsWorksets = new[]
        {
            "QA3_ReferencePlanes",
            "C1_Civils",
            "C2_Topography",
            "C2_ExistingSite",
            "ZLink_CAD_Description",
            "ZLink_RVT_ArchModel",
            "ZLink_RVT_MEPModel",
            "ZLink_RVT_StructModel"
        };

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            // Ensure worksharing is enabled
            if (!doc.IsWorkshared)
            {
                doc.EnableWorksharing("Shared Levels and Grids", "Worksets");
            }

            // Collect user selections via a simple form
            var form = new WorksetSelectionForm();
            if (form.ShowDialog() != WinForms.DialogResult.OK)
            {
                return Result.Cancelled;
            }

            var selections = form.Selections;

            using (var tx = new Transaction(doc, "Setup Worksets"))
            {
                tx.Start();

                // 1) Create discipline worksets based on user selection
                if (selections.TryGetValue("Architecture", out bool doArch) && doArch)
                {
                    foreach (var name in ArchitectureWorksets)
                    {
                        CreateIfMissing(doc, name);
                    }
                }
                if (selections.TryGetValue("MEP", out bool doMEP) && doMEP)
                {
                    foreach (var name in MEPWorksets)
                    {
                        CreateIfMissing(doc, name);
                    }
                }
                if (selections.TryGetValue("Structures", out bool doStruct) && doStruct)
                {
                    foreach (var name in StructureWorksets)
                    {
                        CreateIfMissing(doc, name);
                    }
                }
                if (selections.TryGetValue("Civils", out bool doCiv) && doCiv)
                {
                    foreach (var name in CivilsWorksets)
                    {
                        CreateIfMissing(doc, name);
                    }
                }

                // 2) Rename the default built-in worksets last
                RenameDefaults(doc);

                tx.Commit();
            }

            return Result.Succeeded;
        }

        /// <summary>
        /// Renames Revit’s default worksets using the WorksetTable API (no reflection needed).
        /// </summary>
        private static void RenameDefaults(Document doc)
        {
            // Gather all user worksets in the document
            var userWorksets = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .Cast<Workset>();

            foreach (var ws in userWorksets)
            {
                if (DefaultRenames.TryGetValue(ws.Name, out string newName) &&
                    !ws.Name.Equals(newName, StringComparison.OrdinalIgnoreCase))
                {
                    // Use WorksetTable.RenameWorkset, which is available in RevitAPI.dll
                    WorksetTable.RenameWorkset(doc, ws.Id, newName);
                }
            }
        }

        /// <summary>
        /// Creates a workset only if it does not already exist.
        /// </summary>
        private static void CreateIfMissing(Document doc, string name)
        {
            bool exists = new FilteredWorksetCollector(doc)
                .OfKind(WorksetKind.UserWorkset)
                .Cast<Workset>()
                .Any(ws => ws.Name.Equals(name, StringComparison.OrdinalIgnoreCase));

            if (!exists)
            {
                Workset.Create(doc, name);
            }
        }

        public static Common.ButtonDataClass GetButtonData() =>
            new Common.ButtonDataClass(
                "WorksetCreate",
                "Workset Create",
                "HDR_EMEA.Tools_Core.WorksetCreate",
                Properties.Resources.WorksetCreate_32,
                Properties.Resources.WorksetCreate_16,
                "Rename defaults and create worksets based on user selection"
            );
    }

    /// <summary>
    /// Simple WinForms dialog used to select which disciplines' worksets to create.
    /// </summary>
    internal class WorksetSelectionForm : WinForms.Form
    {
        public Dictionary<string, bool> Selections { get; } = new Dictionary<string, bool>();

        readonly WinForms.CheckBox chkArch = new() { Text = "Architecture", Left = 20, Top = 20, Width = 200 };
        readonly WinForms.CheckBox chkMEP = new() { Text = "MEP", Left = 20, Top = 50, Width = 200 };
        readonly WinForms.CheckBox chkStruct = new() { Text = "Structures", Left = 20, Top = 80, Width = 200 };
        readonly WinForms.CheckBox chkCivils = new() { Text = "Civils", Left = 20, Top = 110, Width = 200 };
        readonly WinForms.Button btnOk = new() { Text = "OK", Left = 50, Top = 150, Width = 75, DialogResult = WinForms.DialogResult.OK };
        readonly WinForms.Button btnCancel = new() { Text = "Cancel", Left = 150, Top = 150, Width = 75, DialogResult = WinForms.DialogResult.Cancel };

        public WorksetSelectionForm()
        {
            Text = "Select Worksets to Create";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition = WinForms.FormStartPosition.CenterScreen;
            ClientSize = new System.Drawing.Size(280, 200);

            Controls.AddRange(new WinForms.Control[] { chkArch, chkMEP, chkStruct, chkCivils, btnOk, btnCancel });
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        protected override void OnFormClosing(WinForms.FormClosingEventArgs e)
        {
            if (DialogResult == WinForms.DialogResult.OK)
            {
                Selections["Architecture"] = chkArch.Checked;
                Selections["MEP"] = chkMEP.Checked;
                Selections["Structures"] = chkStruct.Checked;
                Selections["Civils"] = chkCivils.Checked;
            }
            base.OnFormClosing(e);
        }
    }
}
