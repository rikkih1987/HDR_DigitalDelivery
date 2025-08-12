using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using WinForms = System.Windows.Forms;

namespace HDR_EMEA.Tools_Structural
{
    [Transaction(TransactionMode.Manual)]
    public class ElevCreator : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1) Selected walls or prompt
                var wallIds = GetOrPickWalls(uidoc);
                if (wallIds.Count == 0)
                {
                    TaskDialog.Show("Elev Creator", "No walls selected.");
                    return Result.Cancelled;
                }

                // 2) Section types
                var sectionTypes = new FilteredElementCollector(doc)
                    .OfClass(typeof(ViewFamilyType))
                    .Cast<ViewFamilyType>()
                    .Where(vft => vft.ViewFamily == ViewFamily.Section)
                    .OrderBy(t => t.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (sectionTypes.Count == 0)
                {
                    TaskDialog.Show("Elev Creator", "This project has no Section view types.");
                    return Result.Cancelled;
                }

                // 3) Load last-used options from AppData
                var opts = LoadOptionsFromDisk(doc, sectionTypes);

                // 4) Options dialog (type + padding + default flip + name option)
                using (var dlg = new OptionsDialog(sectionTypes, opts))
                {
                    if (dlg.ShowDialog() != WinForms.DialogResult.OK)
                        return Result.Cancelled;

                    opts = dlg.GetOptions();
                }

                // 5) Save options
                SaveOptionsToDisk(opts);

                // 6) Preview with arrows; then prompt to proceed/flip/cancel
                bool flip = opts.FlipDirection;
                var previewIds = DrawPreview(uidoc, wallIds, opts.SectionType, opts.PaddingMm, flip);

                var td = new TaskDialog("Preview sections");
                td.MainInstruction = "Do the preview arrows look correct?";
                td.MainContent = "Choose whether to keep this direction or flip it. Preview lines will be removed automatically.";
                td.CommonButtons = TaskDialogCommonButtons.None;
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Looks correct");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "Flip and create");
                td.AddCommandLink(TaskDialogCommandLinkId.CommandLink3, "Cancel");
                var choice = td.Show();

                // Remove preview
                if (previewIds.Count > 0)
                {
                    using (var t = new Transaction(doc, "Remove preview"))
                    {
                        t.Start();
                        foreach (var id in previewIds)
                        {
                            if (doc.GetElement(id) != null)
                                doc.Delete(id);
                        }
                        t.Commit();
                    }
                }

                if (choice == TaskDialogResult.CommandLink3) return Result.Cancelled;
                if (choice == TaskDialogResult.CommandLink2) flip = !flip;

                // 7) Create sections
                int created = 0;
                using (var tg = new TransactionGroup(doc, "Create wall sections"))
                {
                    tg.Start();

                    var autoTemplate = FindTemplateForSectionType(doc, opts.SectionType);

                    foreach (var id in wallIds)
                    {
                        var wall = doc.GetElement(id) as Wall;
                        if (wall == null) continue;

                        using (var t = new Transaction(doc, "Create Section"))
                        {
                            t.Start();

                            var vs = CreateSectionForWall(doc, wall, opts.SectionType.Id, opts.PaddingMm, flip);
                            if (vs != null)
                            {
                                if (autoTemplate != null)
                                    vs.ViewTemplateId = autoTemplate.Id;

                                // Name: SEC - [Mark - ]Type - Level
                                string mark = wall.get_Parameter(BuiltInParameter.ALL_MODEL_MARK)?.AsString();
                                string levelName = doc.GetElement(wall.LevelId) is Level lvl ? lvl.Name : string.Empty;
                                string baseName = string.IsNullOrWhiteSpace(mark) || !opts.IncludeWallMarkInName
                                    ? $"{wall.Name}"
                                    : $"{mark} - {wall.Name}";

                                vs.Name = MakeUniqueViewName(doc, $"SEC - {baseName} - {levelName}");
                                created++;
                            }

                            t.Commit();
                        }
                    }

                    tg.Assimilate();
                }

                TaskDialog.Show("Elev Creator", $"Created {created} section view(s).");
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        public static HDR_EMEA.Common.ButtonDataClass GetButtonData()
        {
            return new HDR_EMEA.Common.ButtonDataClass(
                "ElevCreator",
                "Elev Creator",
                "HDR_EMEA.Tools_Structural.ElevCreator",
                Properties.Resources.ElevCreator_32,
                Properties.Resources.ElevCreator_16,
                "Create section views along selected walls with padding and flip, with a preview step."
            );
        }

        // Selection helpers
        private static IList<ElementId> GetOrPickWalls(UIDocument uidoc)
        {
            var pre = uidoc.Selection.GetElementIds()
                .Select(id => uidoc.Document.GetElement(id))
                .OfType<Wall>()
                .Select(w => w.Id)
                .ToList();

            if (pre.Count > 0)
                return pre;

            var picked = uidoc.Selection.PickObjects(ObjectType.Element, new WallSelectionFilter(), "Select walls for sections");
            return picked
                .Select(r => uidoc.Document.GetElement(r.ElementId))
                .OfType<Wall>()
                .Select(w => w.Id)
                .ToList();
        }

        private class WallSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem is Wall;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        // Create the actual section
        private static ViewSection CreateSectionForWall(
            Document doc,
            Wall wall,
            ElementId viewFamilyTypeId,
            double paddingMm,
            bool flip)
        {
            var lc = wall.Location as LocationCurve;
            if (lc == null) return null;

            var curve = lc.Curve;
            XYZ p = curve.GetEndPoint(0);
            XYZ q = curve.GetEndPoint(1);

            XYZ v = flip ? (p - q) : (q - p);

            BoundingBoxXYZ bb = wall.get_BoundingBox(null);
            double minZ = bb?.Min.Z ?? 0.0;
            double maxZ = bb?.Max.Z ?? 0.0;

            double w = v.GetLength();
            double wallWidth = wall.WallType.Width;
            double pad = MmToFt(paddingMm);

            double baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0.0;
            double userHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? (maxZ - minZ);

            // Section coordinate system: X along wall, Y up, Z facing direction
            XYZ min = new XYZ(-0.5 * w - pad, baseOffset - pad, -pad - 0.5 * wallWidth);
            XYZ max = new XYZ(0.5 * w + pad, baseOffset + userHeight + pad, pad + 0.5 * wallWidth);

            int fc = (p.X > q.X || (Math.Abs(p.X - q.X) < 1e-9 && p.Y < q.Y)) ? 1 : -1;
            XYZ midpoint = flip ? q + 0.5 * v : p + 0.5 * v;

            XYZ walldir = (fc * v.Normalize());
            XYZ up = XYZ.BasisZ;
            XYZ viewdir = walldir.CrossProduct(up);

            Transform tr = Transform.Identity;
            tr.Origin = midpoint;
            tr.BasisX = walldir;
            tr.BasisY = up;
            tr.BasisZ = viewdir;

            var sectionBox = new BoundingBoxXYZ
            {
                Transform = tr,
                Min = min,
                Max = max
            };

            return ViewSection.CreateSection(doc, viewFamilyTypeId, sectionBox);
        }

        // Auto-pick a section view template by name based on the selected section type
        private static View FindTemplateForSectionType(Document doc, ViewFamilyType sectionType)
        {
            string name = sectionType?.Name ?? string.Empty;
            if (string.IsNullOrWhiteSpace(name)) return null;

            var templates = new FilteredElementCollector(doc)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(v => v.IsTemplate && v.ViewType == ViewType.Section)
                .ToList();

            var exact = templates.FirstOrDefault(v => v.Name == name);
            if (exact != null) return exact;

            var ci = templates.FirstOrDefault(v => v.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
            if (ci != null) return ci;

            var contains = templates.FirstOrDefault(v => v.Name.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0);
            return contains;
        }

        private static string MakeUniqueViewName(Document doc, string baseName)
        {
            string name = baseName;
            int i = 1;
            while (new FilteredElementCollector(doc)
                   .OfClass(typeof(View))
                   .Cast<View>()
                   .Any(v => !v.IsTemplate && v.Name.Equals(name, StringComparison.OrdinalIgnoreCase)))
            {
                i++;
                name = $"{baseName} ({i})";
            }
            return name;
        }

        private static double MmToFt(double mm) => mm / 304.8;

        // Preview drawing with facing-direction arrows
        private static IList<ElementId> DrawPreview(
            UIDocument uidoc,
            IList<ElementId> wallIds,
            ViewFamilyType sectionType,
            double paddingMm,
            bool flip)
        {
            var doc = uidoc.Document;
            var view = uidoc.ActiveView;
            var ids = new List<ElementId>();
            bool useDetail = ViewSupportsDetailLines(view);

            // Optional custom line style to improve visibility (category Lines → subcategory HDR_Preview_Section)
            GraphicsStyle gs = TryGetLineStyle(doc, "HDR_Preview_Section");

            // Sketch plane fallback (for model lines)
            SketchPlane modelPlane = null;
            if (!useDetail)
            {
                var plane = Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero);
                using (var t = new Transaction(doc, "Create preview sketch plane"))
                {
                    t.Start();
                    modelPlane = SketchPlane.Create(doc, plane);
                    t.Commit();
                }
            }

            // Sizes (mm)
            double pad = MmToFt(Math.Max(0, paddingMm));
            double arrowLen = MmToFt(350);   // arrow shaft
            double wingLen = MmToFt(150);   // arrow wings
            double wingBack = MmToFt(100);   // back from head for wings
            double endTick = MmToFt(200);   // end ticks on section line

            using (var t = new Transaction(doc, "Preview section direction"))
            {
                t.Start();

                Func<Curve, ElementId> addCurve = c =>
                {
                    Element e;
                    if (useDetail)
                    {
                        e = doc.Create.NewDetailCurve(view, c);
                    }
                    else
                    {
                        e = doc.Create.NewModelCurve(c, modelPlane);
                    }

                    if (gs != null && e is CurveElement ce)
                        ce.LineStyle = gs;

                    return e.Id;
                };

                foreach (var id in wallIds)
                {
                    var wall = doc.GetElement(id) as Wall;
                    var lc = wall?.Location as LocationCurve;
                    if (lc == null) continue;

                    var curve = lc.Curve;
                    XYZ p = curve.GetEndPoint(0);
                    XYZ q = curve.GetEndPoint(1);
                    XYZ v = flip ? (p - q) : (q - p);

                    int fc = (p.X > q.X || (Math.Abs(p.X - q.X) < 1e-9 && p.Y < q.Y)) ? 1 : -1;
                    XYZ walldir = (fc * v.Normalize());
                    XYZ up = XYZ.BasisZ;
                    XYZ viewdir = walldir.CrossProduct(up).Normalize();

                    double half = v.GetLength() * 0.5;
                    XYZ mid = flip ? q + 0.5 * v : p + 0.5 * v;

                    // Section line (slightly inset so ticks are visible)
                    XYZ a = mid - walldir * (half - pad * 0.5);
                    XYZ b = mid + walldir * (half - pad * 0.5);
                    ids.Add(addCurve(Line.CreateBound(a, b)));

                    // End ticks
                    XYZ tickOff = viewdir * endTick * 0.5;
                    ids.Add(addCurve(Line.CreateBound(a - tickOff, a + tickOff)));
                    ids.Add(addCurve(Line.CreateBound(b - tickOff, b + tickOff)));

                    // Arrow at midpoint in facing direction
                    XYZ head = mid + viewdir * arrowLen;
                    XYZ leftWing = head + (walldir * wingLen - viewdir * wingBack);
                    XYZ rightWing = head + (-walldir * wingLen - viewdir * wingBack);

                    ids.Add(addCurve(Line.CreateBound(mid, head)));         // shaft
                    ids.Add(addCurve(Line.CreateBound(head, leftWing)));    // wing
                    ids.Add(addCurve(Line.CreateBound(head, rightWing)));   // wing
                }

                t.Commit();
            }

            return ids;
        }

        private static bool ViewSupportsDetailLines(View v)
        {
            switch (v.ViewType)
            {
                case ViewType.FloorPlan:
                case ViewType.CeilingPlan:
                case ViewType.AreaPlan:
                case ViewType.EngineeringPlan:
                case ViewType.Section:
                case ViewType.Elevation:
                case ViewType.DraftingView:
                case ViewType.Detail:
                    return true;
                default:
                    return false;
            }
        }

        // Try to find a GraphicsStyle called HDR_Preview_Section under Lines
        private static GraphicsStyle TryGetLineStyle(Document doc, string name)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(GraphicsStyle))
                .Cast<GraphicsStyle>()
                .FirstOrDefault(gs =>
                    gs.GraphicsStyleCategory?.Parent?.Name == "Lines" &&
                    gs.Name.Equals(name, StringComparison.OrdinalIgnoreCase));
        }

        // Lightweight settings storage (per-user)
        private class ElevOptions
        {
            public ViewFamilyType SectionType { get; set; }
            public double PaddingMm { get; set; }
            public bool FlipDirection { get; set; }
            public bool IncludeWallMarkInName { get; set; }
        }

        private class ElevOptionsPortable
        {
            public string SectionTypeName { get; set; } = "";
            public double PaddingMm { get; set; } = 500;
            public bool FlipDirection { get; set; } = false;
            public bool IncludeWallMarkInName { get; set; } = true;
        }

        private static string SettingsFolder =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "HDR_EMEA");

        private static string SettingsPath =>
            Path.Combine(SettingsFolder, "ElevCreator.settings");

        private static ElevOptions LoadOptionsFromDisk(Document doc, IList<ViewFamilyType> sectionTypes)
        {
            var portable = new ElevOptionsPortable();

            try
            {
                if (File.Exists(SettingsPath))
                {
                    var lines = File.ReadAllLines(SettingsPath);
                    foreach (var line in lines)
                    {
                        var idx = line.IndexOf('=');
                        if (idx <= 0) continue;
                        var key = line.Substring(0, idx).Trim();
                        var val = line.Substring(idx + 1).Trim();

                        switch (key)
                        {
                            case "SectionTypeName": portable.SectionTypeName = val; break;
                            case "PaddingMm": if (double.TryParse(val, out var d)) portable.PaddingMm = d; break;
                            case "FlipDirection": if (bool.TryParse(val, out var b1)) portable.FlipDirection = b1; break;
                            case "IncludeWallMarkInName": if (bool.TryParse(val, out var b2)) portable.IncludeWallMarkInName = b2; break;
                        }
                    }
                }
            }
            catch { /* ignore and keep defaults */ }

            var opts = new ElevOptions
            {
                SectionType = sectionTypes.FirstOrDefault(t => string.Equals(t.Name, portable.SectionTypeName, StringComparison.OrdinalIgnoreCase))
                               ?? sectionTypes.FirstOrDefault(),
                PaddingMm = Math.Max(0, portable.PaddingMm),
                FlipDirection = portable.FlipDirection,
                IncludeWallMarkInName = portable.IncludeWallMarkInName
            };

            return opts;
        }

        private static void SaveOptionsToDisk(ElevOptions opts)
        {
            try
            {
                if (!Directory.Exists(SettingsFolder))
                    Directory.CreateDirectory(SettingsFolder);

                var lines = new[]
                {
                    $"SectionTypeName={(opts.SectionType?.Name ?? "")}",
                    $"PaddingMm={opts.PaddingMm}",
                    $"FlipDirection={opts.FlipDirection}",
                    $"IncludeWallMarkInName={opts.IncludeWallMarkInName}"
                };
                File.WriteAllLines(SettingsPath, lines);
            }
            catch
            {
                // Non-fatal if we can't save preferences
            }
        }

        // Simple options dialog (no template picker)
        private class OptionsDialog : WinForms.Form
        {
            private WinForms.ComboBox cboType;
            private WinForms.NumericUpDown nudPadding;
            private WinForms.CheckBox chkFlip;
            private WinForms.CheckBox chkIncludeMark;
            private WinForms.Button btnOk;
            private WinForms.Button btnCancel;

            private readonly IList<ViewFamilyType> _types;
            private readonly ElevOptions _initial;

            public OptionsDialog(IList<ViewFamilyType> sectionTypes, ElevOptions initial)
            {
                _types = sectionTypes;
                _initial = initial;

                Text = "Elev Creator";
                StartPosition = WinForms.FormStartPosition.CenterParent;
                MinimizeBox = false;
                MaximizeBox = false;
                FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
                AutoSize = true;
                AutoSizeMode = WinForms.AutoSizeMode.GrowAndShrink;
                Padding = new WinForms.Padding(10);

                var layout = new WinForms.TableLayoutPanel
                {
                    ColumnCount = 2,
                    RowCount = 5,
                    AutoSize = true,
                    Dock = WinForms.DockStyle.Fill
                };
                layout.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.AutoSize));
                layout.ColumnStyles.Add(new WinForms.ColumnStyle(WinForms.SizeType.Percent, 100));

                // Section Type
                layout.Controls.Add(new WinForms.Label { Text = "Section type", AutoSize = true }, 0, 0);
                cboType = new WinForms.ComboBox { DropDownStyle = WinForms.ComboBoxStyle.DropDownList, Width = 260 };
                foreach (var t in _types) cboType.Items.Add(t);
                cboType.DisplayMember = "Name";
                layout.Controls.Add(cboType, 1, 0);

                // Padding
                layout.Controls.Add(new WinForms.Label { Text = "Padding [mm]", AutoSize = true }, 0, 1);
                nudPadding = new WinForms.NumericUpDown
                {
                    Minimum = 0,
                    Maximum = 10000,
                    Increment = 50,
                    Value = (decimal)Math.Max(0, _initial.PaddingMm),
                    Width = 120
                };
                layout.Controls.Add(nudPadding, 1, 1);

                // Flip default
                layout.Controls.Add(new WinForms.Label { Text = "Default flip", AutoSize = true }, 0, 2);
                chkFlip = new WinForms.CheckBox { Checked = _initial.FlipDirection, AutoSize = true };
                layout.Controls.Add(chkFlip, 1, 2);

                // Include Mark
                layout.Controls.Add(new WinForms.Label { Text = "Include wall Mark in name", AutoSize = true }, 0, 3);
                chkIncludeMark = new WinForms.CheckBox { Checked = _initial.IncludeWallMarkInName, AutoSize = true };
                layout.Controls.Add(chkIncludeMark, 1, 3);

                // Buttons
                var btnPanel = new WinForms.FlowLayoutPanel
                {
                    FlowDirection = WinForms.FlowDirection.RightToLeft,
                    Dock = WinForms.DockStyle.Fill,
                    AutoSize = true
                };
                btnOk = new WinForms.Button { Text = "OK", DialogResult = WinForms.DialogResult.OK };
                btnCancel = new WinForms.Button { Text = "Cancel", DialogResult = WinForms.DialogResult.Cancel };
                btnPanel.Controls.AddRange(new WinForms.Control[] { btnOk, btnCancel });
                layout.Controls.Add(btnPanel, 0, 4);
                layout.SetColumnSpan(btnPanel, 2);

                Controls.Add(layout);

                // Initialise selections
                if (_initial.SectionType != null)
                    cboType.SelectedItem = _types.FirstOrDefault(t => t.Id == _initial.SectionType.Id) ?? _types.First();

                AcceptButton = btnOk;
                CancelButton = btnCancel;

                btnOk.Click += (s, e) =>
                {
                    if (cboType.SelectedItem == null)
                    {
                        WinForms.MessageBox.Show("Please choose a section type.", "Elev Creator");
                        DialogResult = WinForms.DialogResult.None;
                    }
                };
            }

            public ElevOptions GetOptions()
            {
                return new ElevOptions
                {
                    SectionType = (ViewFamilyType)cboType.SelectedItem,
                    PaddingMm = (double)nudPadding.Value,
                    FlipDirection = chkFlip.Checked,
                    IncludeWallMarkInName = chkIncludeMark.Checked
                };
            }
        }
    }
}
