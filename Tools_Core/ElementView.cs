// This file contains a complete implementation of the ElementView command with
// persistent settings and improved view cropping for cross sections and elevations.
// The command allows users to create cross sections, elevations, and callouts
// for selected elements in Revit, with per-view templates, custom scale, padding
// and depth specified in millimetres, optional sheet placement with a chosen title
// block, and persistent storage of the last-used options.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using WinForms = System.Windows.Forms;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class ElementView : IExternalCommand
    {
        // -------------------------------------------------------------------------
        // Persistent settings
        // The tool persists its last-used settings in a simple key=value file in
        // the user's AppData folder. Each run will reload the file and apply the
        // stored defaults to the UI, and save the current selections back on exit.
        // -------------------------------------------------------------------------

        private static readonly string SettingsFolder =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "HDR_EMEA");

        private static readonly string SettingsFile =
            Path.Combine(SettingsFolder, "ElementViewSettings.txt");

        private static Dictionary<string, string> LoadSettings()
        {
            var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                if (File.Exists(SettingsFile))
                {
                    foreach (var line in File.ReadAllLines(SettingsFile))
                    {
                        var parts = line.Split(new[] { '=' }, 2);
                        if (parts.Length == 2)
                            dict[parts[0].Trim()] = parts[1].Trim();
                    }
                }
            }
            catch
            {
                // Ignore any errors reading settings.
            }
            return dict;
        }

        private static void SaveSettings(Dictionary<string, string> dict)
        {
            try
            {
                Directory.CreateDirectory(SettingsFolder);
                var lines = dict.Select(kvp => $"{kvp.Key}={kvp.Value}").ToArray();
                File.WriteAllLines(SettingsFile, lines);
            }
            catch
            {
                // Ignore errors writing settings.
            }
        }

        // -------------------------------------------------------------------------
        // UI dialog class
        // Defines the controls for the command: toggles for cross section,
        // elevation and callout; template selectors; scale; padding and depth in
        // millimetres; sheet placement and title block. This class exposes
        // helpers to read the user's selections.
        // -------------------------------------------------------------------------
        private class ElementViewDialog : WinForms.Form
        {
            // View toggles
            public WinForms.CheckBox ChkCreateCrossSection;
            public WinForms.CheckBox ChkCreateElevation;
            public WinForms.CheckBox ChkCreateCallout;

            // Templates per view type
            public WinForms.ComboBox CmbTemplateCross;
            public WinForms.ComboBox CmbTemplateElevation;
            public WinForms.ComboBox CmbTemplateCallout;

            // Scale (single, applied to all)
            public WinForms.ComboBox CmbScale;

            // Padding and depth (mm)
            public WinForms.NumericUpDown NudPaddingMM;
            public WinForms.NumericUpDown NudDepthForwardMM;
            public WinForms.NumericUpDown NudDepthBackMM;

            // Sheet placement
            public WinForms.CheckBox ChkPlaceOnSheet;
            public WinForms.ComboBox CmbTitleBlock;

            public WinForms.Button BtnOk;
            public WinForms.Button BtnCancel;

            public ElementViewDialog(IEnumerable<string> templateNames, IEnumerable<string> titleblockNames)
            {
                Text = "Element View";
                StartPosition = FormStartPosition.CenterParent;
                MinimizeBox = false;
                MaximizeBox = false;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                Width = 640;
                Height = 560;

                int left = 12;
                int labelW = 260;
                int inputW = 330;

                // Views group
                var grpViews = new WinForms.GroupBox
                {
                    Text = "Views to create",
                    Left = left,
                    Top = 12,
                    Width = 600,
                    Height = 80
                };
                ChkCreateCrossSection = new WinForms.CheckBox
                {
                    Text = "Cross Section (perpendicular)",
                    Left = 12,
                    Top = 24,
                    Width = 240,
                    Checked = true
                };
                ChkCreateElevation = new WinForms.CheckBox
                {
                    Text = "Elevation (requires active Plan view)",
                    Left = 260,
                    Top = 24,
                    Width = 300,
                    Checked = true
                };
                ChkCreateCallout = new WinForms.CheckBox
                {
                    Text = "Callout (requires active Plan view)",
                    Left = 12,
                    Top = 50,
                    Width = 280,
                    Checked = false
                };
                grpViews.Controls.AddRange(new WinForms.Control[]
                {
                    ChkCreateCrossSection,
                    ChkCreateElevation,
                    ChkCreateCallout
                });

                int rowY = 102;

                // Templates
                AddLabel("View template for Cross Section:", left, rowY, labelW);
                CmbTemplateCross = AddTemplateCombo(left + labelW, rowY, inputW, templateNames);

                rowY += 28;
                AddLabel("View template for Elevation:", left, rowY, labelW);
                CmbTemplateElevation = AddTemplateCombo(left + labelW, rowY, inputW, templateNames);

                rowY += 28;
                AddLabel("View template for Callout:", left, rowY, labelW);
                CmbTemplateCallout = AddTemplateCombo(left + labelW, rowY, inputW, templateNames);

                // Scale
                rowY += 36;
                AddLabel("Scale (denominator applied to all views):", left, rowY, labelW);
                CmbScale = new WinForms.ComboBox
                {
                    Left = left + labelW,
                    Top = rowY - 3,
                    Width = 160,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                int[] scales = { 1, 2, 5, 10, 20, 25, 50, 100, 200 };
                foreach (var s in scales)
                    CmbScale.Items.Add($"1:{s}");
                CmbScale.SelectedItem = "1:50";
                Controls.Add(CmbScale);

                // Padding and depths in mm
                rowY += 34;
                AddLabel("Padding around width/height (mm):", left, rowY, labelW);
                NudPaddingMM = MkMmNumeric(left + labelW, rowY, 50);
                Controls.Add(NudPaddingMM);

                rowY += 28;
                AddLabel("Depth forward (mm):", left, rowY, labelW);
                NudDepthForwardMM = MkMmNumeric(left + labelW, rowY, 300);
                Controls.Add(NudDepthForwardMM);

                rowY += 28;
                AddLabel("Depth back (mm):", left, rowY, labelW);
                NudDepthBackMM = MkMmNumeric(left + labelW, rowY, 300);
                Controls.Add(NudDepthBackMM);

                // Sheet options
                rowY += 34;
                ChkPlaceOnSheet = new WinForms.CheckBox
                {
                    Left = left,
                    Top = rowY,
                    Width = 600,
                    Text = "Place created views on a new sheet",
                    Checked = true
                };
                Controls.Add(ChkPlaceOnSheet);

                rowY += 28;
                AddLabel("Title block family type:", left, rowY, labelW);
                CmbTitleBlock = new WinForms.ComboBox
                {
                    Left = left + labelW,
                    Top = rowY - 3,
                    Width = inputW,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                CmbTitleBlock.Items.Add("None");
                foreach (var tb in titleblockNames)
                    CmbTitleBlock.Items.Add(tb);
                CmbTitleBlock.SelectedIndex = CmbTitleBlock.Items.Count > 1 ? 1 : 0;
                Controls.Add(CmbTitleBlock);

                // Buttons
                BtnOk = new WinForms.Button
                {
                    Text = "OK",
                    Left = 410,
                    Width = 90,
                    Top = 470,
                    DialogResult = DialogResult.OK
                };
                BtnCancel = new WinForms.Button
                {
                    Text = "Cancel",
                    Left = 506,
                    Width = 90,
                    Top = 470,
                    DialogResult = DialogResult.Cancel
                };
                Controls.Add(grpViews);
                Controls.Add(BtnOk);
                Controls.Add(BtnCancel);
                AcceptButton = BtnOk;
                CancelButton = BtnCancel;
            }

            private static WinForms.NumericUpDown MkMmNumeric(int left, int top, int def)
            {
                return new WinForms.NumericUpDown
                {
                    Left = left,
                    Top = top - 3,
                    Width = 100,
                    DecimalPlaces = 0,
                    Minimum = 0,
                    Maximum = 100000,
                    Increment = 10,
                    Value = def
                };
            }

            private void AddLabel(string text, int left, int top, int width)
            {
                Controls.Add(new WinForms.Label
                {
                    Text = text,
                    Left = left,
                    Top = top,
                    Width = width
                });
            }

            private WinForms.ComboBox AddTemplateCombo(int left, int top, int width, IEnumerable<string> templateNames)
            {
                var cmb = new WinForms.ComboBox
                {
                    Left = left,
                    Top = top - 3,
                    Width = width,
                    DropDownStyle = ComboBoxStyle.DropDownList
                };
                cmb.Items.Add("None");
                foreach (var n in templateNames)
                    cmb.Items.Add(n);
                cmb.SelectedIndex = 0;
                Controls.Add(cmb);
                return cmb;
            }

            // Expose selections
            public bool WantCross() => ChkCreateCrossSection.Checked;
            public bool WantElevation() => ChkCreateElevation.Checked;
            public bool WantCallout() => ChkCreateCallout.Checked;

            public string TemplateNameCross() => CmbTemplateCross.SelectedItem?.ToString() ?? "None";
            public string TemplateNameElevation() => CmbTemplateElevation.SelectedItem?.ToString() ?? "None";
            public string TemplateNameCallout() => CmbTemplateCallout.SelectedItem?.ToString() ?? "None";

            public int SelectedScaleDenominator()
            {
                var text = CmbScale.SelectedItem?.ToString() ?? "1:50";
                var parts = text.Split(':');
                return (parts.Length == 2 && int.TryParse(parts[1], out int d) && d > 0) ? d : 50;
            }

            // mm -> values (convert to feet outside the dialog)
            public double PaddingMM() => (double)NudPaddingMM.Value;
            public double DepthForwardMM() => (double)NudDepthForwardMM.Value;
            public double DepthBackMM() => (double)NudDepthBackMM.Value;

            public bool PlaceOnSheet() => ChkPlaceOnSheet.Checked;
            public string SelectedTitleBlockName() => CmbTitleBlock.SelectedItem?.ToString() ?? "None";
        }

        // -------------------------------------------------------------------------
        // Selection filter: allows selection of any non-view-specific element.
        // -------------------------------------------------------------------------
        private class LooseSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => elem != null && !elem.ViewSpecific;
            public bool AllowReference(Reference reference, XYZ position) => true;
        }

        // -------------------------------------------------------------------------
        // Element properties helper: extracts position, main vector, width, height and depth
        // for an element to base the generated views on. Handles walls, curve-based elements
        // and generic families.
        // -------------------------------------------------------------------------
        private class ElementProperties
        {
            public XYZ Origin;
            public XYZ Vector;   // primary direction in XY
            public double Width; // along Vector
            public double Height;
            public double Depth;

            public ElementProperties(Document doc, Element el)
            {
                if (el is Wall w)
                {
                    GetWallProperties(w);
                    return;
                }
                if (el.Location is LocationCurve lc)
                {
                    GetCurveBasedProperties(el, lc);
                    return;
                }
                GetGenericFamilyProperties(doc, el);
            }

            private void GetWallProperties(Wall w)
            {
                var curve = (w.Location as LocationCurve)?.Curve;
                if (curve != null)
                {
                    var s = curve.GetEndPoint(0);
                    var e = curve.GetEndPoint(1);
                    Vector = e - s;
                    var bb = w.get_BoundingBox(null);
                    Origin = (bb.Max + bb.Min) / 2.0;
                    Width = Vector.GetLength();

                    var hParam = w.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                    Height = (hParam != null && hParam.HasValue) ? hParam.AsDouble() : (bb.Max.Z - bb.Min.Z);
                    Depth = Math.Max(bb.Max.Y - bb.Min.Y, 0.1);
                }
                else
                {
                    var bb = w.get_BoundingBox(null);
                    Origin = (bb.Max + bb.Min) / 2.0;
                    var size = bb.Max - bb.Min;
                    Width = Math.Abs(size.X);
                    Height = Math.Abs(size.Z);
                    Depth = Math.Abs(size.Y);
                    Vector = XYZ.BasisX;
                }
            }

            private void GetCurveBasedProperties(Element el, LocationCurve lc)
            {
                var curve = lc.Curve;
                var s = curve.GetEndPoint(0);
                var e = curve.GetEndPoint(1);
                // Flatten to XY
                if (Math.Abs(s.Z - e.Z) > 1e-8)
                    e = new XYZ(e.X, e.Y, s.Z);

                Vector = e - s;

                var bb = el.get_BoundingBox(null);
                Origin = (bb.Max + bb.Min) / 2.0;
                Width = Vector.GetLength();
                Height = bb.Max.Z - bb.Min.Z;
                Depth = Math.Max(bb.Max.Y - bb.Min.Y, 0.1);
            }

            private void GetGenericFamilyProperties(Document doc, Element el)
            {
                var bb = el.get_BoundingBox(null);
                Origin = (bb.Max + bb.Min) / 2.0;

                FamilySymbol sym = null;
                try { sym = doc.GetElement(el.GetTypeId()) as FamilySymbol; } catch { }

                BoundingBoxXYZ bbTyp = null;
                try { if (sym != null) bbTyp = sym.get_BoundingBox(null); } catch { }

                if (bbTyp != null)
                {
                    Width = bbTyp.Max.X - bbTyp.Min.X;
                    Height = bbTyp.Max.Z - bbTyp.Min.Z;
                    Depth = Math.Max(bbTyp.Max.Y - bbTyp.Min.Y, 0.1);

                    var ptStart = new XYZ(bbTyp.Min.X, (bbTyp.Min.Y + bbTyp.Max.Y) / 2.0, bbTyp.Min.Z);
                    var ptEnd = new XYZ(bbTyp.Max.X, (bbTyp.Min.Y + bbTyp.Max.Y) / 2.0, bbTyp.Min.Z);
                    Vector = ptEnd - ptStart;

                    // Apply family instance rotation if available
                    try
                    {
                        if (el.Location is LocationPoint lp && Math.Abs(lp.Rotation) > 1e-12)
                            Vector = RotateZ(Vector, lp.Rotation);
                    }
                    catch { }
                }
                else
                {
                    var size = bb.Max - bb.Min;
                    Width = Math.Abs(size.X);
                    Height = Math.Abs(size.Z);
                    Depth = Math.Abs(size.Y);
                    Vector = XYZ.BasisX;
                }
            }

            private static XYZ RotateZ(XYZ v, double a)
            {
                var x = v.X * Math.Cos(a) - v.Y * Math.Sin(a);
                var y = v.X * Math.Sin(a) + v.Y * Math.Cos(a);
                return new XYZ(x, y, v.Z);
            }
        }

        // -------------------------------------------------------------------------
        // Main Execute method
        // -------------------------------------------------------------------------
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Millimetres to feet conversion factor
            const double MM_TO_FT = 1.0 / 304.8;

            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;

            try
            {
                // Gather templates and title blocks from the document
                var allTemplates = new FilteredElementCollector(doc)
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Where(v => v.IsTemplate)
                    .OrderBy(v => v.Name)
                    .ToList();
                var templateNameToId = new SortedDictionary<string, ElementId>(StringComparer.OrdinalIgnoreCase);
                foreach (var v in allTemplates)
                    if (!templateNameToId.ContainsKey(v.Name))
                        templateNameToId.Add(v.Name, v.Id);

                var titleBlocks = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_TitleBlocks)
                    .Cast<FamilySymbol>()
                    .OrderBy(tb => tb.FamilyName + " : " + tb.Name)
                    .ToList();
                var tbNameToId = new SortedDictionary<string, ElementId>(StringComparer.OrdinalIgnoreCase);
                foreach (var tb in titleBlocks)
                {
                    var label = $"{tb.FamilyName} : {tb.Name}";
                    if (!tbNameToId.ContainsKey(label))
                        tbNameToId.Add(label, tb.Id);
                }

                // Load persisted settings
                var settings = LoadSettings();

                using (var dlg = new ElementViewDialog(templateNameToId.Keys, tbNameToId.Keys))
                {
                    // Apply saved settings to dialog controls
                    ApplySettingsToDialog(dlg, settings);

                    // Show the dialog and cancel if user cancels
                    if (dlg.ShowDialog() != DialogResult.OK)
                        return Result.Cancelled;

                    bool wantCross = dlg.WantCross();
                    bool wantElevation = dlg.WantElevation();
                    bool wantCallout = dlg.WantCallout();
                    if (!wantCross && !wantElevation && !wantCallout)
                    {
                        TaskDialog.Show("Element View", "No view types selected.");
                        return Result.Cancelled;
                    }

                    // Save settings for next run
                    PersistSettingsFromDialog(dlg, settings);
                    SaveSettings(settings);

                    // Resolve templates based on user selections
                    var templateIdCross = GetTemplateId(templateNameToId, dlg.TemplateNameCross());
                    var templateIdElevation = GetTemplateId(templateNameToId, dlg.TemplateNameElevation());
                    var templateIdCallout = GetTemplateId(templateNameToId, dlg.TemplateNameCallout());

                    // Parse scale and convert mm to feet for pad and depth
                    int scaleDenom = dlg.SelectedScaleDenominator();
                    double padFt = dlg.PaddingMM() * MM_TO_FT;
                    double depthFFt = dlg.DepthForwardMM() * MM_TO_FT;
                    double depthBFt = dlg.DepthBackMM() * MM_TO_FT;

                    // Determine sheet placement and title block
                    bool placeOnSheet = dlg.PlaceOnSheet();
                    ElementId selectedTbId = ElementId.InvalidElementId;
                    if (placeOnSheet)
                    {
                        var tbName = dlg.SelectedTitleBlockName();
                        if (!string.Equals(tbName, "None", StringComparison.OrdinalIgnoreCase))
                        {
                            tbNameToId.TryGetValue(tbName, out selectedTbId);
                            if (selectedTbId == ElementId.InvalidElementId)
                            {
                                TaskDialog.Show("Element View", "No valid title block selected. Sheets will not be created.");
                                placeOnSheet = false;
                            }
                        }
                        else
                            placeOnSheet = false;
                    }

                    // Prompt the user to select elements
                    IList<Reference> pickedRefs;
                    try
                    {
                        pickedRefs = uidoc.Selection.PickObjects(
                            ObjectType.Element,
                            new LooseSelectionFilter(),
                            "Select elements, then click Finish");
                    }
                    catch
                    {
                        return Result.Cancelled;
                    }
                    if (pickedRefs == null || pickedRefs.Count == 0)
                    {
                        TaskDialog.Show("Element View", "No elements selected.");
                        return Result.Cancelled;
                    }

                    // Translate references to elements
                    var elems = pickedRefs
                        .Select(r => doc.GetElement(r))
                        .Where(e => e != null)
                        .ToList();

                    // Acquire necessary view family types
                    var sectionType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Section);
                    var elevationType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Elevation);
                    var calloutType = new FilteredElementCollector(doc)
                        .OfClass(typeof(ViewFamilyType))
                        .Cast<ViewFamilyType>()
                        .FirstOrDefault(vft => vft.ViewFamily == ViewFamily.Detail);

                    if (wantCross && sectionType == null)
                    {
                        TaskDialog.Show("Element View", "No Section view type found.");
                        return Result.Failed;
                    }
                    if (wantElevation && elevationType == null)
                    {
                        TaskDialog.Show("Element View", "No Elevation view type found.");
                        return Result.Failed;
                    }
                    if (wantCallout && calloutType == null)
                    {
                        TaskDialog.Show("Element View", "No Detail (callout) view type found.");
                        return Result.Failed;
                    }

                    // Determine if we can host elevation/callout on the active plan
                    var hostPlan = uidoc.ActiveView as ViewPlan;
                    bool canHostPlan = hostPlan != null && hostPlan.CanBePrinted;

                    using (var tx = new Transaction(doc, "Create Views from Selection"))
                    {
                        tx.Start();

                        foreach (var el in elems)
                        {
                            var ep = new ElementProperties(doc, el);
                            var mainDir = ep.Vector;
                            if (mainDir == null || mainDir.GetLength() < 1e-9)
                                mainDir = XYZ.BasisX;
                            mainDir = new XYZ(mainDir.X, mainDir.Y, 0).Normalize();

                            // Cross Section
                            ViewSection vCross = null;
                            if (wantCross)
                            {
                                // The cross section view direction is perpendicular to the main orientation in the XY plane.
                                var crossDir = Perp90(mainDir);
                                // Compute oriented extents for the cross orientation. This returns the width, height
                                // and forward/back depths based on the element's bounding box projected into the cross
                                // section coordinate system. Using these values ensures the cross section covers the
                                // full extents of the element when viewed in this orientation.
                                var dimsCrossOrient = ComputeDimsForOrientation(el, ep.Origin, crossDir, padFt, depthFFt, depthBFt);
                                var trCross = dimsCrossOrient.tr;
                                vCross = CreateSection(doc, sectionType, ep.Origin, crossDir,
                                                       dimsCrossOrient.width, dimsCrossOrient.height,
                                                       dimsCrossOrient.forward, dimsCrossOrient.back);
                                SetCropBoxTo(vCross, trCross, dimsCrossOrient.width, dimsCrossOrient.height,
                                             dimsCrossOrient.forward, dimsCrossOrient.back);
                                ApplyTemplateAndScale(vCross, templateIdCross, scaleDenom);
                            }

                            // Elevation (requires host plan and cardinal orientation)
                            ViewSection vElevation = null;
                            if (wantElevation && canHostPlan)
                            {
                                // Determine the cardinal orientation for the elevation (E/N/W/S) based on the main direction.
                                int idx = PickElevationIndex(mainDir);
                                XYZ elevDir = idx == 0 ? new XYZ(1, 0, 0)
                                         : idx == 1 ? new XYZ(0, 1, 0)
                                         : idx == 2 ? new XYZ(-1, 0, 0)
                                         : new XYZ(0, -1, 0);
                                // Compute oriented dims for this elevation orientation. Use these dims directly
                                // (without merging with callout extents) so the elevation fully covers the element when
                                // viewed in the cardinal direction.
                                var dimsElevOrient = ComputeDimsForOrientation(el, ep.Origin, elevDir, padFt, depthFFt, depthBFt);
                                var trElev = dimsElevOrient.tr;
                                vElevation = CreateElevationAt(doc, elevationType, hostPlan, ep.Origin, mainDir, scaleDenom);
                                SetCropBoxTo(vElevation, trElev, dimsElevOrient.width, dimsElevOrient.height,
                                             dimsElevOrient.forward, dimsElevOrient.back);
                                ApplyTemplateAndScale(vElevation, templateIdElevation, scaleDenom);
                            }

                            // Callout (requires plan)
                            View vCallout = null;
                            if (wantCallout && canHostPlan)
                            {
                                var bb = el.get_BoundingBox(null);
                                var halfX = Math.Max((bb.Max.X - bb.Min.X) / 2.0 + padFt, 1.0 / 12.0);
                                var halfY = Math.Max((bb.Max.Y - bb.Min.Y) / 2.0 + padFt, 1.0 / 12.0);
                                var min3d = new XYZ(ep.Origin.X - halfX, ep.Origin.Y - halfY, 0);
                                var max3d = new XYZ(ep.Origin.X + halfX, ep.Origin.Y + halfY, 0);
                                try
                                {
                                    vCallout = ViewSection.CreateCallout(doc, hostPlan.Id, calloutType.Id, min3d, max3d);
                                    if (vCallout != null)
                                        ApplyTemplateAndScale(vCallout, templateIdCallout, scaleDenom);
                                }
                                catch
                                {
                                    vCallout = null;
                                }
                            }

                            // Assign names
                            var elType = doc.GetElement(el.GetTypeId());
                            var typeName = elType != null ? elType.Name : el.Name;
                            var baseName = $"{typeName}_{el.Id.IntegerValue}";
                            SafeRenameView(vCross, $"{baseName}_Cross");
                            SafeRenameView(vElevation, $"{baseName}_Elevation");
                            SafeRenameView(vCallout, $"{baseName}_Callout");

                            // Sheet placement
                            if (placeOnSheet && selectedTbId != ElementId.InvalidElementId)
                            {
                                var sheet = ViewSheet.Create(doc, selectedTbId);
                                // Provide up to three slots for the views: cross, elevation, callout
                                var slots = new[]
                                {
                                    new XYZ(1.85, 1.65, 0),
                                    new XYZ(1.50, 1.65, 0),
                                    new XYZ(1.85, 1.35, 0)
                                };
                                int slot = 0;
                                TryAddToSheet(doc, sheet, vCross, slots[slot++]);
                                TryAddToSheet(doc, sheet, vElevation, slots[slot++]);
                                TryAddToSheet(doc, sheet, vCallout, slots[slot++]);

                                var catName = el.Category?.Name ?? "Element";
                                var proposedNumber = $"{typeName}_{el.Id.IntegerValue}";
                                var proposedName = $"{catName} - Auto Views";
                                SafeSetSheetNumberAndName(sheet, proposedNumber, proposedName);
                            }
                        }

                        tx.Commit();
                    }

                    return Result.Succeeded;
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }

        // -------------------------------------------------------------------------
        // Apply previously saved settings to the dialog controls
        // -------------------------------------------------------------------------
        private static void ApplySettingsToDialog(ElementViewDialog dlg, Dictionary<string, string> settings)
        {
            // View templates
            if (settings.TryGetValue("CrossTemplate", out var tmp))
            {
                int idx = dlg.CmbTemplateCross.Items.IndexOf(tmp);
                if (idx >= 0) dlg.CmbTemplateCross.SelectedIndex = idx;
            }
            if (settings.TryGetValue("ElevationTemplate", out tmp))
            {
                int idx = dlg.CmbTemplateElevation.Items.IndexOf(tmp);
                if (idx >= 0) dlg.CmbTemplateElevation.SelectedIndex = idx;
            }
            if (settings.TryGetValue("CalloutTemplate", out tmp))
            {
                int idx = dlg.CmbTemplateCallout.Items.IndexOf(tmp);
                if (idx >= 0) dlg.CmbTemplateCallout.SelectedIndex = idx;
            }
            // Scale
            if (settings.TryGetValue("Scale", out tmp))
            {
                foreach (var item in dlg.CmbScale.Items)
                {
                    if (item.ToString().EndsWith(tmp))
                    {
                        dlg.CmbScale.SelectedItem = item;
                        break;
                    }
                }
            }
            // Padding & depth
            if (settings.TryGetValue("PadMM", out tmp) && double.TryParse(tmp, out double mm))
            {
                try { dlg.NudPaddingMM.Value = (decimal)mm; } catch { }
            }
            if (settings.TryGetValue("DepthForwardMM", out tmp) && double.TryParse(tmp, out mm))
            {
                try { dlg.NudDepthForwardMM.Value = (decimal)mm; } catch { }
            }
            if (settings.TryGetValue("DepthBackMM", out tmp) && double.TryParse(tmp, out mm))
            {
                try { dlg.NudDepthBackMM.Value = (decimal)mm; } catch { }
            }
            // Sheet placement
            if (settings.TryGetValue("PlaceOnSheet", out tmp) && bool.TryParse(tmp, out bool b))
            {
                dlg.ChkPlaceOnSheet.Checked = b;
            }
            if (settings.TryGetValue("TitleBlock", out tmp))
            {
                int idx = dlg.CmbTitleBlock.Items.IndexOf(tmp);
                if (idx >= 0) dlg.CmbTitleBlock.SelectedIndex = idx;
            }
            // Toggles
            if (settings.TryGetValue("CreateCross", out tmp) && bool.TryParse(tmp, out bool bc))
            {
                dlg.ChkCreateCrossSection.Checked = bc;
            }
            if (settings.TryGetValue("CreateElevation", out tmp) && bool.TryParse(tmp, out bool be))
            {
                dlg.ChkCreateElevation.Checked = be;
            }
            if (settings.TryGetValue("CreateCallout", out tmp) && bool.TryParse(tmp, out bool bo))
            {
                dlg.ChkCreateCallout.Checked = bo;
            }
        }

        // -------------------------------------------------------------------------
        // Persist user selections from the dialog to the settings dictionary
        // -------------------------------------------------------------------------
        private static void PersistSettingsFromDialog(ElementViewDialog dlg, Dictionary<string, string> settings)
        {
            settings["CrossTemplate"] = dlg.TemplateNameCross();
            settings["ElevationTemplate"] = dlg.TemplateNameElevation();
            settings["CalloutTemplate"] = dlg.TemplateNameCallout();
            settings["Scale"] = dlg.SelectedScaleDenominator().ToString();
            settings["PadMM"] = dlg.PaddingMM().ToString();
            settings["DepthForwardMM"] = dlg.DepthForwardMM().ToString();
            settings["DepthBackMM"] = dlg.DepthBackMM().ToString();
            settings["PlaceOnSheet"] = dlg.PlaceOnSheet().ToString();
            settings["TitleBlock"] = dlg.SelectedTitleBlockName();
            settings["CreateCross"] = dlg.WantCross().ToString();
            settings["CreateElevation"] = dlg.WantElevation().ToString();
            settings["CreateCallout"] = dlg.WantCallout().ToString();
        }

        // -------------------------------------------------------------------------
        // Helper functions
        // -------------------------------------------------------------------------
        private static (double width, double back, double forward, double height, Transform tr)
            ComputeDimsForOrientation(Element el, XYZ origin, XYZ viewDir, double padFt, double depthFFt, double depthBFt)
        {
            var dir = (viewDir == null || viewDir.GetLength() < 1e-9) ? XYZ.BasisX : new XYZ(viewDir.X, viewDir.Y, 0).Normalize();
            var up = Math.Abs(dir.DotProduct(XYZ.BasisZ)) > 0.95 ? XYZ.BasisY : XYZ.BasisZ;
            var right = up.CrossProduct(dir).Normalize();
            up = dir.CrossProduct(right).Normalize();

            var tr = Transform.Identity;
            tr.Origin = origin;
            tr.BasisX = right;  // local X (width)
            tr.BasisY = dir;    // local Y (forward/back)
            tr.BasisZ = up;     // local Z (height)

            var inv = tr.Inverse;
            var bb = el.get_BoundingBox(null);

            double minX = double.PositiveInfinity, minY = double.PositiveInfinity, minZ = double.PositiveInfinity;
            double maxX = double.NegativeInfinity, maxY = double.NegativeInfinity, maxZ = double.NegativeInfinity;

            foreach (var p in EnumerateCorners(bb.Min, bb.Max))
            {
                var q = inv.OfPoint(p);
                if (q.X < minX) minX = q.X; if (q.X > maxX) maxX = q.X;
                if (q.Y < minY) minY = q.Y; if (q.Y > maxY) maxY = q.Y;
                if (q.Z < minZ) minZ = q.Z; if (q.Z > maxZ) maxZ = q.Z;
            }

            var xExtent = Math.Max(0.1, maxX - minX);
            var yBack = Math.Max(0.0, -minY);
            var yFwd = Math.Max(0.0, maxY);
            var zExtent = Math.Max(0.1, maxZ - minZ);

            var width = xExtent + 2 * padFt;
            var height = zExtent + 2 * padFt;
            var forward = yFwd + depthFFt;
            var back = yBack + depthBFt;

            return (width, back, forward, height, tr);
        }

        private static IEnumerable<XYZ> EnumerateCorners(XYZ min, XYZ max)
        {
            yield return new XYZ(min.X, min.Y, min.Z);
            yield return new XYZ(max.X, min.Y, min.Z);
            yield return new XYZ(min.X, max.Y, min.Z);
            yield return new XYZ(max.X, max.Y, min.Z);
            yield return new XYZ(min.X, min.Y, max.Z);
            yield return new XYZ(max.X, min.Y, max.Z);
            yield return new XYZ(min.X, max.Y, max.Z);
            yield return new XYZ(max.X, max.Y, max.Z);
        }

        private static void SetCropBoxTo(ViewSection view, Transform tr, double width, double height, double forward, double back)
        {
            if (view == null) return;
            try
            {
                var halfW = width / 2.0;
                var bb = view.CropBox ?? new BoundingBoxXYZ();
                bb.Transform = tr;
                bb.Min = new XYZ(-halfW, -back, 0.0);
                bb.Max = new XYZ(+halfW, +forward, height);
                view.CropBox = bb;
                if (!view.CropBoxActive) view.CropBoxActive = true;
                if (!view.CropBoxVisible) view.CropBoxVisible = true;
            }
            catch { }
        }

        private static ElementId GetTemplateId(SortedDictionary<string, ElementId> map, string nameOrNone)
        {
            if (string.IsNullOrWhiteSpace(nameOrNone) || string.Equals(nameOrNone, "None", StringComparison.OrdinalIgnoreCase))
                return ElementId.InvalidElementId;
            return map.TryGetValue(nameOrNone, out var id) ? id : ElementId.InvalidElementId;
        }

        private static void ApplyTemplateAndScale(View v, ElementId tpl, int scale)
        {
            if (v == null) return;
            try
            {
                if (tpl != ElementId.InvalidElementId)
                    v.ViewTemplateId = tpl;
            }
            catch { }
            try
            {
                if (scale > 0)
                    v.Scale = scale;
            }
            catch { }
            try
            {
                if (!v.CropBoxActive) v.CropBoxActive = true;
                if (!v.CropBoxVisible) v.CropBoxVisible = true;
            }
            catch { }
        }

        private static void SafeRenameView(View v, string name)
        {
            if (v == null) return;
            var attempt = name;
            for (int i = 0; i < 30; i++)
            {
                try
                {
                    v.Name = attempt;
                    break;
                }
                catch
                {
                    attempt += "*";
                }
            }
        }

        private static void SafeSetSheetNumberAndName(ViewSheet sheet, string number, string name)
        {
            if (sheet == null) return;
            var attempt = number;
            for (int i = 0; i < 30; i++)
            {
                try
                {
                    sheet.SheetNumber = attempt;
                    sheet.Name = name;
                    break;
                }
                catch
                {
                    attempt += "*";
                }
            }
        }

        private static void TryAddToSheet(Document doc, ViewSheet sheet, View view, XYZ ptOnSheet)
        {
            try
            {
                if (sheet != null && view != null && Viewport.CanAddViewToSheet(doc, sheet.Id, view.Id))
                    Viewport.Create(doc, sheet.Id, view.Id, ptOnSheet);
            }
            catch { }
        }

        // Compute dimensions for cross section cropping that match callout plan extents.
        // This returns a tuple containing the full width, back depth, forward depth,
        // and height for a section. Width is based on the global X span of the
        // element’s bounding box (plus padding); forward/back depth are based on
        // the global Y span (plus padding and respective depth offsets); height is
        // the global Z span (plus padding). This ensures the cross section crop
        // region matches the plan callout cropping extents.
        private static (double width, double back, double forward, double height) ComputeCalloutMatchingDims(
            Element el,
            double padFt,
            double depthFFt,
            double depthBFt)
        {
            var bb = el.get_BoundingBox(null);
            double xExtent = bb.Max.X - bb.Min.X;
            double yExtent = bb.Max.Y - bb.Min.Y;
            double zExtent = bb.Max.Z - bb.Min.Z;
            double width = xExtent + 2 * padFt;
            double height = zExtent + 2 * padFt;
            double forward = (yExtent / 2.0) + padFt + depthFFt;
            double back = (yExtent / 2.0) + padFt + depthBFt;
            return (width, back, forward, height);
        }

        // Create a section view with an oriented crop box
        private static ViewSection CreateSection(
            Document doc,
            ViewFamilyType sectionType,
            XYZ origin,
            XYZ viewDir,
            double width,
            double height,
            double forward,
            double back)
        {
            try
            {
                if (sectionType == null) return null;
                var dir = (viewDir == null || viewDir.GetLength() < 1e-9) ? XYZ.BasisX : new XYZ(viewDir.X, viewDir.Y, 0).Normalize();
                var up = Math.Abs(dir.DotProduct(XYZ.BasisZ)) > 0.95 ? XYZ.BasisY : XYZ.BasisZ;
                var right = up.CrossProduct(dir).Normalize();
                up = dir.CrossProduct(right).Normalize();

                var tr = Transform.Identity;
                tr.Origin = origin;
                tr.BasisX = right;
                tr.BasisY = dir;
                tr.BasisZ = up;

                var halfW = width / 2.0;
                var min = new XYZ(-halfW, -back, 0.0);
                var max = new XYZ(+halfW, +forward, height);

                var bb = new BoundingBoxXYZ { Transform = tr, Min = min, Max = max };
                return ViewSection.CreateSection(doc, sectionType.Id, bb);
            }
            catch
            {
                return null;
            }
        }

        // Create an elevation view via an elevation marker
        private static ViewSection CreateElevationAt(
            Document doc,
            ViewFamilyType elevationType,
            ViewPlan hostPlan,
            XYZ origin,
            XYZ facingDir,
            int scaleDenominator)
        {
            try
            {
                if (elevationType == null || hostPlan == null) return null;
                var marker = ElevationMarker.CreateElevationMarker(doc, elevationType.Id, origin, scaleDenominator);
                int index = PickElevationIndex(facingDir);
                return marker.CreateElevation(doc, hostPlan.Id, index);
            }
            catch
            {
                return null;
            }
        }

        // Map a direction vector to one of four cardinal elevation indices (0..3)
        private static int PickElevationIndex(XYZ dir)
        {
            var d = new XYZ(dir.X, dir.Y, 0);
            if (d.GetLength() < 1e-9)
                d = XYZ.BasisX;
            d = d.Normalize();
            if (Math.Abs(d.X) >= Math.Abs(d.Y))
                return d.X >= 0 ? 0 : 2; // East or West
            else
                return d.Y >= 0 ? 1 : 3; // North or South
        }

        // Compute a perpendicular direction in the XY plane
        private static XYZ Perp90(XYZ v)
        {
            var n = v;
            if (n == null || n.GetLength() < 1e-9)
                n = XYZ.BasisX;
            n = n.Normalize();
            var perp = new XYZ(-n.Y, n.X, 0.0);
            if (perp.GetLength() < 1e-9)
                perp = XYZ.BasisY;
            return perp.Normalize();
        }

        // -------------------------------------------------------------------------
        // Provide button data for ribbon integration
        // -------------------------------------------------------------------------
        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "ElementView",
                "Element View",
                "HDR_EMEA.Tools_Core.ElementView",
                Properties.Resources.ElementView_32,   // largeImage (32x32)
                Properties.Resources.ElementView_16,   // smallImage (16x16)
                "Create Cross/Elevation/Callout for picked elements; crop tightly using mm padding/depth; apply per‑type templates and scale; place on chosen title block sheet");
        }
    }
}