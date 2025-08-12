using System;
using System.Linq;
using System.Collections.Generic;
using System.Globalization;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

using WinForms = System.Windows.Forms;

namespace HDR_EMEA.Tools_Structural
{
    [Transaction(TransactionMode.Manual)]
    public class ElementSplitter : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // 1) Select elements (filter to walls + structural columns)
                IList<Reference> picked;
                var preselectedIds = uidoc.Selection.GetElementIds();
                if (preselectedIds != null && preselectedIds.Count > 0)
                {
                    var allowed = preselectedIds
                        .Select(id => doc.GetElement(id))
                        .Where(IsWallOrStructuralColumn)
                        .Select(e => e.Id)
                        .ToList();

                    if (allowed.Count == 0)
                    {
                        TaskDialog.Show("Element Splitter", "No structural walls or columns found in the current selection. Please pick elements.");
                        picked = uidoc.Selection.PickObjects(ObjectType.Element, new WallColumnSelectionFilter(), "Select structural walls and/or columns to split");
                    }
                    else
                    {
                        uidoc.Selection.SetElementIds(allowed);
                        picked = allowed.Select(id => new Reference(doc.GetElement(id))).ToList();
                    }
                }
                else
                {
                    picked = uidoc.Selection.PickObjects(ObjectType.Element, new WallColumnSelectionFilter(), "Select structural walls and/or columns to split");
                }

                if (picked == null || picked.Count == 0)
                    return Result.Cancelled;

                var elems = picked
                    .Select(r => doc.GetElement(r.ElementId))
                    .Where(IsWallOrStructuralColumn)
                    .Distinct(new ElementIdComparer())
                    .ToList();

                if (elems.Count == 0)
                {
                    TaskDialog.Show("Element Splitter", "Nothing to process after filtering. Only structural columns and walls are supported.");
                    return Result.Cancelled;
                }

                // 2) Get levels and show dialog
                var levels = new FilteredElementCollector(doc)
                    .OfClass(typeof(Level))
                    .Cast<Level>()
                    .OrderBy(l => l.Elevation)
                    .ToList();

                if (levels.Count == 0)
                {
                    TaskDialog.Show("Element Splitter", "This model contains no levels.");
                    return Result.Cancelled;
                }

                var dlg = new LevelPickerWindow(levels);
                bool? ok = dlg.ShowDialog();
                if (ok != true || dlg.SelectedLevels.Count == 0)
                    return Result.Cancelled;

                // Offset entered in mm; convert to internal feet
                double offsetFeet = UnitUtils.ConvertToInternalUnits(dlg.OffsetMillimetres, UnitTypeId.Millimeters);

                // Quick lookups
                var levelById = levels.ToDictionary(l => l.Id, l => l);
                var allLevelsByElev = levels.Select(l => (level: l, z: l.Elevation)).OrderBy(t => t.z).ToList();

                // 3) Split
                var failures = new List<string>();
                using (var tg = new TransactionGroup(doc, "Split elements by levels"))
                {
                    tg.Start();

                    foreach (var e in elems)
                    {
                        try
                        {
                            if (e is Wall w)
                                SplitWall_LevelToLevel(doc, w, dlg.SelectedLevels, offsetFeet, levelById, allLevelsByElev);
                            else if (IsStructuralColumn(e))
                                SplitColumn(doc, (FamilyInstance)e, dlg.SelectedLevels, offsetFeet, levelById, allLevelsByElev);
                            else
                                failures.Add(FormatElem(doc, e) + " — unsupported type after filter.");
                        }
                        catch (Exception ex)
                        {
                            failures.Add(FormatElem(doc, e) + $" — {ex.Message}");
                        }
                    }

                    tg.Assimilate();
                }

                if (failures.Count > 0)
                {
                    TaskDialog.Show("Element Splitter", $"Completed with issues on {failures.Count} element(s):\n\n" + string.Join("\n", failures.Take(20)) + (failures.Count > 20 ? "\n…" : ""));
                }
                else
                {
                    TaskDialog.Show("Element Splitter", "Done.");
                }

                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.ToString();
                return Result.Failed;
            }
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "ElementSplitter",
                "Element Splitter",
                "HDR_EMEA.Tools_Structural.ElementSplitter",
                Properties.Resources.ElementSplitter_32,
                Properties.Resources.ElementSplitter_16,
                "Split structural walls and columns at selected levels with an optional offset."
            );
        }

        // ========= Updated WALL logic: true level-to-level segments =========

        private static void SplitWall_LevelToLevel(
            Document doc,
            Wall wall,
            IList<Level> selectedLevels,
            double offsetFeet,
            Dictionary<ElementId, Level> levelById,
            List<(Level level, double z)> allLevelsByElev)
        {
            if (wall.WallType == null || wall.CurtainGrid != null)
                throw new InvalidOperationException("Curtain or special wall types are not supported.");

            // Original extents and constraints
            GetWallExtents(doc, wall, levelById, out double zBase, out double zTop);
            GetWallConstraintInfo(wall, levelById,
                out Level baseLevelOrig, out double baseOffsetOrig,
                out bool hasTopConstraint, out Level topLevelOrig, out double topOffsetOrig);

            // Build the ordered cut planes with their associated levels (selected + offset)
            var cuts = selectedLevels
                .Select(l => new { Level = l, Z = l.Elevation + offsetFeet })
                .Where(x => x.Z > zBase + 1e-06 && x.Z < zTop - 1e-06)
                .GroupBy(x => x.Z) // de-dup on elevation in case of duplicate names
                .Select(g => g.First())
                .OrderBy(x => x.Z)
                .ToList();

            if (cuts.Count == 0)
                return;

            // Segments are: [zBase -> cut0], [cut0 -> cut1], ..., [cutN-1 -> cutN], [cutN -> zTop]
            // For each segment we apply proper base/top constraints:
            //  - First segment base = original base level/offset; top = cut0 level with offset
            //  - Middle segments base = lower cut level; top = upper cut level (both with offset)
            //  - Last segment top = original top (if constrained) or unconnected to zTop

            using (var t = new Transaction(doc, $"Split wall {wall.Id.IntegerValue} (level-to-level)"))
            {
                t.Start();

                // we reuse the original for the first segment
                Wall prototype = wall;

                int segmentCount = cuts.Count + 1;
                for (int i = 0; i < segmentCount; i++)
                {
                    double segStartZ = (i == 0) ? zBase : cuts[i - 1].Z;
                    double segEndZ = (i == cuts.Count) ? zTop : cuts[i].Z;

                    // guard against degenerate
                    if (segEndZ <= segStartZ + 1e-07)
                        continue;

                    Wall target = (i == 0)
                        ? prototype
                        : (Wall)doc.GetElement(ElementTransformUtils.CopyElement(doc, prototype.Id, XYZ.Zero).First());

                    // Base
                    Level baseLevel;
                    double baseOffset;
                    if (i == 0)
                    {
                        baseLevel = baseLevelOrig;
                        baseOffset = segStartZ - baseLevelOrig.Elevation;   // precise start
                    }
                    else
                    {
                        baseLevel = cuts[i - 1].Level;
                        baseOffset = segStartZ - baseLevel.Elevation;       // equals your offset
                    }

                    SetParamElementId(target, BuiltInParameter.WALL_BASE_CONSTRAINT, baseLevel?.Id ?? ElementId.InvalidElementId);
                    SetParamDouble(target, BuiltInParameter.WALL_BASE_OFFSET, baseOffset);

                    // Top
                    if (i == cuts.Count)
                    {
                        // last segment to original top
                        if (hasTopConstraint && topLevelOrig != null)
                        {
                            double topOffset = zTop - topLevelOrig.Elevation;
                            SetParamElementId(target, BuiltInParameter.WALL_HEIGHT_TYPE, topLevelOrig.Id);
                            SetParamDouble(target, BuiltInParameter.WALL_TOP_OFFSET, topOffset);
                            // Clear unconnected height
                            SetParamDouble(target, BuiltInParameter.WALL_USER_HEIGHT_PARAM, 0.0);
                        }
                        else
                        {
                            // unconnected for the remainder
                            SetParamElementId(target, BuiltInParameter.WALL_HEIGHT_TYPE, ElementId.InvalidElementId);
                            SetParamDouble(target, BuiltInParameter.WALL_USER_HEIGHT_PARAM, segEndZ - segStartZ);
                            SetParamDouble(target, BuiltInParameter.WALL_TOP_OFFSET, 0.0);
                        }
                    }
                    else
                    {
                        // to the next cut's level (associated)
                        Level topLevel = cuts[i].Level;
                        double topOffset = segEndZ - topLevel.Elevation;
                        SetParamElementId(target, BuiltInParameter.WALL_HEIGHT_TYPE, topLevel.Id);
                        SetParamDouble(target, BuiltInParameter.WALL_TOP_OFFSET, topOffset);
                        // Clear unconnected height
                        SetParamDouble(target, BuiltInParameter.WALL_USER_HEIGHT_PARAM, 0.0);
                    }
                }

                t.Commit();
            }
        }

        private static void GetWallConstraintInfo(
            Wall wall,
            Dictionary<ElementId, Level> levelById,
            out Level baseLevel, out double baseOffset,
            out bool hasTopConstraint, out Level topLevel, out double topOffset)
        {
            var baseLevelId = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId() ?? ElementId.InvalidElementId;
            baseLevel = (baseLevelId != ElementId.InvalidElementId && levelById.TryGetValue(baseLevelId, out var b)) ? b : null;
            baseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0.0;

            var topTypeId = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)?.AsElementId() ?? ElementId.InvalidElementId;
            hasTopConstraint = topTypeId != ElementId.InvalidElementId;

            if (hasTopConstraint && levelById.TryGetValue(topTypeId, out var tLevel))
            {
                topLevel = tLevel;
                topOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)?.AsDouble() ?? 0.0;
            }
            else
            {
                topLevel = null;
                topOffset = 0.0;
            }
        }

        // ========= Column logic (unchanged; already level-to-level) =========

        private static void SplitColumn(
            Document doc,
            FamilyInstance col,
            IList<Level> selectedLevels,
            double offsetFeet,
            Dictionary<ElementId, Level> levelById,
            List<(Level level, double z)> allLevelsByElev)
        {
            if (!IsStructuralColumn(col))
                throw new InvalidOperationException("Not a structural column.");

            if (col.Location is LocationCurve lc && lc.Curve is Line line)
            {
                XYZ dir = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
                if (Math.Abs(dir.Z) < 0.999)
                    throw new InvalidOperationException("Slanted columns are not supported.");
            }
            else if (!(col.Location is LocationPoint))
            {
                throw new InvalidOperationException("Unsupported column location.");
            }

            GetColumnExtents(doc, col, levelById, out double zBase, out double zTop);

            var cutZs = selectedLevels
                .Select(l => l.Elevation + offsetFeet)
                .Where(z => z > zBase + 1e-06 && z < zTop - 1e-06)
                .Distinct()
                .OrderBy(z => z)
                .ToList();

            if (cutZs.Count == 0)
                return;

            var segments = BuildSegments(zBase, zTop, cutZs);

            using (var t = new Transaction(doc, $"Split column {col.Id.IntegerValue}"))
            {
                t.Start();

                FamilyInstance first = col;

                for (int i = 0; i < segments.Count; i++)
                {
                    var (z0, z1) = segments[i];
                    FamilyInstance target = (i == 0) ? first : (FamilyInstance)doc.GetElement(ElementTransformUtils.CopyElement(doc, first.Id, XYZ.Zero).First());

                    var baseLevel = FindLevelAtOrBelow(allLevelsByElev, z0) ?? allLevelsByElev.First().level;
                    var topLevel = FindLevelAtOrBelow(allLevelsByElev, z1 - 1e-09) ?? baseLevel;

                    double baseOffset = z0 - baseLevel.Elevation;
                    double topOffset = z1 - topLevel.Elevation;

                    SetParamElementId(target, BuiltInParameter.FAMILY_BASE_LEVEL_PARAM, baseLevel.Id);
                    SetParamElementId(target, BuiltInParameter.FAMILY_TOP_LEVEL_PARAM, topLevel.Id);
                    SetParamDouble(target, BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM, baseOffset);
                    SetParamDouble(target, BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM, topOffset);
                }

                t.Commit();
            }
        }

        private static List<(double z0, double z1)> BuildSegments(double zBase, double zTop, List<double> cutsSorted)
        {
            var segments = new List<(double z0, double z1)>();
            double start = zBase;
            foreach (var z in cutsSorted)
            {
                segments.Add((start, z));
                start = z;
            }
            segments.Add((start, zTop));
            segments = segments.Where(s => s.z1 - s.z0 > 1e-06).ToList();
            return segments;
        }

        private static void GetWallExtents(Document doc, Wall wall, Dictionary<ElementId, Level> levelById, out double zBase, out double zTop)
        {
            var pBaseLevelId = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT)?.AsElementId() ?? ElementId.InvalidElementId;
            var pBaseOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET)?.AsDouble() ?? 0.0;

            double baseElev = 0.0;
            if (levelById.TryGetValue(pBaseLevelId, out var baseLevel))
                baseElev = baseLevel.Elevation;

            zBase = baseElev + pBaseOffset;

            var topConstraintId = wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE)?.AsElementId() ?? ElementId.InvalidElementId;
            if (topConstraintId == ElementId.InvalidElementId)
            {
                double h = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM)?.AsDouble() ?? 0.0;
                zTop = zBase + h;
            }
            else
            {
                double topOffset = wall.get_Parameter(BuiltInParameter.WALL_TOP_OFFSET)?.AsDouble() ?? 0.0;
                double topElev = 0.0;
                if (levelById.TryGetValue(topConstraintId, out var topLevel))
                    topElev = topLevel.Elevation;
                zTop = topElev + topOffset;
            }

            if (zTop < zBase)
            {
                var tmp = zBase; zBase = zTop; zTop = tmp;
            }
        }

        private static void GetColumnExtents(Document doc, FamilyInstance col, Dictionary<ElementId, Level> levelById, out double zBase, out double zTop)
        {
            var baseLevelId = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM)?.AsElementId() ?? ElementId.InvalidElementId;
            var topLevelId = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_PARAM)?.AsElementId() ?? ElementId.InvalidElementId;

            double baseElev = 0.0, topElev = 0.0;
            if (levelById.TryGetValue(baseLevelId, out var baseLevel)) baseElev = baseLevel.Elevation;
            if (levelById.TryGetValue(topLevelId, out var topLevel)) topElev = topLevel.Elevation;

            double baseOffset = col.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_OFFSET_PARAM)?.AsDouble() ?? 0.0;
            double topOffset = col.get_Parameter(BuiltInParameter.FAMILY_TOP_LEVEL_OFFSET_PARAM)?.AsDouble() ?? 0.0;

            zBase = baseElev + baseOffset;
            zTop = topElev + topOffset;

            if (zTop < zBase)
            {
                var tmp = zBase; zBase = zTop; zTop = tmp;
            }
        }

        private static Level FindLevelAtOrBelow(List<(Level level, double z)> levelsByElev, double zQuery)
        {
            Level candidate = null;
            foreach (var (level, z) in levelsByElev)
            {
                if (z <= zQuery + 1e-09)
                    candidate = level;
                else
                    break;
            }
            return candidate ?? levelsByElev.First().level;
        }

        private static void SetParamDouble(Element e, BuiltInParameter bip, double value)
        {
            var p = e.get_Parameter(bip);
            if (p == null || p.StorageType != StorageType.Double || p.IsReadOnly) return;
            p.Set(value);
        }

        private static void SetParamElementId(Element e, BuiltInParameter bip, ElementId id)
        {
            var p = e.get_Parameter(bip);
            if (p == null || p.StorageType != StorageType.ElementId || p.IsReadOnly) return;
            p.Set(id);
        }

        private static bool IsWallOrStructuralColumn(Element e)
            => (e is Wall) || IsStructuralColumn(e);

        private static bool IsStructuralColumn(Element e)
        {
            if (e is FamilyInstance fi)
            {
                return fi.Category != null &&
                       fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns;
            }
            return false;
        }

        private static string FormatElem(Document doc, Element e)
        {
            try
            {
                return $"{e.Category?.Name ?? "Category?"} • {e.Name} [Id {e.Id.IntegerValue}]";
            }
            catch
            {
                return $"[Id {e.Id.IntegerValue}]";
            }
        }

        private class ElementIdComparer : IEqualityComparer<Element>
        {
            public bool Equals(Element x, Element y) => x?.Id == y?.Id;
            public int GetHashCode(Element obj) => obj.Id.IntegerValue.GetHashCode();
        }

        private class WallColumnSelectionFilter : ISelectionFilter
        {
            public bool AllowElement(Element elem) => IsWallOrStructuralColumn(elem);
            public bool AllowReference(Reference reference, XYZ position) => true;
        }
    }

    // ---------- UI (same improved dialog as last version) ----------
    internal class LevelPickerWindow : System.Windows.Window
    {
        private readonly List<Level> _allLevels;
        private readonly System.Windows.Controls.ListView _listView;
        private readonly System.Windows.Controls.TextBox _searchBox;
        private readonly System.Windows.Controls.TextBox _offsetBox;

        public List<Level> SelectedLevels { get; } = new List<Level>();
        public double OffsetMillimetres { get; private set; } = 0.0;

        public LevelPickerWindow(IList<Level> levels)
        {
            _allLevels = new List<Level>(levels);

            Title = "Split at Levels";
            Width = 520;
            Height = 560;
            WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner;
            ResizeMode = System.Windows.ResizeMode.CanResize;

            var root = new System.Windows.Controls.Grid { Margin = new System.Windows.Thickness(12) };
            root.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            root.ColumnDefinitions.Add(new System.Windows.Controls.ColumnDefinition { Width = new System.Windows.GridLength(210) });
            root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            root.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            Content = root;

            var leftPane = new System.Windows.Controls.Grid { Margin = new System.Windows.Thickness(0, 0, 12, 0) };
            leftPane.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = System.Windows.GridLength.Auto });
            leftPane.RowDefinitions.Add(new System.Windows.Controls.RowDefinition { Height = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            System.Windows.Controls.Grid.SetColumn(leftPane, 0);
            root.Children.Add(leftPane);

            var searchRow = new System.Windows.Controls.DockPanel { LastChildFill = true, Margin = new System.Windows.Thickness(0, 0, 0, 6) };
            var searchLbl = new System.Windows.Controls.TextBlock { Text = "Filter levels:", VerticalAlignment = System.Windows.VerticalAlignment.Center, Margin = new System.Windows.Thickness(0, 0, 6, 0) };
            _searchBox = new System.Windows.Controls.TextBox { MinWidth = 120 };
            _searchBox.TextChanged += (s, e) => RebuildList();
            searchRow.Children.Add(searchLbl);
            searchRow.Children.Add(_searchBox);
            System.Windows.Controls.Grid.SetRow(searchRow, 0);
            leftPane.Children.Add(searchRow);

            _listView = new System.Windows.Controls.ListView
            {
                SelectionMode = System.Windows.Controls.SelectionMode.Multiple
            };
            var gridView = new System.Windows.Controls.GridView();
            gridView.Columns.Add(new System.Windows.Controls.GridViewColumn
            {
                Header = "Select",
                Width = 60,
                CellTemplate = CreateCheckBoxTemplate()
            });
            gridView.Columns.Add(new System.Windows.Controls.GridViewColumn
            {
                Header = "Level",
                DisplayMemberBinding = new System.Windows.Data.Binding("Name"),
                Width = 180
            });
            gridView.Columns.Add(new System.Windows.Controls.GridViewColumn
            {
                Header = "Elevation (mm)",
                DisplayMemberBinding = new System.Windows.Data.Binding("ElevationMm"),
                Width = 120
            });
            _listView.View = gridView;
            System.Windows.Controls.Grid.SetRow(_listView, 1);
            leftPane.Children.Add(_listView);

            var rightPane = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Margin = new System.Windows.Thickness(0),
                VerticalAlignment = System.Windows.VerticalAlignment.Stretch
            };
            System.Windows.Controls.Grid.SetColumn(rightPane, 1);
            root.Children.Add(rightPane);

            var hint = new System.Windows.Controls.TextBlock
            {
                Text = "Choose levels to split at. Use a positive or negative offset.",
                TextWrapping = System.Windows.TextWrapping.Wrap,
                Margin = new System.Windows.Thickness(0, 0, 0, 8)
            };
            rightPane.Children.Add(hint);

            var offsetLabel = new System.Windows.Controls.TextBlock
            {
                Text = "Offset (mm):",
                Margin = new System.Windows.Thickness(0, 6, 0, 4)
            };
            rightPane.Children.Add(offsetLabel);

            var offsetRow = new System.Windows.Controls.DockPanel { LastChildFill = true };
            _offsetBox = new System.Windows.Controls.TextBox { Width = 90, Text = "0" };
            _offsetBox.PreviewTextInput += (s, e) => { e.Handled = !IsValidNumericInput(e.Text); };
            System.Windows.DataObject.AddPastingHandler(_offsetBox, (s, e) =>
            {
                if (e.DataObject.GetDataPresent(System.Windows.DataFormats.Text))
                {
                    var text = e.DataObject.GetData(System.Windows.DataFormats.Text) as string;
                    if (!IsValidNumericPaste(text)) e.CancelCommand();
                }
                else e.CancelCommand();
            });
            offsetRow.Children.Add(_offsetBox);

            var stepPanel = new System.Windows.Controls.WrapPanel { Margin = new System.Windows.Thickness(6, 0, 0, 0) };
            stepPanel.Children.Add(MakeStepBtn("-100", -100));
            stepPanel.Children.Add(MakeStepBtn("-10", -10));
            stepPanel.Children.Add(MakeStepBtn("-1", -1));
            stepPanel.Children.Add(MakeStepBtn("+1", +1));
            stepPanel.Children.Add(MakeStepBtn("+10", +10));
            stepPanel.Children.Add(MakeStepBtn("+100", +100));
            offsetRow.Children.Add(stepPanel);
            rightPane.Children.Add(offsetRow);

            var actionsLbl = new System.Windows.Controls.TextBlock
            {
                Text = "Selection actions:",
                Margin = new System.Windows.Thickness(0, 12, 0, 4)
            };
            rightPane.Children.Add(actionsLbl);

            var actions = new System.Windows.Controls.StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal };
            var btnSelectVisible = new System.Windows.Controls.Button { Content = "Select visible", Margin = new System.Windows.Thickness(0, 0, 6, 0) };
            var btnClear = new System.Windows.Controls.Button { Content = "Clear", Margin = new System.Windows.Thickness(0, 0, 6, 0) };
            var btnInvert = new System.Windows.Controls.Button { Content = "Invert" };

            btnSelectVisible.Click += (s, e) => SetAllVisible(true);
            btnClear.Click += (s, e) => SetAllVisible(false);
            btnInvert.Click += (s, e) => InvertVisible();

            actions.Children.Add(btnSelectVisible);
            actions.Children.Add(btnClear);
            actions.Children.Add(btnInvert);
            rightPane.Children.Add(actions);

            var footer = new System.Windows.Controls.StackPanel
            {
                Orientation = System.Windows.Controls.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Right,
                Margin = new System.Windows.Thickness(0, 12, 0, 0)
            };
            var okBtn = new System.Windows.Controls.Button { Content = "OK", IsDefault = true, Margin = new System.Windows.Thickness(0, 0, 6, 0), Width = 88 };
            var cancelBtn = new System.Windows.Controls.Button { Content = "Cancel", IsCancel = true, Width = 88 };
            okBtn.Click += OkClick;
            footer.Children.Add(okBtn);
            footer.Children.Add(cancelBtn);

            System.Windows.Controls.Grid.SetRow(footer, 1);
            System.Windows.Controls.Grid.SetColumnSpan(footer, 2);
            root.Children.Add(footer);

            RebuildList();
        }

        private System.Windows.DataTemplate CreateCheckBoxTemplate()
        {
            var fef = new System.Windows.FrameworkElementFactory(typeof(System.Windows.Controls.CheckBox));
            fef.SetBinding(System.Windows.Controls.CheckBox.IsCheckedProperty, new System.Windows.Data.Binding("Checked") { Mode = System.Windows.Data.BindingMode.TwoWay });
            var template = new System.Windows.DataTemplate { VisualTree = fef };
            return template;
        }

        private class LevelRow
        {
            public bool Checked { get; set; }
            public string Name { get; set; }
            public string ElevationMm { get; set; }
            public Level Level { get; set; }
        }

        private void RebuildList()
        {
            string filter = _searchBox?.Text?.Trim() ?? string.Empty;
            var items = new List<LevelRow>();
            foreach (var lv in _allLevels)
            {
                if (!string.IsNullOrEmpty(filter) && lv.Name?.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                double mm = Autodesk.Revit.DB.UnitUtils.ConvertFromInternalUnits(lv.Elevation, Autodesk.Revit.DB.UnitTypeId.Millimeters);
                items.Add(new LevelRow
                {
                    Checked = false,
                    Name = lv.Name,
                    ElevationMm = $"{mm:0.##}",
                    Level = lv
                });
            }
            _listView.ItemsSource = items;
        }

        private System.Windows.Controls.Button MakeStepBtn(string label, int stepMm)
        {
            var b = new System.Windows.Controls.Button { Content = label, Margin = new System.Windows.Thickness(2, 0, 0, 0), Width = 44 };
            b.Click += (s, e) =>
            {
                if (!double.TryParse(_offsetBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double val))
                {
                    if (!double.TryParse(_offsetBox.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out val))
                        val = 0;
                }
                val += stepMm;
                _offsetBox.Text = val.ToString(CultureInfo.InvariantCulture);
            };
            return b;
        }

        private void SetAllVisible(bool state)
        {
            if (_listView.ItemsSource is System.Collections.IEnumerable rows)
            {
                foreach (var r in rows)
                {
                    if (r is LevelRow lr) lr.Checked = state;
                }
                _listView.Items.Refresh();
            }
        }

        private void InvertVisible()
        {
            if (_listView.ItemsSource is System.Collections.IEnumerable rows)
            {
                foreach (var r in rows)
                {
                    if (r is LevelRow lr) lr.Checked = !lr.Checked;
                }
                _listView.Items.Refresh();
            }
        }

        private void OkClick(object sender, System.Windows.RoutedEventArgs e)
        {
            SelectedLevels.Clear();

            if (_listView.ItemsSource is System.Collections.IEnumerable rows)
            {
                foreach (var r in rows)
                {
                    if (r is LevelRow lr && lr.Checked && lr.Level != null)
                        SelectedLevels.Add(lr.Level);
                }
            }

            if (SelectedLevels.Count == 0)
            {
                System.Windows.MessageBox.Show("Please select at least one level.", "Element Splitter", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
                return;
            }

            if (!double.TryParse(_offsetBox.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out double mm))
            {
                if (!double.TryParse(_offsetBox.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out mm))
                {
                    System.Windows.MessageBox.Show("Offset is not a valid number.", "Element Splitter", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }
            }

            OffsetMillimetres = mm;
            DialogResult = true;
            Close();
        }

        private static bool IsValidNumericInput(string s)
        {
            foreach (char c in s)
            {
                if (!(char.IsDigit(c) || c == '-' || c == '.' || c == ',')) return false;
            }
            return true;
        }

        private static bool IsValidNumericPaste(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return false;
            s = s.Trim();
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out _) ||
                   double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out _);
        }
    }
}
