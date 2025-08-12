using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.UI;
using WinForms = System.Windows.Forms;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class FamilyRenamer : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var uiapp = commandData.Application;
            var uidoc = uiapp.ActiveUIDocument;
            var doc = uidoc.Document;

            var allTypes = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .ToElements();

            var include = new HashSet<Type>
            {
                typeof(FamilySymbol),
                typeof(WallType),
                typeof(FloorType),
                typeof(CeilingType),
                typeof(RoofType),
                typeof(FilledRegionType),
                typeof(TextNoteType),
                typeof(AnnotationSymbolType),
                typeof(DimensionType),
                typeof(SpotDimensionType),
                typeof(GridType),
                typeof(CurtainSystemType),
                typeof(MullionType),
                typeof(GroupType),
                typeof(FlexPipeType),
                typeof(FlexDuctType),
                typeof(RailingType),
                typeof(HandRailType),
                typeof(CableTrayType),
                typeof(ConduitType),
                typeof(DuctSystemType),
                typeof(DuctType),
                typeof(MechanicalSystemType),
                typeof(DuctInsulationType),
                typeof(PipingSystemType),
                typeof(PipeInsulationType),
                typeof(PipeType),
                typeof(StairsType),
                typeof(BeamSystemType)
            };

            var filtered = allTypes.Where(t => include.Contains(t.GetType())).ToList();
            if (filtered.Count == 0)
            {
                WinForms.MessageBox.Show("No supported element types were found in this model.", "Family Renamer");
                return Result.Cancelled;
            }

            var items = filtered.Select(t =>
            {
                var famName = TryGetFamilyName(t, out bool isLoadable, out ElementId famId, out ElementId famCatId);
                return new TypeItem
                {
                    Id = t.Id,
                    ClassName = t.GetType().Name,
                    FamilyName = famName,
                    FamilyId = famId,
                    FamilyCategoryId = famCatId,
                    CurrentName = GetTypeName(t),
                    IsLoadable = isLoadable,
                    Selected = false
                };
            })
            .OrderBy(i => i.FamilyName, StringComparer.OrdinalIgnoreCase)
            .ThenBy(i => i.CurrentName, StringComparer.OrdinalIgnoreCase)
            .ToList();

            using (var dlg = new FamilyRenamerForm(items))
            {
                if (dlg.ShowDialog() != WinForms.DialogResult.OK)
                    return Result.Cancelled;

                var find = dlg.FindText ?? string.Empty;
                var replace = dlg.ReplaceText ?? string.Empty;
                var prefix = dlg.PrefixText ?? string.Empty;
                var suffix = dlg.SuffixText ?? string.Empty;
                var caseSensitive = dlg.CaseSensitive;
                var mode = dlg.RenameMode;

                var selected = dlg.GetAllItems().Where(x => x.Selected).ToList();
                if (selected.Count == 0)
                {
                    WinForms.MessageBox.Show("No types selected.", "Family Renamer");
                    return Result.Cancelled;
                }

                var typeNamesByClass = BuildTypeNamesCache(doc, selected.Select(s => s.Id));
                var familyNamesByCat = BuildFamilyNamesCache(doc);

                var familiesRenamed = new Dictionary<ElementId, string>();

                int renamedTypes = 0, renamedFamilies = 0, unchanged = 0, conflicts = 0, errors = 0, unsupported = 0;

                using (var tx = new Transaction(doc, "Rename Family/Type Names"))
                {
                    tx.Start();

                    foreach (var it in selected)
                    {
                        try
                        {
                            var elType = doc.GetElement(it.Id) as ElementType;
                            if (elType == null) { errors++; continue; }

                            if (mode == RenameTarget.TypeOnly || mode == RenameTarget.Both)
                            {
                                string currentTypeName = GetTypeName(elType);
                                string proposedType = ComputeNewName(currentTypeName, find, replace, prefix, suffix, caseSensitive);

                                if (!string.Equals(currentTypeName, proposedType, StringComparison.Ordinal))
                                {
                                    var classKey = elType.GetType().FullName;
                                    var existingTypeNames = typeNamesByClass[classKey];
                                    var uniqueType = EnsureUnique(proposedType, existingTypeNames);

                                    try
                                    {
                                        elType.Name = uniqueType;
                                        existingTypeNames.Add(uniqueType);
                                        renamedTypes++;
                                    }
                                    catch
                                    {
                                        conflicts++;
                                    }
                                }
                                else
                                {
                                    unchanged++;
                                }
                            }

                            if ((mode == RenameTarget.FamilyOnly || mode == RenameTarget.Both) && it.IsLoadable && it.FamilyId != ElementId.InvalidElementId)
                            {
                                if (!familiesRenamed.ContainsKey(it.FamilyId))
                                {
                                    var fam = doc.GetElement(it.FamilyId) as Family;
                                    if (fam == null) { unsupported++; continue; }

                                    string currentFamilyName = fam.Name ?? string.Empty;
                                    string proposedFamily = ComputeNewName(currentFamilyName, find, replace, prefix, suffix, caseSensitive);

                                    if (!string.Equals(currentFamilyName, proposedFamily, StringComparison.Ordinal))
                                    {
                                        HashSet<string> nameSet;
                                        var catKey = it.FamilyCategoryId.IntegerValue;
                                        if (!familyNamesByCat.TryGetValue(catKey, out nameSet))
                                        {
                                            if (!familyNamesByCat.TryGetValue(int.MinValue, out nameSet))
                                            {
                                                nameSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                                familyNamesByCat[int.MinValue] = nameSet;
                                            }
                                        }

                                        var uniqueFamily = EnsureUnique(proposedFamily, nameSet);
                                        try
                                        {
                                            fam.Name = uniqueFamily;
                                            nameSet.Add(uniqueFamily);
                                            familiesRenamed[it.FamilyId] = uniqueFamily;
                                            renamedFamilies++;
                                        }
                                        catch
                                        {
                                            conflicts++;
                                        }
                                    }
                                    else
                                    {
                                        unchanged++;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            errors++;
                        }
                    }

                    tx.Commit();
                }

                WinForms.MessageBox.Show(
                    $"Types renamed: {renamedTypes}\nFamilies renamed: {renamedFamilies}\nUnchanged: {unchanged}\nConflicts/invalid: {conflicts}\nUnsupported: {unsupported}\nErrors: {errors}",
                    "Family Renamer");

                return Result.Succeeded;
            }
        }

        private static Dictionary<string, HashSet<string>> BuildTypeNamesCache(Document doc, IEnumerable<ElementId> selectionIds)
        {
            var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            var universe = new FilteredElementCollector(doc).WhereElementIsElementType().ToElements();

            var classesNeeded = new HashSet<Type>();
            foreach (var id in selectionIds)
            {
                var e = doc.GetElement(id);
                if (e != null) classesNeeded.Add(e.GetType());
            }

            foreach (var cls in classesNeeded)
            {
                var names = universe
                    .Where(e => e.GetType() == cls)
                    .Select(GetTypeName)
                    .Where(n => !string.IsNullOrWhiteSpace(n));
                result[cls.FullName] = new HashSet<string>(names, StringComparer.OrdinalIgnoreCase);
            }

            return result;
        }

        private static Dictionary<int, HashSet<string>> BuildFamilyNamesCache(Document doc)
        {
            var dict = new Dictionary<int, HashSet<string>>();

            var allFamilies = new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .ToList();

            foreach (var grp in allFamilies.GroupBy(f => f.FamilyCategory?.Id.IntegerValue ?? int.MinValue))
            {
                dict[grp.Key] = new HashSet<string>(
                    grp.Select(f => f.Name).Where(n => !string.IsNullOrWhiteSpace(n)),
                    StringComparer.OrdinalIgnoreCase);
            }

            return dict;
        }

        private static string TryGetFamilyName(Element t, out bool isLoadable, out ElementId familyId, out ElementId familyCategoryId)
        {
            isLoadable = false;
            familyId = ElementId.InvalidElementId;
            familyCategoryId = ElementId.InvalidElementId;

            try
            {
                if (t is FamilySymbol fs && fs.Family != null)
                {
                    isLoadable = true;
                    familyId = fs.Family.Id;
                    familyCategoryId = fs.Family.FamilyCategory?.Id ?? ElementId.InvalidElementId;
                    return fs.Family.Name ?? string.Empty;
                }
            }
            catch { }

            try
            {
                var famNameProp = t.GetType().GetProperty("FamilyName");
                if (famNameProp != null)
                {
                    var v = famNameProp.GetValue(t) as string;
                    if (!string.IsNullOrEmpty(v)) return v;
                }
            }
            catch { }

            return string.Empty;
        }

        private static string GetTypeName(Element t)
        {
            var p = t.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME);
            if (p != null && p.StorageType == StorageType.String)
            {
                var v = p.AsString();
                if (!string.IsNullOrEmpty(v)) return v;
            }
            return t.Name ?? string.Empty;
        }

        private static string ComputeNewName(string current, string find, string repl, string prefix, string suffix, bool caseSensitive)
        {
            string replaced = current;

            if (!string.IsNullOrEmpty(find))
            {
                if (caseSensitive)
                {
                    replaced = current.Replace(find, repl ?? string.Empty);
                }
                else
                {
                    var pattern = Regex.Escape(find);
                    replaced = Regex.Replace(current, pattern, repl ?? string.Empty, RegexOptions.IgnoreCase);
                }
            }

            return (prefix ?? string.Empty) + replaced + (suffix ?? string.Empty);
        }

        private static string EnsureUnique(string baseName, HashSet<string> existing)
        {
            if (!existing.Contains(baseName))
                return baseName;

            int i = 1;
            while (true)
            {
                var candidate = $"{baseName} ({i})";
                if (!existing.Contains(candidate))
                    return candidate;
                i++;
            }
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "FamilyRenamer",
                "Family Renamer",
                "HDR_EMEA.Tools_Core.FamilyRenamer",
                Properties.Resources.FamilyRenamer_32,
                Properties.Resources.FamilyRenamer_16,
                "Find/Replace, prefix and suffix across multiple family and system types. Choose whether to rename Types, Families, or both."
            );
        }
    }

    internal enum RenameTarget
    {
        TypeOnly,
        FamilyOnly,
        Both
    }

    internal class TypeItem
    {
        public ElementId Id { get; set; }
        public string ClassName { get; set; }
        public string FamilyName { get; set; }
        public ElementId FamilyId { get; set; }
        public ElementId FamilyCategoryId { get; set; }
        public string CurrentName { get; set; }
        public bool IsLoadable { get; set; }

        public bool Selected { get; set; }
        public string PreviewType { get; set; }
        public string PreviewFamily { get; set; }

        public override string ToString() => $"{FamilyName} : {CurrentName}";
    }

    internal class FamilyRenamerForm : WinForms.Form
    {
        private readonly List<TypeItem> _items;

        private readonly WinForms.ListView _list = new WinForms.ListView();
        private readonly WinForms.TextBox _txtFind = new WinForms.TextBox();
        private readonly WinForms.TextBox _txtReplace = new WinForms.TextBox();
        private readonly WinForms.TextBox _txtPrefix = new WinForms.TextBox();
        private readonly WinForms.TextBox _txtSuffix = new WinForms.TextBox();
        private readonly WinForms.CheckBox _chkCase = new WinForms.CheckBox();
        private readonly WinForms.TextBox _txtFilter = new WinForms.TextBox();
        private readonly WinForms.Button _btnAll = new WinForms.Button();
        private readonly WinForms.Button _btnNone = new WinForms.Button();
        private readonly WinForms.Button _btnOk = new WinForms.Button();
        private readonly WinForms.Button _btnCancel = new WinForms.Button();

        private readonly WinForms.GroupBox _grpTarget = new WinForms.GroupBox();
        private readonly WinForms.RadioButton _rbTypeOnly = new WinForms.RadioButton();
        private readonly WinForms.RadioButton _rbFamilyOnly = new WinForms.RadioButton();
        private readonly WinForms.RadioButton _rbBoth = new WinForms.RadioButton();

        private readonly WinForms.ListView _preview = new WinForms.ListView();

        public string FindText => _txtFind.Text;
        public string ReplaceText => _txtReplace.Text;
        public string PrefixText => _txtPrefix.Text;
        public string SuffixText => _txtSuffix.Text;
        public bool CaseSensitive => _chkCase.Checked;
        public RenameTarget RenameMode
            => _rbBoth.Checked ? RenameTarget.Both
             : _rbFamilyOnly.Checked ? RenameTarget.FamilyOnly
             : RenameTarget.TypeOnly;

        public FamilyRenamerForm(List<TypeItem> items)
        {
            _items = items;

            Text = "Family Renamer";
            StartPosition = WinForms.FormStartPosition.CenterParent;
            Width = 1050;
            Height = 640;
            MinimumSize = new System.Drawing.Size(900, 520);
            FormBorderStyle = WinForms.FormBorderStyle.Sizable;
            MaximizeBox = true;
            MinimizeBox = true;

            var lblFilter = new WinForms.Label { Text = "Filter:", Left = 10, Top = 12, Width = 40 };
            _txtFilter.Left = 55; _txtFilter.Top = 8; _txtFilter.Width = 250;

            var lblFind = new WinForms.Label { Text = "Find:", Left = 315, Top = 12, Width = 35 };
            _txtFind.Left = 350; _txtFind.Top = 8; _txtFind.Width = 120;

            var lblReplace = new WinForms.Label { Text = "Replace:", Left = 480, Top = 12, Width = 60 };
            _txtReplace.Left = 540; _txtReplace.Top = 8; _txtReplace.Width = 130;

            _chkCase.Text = "Case sensitive";
            _chkCase.Left = 680; _chkCase.Top = 10; _chkCase.Width = 120;

            var lblPrefix = new WinForms.Label { Text = "Prefix:", Left = 810, Top = 12, Width = 45 };
            _txtPrefix.Left = 855; _txtPrefix.Top = 8; _txtPrefix.Width = 80;

            var lblSuffix = new WinForms.Label { Text = "Suffix:", Left = 940, Top = 12, Width = 45 };
            _txtSuffix.Left = 985; _txtSuffix.Top = 8; _txtSuffix.Width = 50;

            _grpTarget.Text = "Rename";
            _grpTarget.Left = 10; _grpTarget.Top = 36; _grpTarget.Width = 390; _grpTarget.Height = 42;

            _rbTypeOnly.Text = "Type names";
            _rbTypeOnly.Left = 10; _rbTypeOnly.Top = 16; _rbTypeOnly.Checked = true;

            _rbFamilyOnly.Text = "Family names (loadable only)";
            _rbFamilyOnly.Left = 110; _rbFamilyOnly.Top = 16;

            _rbBoth.Text = "Both where possible";
            _rbBoth.Left = 295; _rbBoth.Top = 16;

            _grpTarget.Controls.AddRange(new WinForms.Control[] { _rbTypeOnly, _rbFamilyOnly, _rbBoth });

            _btnAll.Text = "Select all";
            _btnAll.Left = 410; _btnAll.Top = 42; _btnAll.Width = 90;
            _btnNone.Text = "Select none";
            _btnNone.Left = 510; _btnNone.Top = 42; _btnNone.Width = 90;

            // Left list (resizable)
            _list.Left = 10; _list.Top = 84; _list.Width = 520; _list.Height = 460;
            _list.View = WinForms.View.Details;
            _list.CheckBoxes = true;
            _list.FullRowSelect = true;
            _list.DoubleBuffered(true);
            _list.Columns.Add("Family", 220);
            _list.Columns.Add("Type", 220);
            _list.Columns.Add("Class", 70);

            // Right preview (resizable)
            _preview.Left = 540; _preview.Top = 84; _preview.Width = 490; _preview.Height = 460;
            _preview.View = WinForms.View.Details;
            _preview.FullRowSelect = true;
            _preview.DoubleBuffered(true);
            _preview.Columns.Add("Current Type", 200);
            _preview.Columns.Add("Preview Type", 140);
            _preview.Columns.Add("Preview Family", 130);

            _btnOk.Text = "OK";
            _btnOk.Left = 870; _btnOk.Top = 555; _btnOk.Width = 75;
            _btnCancel.Text = "Cancel";
            _btnCancel.Left = 955; _btnCancel.Top = 555; _btnCancel.Width = 75;

            Controls.AddRange(new WinForms.Control[]
            {
                lblFilter, _txtFilter, lblFind, _txtFind, lblReplace, _txtReplace, _chkCase,
                lblPrefix, _txtPrefix, lblSuffix, _txtSuffix, _grpTarget,
                _btnAll, _btnNone, _list, _preview, _btnOk, _btnCancel
            });

            // Anchors for resize behaviour
            lblFilter.Anchor = _txtFilter.Anchor = lblFind.Anchor = _txtFind.Anchor =
            lblReplace.Anchor = _txtReplace.Anchor = _chkCase.Anchor =
            lblPrefix.Anchor = _txtPrefix.Anchor = lblSuffix.Anchor = _txtSuffix.Anchor =
                WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left;

            _grpTarget.Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left;
            _btnAll.Anchor = _btnNone.Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Left;

            _list.Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Left | WinForms.AnchorStyles.Right;
            _preview.Anchor = WinForms.AnchorStyles.Top | WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;

            _btnOk.Anchor = _btnCancel.Anchor = WinForms.AnchorStyles.Bottom | WinForms.AnchorStyles.Right;

            // Wire up events
            _btnOk.DialogResult = WinForms.DialogResult.OK;
            _btnCancel.DialogResult = WinForms.DialogResult.Cancel;
            AcceptButton = _btnOk;
            CancelButton = _btnCancel;

            _btnAll.Click += (s, e) => SetAllFiltered(true);
            _btnNone.Click += (s, e) => SetAllFiltered(false);

            _txtFind.TextChanged += (s, e) => RefreshPreview();
            _txtReplace.TextChanged += (s, e) => RefreshPreview();
            _txtPrefix.TextChanged += (s, e) => RefreshPreview();
            _txtSuffix.TextChanged += (s, e) => RefreshPreview();
            _chkCase.CheckedChanged += (s, e) => RefreshPreview();
            _rbTypeOnly.CheckedChanged += (s, e) => RefreshPreview();
            _rbFamilyOnly.CheckedChanged += (s, e) => RefreshPreview();
            _rbBoth.CheckedChanged += (s, e) => RefreshPreview();
            _txtFilter.TextChanged += (s, e) => PopulateList();

            _list.ItemCheck += (s, e) =>
            {
                BeginInvoke((Action)(() =>
                {
                    if (e.Index >= 0 && e.Index < _list.Items.Count)
                    {
                        var item = (TypeItem)_list.Items[e.Index].Tag;
                        item.Selected = _list.Items[e.Index].Checked;
                        RefreshPreview();
                    }
                }));
            };

            // Auto-resize columns when the window resizes
            SizeChanged += (s, e) =>
            {
                ResizeColumns(_list, 0.43, 0.43, 0.14);
                ResizeColumns(_preview, 0.42, 0.30, 0.28);
            };

            PopulateList();
            RefreshPreview();

            // Initial column sizing
            ResizeColumns(_list, 0.43, 0.43, 0.14);
            ResizeColumns(_preview, 0.42, 0.30, 0.28);
        }

        public IEnumerable<TypeItem> GetAllItems() => _items;

        private IEnumerable<TypeItem> VisibleItems()
        {
            foreach (WinForms.ListViewItem lvi in _list.Items)
                if (lvi.Tag is TypeItem ti) yield return ti;
        }

        private void PopulateList()
        {
            string filter = _txtFilter.Text?.Trim() ?? string.Empty;
            _list.BeginUpdate();
            _list.Items.Clear();

            foreach (var item in _items)
            {
                if (!string.IsNullOrEmpty(filter))
                {
                    var hay = $"{item.FamilyName} {item.CurrentName} {item.ClassName}";
                    if (hay.IndexOf(filter, StringComparison.OrdinalIgnoreCase) < 0)
                        continue;
                }

                var lvi = new WinForms.ListViewItem(item.FamilyName);
                lvi.SubItems.Add(item.CurrentName);
                lvi.SubItems.Add(item.ClassName);
                lvi.Tag = item;
                lvi.Checked = item.Selected;
                _list.Items.Add(lvi);
            }

            _list.EndUpdate();
        }

        private void SetAllFiltered(bool state)
        {
            _list.BeginUpdate();
            foreach (WinForms.ListViewItem lvi in _list.Items)
            {
                lvi.Checked = state;
                if (lvi.Tag is TypeItem ti) ti.Selected = state;
            }
            _list.EndUpdate();
            RefreshPreview();
        }

        private void RefreshPreview()
        {
            _preview.BeginUpdate();
            _preview.Items.Clear();

            string find = _txtFind.Text ?? string.Empty;
            string repl = _txtReplace.Text ?? string.Empty;
            string pre = _txtPrefix.Text ?? string.Empty;
            string suf = _txtSuffix.Text ?? string.Empty;
            bool cs = _chkCase.Checked;
            bool showFamilyPreview = _rbFamilyOnly.Checked || _rbBoth.Checked;

            foreach (var it in _items.Where(x => x.Selected))
            {
                string newType = ComputePreview(it.CurrentName, find, repl, pre, suf, cs);
                string newFam = showFamilyPreview && it.IsLoadable
                               ? ComputePreview(it.FamilyName ?? string.Empty, find, repl, pre, suf, cs)
                               : string.Empty;

                it.PreviewType = newType;
                it.PreviewFamily = newFam;

                var lvi = new WinForms.ListViewItem(it.CurrentName);
                lvi.SubItems.Add(newType);
                lvi.SubItems.Add(newFam);
                _preview.Items.Add(lvi);
            }

            _preview.EndUpdate();
        }

        private static string ComputePreview(string current, string find, string repl, string prefix, string suffix, bool caseSensitive)
        {
            string replaced = current;

            if (!string.IsNullOrEmpty(find))
            {
                if (caseSensitive)
                    replaced = current.Replace(find, repl ?? string.Empty);
                else
                {
                    var pattern = Regex.Escape(find);
                    replaced = Regex.Replace(current, pattern, repl ?? string.Empty, RegexOptions.IgnoreCase);
                }
            }

            return (prefix ?? string.Empty) + replaced + (suffix ?? string.Empty);
        }

        // Proportional auto-sizing for ListView columns
        private static void ResizeColumns(WinForms.ListView lv, params double[] proportions)
        {
            if (lv.Columns.Count == 0 || proportions.Length != lv.Columns.Count) return;

            // subtract some padding for scrollbars/borders
            int usable = Math.Max(0, lv.ClientSize.Width - 6);
            for (int i = 0; i < lv.Columns.Count; i++)
                lv.Columns[i].Width = (int)Math.Round(usable * proportions[i]);
        }
    }

    // Small helper to reduce flicker on ListView
    internal static class ListViewExtensions
    {
        public static void DoubleBuffered(this WinForms.ListView lv, bool enable)
        {
            try
            {
                var prop = typeof(WinForms.Control).GetProperty("DoubleBuffered",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                prop?.SetValue(lv, enable, null);
            }
            catch { }
        }
    }
}
