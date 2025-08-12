using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;                                 // WPF Window
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using HDR_EMEA.Common;
// Alias WPF controls to avoid WinForms ambiguity
using Wpf = System.Windows.Controls;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.Manual)]
    public class TextTransform : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData,
                              ref string message,
                              ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            if (uidoc == null)
            {
                TaskDialog.Show("Text Transform", "No active document.");
                return Result.Cancelled;
            }

            Document doc = uidoc.Document;

            // Gather initial selection; if none, prompt to pick
            var textNoteIds = GetSelectedTextNoteIds(uidoc);
            if (textNoteIds.Count == 0)
            {
                var td = new TaskDialog("Text Transform")
                {
                    MainInstruction = "No Text Notes selected.",
                    MainContent = "Select one or more Text Notes to transform.",
                    CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel,
                    DefaultButton = TaskDialogResult.Ok
                };
                if (td.Show() == TaskDialogResult.Ok)
                {
                    try
                    {
                        var picked = uidoc.Selection.PickObjects(ObjectType.Element,
                            new TextNoteSelectionFilter(), "Select Text Notes");
                        textNoteIds = picked.Select(r => r.ElementId).ToList();
                    }
                    catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                    {
                        return Result.Cancelled;
                    }
                }
                else
                    return Result.Cancelled;
            }

            var win = new TextTransformWindow(uidoc, textNoteIds);
            var helper = new System.Windows.Interop.WindowInteropHelper(win)
            {
                Owner = System.Diagnostics.Process.GetCurrentProcess().MainWindowHandle
            };
            win.ShowDialog();
            return Result.Succeeded;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            string tooltip =
                "Transform selected Text Notes:\n" +
                "• Upper / Lower / Title / Capitalise\n" +
                "• Swap case / Reverse / Trim\n" +
                "• Find & Replace with Prefix/Suffix, Case-sensitive, Whole-word\n\n" +
                "Tip: If nothing is selected, you will be prompted to pick Text Notes.";

            return new Common.ButtonDataClass(
                "TextTransform",
                "Text Transform",
                "HDR_EMEA.Tools_Core.TextTransform",
                Properties.Resources.TextTransform_32,
                Properties.Resources.TextTransform_16,
                tooltip
            );
        }

        private static List<ElementId> GetSelectedTextNoteIds(UIDocument uidoc)
        {
            var ids = uidoc.Selection.GetElementIds();
            if (ids == null || ids.Count == 0) return new List<ElementId>();

            var doc = uidoc.Document;
            return ids
                .Select(id => doc.GetElement(id))
                .OfType<TextNote>()
                .Select(tn => tn.Id)
                .ToList();
        }
    }

    // Public selection filter so it is accessible across classes
    public class TextNoteSelectionFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem) => elem is TextNote;
        public bool AllowReference(Reference reference, XYZ position) => false;
    }

    internal class TextTransformWindow : Window
    {
        private readonly UIDocument _uidoc;
        private readonly Document _doc;
        private List<ElementId> _textNoteIds;

        // Controls
        private readonly Wpf.TextBox _txtFind = new Wpf.TextBox();
        private readonly Wpf.TextBox _txtReplace = new Wpf.TextBox();
        private readonly Wpf.TextBox _txtPrefix = new Wpf.TextBox();
        private readonly Wpf.TextBox _txtSuffix = new Wpf.TextBox();
        private readonly Wpf.CheckBox _chkCaseSensitive = new Wpf.CheckBox();
        private readonly Wpf.CheckBox _chkWholeWord = new Wpf.CheckBox();
        private readonly Wpf.TextBlock _lblCount = new Wpf.TextBlock();

        public TextTransformWindow(UIDocument uidoc, List<ElementId> textNoteIds)
        {
            _uidoc = uidoc;
            _doc = uidoc.Document;
            _textNoteIds = textNoteIds != null ? new List<ElementId>(textNoteIds) : new List<ElementId>();

            Title = "Text Transform";
            WindowStartupLocation = WindowStartupLocation.CenterOwner;
            ResizeMode = ResizeMode.CanResize;
            MinWidth = 560;
            MinHeight = 340;
            SizeToContent = SizeToContent.WidthAndHeight;

            Content = BuildUi();
            UpdateCountLabel();
        }

        private FrameworkElement BuildUi()
        {
            var root = new Wpf.Grid { Margin = new Thickness(12) };

            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // header
            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // case buttons
            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // find/replace
            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // prefix/suffix
            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // options
            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // actions
            root.RowDefinitions.Add(new Wpf.RowDefinition { Height = System.Windows.GridLength.Auto }); // footer

            // Header
            var header = new Wpf.StackPanel { Orientation = Wpf.Orientation.Horizontal };
            var title = new Wpf.TextBlock
            {
                Text = "Transform selected Text Notes",
                FontSize = 16,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 0, 12, 8)
            };
            _lblCount.Margin = new Thickness(0, 0, 0, 8);
            header.Children.Add(title);
            header.Children.Add(_lblCount);
            root.Children.Add(header);
            Wpf.Grid.SetRow(header, 0);

            // Case/format buttons
            var btnPanel = new Wpf.WrapPanel { Margin = new Thickness(0, 4, 0, 8) };
            btnPanel.Children.Add(MakeButton("UPPER", OnUpper));
            btnPanel.Children.Add(MakeButton("lower", OnLower));
            btnPanel.Children.Add(MakeButton("Title Case", OnTitle));
            btnPanel.Children.Add(MakeButton("Capitalise", OnCapitalise));
            btnPanel.Children.Add(MakeButton("Swap case", OnSwapCase));
            btnPanel.Children.Add(MakeButton("Reverse", OnReverse));
            btnPanel.Children.Add(MakeButton("Trim", OnTrim));
            root.Children.Add(btnPanel);
            Wpf.Grid.SetRow(btnPanel, 1);

            // Find / Replace line
            var frGrid = new Wpf.Grid { Margin = new Thickness(0, 0, 0, 6) };
            frGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = System.Windows.GridLength.Auto });
            frGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            frGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = System.Windows.GridLength.Auto });
            frGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });

            frGrid.Children.Add(new Wpf.TextBlock { Text = "Find", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) });
            frGrid.Children.Add(_txtFind);
            Wpf.Grid.SetColumn(_txtFind, 1);

            var lblReplace = new Wpf.TextBlock { Text = "Replace", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 0, 6, 0) };
            frGrid.Children.Add(lblReplace);
            Wpf.Grid.SetColumn(lblReplace, 2);

            frGrid.Children.Add(_txtReplace);
            Wpf.Grid.SetColumn(_txtReplace, 3);

            root.Children.Add(frGrid);
            Wpf.Grid.SetRow(frGrid, 2);

            // Prefix / Suffix line
            var psGrid = new Wpf.Grid { Margin = new Thickness(0, 0, 0, 6) };
            psGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = System.Windows.GridLength.Auto });
            psGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });
            psGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = System.Windows.GridLength.Auto });
            psGrid.ColumnDefinitions.Add(new Wpf.ColumnDefinition { Width = new System.Windows.GridLength(1, System.Windows.GridUnitType.Star) });

            psGrid.Children.Add(new Wpf.TextBlock { Text = "Prefix", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(0, 0, 6, 0) });
            psGrid.Children.Add(_txtPrefix);
            Wpf.Grid.SetColumn(_txtPrefix, 1);

            var lblSuffix = new Wpf.TextBlock { Text = "Suffix", VerticalAlignment = VerticalAlignment.Center, Margin = new Thickness(12, 0, 6, 0) };
            psGrid.Children.Add(lblSuffix);
            Wpf.Grid.SetColumn(lblSuffix, 2);

            psGrid.Children.Add(_txtSuffix);
            Wpf.Grid.SetColumn(_txtSuffix, 3);

            root.Children.Add(psGrid);
            Wpf.Grid.SetRow(psGrid, 3);

            // Options
            var optPanel = new Wpf.StackPanel
            {
                Orientation = Wpf.Orientation.Horizontal,
                Margin = new Thickness(0, 0, 0, 10)
            };
            _chkCaseSensitive.Content = "Case-sensitive";
            _chkWholeWord.Content = "Whole-word";
            _chkCaseSensitive.Margin = new Thickness(0, 0, 12, 0);
            optPanel.Children.Add(_chkCaseSensitive);
            optPanel.Children.Add(_chkWholeWord);
            root.Children.Add(optPanel);
            Wpf.Grid.SetRow(optPanel, 4);

            // Action buttons
            var actionPanel = new Wpf.StackPanel
            {
                Orientation = Wpf.Orientation.Horizontal,
                HorizontalAlignment = System.Windows.HorizontalAlignment.Left
            };
            actionPanel.Children.Add(MakeButton("Find & Replace", OnFindReplace));
            actionPanel.Children.Add(MakeButton("Refresh Selection", OnRefreshSelection));
            actionPanel.Children.Add(MakeButton("Close", OnClose));
            root.Children.Add(actionPanel);
            Wpf.Grid.SetRow(actionPanel, 5);

            // Footer hint
            var hint = new Wpf.TextBlock
            {
                Margin = new Thickness(0, 8, 0, 0),
                Opacity = 0.7,
                Text = "Each button applies to the currently selected Text Notes. Use Refresh Selection if you change the selection."
            };
            root.Children.Add(hint);
            Wpf.Grid.SetRow(hint, 6);

            foreach (var tb in new[] { _txtFind, _txtReplace, _txtPrefix, _txtSuffix })
                tb.MinWidth = 160;

            return root;
        }

        private Wpf.Button MakeButton(string text, RoutedEventHandler onClick)
        {
            var b = new Wpf.Button
            {
                Content = text,
                Margin = new Thickness(0, 0, 8, 8),
                Padding = new Thickness(12, 6, 12, 6)
            };
            b.Click += onClick;
            return b;
        }

        private void UpdateCountLabel()
        {
            _lblCount.Text = $"({_textNoteIds.Count} selected)";
            _lblCount.VerticalAlignment = VerticalAlignment.Bottom;
            _lblCount.Opacity = 0.75;
        }

        // Button handlers (no inline lambdas)

        private void OnUpper(object sender, RoutedEventArgs e) => ApplyTransform(StrToUpper);
        private void OnLower(object sender, RoutedEventArgs e) => ApplyTransform(StrToLower);
        private void OnTitle(object sender, RoutedEventArgs e) => ApplyTransform(StrToTitle);
        private void OnCapitalise(object sender, RoutedEventArgs e) => ApplyTransform(StrCapitaliseFirst);
        private void OnSwapCase(object sender, RoutedEventArgs e) => ApplyTransform(StrSwapCase);
        private void OnReverse(object sender, RoutedEventArgs e) => ApplyTransform(StrReverse);
        private void OnTrim(object sender, RoutedEventArgs e) => ApplyTransform(StrTrim);

        private void OnClose(object sender, RoutedEventArgs e) => Close();

        private void OnRefreshSelection(object sender, RoutedEventArgs e)
        {
            _textNoteIds = _uidoc.Selection.GetElementIds()
                .Select(id => _doc.GetElement(id))
                .OfType<TextNote>()
                .Select(tn => tn.Id)
                .ToList();

            if (_textNoteIds.Count == 0)
            {
                try
                {
                    var picked = _uidoc.Selection.PickObjects(ObjectType.Element,
                        new TextNoteSelectionFilter(), "Select Text Notes");
                    _textNoteIds = picked.Select(r => r.ElementId).ToList();
                }
                catch (Autodesk.Revit.Exceptions.OperationCanceledException)
                {
                    // keep previous selection set
                }
            }

            UpdateCountLabel();
        }

        private void OnFindReplace(object sender, RoutedEventArgs e)
        {
            string find = _txtFind.Text ?? string.Empty;
            string replace = _txtReplace.Text ?? string.Empty;
            string prefix = _txtPrefix.Text ?? string.Empty;
            string suffix = _txtSuffix.Text ?? string.Empty;

            if (string.IsNullOrEmpty(find) && string.IsNullOrEmpty(prefix) && string.IsNullOrEmpty(suffix))
            {
                TaskDialog.Show("Text Transform", "Enter a Find term and/or Prefix/Suffix.");
                return;
            }

            int changes = 0;

            using (var t = new Transaction(_doc, "Text Transform – Find/Replace"))
            {
                t.Start();

                foreach (var tn in GetLiveTextNotes())
                {
                    string original = tn.Text ?? string.Empty;
                    string replaced = ReplaceCore(original, find, replace,
                        _chkCaseSensitive.IsChecked == true,
                        _chkWholeWord.IsChecked == true);

                    string final = prefix + replaced + suffix;

                    if (!string.Equals(final, original, StringComparison.Ordinal))
                    {
                        tn.Text = final;
                        changes++;
                    }
                }

                t.Commit();
            }

            _doc.Regenerate();
            TaskDialog.Show("Text Transform", $"Updated {changes} Text Note(s).");
        }

        private void ApplyTransform(Func<string, string> transform)
        {
            if (_textNoteIds.Count == 0)
            {
                TaskDialog.Show("Text Transform", "No Text Notes selected.");
                return;
            }

            int changes = 0;

            using (var t = new Transaction(_doc, "Text Transform"))
            {
                t.Start();

                foreach (var tn in GetLiveTextNotes())
                {
                    string original = tn.Text ?? string.Empty;
                    string updated = transform(original) ?? string.Empty;

                    if (!string.Equals(updated, original, StringComparison.Ordinal))
                    {
                        tn.Text = updated;
                        changes++;
                    }
                }

                t.Commit();
            }

            _doc.Regenerate();
            TaskDialog.Show("Text Transform", $"Updated {changes} Text Note(s).");
        }

        private IEnumerable<TextNote> GetLiveTextNotes()
        {
            foreach (var id in _textNoteIds)
            {
                var e = _doc.GetElement(id) as TextNote;
                if (e != null) yield return e;
            }
        }

        // Replace logic supporting case-sensitivity and whole-word
        private static string ReplaceCore(string input, string find, string replace, bool caseSensitive, bool wholeWord)
        {
            if (string.IsNullOrEmpty(find)) return input ?? string.Empty;

            // Fast path: simple Replace when no flags set
            if (!caseSensitive && !wholeWord)
                return (input ?? string.Empty).Replace(find, replace ?? string.Empty);

            // Build regex pattern
            string pattern = Regex.Escape(find);
            if (wholeWord) pattern = @"\b" + pattern + @"\b";

            var options = caseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase;
            return Regex.Replace(input ?? string.Empty, pattern, replace ?? string.Empty, options);
        }

        // String helpers

        private static string StrToUpper(string s) => s != null ? s.ToUpper(CultureInfo.CurrentCulture) : null;
        private static string StrToLower(string s) => s != null ? s.ToLower(CultureInfo.CurrentCulture) : null;

        private static string StrToTitle(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var ti = CultureInfo.CurrentCulture.TextInfo;
            return ti.ToTitleCase(s.ToLower(CultureInfo.CurrentCulture));
        }

        private static string StrCapitaliseFirst(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            if (s.Length == 1) return s.ToUpper(CultureInfo.CurrentCulture);
            return char.ToUpper(s[0], CultureInfo.CurrentCulture) + s.Substring(1);
        }

        private static string StrReverse(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var arr = s.ToCharArray();
            Array.Reverse(arr);
            return new string(arr);
        }

        private static string StrSwapCase(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            var sb = new StringBuilder(s.Length);
            foreach (var c in s)
            {
                if (char.IsLetter(c))
                {
                    if (char.IsUpper(c))
                        sb.Append(char.ToLower(c, CultureInfo.CurrentCulture));
                    else
                        sb.Append(char.ToUpper(c, CultureInfo.CurrentCulture));
                }
                else
                {
                    sb.Append(c);
                }
            }
            return sb.ToString();
        }

        private static string StrTrim(string s) => s != null ? s.Trim() : null;
    }
}
