using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using HDR_EMEA.Forms;

namespace HDR_EMEA.Tools_Core
{
    [Transaction(TransactionMode.ReadOnly)]
    public class ScheduleExport : IExternalCommand
    {
        public Result Execute(ExternalCommandData data,
                              ref string message,
                              ElementSet elements)
        {
            var uiDoc = data.Application.ActiveUIDocument;

            // 1) pick schedules
            var picker = new SchedulePickerWindow(uiDoc);
            bool? ok = picker.ShowDialog();
            if (ok != true || picker.SelectedSchedules.Length == 0)
                return Result.Cancelled;
            var schedules = picker.SelectedSchedules;

            // 2) choose output file
            var dlg = new Microsoft.Win32.SaveFileDialog
            {
                Title = "Export Schedules to Excel",
                Filter = "Excel Workbook|*.xlsx",
                FileName = "Schedules.xlsx",
                DefaultExt = ".xlsx",
                AddExtension = true
            };
            if (dlg.ShowDialog() != true) return Result.Cancelled;
            string filePath = dlg.FileName;

            // helper to strip CSV quotes
            static string Unquote(string s)
            {
                if (s.Length >= 2 && s[0] == '"' && s[s.Length - 1] == '"')
                    s = s.Substring(1, s.Length - 2).Replace("\"\"", "\"");
                return s;
            }

            using (var wb = new XLWorkbook())
            {
                foreach (var sched in schedules)
                {
                    // a) export raw CSV
                    string tmpBase = Guid.NewGuid().ToString();
                    string tmpFile = Path.Combine(Path.GetTempPath(), tmpBase);
                    var opts = new ViewScheduleExportOptions
                    {
                        ColumnHeaders = ExportColumnHeaders.MultipleRows,
                        FieldDelimiter = ",",
                        TextQualifier = ExportTextQualifier.DoubleQuote,
                        HeadersFootersBlanks = true,
                        Title = true
                    };
                    sched.Export(Path.GetDirectoryName(tmpFile),
                                 Path.GetFileName(tmpFile),
                                 opts);

                    // b) read & delete
                    var lines = File.ReadAllLines(tmpFile);
                    File.Delete(tmpFile);
                    if (lines.Length == 0) continue;

                    // c) add safe sheet
                    string sheetName = SafeSheetName(sched.Name);
                    var ws = wb.Worksheets.Add(sheetName);

                    // d) write exact CSV fields
                    for (int r = 0; r < lines.Length; r++)
                    {
                        var parts = lines[r].Split(',');
                        for (int c = 0; c < parts.Length; c++)
                            ws.Cell(r + 1, c + 1).Value = Unquote(parts[c]);
                    }

                    // e) optional table formatting
                    int lastRow = lines.Length;
                    int lastCol = lines[0].Split(',').Length;
                    var tbl = ws.Range(1, 1, lastRow, lastCol).CreateTable();
                    tbl.Theme = XLTableTheme.TableStyleMedium2;
                }

                // 4) save
                wb.SaveAs(filePath);
            }

            TaskDialog.Show("Export Complete",
                            $"Wrote {schedules.Length} sheet(s) to:\n{filePath}");
            return Result.Succeeded;
        }

        /// <summary>
        /// Ensures the sheet name is valid in Excel: strips invalid chars & max 31 chars.
        /// </summary>
        private string SafeSheetName(string name)
        {
            var invalid = new[] { '\\', '/', '?', '*', '[', ']', ':' };
            string cleaned = new string(name
                .Where(c => !invalid.Contains(c))
                .ToArray())
                .Trim(' ', '\'');

            if (string.IsNullOrWhiteSpace(cleaned))
                cleaned = "Schedule";

            return cleaned.Length > 31
                ? cleaned.Substring(0, 31)
                : cleaned;
        }

        public static Common.ButtonDataClass GetButtonData()
        {
            return new Common.ButtonDataClass(
                "ScheduleExport",
                "Export\nSchedules",
                "HDR_EMEA.Tools_Core.ScheduleExport",
                Properties.Resources.Excel_32,
                Properties.Resources.Excel_16,
                "Export multiple schedules to a single Excel file"
            );
        }
    }
}
