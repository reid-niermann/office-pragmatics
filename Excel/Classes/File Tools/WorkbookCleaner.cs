// WorkbookCleaner.cs – OfficePragmatics_Excel  (C# 7.3)

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using WinForms = System.Windows.Forms;

namespace OfficePragmatics_Excel
{
    internal static class WorkbookCleaner
    {
        /*────────────────────────────  PUBLIC ENTRY POINTS  ─────────────────────────────*/

        public static void CleanAllNames()
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null) return;

            if (string.IsNullOrEmpty(wb.FullName))
            {
                WinForms.MessageBox.Show(
                    "Please save the workbook before running the name cleaner.",
                    "Workbook Cleaner",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Information);
                return;
            }

            // Backup first (local or SharePoint-aware)
            BackupWorkbook(wb);

            int count = SafeCount(() => wb.Names.Count);

            // Always use the fast XML-based cleaner, regardless of size.
            RunWithProgress($"Cleaning workbook names ({count:n0})", _ =>
            {
                FastCleanNames_OpenXml(wb, count);
                return count;
            }, showDoneDialog: false);
        }

        public static void ComprehensiveClean()
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook;
            if (wb == null) return;

            // Starting file size – right at the beginning
            long startBytes;
            bool haveStartSize = TryGetWorkbookFileSizeBytes(wb, out startBytes);

            // Run all cleaners
            CleanAllNames();
            CleanNumberFormats(false);
            CleanCellRanges(false);
            CleanThemes(false);

            // Ensure the final state is saved before measuring ending size
            try
            {
                app.ActiveWorkbook?.Save();
            }
            catch
            {
                // Ignore save failures for purposes of the size message
            }

            // Ending file size – after everything is done
            wb = app.ActiveWorkbook;
            long endBytes = 0;
            bool haveEndSize = wb != null && TryGetWorkbookFileSizeBytes(wb, out endBytes);

            if (haveStartSize && haveEndSize)
            {
                long startKB = startBytes / 1024;
                long endKB = endBytes / 1024;

                string startText = $"{startKB:n0} KB";
                string endText = $"{endKB:n0} KB";

                string sizeClause;
                if (endKB < startKB)
                {
                    sizeClause = $"Decreased file size from {startText} to {endText}.";
                }
                else if (endKB > startKB)
                {
                    sizeClause = $"Changed file size from {startText} to {endText}.";
                }
                else
                {
                    sizeClause = $"Kept file size at {startText}.";
                }

                WinForms.MessageBox.Show(
                    "File optimization complete. Eliminated performance drags. " +
                    sizeClause,
                    "Workbook Cleaner",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Information);
            }
            else
            {
                // Fallback if we can't reliably measure file size (e.g., non-file path)
                WinForms.MessageBox.Show(
                    "File optimization complete. Eliminated performance drags.",
                    "Workbook Cleaner",
                    WinForms.MessageBoxButtons.OK,
                    WinForms.MessageBoxIcon.Information);
            }
        }

        /*────────────────────  FAST CLEANER (used for all sizes)  ───────────────────*/

        //  For name cleaning we operate on a local file:
        //  - Local/UNC workbook: use its real path
        //  - SharePoint/WOPI: SaveCopyAs to a temp file, clean that, then SaveAs back
        private static void FastCleanNames_OpenXml(Excel.Workbook wb, int originalCount)
        {
            var app = wb.Application;
            string originalFullName = wb.FullName ?? string.Empty;
            string originalPath = wb.Path ?? string.Empty;

            bool isLocal = IsLocalFilePath(originalFullName);
            bool isRemote = !isLocal;

            string workingPath;
            string remoteTarget = null;

            try { if (!wb.Saved) wb.Save(); } catch { }

            // Capture the window caption so we can restore nicely
            string winCaption = null;
            try { if (wb.Windows.Count > 0) winCaption = wb.Windows[1].Caption; } catch { }

            if (isLocal)
            {
                // Normal file system – operate directly on the file
                workingPath = originalFullName;
            }
            else
            {
                // SharePoint / OneDrive / WOPI – create a local working copy
                workingPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                try
                {
                    wb.SaveCopyAs(workingPath);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        "Workbook Cleaner could not prepare the file for cleaning. " + ex.Message, ex);
                }

                // Best-effort: figure out where to SaveAs back to
                try
                {
                    if (!string.IsNullOrEmpty(originalFullName) &&
                        (originalFullName.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                         originalFullName.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                    {
                        // FullName is already a URL
                        remoteTarget = originalFullName;
                    }
                    else if (!string.IsNullOrEmpty(originalPath) &&
                             (originalPath.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                              originalPath.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                    {
                        // Path is the folder URL; rebuild a URL for this workbook
                        string ext = Path.GetExtension(wb.Name);
                        if (string.IsNullOrEmpty(ext)) ext = ".xlsx";
                        string baseName = Path.GetFileNameWithoutExtension(wb.Name);
                        string fileName = baseName + ext;
                        remoteTarget = originalPath.TrimEnd('/') + "/" +
                                       Uri.EscapeDataString(fileName);
                    }
                }
                catch
                {
                    remoteTarget = null;
                }
            }

            // Close the original workbook before we mutate the file on disk
            wb.Close(SaveChanges: true);

            var dlg = ProgressDialog.Current;
            if (dlg != null)
            {
                dlg.SetMaximum(5);
                dlg.Step("Preparing workbook for name cleanup…");
            }

            HashSet<string> allNames;
            HashSet<string> usedNames;
            HashSet<string> sheetNames;
            List<DefinedName> xmlNames;

            OpenSettings opts = new OpenSettings { AutoSave = true };

            // FIX: Opened directly by path to avoid VSTO IsolatedStorageException (Identity of domain)
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(workingPath, true, opts))
            {
                WorkbookPart wp = doc.WorkbookPart;
                Workbook wbXml = wp.Workbook;

                xmlNames = wbXml.DefinedNames?.Elements<DefinedName>().ToList()
                           ?? new List<DefinedName>();

                sheetNames = new HashSet<string>(
                                wbXml.Descendants<Sheet>()
                                     .Select(s => s.Name?.Value ?? ""),
                                StringComparer.OrdinalIgnoreCase);

                allNames = new HashSet<string>(
                                xmlNames.Select(n => n.Name?.Value ?? "")
                                        .Where(s => !string.IsNullOrEmpty(s)),
                                StringComparer.OrdinalIgnoreCase);

                usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                if (dlg != null) dlg.Step("Reviewing worksheets for named references…");

                foreach (WorksheetPart wsp in wp.WorksheetParts)
                    ScanWorksheetPartForNames(wsp, allNames, usedNames);

                foreach (var dn in xmlNames)
                    ScanFormulaForNames(dn.InnerText, allNames, usedNames);

                if (dlg != null) dlg.Step("Checking which names are still in use…");

                var root = wbXml.DefinedNames ?? wbXml.AppendChild(new DefinedNames());
                var keep = new List<DefinedName>();

                foreach (var dn in xmlNames)
                {
                    string nm = dn.Name?.Value ?? "";
                    bool isPA = nm.IndexOf("Print_Area", StringComparison.OrdinalIgnoreCase) >= 0;
                    string refTxt = dn.InnerText ?? "";
                    bool hasErr = refTxt.Contains("#");
                    bool badSheet = !hasErr && !IsSheetReferenceValid(refTxt, sheetNames);

                    if (hasErr || badSheet) continue;                                   // drop
                    if (isPA)
                    {
                        keep.Add((DefinedName)dn.CloneNode(true));
                        continue;
                    }
                    if (usedNames.Contains(nm)) keep.Add((DefinedName)dn.CloneNode(true));
                }

                root.RemoveAllChildren();
                foreach (var k in keep) root.Append(k);
                wbXml.Save();
            }

            if (dlg != null) dlg.Step("Loading cleaned workbook…");
            var reopened = app.Workbooks.Open(workingPath);
            try { reopened.Windows[1].Caption = winCaption ?? reopened.Windows[1].Caption; } catch { }

            // For SharePoint / OneDrive: try to push the cleaned copy back to the original URL
            if (isRemote && !string.IsNullOrEmpty(remoteTarget))
            {
                try
                {
                    reopened.SaveAs(
                        remoteTarget,
                        FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);
                }
                catch (Exception ex)
                {
                    WinForms.MessageBox.Show(
                        "The workbook was cleaned in a working copy but could not be saved back " +
                        "to the original SharePoint / OneDrive location.\r\n\r\n" +
                        "A cleaned copy is open in Excel. You can use File ▸ Save As to store it.\r\n\r\n" +
                        "Error: " + ex.Message,
                        "Workbook Cleaner",
                        WinForms.MessageBoxButtons.OK,
                        WinForms.MessageBoxIcon.Warning);
                }
            }

            int removed = originalCount - (usedNames != null ? usedNames.Count : 0);
            if (removed < 0) removed = 0;
            if (dlg != null) dlg.Step(string.Format("Name cleanup finished (removed {0:n0} names).", removed));

            // For remote cases, try to clean up the local working file
            if (isRemote)
            {
                try { if (File.Exists(workingPath)) File.Delete(workingPath); } catch { }
            }
        }

        /*────────────────────  TOKEN SCAN HELPERS  ───────────────────*/

        private static void ScanWorksheetPartForNames(WorksheetPart wsp,
                                                      HashSet<string> universe,
                                                      HashSet<string> used)
        {
            foreach (Cell c in wsp.Worksheet
                                   .Descendants<Cell>()
                                   .Where(c => c.CellFormula != null))
            {
                ScanFormulaForNames(c.CellFormula.InnerText, universe, used);
            }
        }

        private static void ScanFormulaForNames(string formula,
                                                HashSet<string> universe,
                                                HashSet<string> used)
        {
            if (string.IsNullOrEmpty(formula)) return;
            int len = formula.Length, i = 0;
            while (i < len)
            {
                char c = formula[i];
                if ((c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z') || c == '_')
                {
                    int s = i; i++;
                    while (i < len)
                    {
                        char d = formula[i];
                        if ((d >= 'A' && d <= 'Z') || (d >= 'a' && d <= 'z') ||
                            (d >= '0' && d <= '9') || d == '_' || d == '.') i++;
                        else break;
                    }
                    string tok = formula.Substring(s, i - s);
                    if (universe.Contains(tok)) used.Add(tok);
                }
                else i++;
            }
        }

        /*────────────────────  SHEET VALIDITY CHECK  ───────────────────*/

        private static readonly Regex SheetRefRegex =
            new Regex(@"'?<sheet>[^\s'\[\]!]+'?!", RegexOptions.Compiled);

        private static bool IsSheetReferenceValid(string refersTo, HashSet<string> sheetNames)
        {
            foreach (Match m in SheetRefRegex.Matches(refersTo))
            {
                string sheet = m.Groups["sheet"].Value;
                if (!sheetNames.Contains(sheet)) return false;   // bad sheet
            }
            return true;    // all sheet refs exist (or no sheet refs)
        }

        /*────────────────────  OTHER CLEANERS  ───────────────────*/

        public static void CleanNumberFormats(bool showDone = true) =>
            RunWithProgress("Tidying up number formats", wb =>
            {
                var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (Excel.Worksheet ws in wb.Worksheets)
                {
                    try
                    {
                        var ur = ws.UsedRange; if (ur == null) continue;
                        object v = ur.Style;
                        if (v is object[,] arr) foreach (var s in arr) used.Add(s?.ToString() ?? "");
                        else used.Add(v?.ToString() ?? "");
                    }
                    catch { }
                }
                var styles = wb.Styles.Cast<Excel.Style>().Where(s => !s.BuiltIn).ToList();
                return Enumerate(styles.Count, (p, i) =>
                {
                    var s = styles[i]; string n = SafeStyleName(s);
                    try { if (!used.Contains(n)) s.Delete(); } catch { }
                    p.Step($"Reviewing format “{n}”…");
                });
            }, showDone);

        public static void CleanCellRanges(bool showDone = true) =>
            RunWithProgress("Clearing unused rows and columns", wb =>
            {
                var sheets = wb.Worksheets.Cast<Excel.Worksheet>().ToList();
                return Enumerate(sheets.Count, (p, i) =>
                {
                    var ws = sheets[i]; Excel.Range ur = null;
                    try { ur = ws.UsedRange; } catch { }
                    int lr = ur?.Row + ur?.Rows.Count - 1 ?? 1;
                    int lc = ur?.Column + ur?.Columns.Count - 1 ?? 1;
                    try
                    {
                        if (lr < ws.Rows.Count)
                            ws.Range[ws.Cells[lr + 1, 1], ws.Cells[ws.Rows.Count, ws.Columns.Count]].Clear();
                        if (lc < ws.Columns.Count)
                            ws.Range[ws.Cells[1, lc + 1], ws.Cells[lr, ws.Columns.Count]].Clear();
                    }
                    catch { }
                    p.Step($"Tidying sheet “{ws.Name}”…");
                });
            }, showDone);

        public static void CleanThemes(bool showDone = true) =>
            RunWithProgress("Resetting workbook theme", wb =>
            {
                try { string t = FindOfficeTheme(); if (t != null) wb.ApplyTheme(t); } catch { }
                return 1;
            }, showDone);

        /*────────────────────  helper utilities  ───────────────────*/

        private static bool IsLocalFilePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            if (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return false;

            Uri uri;
            if (Uri.TryCreate(path, UriKind.Absolute, out uri))
                return uri.IsFile;

            return Path.IsPathRooted(path);
        }

        private static bool TryGetWorkbookFileSizeBytes(Excel.Workbook wb, out long bytes)
        {
            bytes = 0;
            try
            {
                string fullName = wb?.FullName;
                if (string.IsNullOrWhiteSpace(fullName)) return false;
                if (!IsLocalFilePath(fullName)) return false;
                if (!File.Exists(fullName)) return false;

                bytes = new FileInfo(fullName).Length;
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string SafeStyleName(Excel.Style s)
        {
            try { return s?.Name ?? "<unknown>"; }
            catch { return "<unknown>"; }
        }

        private static int SafeCount(Func<int> f)
        {
            try { return f(); }
            catch { return 0; }
        }

        /*──────────────  BACKUP UTILITIES  ──────────────*/

        private static void BackupWorkbook(Excel.Workbook wb)
        {
            try
            {
                string original = wb.FullName;
                if (string.IsNullOrEmpty(original)) return; // unsaved workbook

                string backupPath;

                if (IsLocalFilePath(original))
                {
                    // Normal local/UNC file – keep backup next to original
                    backupPath = GetBackupPath(original);
                }
                else
                {
                    // SharePoint/WOPI – put backup in Documents
                    string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    string baseName = Path.GetFileNameWithoutExtension(wb.Name)
                                            .Replace("#", "_").Replace("%", "_");
                    string ext = Path.GetExtension(wb.Name);
                    if (string.IsNullOrEmpty(ext)) ext = ".xlsx";

                    backupPath = Path.Combine(docs, baseName + "_OLD" + ext);
                    int i = 1;
                    while (File.Exists(backupPath))
                        backupPath = Path.Combine(docs, $"{baseName}_OLD ({i++}){ext}");
                }

                wb.SaveCopyAs(backupPath);
            }
            catch
            {
                // backup failure should not abort cleaning
            }
        }

        private static string GetBackupPath(string original)
        {
            string dir = Path.GetDirectoryName(original);
            string b = Path.GetFileNameWithoutExtension(original)
                          .Replace("#", "_").Replace("%", "_");
            string ext = Path.GetExtension(original);
            string c = Path.Combine(dir, b + "_OLD" + ext);
            int i = 1;
            while (File.Exists(c)) c = Path.Combine(dir, $"{b}_OLD ({i++}){ext}");
            return c;
        }

        private static string FindOfficeTheme()
        {
            string[] p =
            {
                Environment.ExpandEnvironmentVariables(
                    @"%ProgramFiles(x86)%\Microsoft Office\root\Document Themes 16\Office Theme.thmx"),
                Environment.ExpandEnvironmentVariables(
                    @"%ProgramFiles(x86)%\Microsoft Office\Document Themes 16\Document Themes 16\Office Theme.thmx"),
                Environment.ExpandEnvironmentVariables(
                    @"%ProgramFiles(x86)%\Microsoft Office\Document Themes 16\Office Theme.thmx")
            };
            return Array.Find(p, File.Exists);
        }

        /*────────────────────  progress & runner  ───────────────────*/

        private static int Enumerate(int count, Action<ProgressDialog, int> body)
        {
            var dlg = ProgressDialog.Current;
            if (dlg != null)
            {
                dlg.SetMaximum(count);
                for (int i = 0; i < count; i++)
                {
                    dlg.ThrowIfCancelled();
                    dlg.WaitIfPaused();
                    body(dlg, i);
                }
            }
            else
            {
                // Fallback without status updates
                for (int i = 0; i < count; i++)
                {
                    body(null, i);
                }
            }
            return count;
        }

        private static void RunWithProgress(string title, Func<Excel.Workbook, int> work, bool showDoneDialog = true)
        {
            var app = Globals.ThisAddIn.Application;
            var wb = app.ActiveWorkbook ?? throw new InvalidOperationException("No active workbook.");
            var excelHwnd = new IntPtr(app.Hwnd);

            using (var dlg = new ProgressDialog(title, excelHwnd))
            {
                // No visual form anymore – all progress goes to the status bar.
                dlg.Show(new ExcelWin32Window(excelHwnd));

                int total = 0;
                Exception err = null;
                try
                {
                    total = work(wb);
                }
                catch (Exception ex)
                {
                    err = ex;
                }

                if (err != null)
                {
                    WinForms.MessageBox.Show(
                        $"{title}\n\nError: {err.Message}",
                        "Workbook Cleaner",
                        WinForms.MessageBoxButtons.OK,
                        WinForms.MessageBoxIcon.Error);
                }
                else if (showDoneDialog && total > 0)
                {
                    WinForms.MessageBox.Show(
                        $"{title}\n\nDone ({total} item(s)).",
                        "Workbook Cleaner",
                        WinForms.MessageBoxButtons.OK,
                        WinForms.MessageBoxIcon.Information);
                }
            }
        }

        private sealed class ExcelWin32Window : WinForms.IWin32Window
        {
            public ExcelWin32Window(IntPtr h) { Handle = h; }
            public IntPtr Handle { get; }
        }

        /// <summary>
        /// Lightweight, non-UI progress helper that only writes to Excel's status bar.
        /// Keeps the old API shape (Current/SetMaximum/Step/Show/etc.) so other code
        /// does not need to change.
        /// </summary>
        private sealed class ProgressDialog : IDisposable
        {
            internal static ProgressDialog Current { get; private set; }

            private readonly Excel.Application _app;
            private readonly string _caption;
            private int _maximum = 100;
            private int _value = 0;
            private readonly object _originalStatus;
            private bool _disposed;

            internal ProgressDialog(string caption, IntPtr excelHwnd)
            {
                _app = Globals.ThisAddIn.Application;
                _caption = string.IsNullOrEmpty(caption) ? "Working…" : caption;

                try
                {
                    _originalStatus = _app.StatusBar;
                }
                catch
                {
                    _originalStatus = null;
                }

                Current = this;
                SetStatus("Starting…");
            }

            /// <summary>
            /// Kept for compatibility; does nothing now (no form is shown).
            /// </summary>
            internal void Show(WinForms.IWin32Window owner)
            {
                // No UI – we only use the status bar.
            }

            internal void SetMaximum(int max)
            {
                _maximum = Math.Max(1, max);
                _value = 0;
            }

            internal void Step(string message)
            {
                _value = Math.Min(_value + 1, _maximum);

                string text;
                if (_maximum > 1)
                {
                    text = string.IsNullOrEmpty(message)
                        ? $"{_caption} ({_value}/{_maximum})"
                        : $"{_caption} ({_value}/{_maximum}) – {message}";
                }
                else
                {
                    text = string.IsNullOrEmpty(message)
                        ? _caption
                        : $"{_caption} – {message}";
                }

                try
                {
                    _app.StatusBar = text;
                }
                catch
                {
                    // ignore status bar failures
                }

                WinForms.Application.DoEvents();
            }

            internal void ThrowIfCancelled()
            {
                // Cancel support removed – no UI to capture it.
            }

            internal void WaitIfPaused()
            {
                // Pause support removed – just a no-op now.
            }

            private void SetStatus(string message)
            {
                string text = string.IsNullOrEmpty(message)
                    ? _caption
                    : $"{_caption} – {message}";

                try
                {
                    _app.StatusBar = text;
                }
                catch
                {
                    // ignore
                }
            }

            public void Dispose()
            {
                if (_disposed) return;
                _disposed = true;

                if (Current == this)
                    Current = null;

                try
                {
                    if (_originalStatus == null)
                    {
                        // Reset to Excel's default "Ready"
                        _app.StatusBar = false;
                    }
                    else
                    {
                        _app.StatusBar = _originalStatus;
                    }
                }
                catch
                {
                    // ignore
                }
            }
        }
    }
}