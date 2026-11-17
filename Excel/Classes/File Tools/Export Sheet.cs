using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System.IO.Compression;
using System.IO.MemoryMappedFiles;
using System.Xml;

namespace OfficePragmatics_Excel.Classes
{
    public static class ExportSelectionTool
    {
        internal sealed class ExcelUiSilencer : IDisposable
        {
            private readonly Excel.Application _app;
            private readonly bool _screen, _alerts, _events;
            private readonly Excel.XlCalculation _calc;
            private readonly Excel.XlMousePointer _cursor;

            public ExcelUiSilencer(Excel.Application app)
            {
                _app = app;
                _screen = app.ScreenUpdating;
                _alerts = app.DisplayAlerts;
                _events = app.EnableEvents;
                _calc = app.Calculation;
                _cursor = app.Cursor;

                try
                {
                    app.ScreenUpdating = false;
                    app.DisplayAlerts = false;
                    app.EnableEvents = false;
                    app.Calculation = Excel.XlCalculation.xlCalculationManual;
                    app.Cursor = Excel.XlMousePointer.xlWait;
                }
                catch { /* ignore */ }
            }

            public void Dispose()
            {
                try { _app.Cursor = _cursor; } catch { }
                try { _app.Calculation = _calc; } catch { }
                try { _app.EnableEvents = _events; } catch { }
                try { _app.DisplayAlerts = _alerts; } catch { }
                try { _app.ScreenUpdating = _screen; } catch { }
            }
        }

        private sealed class ExportSelectionOptions
        {
            public string WorkbookName { get; set; }
            public bool ConvertBlue { get; set; }
            public bool ConvertGreen { get; set; }
            public bool ConvertRed { get; set; }
            public bool ConvertPurple { get; set; }
        }

        private sealed class ExportSelectionDialog : Form
        {
            private readonly TextBox _nameBox;
            private readonly CheckBox _chkBlue;
            private readonly CheckBox _chkGreen;
            private readonly CheckBox _chkRed;
            private readonly CheckBox _chkPurple;
            private readonly Button _okButton;
            private readonly Button _cancelButton;

            public ExportSelectionOptions Options { get; private set; }

            public ExportSelectionDialog(string suggestedName)
            {
                Text = "Export Selection";
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterScreen;
                MinimizeBox = false;
                MaximizeBox = false;
                ShowInTaskbar = false;
                AutoSize = true;
                AutoSizeMode = AutoSizeMode.GrowAndShrink;

                var mainLayout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    ColumnCount = 1,
                    RowCount = 3,
                    Padding = new Padding(10)
                };

                var nameLayout = new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    ColumnCount = 1,
                    RowCount = 2
                };

                var lblName = new Label
                {
                    Text = "New workbook name (no extension):",
                    AutoSize = true,
                    Dock = DockStyle.Fill,
                    Margin = new Padding(0, 0, 0, 4)
                };

                _nameBox = new TextBox
                {
                    Text = suggestedName ?? string.Empty,
                    Dock = DockStyle.Top,
                    Width = 260
                };

                nameLayout.Controls.Add(lblName, 0, 0);
                nameLayout.Controls.Add(_nameBox, 0, 1);

                var groupBox = new GroupBox
                {
                    Text = "Convert font colors to black:",
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    Padding = new Padding(10, 18, 10, 10)
                };

                var chkLayout = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.TopDown,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    WrapContents = false
                };

                _chkBlue = new CheckBox
                {
                    Text = "Blue → black",
                    AutoSize = true
                };
                _chkGreen = new CheckBox
                {
                    Text = "Green → black",
                    AutoSize = true
                };
                _chkRed = new CheckBox
                {
                    Text = "Red → black",
                    AutoSize = true
                };
                _chkPurple = new CheckBox
                {
                    Text = "Purple → black",
                    AutoSize = true
                };

                chkLayout.Controls.Add(_chkBlue);
                chkLayout.Controls.Add(_chkGreen);
                chkLayout.Controls.Add(_chkRed);
                chkLayout.Controls.Add(_chkPurple);
                groupBox.Controls.Add(chkLayout);

                var buttonsLayout = new FlowLayoutPanel
                {
                    Dock = DockStyle.Top,
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    Padding = new Padding(0, 10, 0, 0)
                };

                _okButton = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    AutoSize = true
                };
                _cancelButton = new Button
                {
                    Text = "Cancel",
                    DialogResult = DialogResult.Cancel,
                    AutoSize = true
                };

                buttonsLayout.Controls.Add(_okButton);
                buttonsLayout.Controls.Add(_cancelButton);

                AcceptButton = _okButton;
                CancelButton = _cancelButton;

                mainLayout.Controls.Add(nameLayout);
                mainLayout.Controls.Add(groupBox);
                mainLayout.Controls.Add(buttonsLayout);

                Controls.Add(mainLayout);

                _okButton.Click += OkButton_Click;
            }

            private void OkButton_Click(object sender, EventArgs e)
            {
                var name = (_nameBox.Text ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(name))
                {
                    MessageBox.Show(
                        this,
                        "Please enter a workbook name.",
                        "Export Selection",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    DialogResult = DialogResult.None;
                    return;
                }

                foreach (var invalid in Path.GetInvalidFileNameChars())
                {
                    if (name.Contains(invalid.ToString()))
                    {
                        MessageBox.Show(
                            this,
                            "The workbook name contains invalid characters.",
                            "Export Selection",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        DialogResult = DialogResult.None;
                        return;
                    }
                }

                if (name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                {
                    name = Path.GetFileNameWithoutExtension(name);
                }

                Options = new ExportSelectionOptions
                {
                    WorkbookName = name,
                    ConvertBlue = _chkBlue.Checked,
                    ConvertGreen = _chkGreen.Checked,
                    ConvertRed = _chkRed.Checked,
                    ConvertPurple = _chkPurple.Checked
                };
            }
        }

        private static ExportSelectionOptions ShowExportOptionsDialog(string suggestedName)
        {
            using (var dlg = new ExportSelectionDialog(suggestedName))
            {
                var result = dlg.ShowDialog();
                if (result == DialogResult.OK && dlg.Options != null)
                    return dlg.Options;
            }
            return null;
        }

        private static IList<string> BuildFontColorTargets(ExportSelectionOptions options)
        {
            var list = new List<string>();
            if (options == null) return list;

            // Hex colors provided:
            // #800080 (Purple), #FF0000 (Red), #008000 (Green), #0000FF (Blue)
            if (options.ConvertPurple) list.Add("#800080");
            if (options.ConvertRed) list.Add("#FF0000");
            if (options.ConvertGreen) list.Add("#008000");
            if (options.ConvertBlue) list.Add("#0000FF");

            return list;
        }

        public static void ExportSelection(Excel.Application app)
        {
            using (new ExcelUiSilencer(app))
            {
                Excel.Workbook srcWb = null;
                Excel.Sheets selSheets = null;
                string tempSourcePath = null;
                var originalStatus = app.StatusBar;

                try
                {
                    srcWb = app?.ActiveWorkbook;
                    if (srcWb == null)
                    {
                        MessageBox.Show("No active workbook.");
                        return;
                    }

                    // Require the workbook to be saved somewhere (local or SharePoint/OneDrive)
                    if (string.IsNullOrEmpty(srcWb.Path))
                    {
                        MessageBox.Show(
                            "Please save the workbook before exporting selection.",
                            "Export Selection",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        return;
                    }

                    selSheets = app.ActiveWindow?.SelectedSheets as Excel.Sheets;
                    if (selSheets == null || selSheets.Count == 0)
                    {
                        MessageBox.Show("Select at least one sheet first.");
                        return;
                    }

                    var sheetNames = new string[selSheets.Count];
                    for (int i = 1; i <= selSheets.Count; i++)
                    {
                        dynamic sheet = selSheets[i];
                        sheetNames[i - 1] = sheet.Name;

                        // Release the individual sheet COM object
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet); }
                        catch { }
                    }
                    var keptSheets = new HashSet<string>(sheetNames, StringComparer.OrdinalIgnoreCase);

                    string workbookName = srcWb.Name;
                    string suggested = Path.GetFileNameWithoutExtension(workbookName) + "_External";

                    var options = ShowExportOptionsDialog(suggested);
                    if (options == null) return;
                    string outName = options.WorkbookName;
                    var colorTargets = BuildFontColorTargets(options);

                    // ---- Destination handling (local vs SharePoint/OneDrive/WOPI)
                    string destinationPathLocal;
                    string remoteSaveTarget = null;   // non-null => SaveAs back to remote folder
                    bool useTempDest = false;

                    bool isFilePath = IsFileSystemPath(srcWb.FullName);

                    if (isFilePath && !string.IsNullOrEmpty(srcWb.Path) && Directory.Exists(srcWb.Path))
                    {
                        // Normal local/UNC workbook – export into same folder
                        destinationPathLocal = Path.Combine(srcWb.Path, outName + ".xlsx");

                        if (File.Exists(destinationPathLocal))
                        {
                            if (IsFileLocked(destinationPathLocal))
                            {
                                destinationPathLocal = Path.Combine(
                                    Path.GetTempPath(),
                                    outName + "_" + Guid.NewGuid() + ".xlsx");
                                useTempDest = true;

                                MessageBox.Show(
                                    $"The file \"{outName}.xlsx\" is currently open.\r\n\r\nThe export will open in a new window.",
                                    "Export Selection",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }
                            else
                            {
                                var resp = MessageBox.Show(
                                    $"\"{outName}.xlsx\" already exists. Overwrite?",
                                    "Export Selection",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question);
                                if (resp != DialogResult.Yes) return;

                                try { File.Delete(Path.Combine(srcWb.Path, outName + ".xlsx")); }
                                catch (Exception delEx)
                                {
                                    MessageBox.Show(
                                        $"Cannot delete existing file: {delEx.Message}",
                                        "Export Selection",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }
                    else
                    {
                        // SharePoint / OneDrive / WOPI / non-file:
                        // 1) Local temp target for OOXML pipeline
                        destinationPathLocal = Path.Combine(
                            Path.GetTempPath(),
                            outName + "_" + Guid.NewGuid() + ".xlsx");

                        // 2) Remote URL in same folder as the source workbook
                        try
                        {
                            string folderUrl = srcWb.Path;
                            if (string.IsNullOrEmpty(folderUrl))
                            {
                                string srcUrl = srcWb.FullName ?? string.Empty;
                                int lastSlash = srcUrl.LastIndexOf('/');
                                if (lastSlash > 0)
                                    folderUrl = srcUrl.Substring(0, lastSlash);
                            }

                            if (!string.IsNullOrEmpty(folderUrl) &&
                                (folderUrl.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                                 folderUrl.StartsWith("https://", StringComparison.OrdinalIgnoreCase)))
                            {
                                string fileNameComponent = outName + ".xlsx";
                                remoteSaveTarget = folderUrl.TrimEnd('/') + "/" +
                                                   Uri.EscapeDataString(fileNameComponent);
                            }
                        }
                        catch
                        {
                            remoteSaveTarget = null;
                        }
                    }

                    // ---- Make a local working copy (supports SharePoint/WOPI)
                    tempSourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
                    try { app.StatusBar = "Creating local working copy..."; } catch { }

                    try
                    {
                        if (IsFileSystemPath(srcWb.FullName) && File.Exists(srcWb.FullName))
                        {
                            // Local/UNC path – fast binary copy
                            CopyLocalWorkbookFast(srcWb.FullName, tempSourcePath);
                        }
                        else
                        {
                            // SharePoint/WOPI/URL – let Excel write a local copy
                            srcWb.SaveCopyAs(tempSourcePath);
                        }
                    }
                    catch (IOException)
                    {
                        // Fallback to SaveCopyAs on I/O issues
                        try
                        {
                            srcWb.SaveCopyAs(tempSourcePath);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(
                                $"Failed to create working copy: {ex.Message}",
                                "Export Selection",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Error);
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Failed to create working copy: {ex.Message}",
                            "Export Selection",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
                        return;
                    }

                    // ---- OOXML rewrite: formulas + font colors + external links
                    try { app.StatusBar = "Rewriting formulas, fonts, and pruning externals..."; } catch { }
                    OOXmlExportSanitizer.RewriteForExport(tempSourcePath, keptSheets, colorTargets);

                    // Step 1: Copy sanitized temp file to local destination
                    try { app.StatusBar = "Creating destination workbook..."; } catch { }
                    CopyLocalWorkbookFast(tempSourcePath, destinationPathLocal);

                    // Step 2: Prune unwanted sheets and finalize
                    try { app.StatusBar = "Finalizing export (pruning and formatting)..."; } catch { }
                    try
                    {
                        SheetPruner.Prune(destinationPathLocal, keptSheets);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Error during finalization: {ex}");
                        MessageBox.Show(
                            $"An error occurred during finalization.\r\n\r\nError: {ex.Message}\r\n\r\nThe exported file might require repair.",
                            "Export Selection Warning",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning);
                        // Proceed to open the potentially incomplete file.
                    }

                    // ---- Open result and clean up (local vs remote)
                    Excel.Workbook destWb = null;
                    try
                    {
                        destWb = app.Workbooks.Open(destinationPathLocal);

                        if (!string.IsNullOrEmpty(remoteSaveTarget))
                        {
                            // Cloud / SharePoint / OneDrive / WOPI: SaveAs back to same library/folder
                            try
                            {
                                destWb.SaveAs(
                                    remoteSaveTarget,
                                    FileFormat: Excel.XlFileFormat.xlOpenXMLWorkbook);

                                MessageBox.Show(
                                    $"Exported to: {remoteSaveTarget}",
                                    "Export Selection",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);

                                // Try to delete the local temp copy
                                try
                                {
                                    if (File.Exists(destinationPathLocal))
                                        File.Delete(destinationPathLocal);
                                }
                                catch { }
                            }
                            catch (Exception saveEx)
                            {
                                MessageBox.Show(
                                    "The exported workbook was created locally but could not be saved back " +
                                    "to the original SharePoint/OneDrive location.\r\n\r\n" +
                                    "A local copy is open in Excel. You can use File ▸ Save As to store it.\r\n\r\n" +
                                    "Error: " + saveEx.Message,
                                    "Export Selection",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning);
                            }
                        }
                        else
                        {
                            // Purely local/UNC behavior
                            if (useTempDest)
                            {
                                Task.Delay(1500).ContinueWith(_ =>
                                {
                                    try
                                    {
                                        if (File.Exists(destinationPathLocal))
                                            File.Delete(destinationPathLocal);
                                    }
                                    catch { }
                                });

                                MessageBox.Show(
                                    $"Exported: {outName}.xlsx\r\n\r\nPlease save it to your preferred location.",
                                    "Export Selection",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }
                            else
                            {
                                AddToRecentFiles(destinationPathLocal, outName);
                                try { app.StatusBar = $"Exported: {destinationPathLocal}"; } catch { }

                                MessageBox.Show(
                                    $"Exported to: {destinationPathLocal}",
                                    "Export Selection",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Information);
                            }
                        }
                    }
                    finally
                    {
                        try { if (originalStatus != null) app.StatusBar = originalStatus; } catch { }
                    }
                }
                catch (Exception ex)
                {
                    try { if (originalStatus != null) app.StatusBar = originalStatus; } catch { }
                    MessageBox.Show(
                        $"Export failed: {ex.Message}",
                        "ExportSelection Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                finally
                {
                    // Clean up the temporary source file
                    if (!string.IsNullOrEmpty(tempSourcePath) && File.Exists(tempSourcePath))
                    {
                        try { File.Delete(tempSourcePath); } catch { }
                    }
                    if (selSheets != null)
                    {
                        try { System.Runtime.InteropServices.Marshal.ReleaseComObject(selSheets); } catch { }
                    }
                }
            }
        }

        // ---------- OPTIMIZED UTILITIES ----------

        private static bool IsFileLocked(string filePath)
        {
            if (!File.Exists(filePath)) return false;
            try
            {
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                    return false;
            }
            catch (IOException)
            {
                return true;
            }
        }

        private static bool IsFileSystemPath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;

            if (path.StartsWith("http://", StringComparison.OrdinalIgnoreCase) ||
                path.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                return false;

            if (Uri.TryCreate(path, UriKind.Absolute, out var uri))
                return uri.IsFile;

            return Path.IsPathRooted(path);
        }

        // Optimized file copy with smart selection based on file size
        private static void CopyLocalWorkbookFast(string sourcePath, string destPath)
        {
            // Ensure destination is clear before copying
            if (File.Exists(destPath))
            {
                try { File.Delete(destPath); } catch { /* Ignore if delete fails */ }
            }

            const int bufferSize = 1 << 21; // 2 MB buffer

            // For all files, use optimized streaming with permissive FileShare
            // This works better with Excel locks than MemoryMappedFile
            using (var src = new FileStream(
                       sourcePath,
                       FileMode.Open,
                       FileAccess.Read,
                       FileShare.ReadWrite | FileShare.Delete,
                       bufferSize,
                       FileOptions.SequentialScan))
            using (var dst = new FileStream(
                       destPath,
                       FileMode.Create,
                       FileAccess.Write,
                       FileShare.None,
                       bufferSize,
                       FileOptions.SequentialScan))
            {
                src.CopyTo(dst, bufferSize);
            }
        }

        private static void AddToRecentFiles(string filePath, string fileNameNoExt)
        {
            try
            {
                if (filePath.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return;
                string recentPath = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
                string shortcutPath = Path.Combine(recentPath, Path.GetFileName(filePath) + ".lnk");

                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                dynamic shell = Activator.CreateInstance(shellType);
                var shortcut = shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = filePath;
                shortcut.Save();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);
            }
            catch (Exception ex) { Debug.WriteLine($"Failed to add to recent files: {ex.Message}"); }
        }
    }

    /// <summary>
    /// Optimized OOXML processor for export (Sanitization logic)
    /// </summary>
    internal static class OOXmlExportSanitizer
    {
        // Regex cache for performance
        private static readonly ConcurrentDictionary<string, Regex> _regexCache =
            new ConcurrentDictionary<string, Regex>();

        private static readonly Regex _errorTokenRegex = new Regex(
            @"#(REF!|NAME\?|VALUE!|DIV/0!|NULL!|N/A|NUM!)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        // Matches external workbook references like [1], [2], etc.
        private static readonly Regex _externalRefRegex = new Regex(
            @"\[\d+\]",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        public static void RewriteForExport(string filePath, ISet<string> keptSheets, IList<string> colorsToBlack = null)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                var wbEntry = zip.GetEntry("xl/workbook.xml");
                var relEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (wbEntry == null || relEntry == null) return;

                // Step 1: Build all necessary lookup data
                Dictionary<string, string> idToTarget;
                List<string> allSheetNames;
                PrepareRewriteData(zip, wbEntry, relEntry, out idToTarget, out allSheetNames);

                var nonKept = new HashSet<string>(
                    allSheetNames.Where(n => !keptSheets.Contains(n)),
                    StringComparer.OrdinalIgnoreCase);

                // --- FIX START: Identify Structured Tables on the sheets we are about to delete ---
                // We must identify the file paths of the nonKept sheets to parse their relationships.
                var wbXml = LoadXmlFromEntry(wbEntry);
                var relXml = LoadXmlFromEntry(relEntry);
                var ns = new XmlNamespaceManager(wbXml.NameTable);
                ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                var nonKeptSheetPaths = new List<string>();
                foreach (XmlElement sheet in wbXml.SelectNodes("//s:sheets/s:sheet", ns))
                {
                    var name = sheet.GetAttribute("name");
                    if (nonKept.Contains(name))
                    {
                        var rid = sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                        if (idToTarget.TryGetValue(rid, out var path))
                        {
                            nonKeptSheetPaths.Add(path);
                        }
                    }
                }

                // Get table names from these paths and add them to the neutralization list
                var deletedTableNames = GetTableNamesFromSheets(zip, nonKeptSheetPaths);
                var targetsToNeutralize = new HashSet<string>(nonKept, StringComparer.OrdinalIgnoreCase);
                foreach (var tName in deletedTableNames)
                {
                    targetsToNeutralize.Add(tName);
                }
                // --- FIX END ---

                bool hasExternalLinks = wbXml.SelectSingleNode("//s:externalReferences", ns) != null;

                // Quick exit if all sheets are kept (but still apply font conversions / external link pruning)
                if (nonKept.Count == 0)
                {
                    PruneExternalLinks(zip, wbXml, relXml);

                    if (hasExternalLinks)
                    {
                        // Re-fetch entry after PruneExternalLinks may have deleted/recreated it
                        wbEntry = zip.GetEntry("xl/workbook.xml");
                        if (wbEntry != null)
                        {
                            wbXml = LoadXmlFromEntry(wbEntry);
                            // Even if no sheets deleted, we check for external refs
                            ProcessWorksheets(zip, wbXml, keptSheets, idToTarget, null, true);
                        }
                    }

                    if (colorsToBlack != null && colorsToBlack.Count > 0)
                    {
                        ConvertFontColorsToBlack(zip, colorsToBlack);
                    }

                    CleanupContentTypes(zip);
                    return;
                }

                // Build regex including Sheet Names AND Table Names
                var neutralizationRegex = BuildSheetRefRegex(targetsToNeutralize);

                // Step 2: Remove external links (single operation)
                PruneExternalLinks(zip, wbXml, relXml);

                // Re-fetch entries after PruneExternalLinks 
                wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry == null) return;
                wbXml = LoadXmlFromEntry(wbEntry);

                // Step 3: Prune broken defined names (that reference non-kept sheets)
                bool wbChanged = PruneDefinedNames(wbXml, targetsToNeutralize);

                if (hasExternalLinks)
                {
                    if (PruneExternalDefinedNames(wbXml)) wbChanged = true;
                }

                if (wbChanged)
                {
                    SaveXmlToEntry(wbXml, wbEntry, zip);
                    wbEntry = zip.GetEntry("xl/workbook.xml");
                    if (wbEntry != null) wbXml = LoadXmlFromEntry(wbEntry);
                }

                // Step 4: Process all worksheets 
                // Pass the regex that now includes Table Names
                ProcessWorksheets(zip, wbXml, keptSheets, idToTarget, neutralizationRegex, hasExternalLinks);

                // Step 5: Convert selected font colors
                if (colorsToBlack != null && colorsToBlack.Count > 0)
                {
                    ConvertFontColorsToBlack(zip, colorsToBlack);
                }

                CleanupContentTypes(zip);
            }
        }

        /// <summary>
        /// Scans the specified sheet paths for Table definitions and returns their display names.
        /// </summary>
        private static HashSet<string> GetTableNamesFromSheets(ZipArchive zip, List<string> sheetPaths)
        {
            var tableNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var sheetPath in sheetPaths)
            {
                // Construct path to the sheet's relationships file
                // e.g., xl/worksheets/sheet1.xml -> xl/worksheets/_rels/sheet1.xml.rels
                var dir = Path.GetDirectoryName(sheetPath).Replace('\\', '/');
                var fn = Path.GetFileName(sheetPath);
                var relsPath = (string.IsNullOrEmpty(dir) ? "" : dir + "/") + "_rels/" + fn + ".rels";

                var relsEntry = zip.GetEntry(relsPath);
                if (relsEntry == null) continue;

                try
                {
                    var relsXml = LoadXmlFromEntry(relsEntry);
                    var ns = new XmlNamespaceManager(relsXml.NameTable);
                    ns.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships");

                    // Find all relationships of type "table"
                    foreach (XmlElement rel in relsXml.SelectNodes("//r:Relationship", ns))
                    {
                        var type = rel.GetAttribute("Type");
                        if (type.EndsWith("/table", StringComparison.OrdinalIgnoreCase))
                        {
                            var target = rel.GetAttribute("Target");
                            // Resolve target path (usually relative)
                            string tablePath = target;
                            if (!target.StartsWith("/"))
                            {
                                tablePath = (string.IsNullOrEmpty(dir) ? "" : dir + "/") + target;
                            }
                            tablePath = tablePath.TrimStart('/');

                            // Open the table definition to get the Display Name (used in formulas)
                            var tableEntry = zip.GetEntry(tablePath);
                            if (tableEntry != null)
                            {
                                var tableXml = LoadXmlFromEntry(tableEntry);
                                var tRoot = tableXml.DocumentElement;
                                var displayName = tRoot?.GetAttribute("displayName"); // e.g., "Table1"
                                var name = tRoot?.GetAttribute("name"); // Fallback

                                if (!string.IsNullOrEmpty(displayName)) tableNames.Add(displayName);
                                else if (!string.IsNullOrEmpty(name)) tableNames.Add(name);
                            }
                        }
                    }
                }
                catch { /* Ignore errors parsing specific sheets/tables */ }
            }

            return tableNames;
        }

        internal static void CleanupContentTypes(ZipArchive zip)
        {
            var ctEntry = zip.GetEntry("[Content_Types].xml");
            if (ctEntry == null) return;

            var ctXml = LoadXmlFromEntry(ctEntry);
            var ns = new XmlNamespaceManager(ctXml.NameTable);
            ns.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types");

            bool changed = false;
            var overrides = ctXml.SelectNodes("//ct:Override", ns);
            if (overrides != null)
            {
                var toRemove = new List<XmlElement>();
                foreach (XmlElement ov in overrides)
                {
                    var partName = ov.GetAttribute("PartName");
                    if (string.IsNullOrEmpty(partName)) continue;
                    var zipPath = partName.TrimStart('/');
                    if (zip.GetEntry(zipPath) == null)
                    {
                        toRemove.Add(ov);
                    }
                }
                foreach (var ov in toRemove)
                {
                    ov.ParentNode.RemoveChild(ov);
                    changed = true;
                }
            }

            if (changed) SaveXmlToEntry(ctXml, ctEntry, zip);
        }

        private static bool PruneExternalDefinedNames(XmlDocument wbXml)
        {
            var ns = new XmlNamespaceManager(wbXml.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var dnNodes = wbXml.SelectNodes("//s:definedNames/s:definedName", ns);
            if (dnNodes == null || dnNodes.Count == 0) return false;

            var toRemove = new List<XmlElement>();
            foreach (XmlElement dn in dnNodes)
            {
                var text = dn.InnerText ?? string.Empty;
                if (_externalRefRegex.IsMatch(text))
                    toRemove.Add(dn);
            }

            if (toRemove.Count > 0)
            {
                foreach (var dn in toRemove) dn.ParentNode.RemoveChild(dn);
                return true;
            }
            return false;
        }

        private static void PrepareRewriteData(
            ZipArchive zip,
            ZipArchiveEntry wbEntry,
            ZipArchiveEntry relEntry,
            out Dictionary<string, string> idToTarget,
            out List<string> allSheetNames)
        {
            idToTarget = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using (var stream = relEntry.Open())
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true }))
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "Relationship")
                    {
                        var id = reader.GetAttribute("Id");
                        var target = reader.GetAttribute("Target");
                        if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(target))
                        {
                            var t = target.Replace('\\', '/');
                            if (t.StartsWith("/")) t = t.Substring(1);
                            else if (!target.StartsWith("http", StringComparison.OrdinalIgnoreCase) &&
                                     !target.StartsWith("mailto:", StringComparison.OrdinalIgnoreCase))
                            {
                                if (!t.StartsWith("xl/")) t = "xl/" + t;
                            }
                            idToTarget[id] = t;
                        }
                    }
                }
            }

            allSheetNames = new List<string>();
            using (var stream = wbEntry.Open())
            using (var reader = XmlReader.Create(stream, new XmlReaderSettings { IgnoreWhitespace = true }))
            {
                while (reader.ReadToFollowing("sheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main"))
                {
                    var name = reader.GetAttribute("name");
                    if (!string.IsNullOrEmpty(name)) allSheetNames.Add(name);
                }
            }
        }

        private static bool PruneDefinedNames(XmlDocument wbXml, ISet<string> namesToNeutralize)
        {
            var ns = new XmlNamespaceManager(wbXml.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var dnNodes = wbXml.SelectNodes("//s:definedNames/s:definedName", ns);
            if (dnNodes == null || dnNodes.Count == 0) return false;

            var toRemove = new List<XmlElement>();
            foreach (XmlElement dn in dnNodes)
            {
                var text = dn.InnerText ?? string.Empty;
                if (ContainsAnyErrorToken(text) || MentionsAnySheet(text, namesToNeutralize))
                    toRemove.Add(dn);
            }

            if (toRemove.Count > 0)
            {
                foreach (var dn in toRemove) dn.ParentNode.RemoveChild(dn);
                return true;
            }
            return false;
        }

        private static void ProcessWorksheets(
            ZipArchive zip,
            XmlDocument wbXml,
            ISet<string> keptSheets,
            Dictionary<string, string> idToTarget,
            Regex neutralizationRegex,
            bool pruneExternalRefs)
        {
            var ns = new XmlNamespaceManager(wbXml.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            var sheetsToProcess = new List<(string name, string path)>();
            foreach (XmlElement sheet in wbXml.SelectNodes("//s:sheets/s:sheet", ns))
            {
                var name = sheet.GetAttribute("name");
                if (string.IsNullOrEmpty(name) || !keptSheets.Contains(name)) continue;

                var rid = sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                if (string.IsNullOrEmpty(rid)) continue;

                if (!idToTarget.TryGetValue(rid, out var fullPath)) continue;

                var wsEntry = zip.GetEntry(fullPath);
                if (wsEntry != null && fullPath.Contains("/worksheets/"))
                    sheetsToProcess.Add((name, fullPath));
            }

            if (sheetsToProcess.Count > 5)
            {
                var sheetXmlData = new List<(string path, XmlDocument doc)>();
                foreach (var sheetInfo in sheetsToProcess)
                {
                    var entry = zip.GetEntry(sheetInfo.path);
                    if (entry != null)
                    {
                        var wsXml = LoadXmlFromEntry(entry);
                        sheetXmlData.Add((sheetInfo.path, wsXml));
                    }
                }

                var updates = new ConcurrentBag<(string path, XmlDocument doc)>();
                Parallel.ForEach(sheetXmlData,
                    new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount },
                    item =>
                    {
                        if (ProcessWorksheetFormulas(item.doc, neutralizationRegex, pruneExternalRefs))
                            updates.Add((item.path, item.doc));
                    });

                foreach (var item in updates)
                {
                    var entry = zip.GetEntry(item.path);
                    if (entry != null)
                        SaveXmlToEntry(item.doc, entry, zip);
                }
            }
            else
            {
                foreach (var sheetInfo in sheetsToProcess)
                {
                    var entry = zip.GetEntry(sheetInfo.path);
                    if (entry == null) continue;

                    var wsXml = LoadXmlFromEntry(entry);
                    if (ProcessWorksheetFormulas(wsXml, neutralizationRegex, pruneExternalRefs))
                    {
                        entry = zip.GetEntry(sheetInfo.path);
                        if (entry != null)
                            SaveXmlToEntry(wsXml, entry, zip);
                    }
                }
            }
        }

        private static bool ProcessWorksheetFormulas(XmlDocument wsXml, Regex neutralizationRegex, bool pruneExternalRefs)
        {
            var sns = new XmlNamespaceManager(wsXml.NameTable);
            sns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            bool changed = false;
            var formulaNodes = wsXml.SelectNodes("//s:c[s:f]/s:f", sns);
            if (formulaNodes == null || formulaNodes.Count == 0) return false;

            foreach (XmlElement fNode in formulaNodes)
            {
                string txt = fNode.InnerText ?? string.Empty;
                string fType = fNode.GetAttribute("t");
                bool hasRef = fNode.HasAttribute("ref");

                // Check 1: Does it mention a deleted sheet OR a deleted Table Name?
                bool mentionsNeutralizedTarget =
                    neutralizationRegex != null &&
                    neutralizationRegex.IsMatch(txt);

                // Check 2: Does it mention external workbook [n]?
                bool mentionsExternalWorkbook =
                    pruneExternalRefs && _externalRefRegex.IsMatch(txt);

                bool shouldNeutralize = mentionsNeutralizedTarget || mentionsExternalWorkbook;

                if (fType == "array" && hasRef)
                {
                    if (shouldNeutralize)
                    {
                        string spillRef = fNode.GetAttribute("ref");
                        if (!string.IsNullOrWhiteSpace(spillRef))
                        {
                            if (BlankOutFormulaRange(wsXml, spillRef, sns))
                                changed = true;
                        }
                    }
                    continue;
                }

                if (shouldNeutralize)
                {
                    var cell = (XmlElement)fNode.ParentNode;
                    var vNode = (XmlElement)cell.SelectSingleNode("s:v", sns);
                    if (vNode == null)
                    {
                        vNode = wsXml.CreateElement("v", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                        vNode.InnerText = "";
                        cell.AppendChild(vNode);
                    }
                    cell.RemoveChild(fNode);
                    changed = true;
                }
            }

            return changed;
        }

        private static bool BlankOutFormulaRange(XmlDocument wsXml, string rangeRef, XmlNamespaceManager ns)
        {
            if (string.IsNullOrWhiteSpace(rangeRef)) return false;

            if (!TryParseRange(rangeRef, out int row1, out int col1, out int row2, out int col2))
                return false;

            bool changed = false;
            for (int r = row1; r <= row2; r++)
            {
                for (int c = col1; c <= col2; c++)
                {
                    string addr = ColumnIndexToLetters(c) + r.ToString();
                    var cell = wsXml.SelectSingleNode($"//s:c[@r='{addr}']", ns) as XmlElement;
                    if (cell == null) continue;

                    var fNode = cell.SelectSingleNode("s:f", ns) as XmlElement;
                    if (fNode != null)
                    {
                        cell.RemoveChild(fNode);
                        changed = true;
                    }

                    var vNode = cell.SelectSingleNode("s:v", ns) as XmlElement;
                    if (vNode == null)
                    {
                        vNode = wsXml.CreateElement("v", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                        vNode.InnerText = "";
                        cell.AppendChild(vNode);
                        changed = true;
                    }
                }
            }
            return changed;
        }

        private static bool TryParseRange(string rangeRef, out int row1, out int col1, out int row2, out int col2)
        {
            row1 = col1 = row2 = col2 = 0;
            if (string.IsNullOrWhiteSpace(rangeRef)) return false;

            rangeRef = rangeRef.Trim();
            int bangIndex = rangeRef.LastIndexOf('!');
            if (bangIndex >= 0 && bangIndex < rangeRef.Length - 1)
                rangeRef = rangeRef.Substring(bangIndex + 1);

            var parts = rangeRef.Split(':');
            if (parts.Length == 1)
            {
                if (!TryParseCellRef(parts[0], out row1, out col1)) return false;
                row2 = row1;
                col2 = col1;
                return true;
            }
            else if (parts.Length >= 2)
            {
                if (!TryParseCellRef(parts[0], out row1, out col1)) return false;
                if (!TryParseCellRef(parts[1], out row2, out col2)) return false;
                if (row2 < row1) { int tmp = row1; row1 = row2; row2 = tmp; }
                if (col2 < col1) { int tmp = col1; col1 = col2; col2 = tmp; }
                return true;
            }
            return false;
        }

        private static bool TryParseCellRef(string cellRef, out int row, out int col)
        {
            row = 0; col = 0;
            if (string.IsNullOrWhiteSpace(cellRef)) return false;
            cellRef = cellRef.Trim();
            int i = 0;
            while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
            if (i == 0 || i == cellRef.Length) return false;
            string colPart = cellRef.Substring(0, i);
            string rowPart = cellRef.Substring(i);
            if (!int.TryParse(rowPart, out row) || row <= 0) return false;
            col = ColumnLettersToIndex(colPart);
            return col > 0;
        }

        private static int ColumnLettersToIndex(string letters)
        {
            if (string.IsNullOrEmpty(letters)) return -1;
            letters = letters.ToUpperInvariant();
            int col = 0;
            foreach (char ch in letters)
            {
                if (ch < 'A' || ch > 'Z') return -1;
                col = col * 26 + (ch - 'A' + 1);
            }
            return col;
        }

        private static string ColumnIndexToLetters(int columnIndex)
        {
            if (columnIndex <= 0) throw new ArgumentOutOfRangeException(nameof(columnIndex));
            var sb = new StringBuilder();
            while (columnIndex > 0)
            {
                int rem = (columnIndex - 1) % 26;
                sb.Insert(0, (char)('A' + rem));
                columnIndex = (columnIndex - 1) / 26;
            }
            return sb.ToString();
        }

        private static Regex BuildSheetRefRegex(ISet<string> namesToMatch)
        {
            if (namesToMatch == null || namesToMatch.Count == 0) return null;

            // Cache key based on sorted names
            var key = string.Join("|", namesToMatch.OrderBy(s => s, StringComparer.OrdinalIgnoreCase));

            return _regexCache.GetOrAdd(key, _ =>
            {
                var plain = string.Join("|", namesToMatch.Select(Regex.Escape));
                var quoted = string.Join("|", namesToMatch.Select(n => Regex.Escape(n.Replace("'", "''"))));

                // Matches a sheet reference ('Quoted Name'! or PlainName!) or a structured
                // table reference (TableName[...). A bare name isn't enough: it could be a
                // function name or text inside a string.
                // (?:'(?:quoted)'|(?:\bplain\b))(?=\!|\[)

                var pattern = @"(?i)(?:'(?:" + quoted + @")'|(?:\b(?:" + plain + @")\b))(?=\!|\[)";

                return new Regex(pattern, RegexOptions.Compiled | RegexOptions.CultureInvariant);
            });
        }

        private static bool MentionsAnySheet(string s, ISet<string> names)
        {
            if (string.IsNullOrEmpty(s) || names == null || names.Count == 0) return false;
            var regex = BuildSheetRefRegex(names);
            return regex != null && regex.IsMatch(s);
        }

        private static void PruneExternalLinks(ZipArchive zip, XmlDocument wbXml, XmlDocument relXml)
        {
            var ns = new XmlNamespaceManager(wbXml.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
            var rns = new XmlNamespaceManager(relXml.NameTable);
            rns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

            bool wbChanged = false;
            bool relChanged = false;

            var extRefs = wbXml.SelectSingleNode("//s:externalReferences", ns);
            if (extRefs != null)
            {
                extRefs.ParentNode.RemoveChild(extRefs);
                wbChanged = true;
            }

            var toDelete = new List<XmlElement>();
            foreach (XmlElement rel in relXml.SelectNodes("//rel:Relationship", rns))
            {
                var type = rel.GetAttribute("Type");
                if (string.Equals(type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink",
                                  StringComparison.OrdinalIgnoreCase))
                    toDelete.Add(rel);
            }

            if (toDelete.Count > 0)
            {
                foreach (var rel in toDelete)
                {
                    var target = rel.GetAttribute("Target").Replace('\\', '/');
                    string fullPath;
                    if (target.StartsWith("/")) fullPath = target.Substring(1);
                    else fullPath = "xl/" + target;

                    var e = zip.GetEntry(fullPath);
                    if (e != null) e.Delete();

                    var relsDir = Path.GetDirectoryName(fullPath).Replace('\\', '/');
                    var relsPath = $"{relsDir}/_rels/{Path.GetFileName(fullPath)}.rels";
                    var rels = zip.GetEntry(relsPath);
                    if (rels != null) rels.Delete();

                    rel.ParentNode.RemoveChild(rel);
                }
                relChanged = true;
            }

            if (wbChanged)
            {
                var wbEntry = zip.GetEntry("xl/workbook.xml");
                if (wbEntry != null) SaveXmlToEntry(wbXml, wbEntry, zip);
            }
            if (relChanged)
            {
                var relEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (relEntry != null) SaveXmlToEntry(relXml, relEntry, zip);
            }
        }

        private static void ConvertFontColorsToBlack(ZipArchive zip, IList<string> colorsToBlack)
        {
            if (colorsToBlack == null || colorsToBlack.Count == 0) return;

            var argbTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var raw in colorsToBlack)
            {
                if (string.IsNullOrWhiteSpace(raw)) continue;
                var t = raw.Trim();
                if (t.StartsWith("#")) t = t.Substring(1);
                if (t.Length == 6) t = "FF" + t;
                else if (t.Length != 8) continue;
                argbTargets.Add(t.ToUpperInvariant());
            }

            if (argbTargets.Count == 0) return;
            const string BlackArgb = "FF000000";

            var stylesEntry = zip.GetEntry("xl/styles.xml");
            if (stylesEntry != null)
            {
                var stylesXml = LoadXmlFromEntry(stylesEntry);
                var ns = new XmlNamespaceManager(stylesXml.NameTable);
                ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                bool changed = false;
                var fontColorNodes = stylesXml.SelectNodes("//s:fonts/s:font/s:color", ns);
                if (fontColorNodes != null)
                {
                    foreach (XmlElement colorNode in fontColorNodes)
                    {
                        var rgbAttr = colorNode.GetAttribute("rgb");
                        if (!string.IsNullOrEmpty(rgbAttr))
                        {
                            var normalized = rgbAttr.Trim();
                            if (normalized.StartsWith("#")) normalized = normalized.Substring(1);
                            if (normalized.Length == 6) normalized = "FF" + normalized;

                            if (argbTargets.Contains(normalized))
                            {
                                colorNode.RemoveAttribute("theme");
                                colorNode.RemoveAttribute("indexed");
                                colorNode.SetAttribute("rgb", BlackArgb);
                                changed = true;
                            }
                        }
                    }
                }
                if (changed) SaveXmlToEntry(stylesXml, stylesEntry, zip);
            }

            var textEntries = zip.Entries.Where(e =>
            {
                var path = e.FullName.Replace('\\', '/');
                return string.Equals(path, "xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase) ||
                       path.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase);
            }).ToList();

            foreach (var entry in textEntries)
            {
                var doc = LoadXmlFromEntry(entry);
                var ns = new XmlNamespaceManager(doc.NameTable);
                ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

                bool changed = false;
                var colorNodes = doc.SelectNodes("//s:rPr/s:color", ns);
                if (colorNodes != null)
                {
                    foreach (XmlElement colorNode in colorNodes)
                    {
                        var rgbAttr = colorNode.GetAttribute("rgb");
                        if (!string.IsNullOrEmpty(rgbAttr))
                        {
                            var normalized = rgbAttr.Trim();
                            if (normalized.StartsWith("#")) normalized = normalized.Substring(1);
                            if (normalized.Length == 6) normalized = "FF" + normalized;
                            if (argbTargets.Contains(normalized))
                            {
                                colorNode.RemoveAttribute("theme");
                                colorNode.RemoveAttribute("indexed");
                                colorNode.SetAttribute("rgb", BlackArgb);
                                changed = true;
                            }
                        }
                    }
                }
                if (changed) SaveXmlToEntry(doc, entry, zip);
            }
        }

        internal static XmlDocument LoadXmlFromEntry(ZipArchiveEntry entry)
        {
            using (var s = entry.Open())
            using (var ms = new MemoryStream())
            {
                s.CopyTo(ms);
                ms.Position = 0;
                var xd = new XmlDocument { PreserveWhitespace = true };
                xd.Load(ms);
                return xd;
            }
        }

        internal static void SaveXmlToEntry(XmlDocument doc, ZipArchiveEntry entry, ZipArchive zip)
        {
            var path = entry.FullName;
            entry.Delete();
            var newEntry = zip.CreateEntry(path, CompressionLevel.Optimal);
            using (var s = newEntry.Open())
            using (var xw = XmlWriter.Create(s, new XmlWriterSettings { Encoding = new UTF8Encoding(false), Indent = false }))
            {
                doc.Save(xw);
            }
        }

        private static bool ContainsAnyErrorToken(string formula)
        {
            if (string.IsNullOrEmpty(formula)) return false;
            return _errorTokenRegex.IsMatch(formula);
        }
    }

    /// <summary>
    /// Optimized OOXML processor to remove unwanted sheets from a workbook package.
    /// Ensures workbook integrity by recursively deleting parts, updating indices (localSheetId), and handling visibility.
    /// </summary>
    internal static class SheetPruner
    {
        public static void Prune(string filePath, ISet<string> keptSheets)
        {
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            using (var zip = new ZipArchive(fs, ZipArchiveMode.Update))
            {
                var wbEntry = zip.GetEntry("xl/workbook.xml");
                var relEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");
                if (wbEntry == null || relEntry == null) return;

                // Relying on helper methods from OOXmlExportSanitizer for XML handling
                var wbXml = OOXmlExportSanitizer.LoadXmlFromEntry(wbEntry);
                var relXml = OOXmlExportSanitizer.LoadXmlFromEntry(relEntry);

                var ns = new XmlNamespaceManager(wbXml.NameTable);
                ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");
                ns.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

                var rns = new XmlNamespaceManager(relXml.NameTable);
                rns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

                var sheetsNode = wbXml.SelectSingleNode("//s:sheets", ns);
                if (sheetsNode == null) return;
                var sheets = wbXml.SelectNodes("//s:sheets/s:sheet", ns);
                if (sheets == null) return;

                // 1. Analyze sheets and relationships
                var sheetData = new List<(int index, XmlElement sheet, string rid, string targetPath)>();
                var keptSheetTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                for (int i = 0; i < sheets.Count; i++)
                {
                    var sheet = (XmlElement)sheets[i];
                    var rid = sheet.GetAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
                    if (string.IsNullOrEmpty(rid)) continue;

                    var relNode = relXml.SelectSingleNode($"//rel:Relationship[@Id='{rid}']", rns) as XmlElement;
                    if (relNode == null) continue;

                    var target = relNode.GetAttribute("Target");
                    if (string.IsNullOrEmpty(target)) continue;

                    // Resolve path relative to the package root. Relationships in workbook.xml.rels are relative to xl/.
                    string fullPath;
                    if (target.StartsWith("/")) fullPath = target.Substring(1);
                    else fullPath = "xl/" + target;
                    fullPath = fullPath.Replace('\\', '/');

                    sheetData.Add((i, sheet, rid, fullPath));
                }

                // 2. Identify sheets to remove and map indices (Critical for localSheetId)
                var oldToNewIndexMap = new Dictionary<int, int>();
                int newIndex = 0;

                foreach (var item in sheetData)
                {
                    var index = item.index;
                    var sheet = item.sheet;
                    var targetPath = item.targetPath;

                    var name = sheet.GetAttribute("name");
                    // Case-insensitive comparison for sheet names
                    if (keptSheets.Contains(name))
                    {
                        oldToNewIndexMap[index] = newIndex++;
                        keptSheetTargets.Add(targetPath);
                    }
                    else
                    {
                        oldToNewIndexMap[index] = -1; // Mark as removed
                    }
                }

                // Check if we are removing everything (invalid workbook state)
                if (newIndex == 0 && sheetData.Count > 0)
                {
                    throw new InvalidOperationException("Cannot export an empty workbook. At least one sheet must be kept.");
                }

                // Check if we are removing nothing
                if (newIndex == sheetData.Count)
                {
                    return;
                }

                bool wbChanged = false;
                bool relChanged = false;

                // 3. Remove sheets and their parts recursively
                foreach (var item in sheetData)
                {
                    var index = item.index;
                    var sheet = item.sheet;
                    var rid = item.rid;
                    var targetPath = item.targetPath;

                    if (oldToNewIndexMap[index] == -1)
                    {
                        // Remove from workbook.xml
                        sheetsNode.RemoveChild(sheet);
                        wbChanged = true;

                        // Remove from workbook.xml.rels
                        var relNode = relXml.SelectSingleNode($"//rel:Relationship[@Id='{rid}']", rns) as XmlElement;
                        if (relNode != null)
                        {
                            relNode.ParentNode.RemoveChild(relNode);
                            relChanged = true;
                        }

                        // Delete the sheet part and all its related parts recursively.
                        DeletePartAndRels(zip, targetPath);
                    }
                }

                // 4. Prune and update Defined Names (localSheetId indices)
                if (PruneAndUpdateDefinedNames(wbXml, oldToNewIndexMap)) wbChanged = true;

                // 5. Clean up Calculation Chain (xl/calcChain.xml). Removing it forces Excel to rebuild it on open.
                var calcChainEntry = zip.GetEntry("xl/calcChain.xml");
                if (calcChainEntry != null) calcChainEntry.Delete();

                // 6. Ensure visibility and update activeTab index
                EnsureVisibilityAndActiveTab(wbXml);
                wbChanged = true;

                // 7. Save changes to workbook and relationships XML
                if (wbChanged) OOXmlExportSanitizer.SaveXmlToEntry(wbXml, wbEntry, zip);
                if (relChanged) OOXmlExportSanitizer.SaveXmlToEntry(relXml, relEntry, zip);

                // Clean up [Content_Types].xml after all part deletions
                OOXmlExportSanitizer.CleanupContentTypes(zip);
            }
        }

        // --- Workbook Integrity Functions ---

        private static bool PruneAndUpdateDefinedNames(XmlDocument wbXml, Dictionary<int, int> oldToNewIndexMap)
        {
            var ns = new XmlNamespaceManager(wbXml.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            var definedNamesNode = wbXml.SelectSingleNode("//s:definedNames", ns);
            if (definedNamesNode == null) return false;

            var definedNames = wbXml.SelectNodes("//s:definedNames/s:definedName", ns);
            if (definedNames == null) return false;

            bool changed = false;
            var toRemove = new List<XmlElement>();

            foreach (XmlElement dn in definedNames)
            {
                // Process local defined names (scoped to a sheet via localSheetId)
                if (dn.HasAttribute("localSheetId"))
                {
                    var localSheetIdStr = dn.GetAttribute("localSheetId");
                    if (int.TryParse(localSheetIdStr, out int oldIndex))
                    {
                        if (oldToNewIndexMap.TryGetValue(oldIndex, out int newIndex))
                        {
                            if (newIndex == -1)
                            {
                                // Sheet was removed, so remove the defined name
                                toRemove.Add(dn);
                                changed = true;
                            }
                            else
                            {
                                // Sheet was kept, update the localSheetId if index changed
                                if (newIndex != oldIndex)
                                {
                                    dn.SetAttribute("localSheetId", newIndex.ToString());
                                    changed = true;
                                }
                            }
                        }
                        else
                        {
                            // Handle potentially corrupted localSheetId (out of range)
                            toRemove.Add(dn);
                            changed = true;
                        }
                    }
                }
            }

            foreach (var dn in toRemove)
            {
                definedNamesNode.RemoveChild(dn);
            }

            // If no defined names remain, remove the container element
            if (definedNamesNode.ChildNodes.Count == 0)
            {
                definedNamesNode.ParentNode?.RemoveChild(definedNamesNode);
                changed = true;
            }

            return changed;
        }

        private static void EnsureVisibilityAndActiveTab(XmlDocument wbXml)
        {
            var ns = new XmlNamespaceManager(wbXml.NameTable);
            ns.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            // 1. Ensure at least one sheet is visible
            var visibleSheets = wbXml.SelectNodes("//s:sheets/s:sheet[not(@state) or @state='visible']", ns);
            if (visibleSheets == null || visibleSheets.Count == 0)
            {
                // If all remaining sheets are hidden (or very hidden), make the first one visible.
                var firstSheet = wbXml.SelectSingleNode("//s:sheets/s:sheet", ns) as XmlElement;
                if (firstSheet != null)
                {
                    firstSheet.RemoveAttribute("state"); // Default is visible
                }
            }

            // 2. Update activeTab index in bookViews
            var bookViews = wbXml.SelectNodes("//s:bookViews/s:workbookView", ns);
            if (bookViews != null)
            {
                foreach (XmlElement bv in bookViews)
                {
                    // It is safer and simpler to reset the view to the first visible sheet (index 0).
                    bv.SetAttribute("activeTab", "0");
                }
            }
        }

        // --- Recursive Part Deletion Functions ---

        // Recursive function to delete a part and its relationships
        private static void DeletePartAndRels(ZipArchive zip, string partPath)
        {
            // Normalize path
            partPath = partPath.Replace('\\', '/').TrimStart('/');
            var entry = zip.GetEntry(partPath);
            if (entry == null) return;

            // 1. Find the relationships file for this part
            var partDir = Path.GetDirectoryName(partPath).Replace('\\', '/');
            // Handle root directory correctly for path resolution
            var resolutionBaseDir = string.IsNullOrEmpty(partDir) ? "/" : partDir;

            var relsPath = (string.IsNullOrEmpty(partDir) ? "" : partDir + "/") + "_rels/" + Path.GetFileName(partPath) + ".rels";

            var relsEntry = zip.GetEntry(relsPath);

            if (relsEntry != null)
            {
                // 2. Process relationships
                try
                {
                    var relsXml = OOXmlExportSanitizer.LoadXmlFromEntry(relsEntry);
                    var rns = new XmlNamespaceManager(relsXml.NameTable);
                    rns.AddNamespace("rel", "http://schemas.openxmlformats.org/package/2006/relationships");

                    foreach (XmlElement rel in relsXml.SelectNodes("//rel:Relationship", rns))
                    {
                        var target = rel.GetAttribute("Target");
                        if (string.IsNullOrEmpty(target)) continue;

                        // Check if it's an external link
                        var targetMode = rel.GetAttribute("TargetMode");
                        if (string.Equals(targetMode, "External", StringComparison.OrdinalIgnoreCase)) continue;

                        // Resolve path relative to the part directory
                        var resolvedPath = ResolvePartPath(resolutionBaseDir, target);

                        // 3. Recursively delete the target part
                        DeletePartAndRels(zip, resolvedPath);
                    }
                }
                catch (XmlException ex)
                {
                    Debug.WriteLine($"Error processing relationships XML {relsPath}: {ex.Message}. Skipping.");
                }

                // 4. Delete the relationships file itself
                relsEntry.Delete();
            }

            // 5. Delete the part itself
            entry.Delete();
        }

        // Helper to resolve part paths according to OpenXML conventions (OPC)
        private static string ResolvePartPath(string sourceDir, string targetPath)
        {
            targetPath = targetPath.Replace('\\', '/');
            if (targetPath.StartsWith("/"))
            {
                // Absolute path from the root of the package
                return targetPath.Substring(1);
            }
            else
            {
                // Relative path from the source directory
                return CombineUriPaths(sourceDir, targetPath);
            }
        }

        // Helper to combine URI paths (handles .. and .)
        private static string CombineUriPaths(string basePath, string relativePath)
        {
            // Ensure basePath uses forward slashes and remove empty entries (e.g. if basePath is just "/")
            basePath = basePath.Replace('\\', '/');
            var parts = new List<string>(basePath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries));

            foreach (var part in relativePath.Split('/'))
            {
                if (part == "..")
                {
                    if (parts.Count > 0) parts.RemoveAt(parts.Count - 1);
                }
                else if (part != "." && !string.IsNullOrEmpty(part))
                {
                    parts.Add(part);
                }
            }
            return string.Join("/", parts);
        }
    }
}