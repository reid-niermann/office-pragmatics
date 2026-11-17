using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;        // Interaction.InputBox

namespace OfficePragmatics_Excel.Classes
{
    public static class FileTools
    {
        // -------------------- helpers (cloud/WOPI aware) --------------------
        private static bool IsCloudDoc(Excel.Workbook wb)
        {
            string full = wb?.FullName ?? "";
            string path = wb?.Path ?? "";
            return full.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("http", StringComparison.OrdinalIgnoreCase);
        }

        private static string TempPdfPathFor(string fileBaseName)
        {
            string safe = Regex.Replace(fileBaseName ?? "Workbook", @"[^\w\-. ]+", "_");
            return System.IO.Path.Combine(System.IO.Path.GetTempPath(), safe + ".pdf");
        }

        private static void OpenWithShell(string pathOrUrl)
        {
            try { Process.Start(new ProcessStartInfo(pathOrUrl) { UseShellExecute = true }); } catch { }
        }

        // -------------------- VERSION UP --------------------
        public static void VersionUp(Excel.Application app)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                if (wb == null) return;

                if (string.IsNullOrEmpty(wb.Name))
                {
                    MessageBox.Show("This workbook has not been saved yet. Please save it before creating a new version.",
                                    "VersionUp Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string full = wb.FullName;                // may be URL for cloud/WOPI
                string dir = wb.Path;                    // may be URL for cloud/WOPI
                string nameOnly = wb.Name;                // e.g., "Model.xlsx"
                string baseNm = System.IO.Path.GetFileNameWithoutExtension(nameOnly);
                string ext = System.IO.Path.GetExtension(nameOnly);
                bool isCloud = IsCloudDoc(wb);

                var m = Regex.Match(baseNm, @"[ _]v(\d+)$", RegexOptions.IgnoreCase);

                if (!m.Success)
                {
                    var add = MessageBox.Show(
                        $"The file \"{nameOnly}\" is not versioned.\n\n" +
                        $"Add versioning?\n\nThis will create '{baseNm} _v1{ext}' and save current as '{baseNm} _v2{ext}'.",
                        "Add Versioning?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (add == DialogResult.No) return;

                    string v1Name = $"{baseNm} _v1{ext}";
                    string v2Name = $"{baseNm} _v2{ext}";

                    if (isCloud)
                    {
                        // Cloud/WOPI: filename-only; Excel saves next to the current file in the same cloud folder
                        wb.SaveCopyAs(v1Name);             // create _v1
                        wb.SaveAs(v2Name, AddToMru: true); // current becomes _v2
                        MessageBox.Show($"Versioning added.\nNew version: \"{v2Name}\"",
                            "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Local file: archive move, conflict handling, etc.
                        string v2_full = System.IO.Path.Combine(dir, v2Name);
                        string archiveFolder = FindArchiveFolder(dir);
                        if (archiveFolder == null)
                        {
                            var create = MessageBox.Show("No 'Archive' or 'Old' folder found. Create an 'Archive' folder to store the _v1 file?",
                                                         "Create Archive Folder?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (create == DialogResult.Yes)
                            {
                                archiveFolder = System.IO.Path.Combine(dir, "Archive");
                                System.IO.Directory.CreateDirectory(archiveFolder);
                            }
                        }
                        string v1_destination = (archiveFolder != null)
                            ? System.IO.Path.Combine(archiveFolder, v1Name)
                            : System.IO.Path.Combine(dir, v1Name);

                        if (System.IO.File.Exists(v1_destination) || System.IO.File.Exists(v2_full))
                        {
                            MessageBox.Show($"Cannot add versioning because '{System.IO.Path.GetFileName(v1_destination)}' or '{System.IO.Path.GetFileName(v2_full)}' already exists.",
                                            "VersionUp Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        wb.SaveAs(v2_full, AddToMru: true);        // write _v2
                        if (!string.IsNullOrEmpty(full) && System.IO.File.Exists(full))
                        {
                            System.IO.File.Move(full, v1_destination);   // move old to _v1 (Archive if chosen)
                            MoveCorrespondingPdfToArchive(full, archiveFolder);
                        }

                        AddToRecentFiles(v2_full, v2Name.Replace(ext, ""));
                        MessageBox.Show($"Versioning added.\nNew version: \"{v2Name}\"\nOriginal moved to: \"{v1_destination}\"",
                                        "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // already versioned: bump to next
                    int next = int.Parse(m.Groups[1].Value) + 1;
                    string prefix = baseNm.Substring(0, m.Index);
                    string newBase = $"{prefix}_v{next}";
                    string newName = newBase + ext;

                    if (isCloud)
                    {
                        wb.SaveAs(newName, AddToMru: true);
                        MessageBox.Show($"Saved new version: \"{newName}\"",
                                        "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string newFull = System.IO.Path.Combine(dir, newName);
                        wb.SaveAs(newFull, AddToMru: true);

                        string archive = FindArchiveFolder(dir);
                        if (archive != null && !string.IsNullOrEmpty(full) && System.IO.File.Exists(full))
                        {
                            string dest = System.IO.Path.Combine(archive, System.IO.Path.GetFileName(full));
                            System.IO.File.Move(full, dest);
                            MoveCorrespondingPdfToArchive(full, archive);
                        }

                        AddToRecentFiles(newFull, newBase);
                        MessageBox.Show($"Save complete: \"{newName}\"",
                                        "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was a problem saving a new version.\r\n" + ex.Message,
                                "VersionUp Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // -------------------- SAVE PDF --------------------
        public static void SavePDF(Excel.Application app)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                if (wb == null) return;

                if (string.IsNullOrEmpty(wb.Name))
                {
                    MessageBox.Show("This workbook has not been saved yet. Please save it before exporting to PDF.",
                                    "Save PDF Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool isCloud = IsCloudDoc(wb);
                string fileBase = System.IO.Path.GetFileNameWithoutExtension(wb.Name);
                string pdfPath;

                if (isCloud)
                {
                    // Always export to local temp when cloud/WOPI
                    pdfPath = TempPdfPathFor(fileBase);
                }
                else
                {
                    string folder = wb.Path;
                    pdfPath = System.IO.Path.Combine(folder, fileBase + ".pdf");

                    if (System.IO.File.Exists(pdfPath))
                    {
                        var resp = MessageBox.Show("File already exists.\r\nYes = overwrite, No = rename, Cancel = abort",
                                                   "Save PDF", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                        if (resp == DialogResult.Cancel) return;
                        if (resp == DialogResult.No)
                        {
                            string input;
                            do
                            {
                                input = Interaction.InputBox("Enter new PDF name (no extension):", "Save PDF", fileBase);
                                if (string.IsNullOrEmpty(input)) return;
                            }
                            while (string.IsNullOrWhiteSpace(input) || input.IndexOfAny(System.IO.Path.GetInvalidFileNameChars()) >= 0);

                            pdfPath = System.IO.Path.Combine(folder, input + ".pdf");
                            fileBase = input;
                        }
                    }
                }

                // Export active sheet
                wb.ActiveSheet.ExportAsFixedFormat(
                    Type: Excel.XlFixedFormatType.xlTypePDF,
                    Filename: pdfPath,
                    Quality: Excel.XlFixedFormatQuality.xlQualityStandard,
                    IncludeDocProperties: true,
                    IgnorePrintAreas: false,
                    OpenAfterPublish: false);

                MessageBox.Show($"PDF saved to:\r\n{pdfPath}", "Save PDF", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!isCloud) AddToRecentFiles(pdfPath, fileBase);
                OpenWithShell(pdfPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem saving PDF:\r\n" + ex.Message, "Save PDF Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // -------------------- OPEN CONTAINING FOLDER --------------------
        public static void OpenContainingFolder()
        {
            var wb = OfficePragmatics_Excel.Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wb == null || string.IsNullOrEmpty(wb.Name))
            {
                MessageBox.Show("Save the workbook first.");
                return;
            }

            string full = wb.FullName;
            string path = wb.Path;

            bool isLocal = !full.StartsWith("http", StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(full);
            if (isLocal)
            {
                Process.Start("explorer.exe", $"/select,\"{full}\"");
            }
            else if (full.StartsWith("http", StringComparison.OrdinalIgnoreCase))
            {
                // If Excel provides a folder URL, open it; else open the file URL
                string url = (!string.IsNullOrEmpty(path) && path.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                    ? path
                    : full;
                OpenWithShell(url);
            }
            else
            {
                MessageBox.Show("File path not recognized as local or web URL.");
            }
        }

        // -------------------- COPY PATH/URL TO CLIPBOARD --------------------
        public static void CopyFilePathToClipboard()
        {
            var wb = OfficePragmatics_Excel.Globals.ThisAddIn.Application.ActiveWorkbook;
            if (wb == null || string.IsNullOrEmpty(wb.Name))
            {
                MessageBox.Show("Save the workbook first.");
                return;
            }
            string result = wb.FullName; // URL for cloud; full path for local
            Clipboard.SetText(result);
            MessageBox.Show($"Copied:\n{result}", "Path Copied");
        }

        // -------------------- HIDE PAGE BREAKS (ALL SHEETS) --------------------
        public static void HidePageBreaks(Excel.Application app)
        {
            try
            {
                var wb = app.ActiveWorkbook;
                if (wb == null) { MessageBox.Show("No workbook is active."); return; }
                foreach (Excel.Worksheet ws in wb.Worksheets) ws.DisplayPageBreaks = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("HidePageBreaks error:\n" + ex.Message,
                                "HidePageBreaks", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // -------------------- helpers --------------------
        private static void MoveCorrespondingPdfToArchive(string oldXlsxFullPath, string archiveFolderPath)
        {
            try
            {
                if (string.IsNullOrEmpty(archiveFolderPath)) return;
                string pdfToMove = System.IO.Path.ChangeExtension(oldXlsxFullPath, ".pdf");
                if (!System.IO.File.Exists(pdfToMove)) return;

                string destinationPdfPath = System.IO.Path.Combine(archiveFolderPath, System.IO.Path.GetFileName(pdfToMove));
                if (!System.IO.File.Exists(destinationPdfPath))
                    System.IO.File.Move(pdfToMove, destinationPdfPath);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not move corresponding PDF: {ex.Message}");
            }
        }

        private static string FindArchiveFolder(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath) || !System.IO.Directory.Exists(directoryPath)) return null;
            return System.IO.Directory.GetDirectories(directoryPath)
                     .FirstOrDefault(d =>
                     {
                         string folderName = System.IO.Path.GetFileName(d);
                         return folderName != null &&
                                (folderName.EndsWith("archive", StringComparison.OrdinalIgnoreCase) ||
                                 folderName.EndsWith("old", StringComparison.OrdinalIgnoreCase));
                     });
        }

        /// <summary>Adds a Windows "Recent" shortcut (skip for URLs).</summary>
        private static void AddToRecentFiles(string filePath, string fileNameNoExt)
        {
            try
            {
                if (filePath.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return;
                string recentPath = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
                string shortcutPath = System.IO.Path.Combine(recentPath, fileNameNoExt + ".lnk");

                Type shellType = Type.GetTypeFromProgID("WScript.Shell");
                dynamic shell = Activator.CreateInstance(shellType);
                var shortcut = shell.CreateShortcut(shortcutPath);
                shortcut.TargetPath = filePath;
                shortcut.Save();

                System.Runtime.InteropServices.Marshal.ReleaseComObject(shortcut);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(shell);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Failed to add to recent files: {ex.Message}");
            }
        }
    }
}