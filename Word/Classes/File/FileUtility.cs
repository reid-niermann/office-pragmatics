using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic; // Interaction.InputBox
using OfficePragmatics_Word.Utilities; // Utility.IsValidFileName
using OfficePragmatics_Word;

namespace Office_Pragmatics_Word
{
    public static class FileUtility
    {
        // =========================
        // Helpers (cloud-safe)
        // =========================
        private static bool IsCloudDoc(Word.Document doc)
        {
            string f = doc?.FullName ?? "";
            string path = doc?.Path ?? "";
            return f.StartsWith("http", StringComparison.OrdinalIgnoreCase)
                || path.StartsWith("http", StringComparison.OrdinalIgnoreCase);
        }

        private static string TempPdfPathFor(string fileBaseName)
        {
            string safe = Regex.Replace(fileBaseName, @"[^\w\-. ]+", "_");
            return Path.Combine(Path.GetTempPath(), safe + ".pdf");
        }

        private static void OpenWithShell(string pathOrUrl)
        {
            try
            {
                var psi = new ProcessStartInfo(pathOrUrl) { UseShellExecute = true };
                Process.Start(psi);
            }
            catch { /* swallow */ }
        }

        // =========================
        // VERSION UP
        // =========================
        public static void VersionUp(Word.Application app)
        {
            try
            {
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No active document.", "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string full = doc.FullName;                // may be URL for cloud/WOPI
                string dir = doc.Path;                     // may be URL for cloud/WOPI
                string nameOnly = doc.Name;                // e.g., "MyDocument.docx"
                string baseNm = Path.GetFileNameWithoutExtension(nameOnly);
                string ext = Path.GetExtension(nameOnly);
                bool isCloud = IsCloudDoc(doc);

                var m = Regex.Match(baseNm, @"[ _]v(\d+)$", RegexOptions.IgnoreCase);

                if (!m.Success)
                {
                    // Not versioned: propose creating v1 (archive/or copy) and v2 (current)
                    var add = MessageBox.Show(
                        $"The file \"{nameOnly}\" is not versioned.\n\n" +
                        $"Add versioning?\n\nThis will create '{baseNm} _v1{ext}' and save current as '{baseNm} _v2{ext}'.",
                        "Add Versioning?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (add == DialogResult.No) return;

                    string v1Name = $"{baseNm} _v1{ext}";
                    string v2Name = $"{baseNm} _v2{ext}";

                    if (isCloud)
                    {
                        // Cloud/WOPI: Use filename-only so Office writes to same cloud location
                        doc.SaveAs2(v1Name); // copy current as v1
                        doc.SaveAs2(v2Name); // current becomes v2 in same folder
                        MessageBox.Show($"Versioning added.\nNew version: \"{v2Name}\"", "VersionUp",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        // Local file
                        string v2Full = Path.Combine(dir, v2Name);
                        string v1DestDir = FindArchiveFolder(dir);
                        if (v1DestDir == null)
                        {
                            var createArchive = MessageBox.Show(
                                "No 'Archive' or 'Old' folder found. Create an 'Archive' folder to store the _v1 file?",
                                "Create Archive Folder?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                            if (createArchive == DialogResult.Yes)
                            {
                                v1DestDir = Path.Combine(dir, "Archive");
                                Directory.CreateDirectory(v1DestDir);
                            }
                        }

                        string v1FullDest = (v1DestDir != null) ? Path.Combine(v1DestDir, v1Name) : Path.Combine(dir, v1Name);

                        // Conflicts?
                        if (System.IO.File.Exists(v1FullDest) || System.IO.File.Exists(v2Full))
                        {
                            MessageBox.Show(
                                $"Cannot add versioning because a target file already exists:\n\n" +
                                $"- {Path.GetFileName(v1FullDest)} or\n- {Path.GetFileName(v2Full)}\n\n" +
                                "Remove/rename conflicting files and try again.",
                                "VersionUp Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }

                        // 1) Save current as v2
                        doc.SaveAs2(v2Full);
                        // 2) Rename/move original file to v1 destination
                        if (!string.IsNullOrEmpty(full) && System.IO.File.Exists(full))
                        {
                            System.IO.File.Move(full, v1FullDest);
                            MoveCorrespondingPdfToArchive(full, v1DestDir);
                        }

                        AddToRecentFiles(v2Full, $"{baseNm} _v2");
                        MessageBox.Show(
                            $"Versioning added.\nNew version: \"{v2Name}\"\n" +
                            $"Original renamed/moved to: \"{v1FullDest}\"",
                            "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                else
                {
                    // Already versioned: bump to next version
                    int next = int.Parse(m.Groups[1].Value) + 1;
                    string prefix = baseNm.Substring(0, m.Index);
                    string newName = $"{prefix}_v{next}{ext}";

                    if (isCloud)
                    {
                        // Filename-only -> saved in same cloud folder
                        doc.SaveAs2(newName);
                        MessageBox.Show($"Saved new version: \"{newName}\"", "VersionUp",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string newFull = Path.Combine(dir, newName);
                        doc.SaveAs2(newFull);

                        // Optional: move old file to Archive if available
                        string archiveFolder = FindArchiveFolder(dir);
                        if (archiveFolder != null && !string.IsNullOrEmpty(full) && System.IO.File.Exists(full))
                        {
                            string dest = Path.Combine(archiveFolder, Path.GetFileName(full));
                            System.IO.File.Move(full, dest);
                            MoveCorrespondingPdfToArchive(full, archiveFolder);
                        }

                        AddToRecentFiles(newFull, Path.GetFileNameWithoutExtension(newName));
                        MessageBox.Show($"Saved new version: \"{newName}\"", "VersionUp",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("There was a problem saving a new version.\r\n" + ex.Message,
                    "VersionUp Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string FindArchiveFolder(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath) || !Directory.Exists(directoryPath)) return null;
            return Directory.GetDirectories(directoryPath)
                            .FirstOrDefault(d =>
                            {
                                string folderName = Path.GetFileName(d);
                                return folderName != null &&
                                       (folderName.EndsWith("archive", StringComparison.OrdinalIgnoreCase) ||
                                        folderName.EndsWith("old", StringComparison.OrdinalIgnoreCase));
                            });
        }

        /// <summary>
        /// Looks for a PDF with the same base name as the old DOCX file and moves it to the archive folder.
        /// Local-only (no-op for cloud/WOPI).
        /// </summary>
        private static void MoveCorrespondingPdfToArchive(string oldDocxFullPath, string archiveFolderPath)
        {
            try
            {
                if (string.IsNullOrEmpty(archiveFolderPath)) return;
                if (string.IsNullOrEmpty(oldDocxFullPath)) return;

                string pdfToMove = Path.ChangeExtension(oldDocxFullPath, ".pdf");
                if (System.IO.File.Exists(pdfToMove))
                {
                    string pdfFileName = Path.GetFileName(pdfToMove);
                    string destinationPdfPath = Path.Combine(archiveFolderPath, pdfFileName);
                    if (!System.IO.File.Exists(destinationPdfPath))
                    {
                        System.IO.File.Move(pdfToMove, destinationPdfPath);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Could not move corresponding PDF: {ex.Message}");
            }
        }

        // =========================
        // SAVE PDF
        // =========================
        public static void SavePDF(Word.Application app)
        {
            try
            {
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No active document.", "SavePDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool isCloud = IsCloudDoc(doc);
                string nameOnly = doc.Name;
                string baseName = Path.GetFileNameWithoutExtension(nameOnly);

                string pdfPath;
                if (isCloud)
                {
                    // Always export to local temp for cloud/WOPI
                    pdfPath = TempPdfPathFor(baseName);
                }
                else
                {
                    string folder = doc.Path;
                    pdfPath = Path.Combine(folder, baseName + ".pdf");

                    if (System.IO.File.Exists(pdfPath))
                    {
                        var resp = MessageBox.Show(
                            "File already exists.\r\nYes = overwrite, No = rename, Cancel = abort",
                            "SavePDF", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                        if (resp == DialogResult.Cancel) return;
                        if (resp == DialogResult.No)
                        {
                            string input;
                            do
                            {
                                input = Interaction.InputBox("Enter new PDF name (no extension):", "SavePDF", baseName);
                                if (input == "") return;
                            } while (!Utility.IsValidFileName(input));
                            pdfPath = Path.Combine(folder, input + ".pdf");
                        }
                    }
                }

                doc.ExportAsFixedFormat(
                    pdfPath,
                    Word.WdExportFormat.wdExportFormatPDF,
                    false,
                    Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                    Word.WdExportRange.wdExportAllDocument,
                    0, 0,
                    Word.WdExportItem.wdExportDocumentContent,
                    true, true,
                    Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                    true, true, false);

                MessageBox.Show($"PDF saved to:\r\n{pdfPath}", "SavePDF",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!isCloud) AddToRecentFiles(pdfPath, Path.GetFileNameWithoutExtension(pdfPath));
                OpenWithShell(pdfPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problem saving PDF:\r\n" + ex.Message, "SavePDF Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // =========================
        // EXPORT PAGES (PDF/DOCX)
        // =========================
        // Helper: does a range end with a hard page/section break?
        private static bool EndsWithHardPageOrSectionBreak(Word.Range r)
        {
            if (r == null) return false;
            string txt = r.Text ?? string.Empty;
            if (txt.Length == 0) return false;

            // Look in the tail for Chr(12) (manual page break / section break next page acts like this)
            int look = Math.Min(4, txt.Length);
            string tail = txt.Substring(txt.Length - look);
            return tail.IndexOf('\f') >= 0; // '\f' == (char)12
        }

        // Helper: delete a single trailing hard page/section break at end of document (if present)
        private static void DeleteTrailingHardBreak(Word.Document d)
        {
            Word.Range endRng = d.Content;
            endRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
            if (endRng.Start == 0) return;

            int end = endRng.Start;
            // Examine the last few characters
            int start = Math.Max(0, end - 4);
            Word.Range tail = d.Range(start, end);
            string t = tail.Text ?? string.Empty;
            int idx = t.LastIndexOf('\f');
            if (idx >= 0)
            {
                int abs = tail.Start + idx;
                Word.Range br = d.Range(abs, abs + 1);
                br.Delete(); // remove the trailing break
            }
        }

        public static void ExportPages(Word.Application app)
        {
            try
            {
                var doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No active document.", "Export Pages", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                int totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);

                string pageInput = Interaction.InputBox(
                    $"Enter page range to export (e.g., 1-3, 5, 7-9)\nTotal pages: {totalPages}",
                    "Export Pages",
                    "1-" + totalPages);
                if (string.IsNullOrWhiteSpace(pageInput)) return;

                if (!IsValidPageRange(pageInput, totalPages))
                {
                    MessageBox.Show("Invalid page range. Please use format like: 1-3, 5, 7-9",
                        "Invalid Range", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                bool isCloud = IsCloudDoc(doc);
                string baseName = Path.GetFileNameWithoutExtension(doc.Name);

                string choice = Interaction.InputBox("PDF = 1, DOCX = 2", "Export Pages", "1");
                if (choice != "1" && choice != "2") return;

                var pages = ParsePageRange(pageInput)
                    .Where(p => p >= 1 && p <= totalPages)
                    .Distinct()
                    .OrderBy(p => p)
                    .ToList();
                if (pages.Count == 0)
                {
                    MessageBox.Show("No valid pages to export.", "Export Pages",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Save/restore selection; speed up UI
                int selStart = app.Selection?.Range?.Start ?? 0;
                int selEnd = app.Selection?.Range?.End ?? 0;
                bool prevSU = app.ScreenUpdating;
                app.ScreenUpdating = false;

                Word.Document tempDoc = app.Documents.Add();
                object missing = Type.Missing;

                try
                {
                    object what = Word.WdGoToItem.wdGoToPage;
                    object which = Word.WdGoToDirection.wdGoToAbsolute;
                    Word.Range mainStory = doc.StoryRanges[Word.WdStoryType.wdMainTextStory];

                    for (int i = 0; i < pages.Count; i++)
                    {
                        int pageNum = pages[i];
                        bool isLast = (i == pages.Count - 1);

                        // Anchor in MAIN story at start of target page
                        object count = pageNum;
                        Word.Range anchor = mainStory.GoTo(ref what, ref which, ref count, ref missing);

                        // Briefly select so built-in \page resolves for THIS page
                        Word.Range cursor = doc.Range(anchor.Start, anchor.Start);
                        cursor.Select();

                        // Resolve \page
                        Word.Range pageRange = null;
                        try
                        {
                            object bkName = "\\page";
                            pageRange = doc.Bookmarks.get_Item(ref bkName).Range;
                        }
                        catch
                        {
                            continue; // skip if Word can't resolve
                        }

                        // Destination: end-of-document
                        Word.Range endRng = tempDoc.Content;
                        endRng.Collapse(Word.WdCollapseDirection.wdCollapseEnd);

                        if (pageRange != null && pageRange.Start < pageRange.End)
                        {
                            // Remember content bounds before/after insertion
                            int before = tempDoc.Content.End;
                            endRng.FormattedText = pageRange.Duplicate.FormattedText;
                            int after = tempDoc.Content.End;

                            // Inspect inserted chunk
                            Word.Range inserted = tempDoc.Range(before, after);
                            bool endsWithBreak = EndsWithHardPageOrSectionBreak(inserted);

                            if (isLast)
                            {
                                // Trim trailing break from LAST page so we don't get an extra blank page
                                if (endsWithBreak) DeleteTrailingHardBreak(tempDoc);
                            }
                            else
                            {
                                // For non-last pages: only add a break if the chunk didn't already include one
                                if (!endsWithBreak)
                                {
                                    Word.Range br = tempDoc.Content;
                                    br.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                    br.InsertBreak(Word.WdBreakType.wdPageBreak);
                                }
                            }
                        }
                        else
                        {
                            // Truly empty page: represent by a single break (unless last)
                            if (!isLast)
                            {
                                Word.Range br = tempDoc.Content;
                                br.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                                br.InsertBreak(Word.WdBreakType.wdPageBreak);
                            }
                        }
                    }

                    string defName = baseName + " - Pages " + pageInput.Replace(",", "_").Replace("-", "to");

                    if (choice == "1")
                    {
                        string outName = Interaction.InputBox("Enter export name (no extension):", "Export PDF", defName);
                        if (string.IsNullOrWhiteSpace(outName)) return;

                        string pdfPath = isCloud
                            ? TempPdfPathFor(outName)
                            : Path.Combine(doc.Path, outName + ".pdf");

                        tempDoc.ExportAsFixedFormat(
                            pdfPath,
                            Word.WdExportFormat.wdExportFormatPDF,
                            false,
                            Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                            Word.WdExportRange.wdExportAllDocument,
                            1, 1,
                            Word.WdExportItem.wdExportDocumentContent,
                            true, true,
                            Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                            true, true, false);

                        if (!isCloud) AddToRecentFiles(pdfPath, outName);
                        OpenWithShell(pdfPath);
                        MessageBox.Show($"PDF saved to:\r\n{pdfPath}", "Export PDF",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string outName = Interaction.InputBox("Enter DOCX name (no extension):", "Export DOCX", defName);
                        if (string.IsNullOrWhiteSpace(outName)) return;

                        if (isCloud)
                        {
                            tempDoc.SaveAs2(outName + ".docx", Word.WdSaveFormat.wdFormatXMLDocument);
                            MessageBox.Show($"DOCX saved as:\r\n{outName}.docx", "Export DOCX",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        else
                        {
                            string docxPath = Path.Combine(doc.Path, outName + ".docx");
                            tempDoc.SaveAs2(docxPath, Word.WdSaveFormat.wdFormatXMLDocument);
                            AddToRecentFiles(docxPath, outName);
                            OpenWithShell(docxPath);
                            MessageBox.Show($"DOCX saved to:\r\n{docxPath}", "Export DOCX",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }

                    try { doc.Range(selStart, selEnd).Select(); } catch { /* ignore */ }
                }
                finally
                {
                    app.ScreenUpdating = prevSU;
                    tempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting pages:\r\n" + ex.Message, "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Helper method to parse page range string (e.g., "1-3,5,7-9")
        private static System.Collections.Generic.List<int> ParsePageRange(string input)
        {
            var pages = new System.Collections.Generic.List<int>();
            var parts = input.Split(',');

            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.Contains("-"))
                {
                    var range = trimmed.Split('-');
                    if (range.Length == 2 && int.TryParse(range[0], out int start) && int.TryParse(range[1], out int end))
                    {
                        for (int i = start; i <= end; i++)
                        {
                            if (!pages.Contains(i))
                                pages.Add(i);
                        }
                    }
                }
                else if (int.TryParse(trimmed, out int page))
                {
                    if (!pages.Contains(page))
                        pages.Add(page);
                }
            }

            return pages;
        }

        // Helper method to validate page range
        private static bool IsValidPageRange(string input, int maxPages)
        {
            if (string.IsNullOrWhiteSpace(input)) return false;

            var parts = input.Split(',');
            foreach (var part in parts)
            {
                var trimmed = part.Trim();
                if (trimmed.Contains("-"))
                {
                    var range = trimmed.Split('-');
                    if (range.Length != 2) return false;
                    if (!int.TryParse(range[0], out int start) || !int.TryParse(range[1], out int end))
                        return false;
                    if (start < 1 || end > maxPages || start > end)
                        return false;
                }
                else
                {
                    if (!int.TryParse(trimmed, out int page))
                        return false;
                    if (page < 1 || page > maxPages)
                        return false;
                }
            }

            return true;
        }

        // =========================
        // OPEN CONTAINING FOLDER
        // =========================
        public static void OpenContainingFolder(Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No document is currently active.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (!doc.Saved && string.IsNullOrEmpty(doc.Path))
                {
                    MessageBox.Show("The current document has not been saved. Please save it first.",
                        "File Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(doc.FullName) || doc.FullName.Length < 5)
                {
                    if (!doc.Saved)
                    {
                        MessageBox.Show("The current document has not been saved with a valid name or location. Please save it first.",
                            "File Not Properly Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                string full = doc.FullName;
                string path = doc.Path;
                bool local = !full.StartsWith("http", StringComparison.OrdinalIgnoreCase) && System.IO.File.Exists(full);

                if (local)
                {
                    string localDir = Path.GetDirectoryName(full);
                    if (Directory.Exists(localDir))
                        Process.Start("explorer.exe", $"/select,\"{full}\"");
                    else
                        MessageBox.Show($"Local folder '{localDir}' not found for '{full}'.",
                            "Local Folder Not Found", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else if (full.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                {
                    // If Word gives a folder URL, open it; else open the file URL
                    string url = (!string.IsNullOrEmpty(path) && path.StartsWith("http", StringComparison.OrdinalIgnoreCase))
                        ? path
                        : full;
                    try { OpenWithShell(url); }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Could not open web location.\nError: {ex.Message}\nURL: {url}",
                            "Error Opening Web Location", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    if (!doc.Saved)
                        MessageBox.Show("Current document not saved. Save it first.", "File Not Saved",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                        MessageBox.Show($"Path ('{full}') not recognized as local file or web URL.",
                            "Path Not Recognized", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show($"Word error: {ex.Message}", "Word Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (doc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            }
        }

        // =========================
        // COPY FILE PATH / URL
        // =========================
        public static void CopyFilePathToClipboard(Word.Application app)
        {
            Word.Document doc = null;
            try
            {
                doc = app.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No document active.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (!doc.Saved && string.IsNullOrEmpty(doc.Path))
                {
                    MessageBox.Show("Document not saved. Save it first to get path.",
                        "File Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(doc.FullName) || doc.FullName.Length < 5)
                {
                    if (!doc.Saved)
                    {
                        MessageBox.Show("Document not saved with valid name/location. Save it first.",
                            "File Not Properly Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                string result = doc.FullName; // URL for cloud; full path for local
                if (string.IsNullOrEmpty(result))
                {
                    MessageBox.Show("File path unknown. Ensure document is saved.",
                        "Path Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                Clipboard.SetText(result);
                MessageBox.Show($"Path/URL copied:\n{result}", "Path Copied",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show($"Word error: {ex.Message}", "Word Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (doc != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(doc);
            }
        }

        // =========================
        // RECENT FILES SHORTCUT (local only)
        // =========================
        private static void AddToRecentFiles(string filePath, string fileName)
        {
            try
            {
                if (filePath.StartsWith("http", StringComparison.OrdinalIgnoreCase)) return; // URL -> skip
                string recentPath = Environment.GetFolderPath(Environment.SpecialFolder.Recent);
                string shortcutPath = Path.Combine(recentPath, fileName + ".lnk");

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