using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq; // QAT
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic; // Interaction.InputBox
using OfficePragmatics_PowerPoint.Utilities; // Utility.IsValidFileName
using OfficePragmatics_Powerpoint;

namespace OfficePragmatics_PowerPoint.File
{
    public static class FileUtility
    {
        // =========================
        // Helpers (cloud-safe)
        // =========================
        private static bool IsCloudDoc(PowerPoint.Presentation p)
        {
            string f = p?.FullName ?? "";
            string path = p?.Path ?? "";
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
        public static void VersionUp(PowerPoint.Application app)
        {
            try
            {
                var pres = app.ActivePresentation;
                if (pres == null)
                {
                    MessageBox.Show("No active presentation.", "VersionUp", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string full = pres.FullName;                // may be URL for cloud/WOPI
                string dir = pres.Path;                     // may be URL for cloud/WOPI
                string nameOnly = pres.Name;                // e.g., "MyDeck.pptx"
                string baseNm = Path.GetFileNameWithoutExtension(nameOnly);
                string ext = Path.GetExtension(nameOnly);
                bool isCloud = IsCloudDoc(pres);

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
                        pres.SaveCopyAs(v1Name); // copy current as v1
                        pres.SaveAs(v2Name);     // current becomes v2 in same folder
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
                        pres.SaveAs(v2Full);
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
                        pres.SaveAs(newName);
                        MessageBox.Show($"Saved new version: \"{newName}\"", "VersionUp",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string newFull = Path.Combine(dir, newName);
                        pres.SaveAs(newFull);

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
        /// Looks for a PDF with the same base name as the old PPTX file and moves it to the archive folder.
        /// Local-only (no-op for cloud/WOPI).
        /// </summary>
        private static void MoveCorrespondingPdfToArchive(string oldPptxFullPath, string archiveFolderPath)
        {
            try
            {
                if (string.IsNullOrEmpty(archiveFolderPath)) return;
                if (string.IsNullOrEmpty(oldPptxFullPath)) return;

                string pdfToMove = Path.ChangeExtension(oldPptxFullPath, ".pdf");
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
        public static void SavePDF(PowerPoint.Application app)
        {
            try
            {
                var pres = app.ActivePresentation;
                if (pres == null)
                {
                    MessageBox.Show("No active presentation.", "SavePDF", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (pres.Saved == Office.MsoTriState.msoFalse)
                {
                    pres.Save();
                }

                bool isCloud = IsCloudDoc(pres);
                string nameOnly = pres.Name;
                string baseName = Path.GetFileNameWithoutExtension(nameOnly);

                string pdfPath;
                if (isCloud)
                {
                    // Always export to local temp for cloud/WOPI
                    pdfPath = TempPdfPathFor(baseName);
                }
                else
                {
                    string folder = pres.Path;
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

                pres.ExportAsFixedFormat(
                    pdfPath,
                    PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                    PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                    Office.MsoTriState.msoFalse,
                    PowerPoint.PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
                    PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                    Office.MsoTriState.msoFalse,
                    null,
                    PowerPoint.PpPrintRangeType.ppPrintAll,
                    null,
                    false, false, false, false, false);

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
        // OPEN CONTAINING FOLDER
        // =========================
        public static void OpenContainingFolder()
        {
            PowerPoint.Presentation p = null;
            try
            {
                p = OfficePragmatics_Powerpoint.Globals.ThisAddIn.Application.ActivePresentation;
                if (p == null)
                {
                    MessageBox.Show("No presentation is currently active.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (p.Saved == Office.MsoTriState.msoFalse && string.IsNullOrEmpty(p.Path))
                {
                    MessageBox.Show("The current presentation has not been saved. Please save it first.",
                        "File Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(p.FullName) || p.FullName.Length < 5)
                {
                    if (p.Saved == Office.MsoTriState.msoFalse)
                    {
                        MessageBox.Show("The current presentation has not been saved with a valid name or location. Please save it first.",
                            "File Not Properly Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                string full = p.FullName;
                string path = p.Path;
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
                    // If PowerPoint gives a folder URL, open it; else open the file URL
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
                    if (p.Saved == Office.MsoTriState.msoFalse)
                        MessageBox.Show("Current presentation not saved. Save it first.", "File Not Saved",
                            MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    else
                        MessageBox.Show($"Path ('{full}') not recognized as local file or web URL.",
                            "Path Not Recognized", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show($"PowerPoint error: {ex.Message}", "PowerPoint Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (p != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(p);
            }
        }

        // =========================
        // COPY FILE PATH / URL
        // =========================
        public static void CopyFilePathToClipboard()
        {
            PowerPoint.Presentation p = null;
            try
            {
                p = OfficePragmatics_Powerpoint.Globals.ThisAddIn.Application.ActivePresentation;
                if (p == null)
                {
                    MessageBox.Show("No presentation active.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (p.Saved == Office.MsoTriState.msoFalse && string.IsNullOrEmpty(p.Path))
                {
                    MessageBox.Show("Presentation not saved. Save it first to get path.",
                        "File Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                if (string.IsNullOrEmpty(p.FullName) || p.FullName.Length < 5)
                {
                    if (p.Saved == Office.MsoTriState.msoFalse)
                    {
                        MessageBox.Show("Presentation not saved with valid name/location. Save it first.",
                            "File Not Properly Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                }

                string result = p.FullName; // URL for cloud; full path for local
                if (string.IsNullOrEmpty(result))
                {
                    MessageBox.Show("File path unknown. Ensure presentation is saved.",
                        "Path Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                Clipboard.SetText(result);
                MessageBox.Show($"Path/URL copied:\n{result}", "Path Copied",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                MessageBox.Show($"PowerPoint error: {ex.Message}", "PowerPoint Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Unexpected error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (p != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(p);
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

        // =========================
        // QUICK ACCESS TOOLBAR SETUP
        // =========================
        private static readonly string[] DesiredQatCommands =
        {
            "TextFillColorPicker", "ShapeFillColorPicker", "ShapeOutlineColorPicker",
            "TextDirectionMoreOptionsDialog", "ObjectsArrangeMenu", "FontSize",
            "ShapeRectangle", "TextBoxInsertHorizontal", "GalleryAllShapesAndTextboxes",
            "ShapeStraightConnector", "CombineShapesGallery",
        };

        public static void SetDefaultQuickAccessToolbar(PowerPoint.Application app)
        {
            string officeUiFilePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                @"Microsoft\Office\PowerPoint.officeUI");

            XNamespace mso = "http://schemas.microsoft.com/office/2009/07/customui";
            XDocument officeXml = System.IO.File.Exists(officeUiFilePath)
                ? XDocument.Load(officeUiFilePath)
                : new XDocument(new XElement(mso + "customUI"));

            if (officeXml.Root?.Name != mso + "customUI")
            {
                officeXml.RemoveNodes();
                officeXml.Add(new XElement(mso + "customUI"));
            }

            XElement ribbonNode = officeXml.Root.Element(mso + "ribbon") ?? new XElement(mso + "ribbon");
            if (ribbonNode.Parent == null) officeXml.Root.Add(ribbonNode);

            XElement qatNode = ribbonNode.Element(mso + "qat") ?? new XElement(mso + "qat");
            if (qatNode.Parent == null) ribbonNode.Add(qatNode);

            qatNode.Element(mso + "sharedControls")?.Remove();
            XElement sharedControlsNode = new XElement(mso + "sharedControls");
            qatNode.Add(sharedControlsNode);

            foreach (string idMso in DesiredQatCommands)
            {
                sharedControlsNode.Add(new XElement(mso + "control",
                    new XAttribute("idMso", idMso),
                    new XAttribute("visible", "true")));
            }

            Directory.CreateDirectory(Path.GetDirectoryName(officeUiFilePath) ?? "");
            officeXml.Save(officeUiFilePath);

            MessageBox.Show(
                "Quick Access Toolbar updated.\nRestart PowerPoint for changes to take full effect if QAT was customized before.",
                "QAT Update Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}