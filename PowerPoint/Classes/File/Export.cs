// Export.cs
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.VisualBasic; // Interaction.InputBox

namespace OfficePragmatics_PowerPoint.File
{
    public static class ExportUtility
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
        // EXPORT SELECTION (PDF/PPTX)
        // =========================
        public static void ExportSelection(PowerPoint.Application app)
        {
            try
            {
                var pres = app.ActivePresentation;
                if (pres == null)
                {
                    MessageBox.Show("No active presentation.", "Export Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                bool isCloud = IsCloudDoc(pres);
                string baseName = Path.GetFileNameWithoutExtension(pres.Name);

                string choice = Interaction.InputBox("PDF = 1, PPT = 2", "Export Selection", "1");
                if (choice != "1" && choice != "2") return;

                if (choice == "1")
                {
                    // PDF of selection
                    string defName = baseName + " - Selected";
                    string outName = Interaction.InputBox("Enter export name (no extension):", "ExportPDF", defName);
                    if (string.IsNullOrEmpty(outName)) return;

                    string pdfPath = isCloud
                        ? TempPdfPathFor(outName)
                        : Path.Combine(pres.Path, outName + ".pdf");

                    pres.ExportAsFixedFormat(
                        pdfPath,
                        PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF,
                        PowerPoint.PpFixedFormatIntent.ppFixedFormatIntentPrint,
                        Office.MsoTriState.msoFalse,
                        PowerPoint.PpPrintHandoutOrder.ppPrintHandoutHorizontalFirst,
                        PowerPoint.PpPrintOutputType.ppPrintOutputSlides,
                        Office.MsoTriState.msoFalse,
                        null,
                        PowerPoint.PpPrintRangeType.ppPrintSelection,
                        null,
                        false, false, false, false, false);

                    if (!isCloud) AddToRecentFiles(pdfPath, outName);
                    OpenWithShell(pdfPath);
                }
                else // choice == "2"
                {
                    // PPTX of selected slides
                    var sel = app.ActiveWindow?.Selection?.SlideRange;
                    if (sel == null || sel.Count == 0)
                    {
                        MessageBox.Show("No slides selected.", "Export Selection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    var indices = sel.Cast<PowerPoint.Slide>().Select(s => s.SlideIndex).OrderBy(i => i).ToArray();
                    var newPpt = app.Presentations.Add(Office.MsoTriState.msoTrue);

                    // Match geometry
                    newPpt.PageSetup.SlideHeight = pres.PageSetup.SlideHeight;
                    newPpt.PageSetup.SlideWidth = pres.PageSetup.SlideWidth;
                    newPpt.PageSetup.SlideOrientation = pres.PageSetup.SlideOrientation;
                    newPpt.PageSetup.SlideSize = pres.PageSetup.SlideSize;

                    foreach (int idx in indices)
                    {
                        pres.Slides[idx].Copy();
                        var pasted = newPpt.Slides.Paste();
                        pasted[1].Design = pres.Slides[idx].Design;
                        pasted[1].FollowMasterBackground = pres.Slides[idx].FollowMasterBackground;
                    }

                    string defName = baseName + " - Selected";
                    string outName = Interaction.InputBox("Enter PPTX name (no extension):", "ExportPPT", defName);
                    if (string.IsNullOrEmpty(outName)) { newPpt.Close(); return; }

                    if (IsCloudDoc(pres))
                    {
                        // Cloud: filename-only -> saves into same cloud folder
                        newPpt.SaveAs(outName + ".pptx", PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                        MessageBox.Show($"PPTX saved as:\r\n{outName}.pptx", "ExportPPT",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        string pptPath = Path.Combine(pres.Path, outName + ".pptx");
                        newPpt.SaveAs(pptPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                        AddToRecentFiles(pptPath, outName);
                        MessageBox.Show($"PPTX saved to:\r\n{pptPath}", "ExportPPT",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
                        OpenWithShell(pptPath);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error exporting selection:\r\n" + ex.Message, "Export Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}