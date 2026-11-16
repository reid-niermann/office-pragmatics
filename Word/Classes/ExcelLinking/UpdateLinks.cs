// UpdateLinks.cs (.NET Framework 4.8 / C# 7.3)
// Main entry point for updating all Excel links in a Word document.

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// Main class for updating all Excel links in a Word document.
    /// Coordinates chart, range, and text link updates with progress tracking.
    /// </summary>
    public static class UpdateLinks
    {
        /// <summary>
        /// Updates all Excel links in the active Word document.
        /// </summary>
        /// <param name="wordApp">Word application instance</param>
        /// <param name="showProgress">Whether to show progress dialog</param>
        /// <returns>True if any links were updated</returns>
        public static bool UpdateAllLinks(Word.Application wordApp, bool showProgress = true)
        {
            if (wordApp == null)
                return false;

            Word.Document doc = null;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                MessageBox.Show("No active document.", "Update Links", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            if (doc == null)
            {
                MessageBox.Show("No active document.", "Update Links", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            return UpdateAllLinks(doc, showProgress);
        }

        /// <summary>
        /// Updates all Excel links in the specified Word document.
        /// </summary>
        public static bool UpdateAllLinks(Word.Document doc, bool showProgress = true)
        {
            if (doc == null)
                return false;

            // Collect all linked shapes and content controls
            var linkedShapes = new List<Word.Shape>();
            var linkedContentControls = new List<Word.ContentControl>();

            try
            {
                // Find linked shapes in main document body
                foreach (Word.Shape shape in doc.Shapes)
                {
                    if (ChartLinkUtility.IsChartLink(shape) || RangeLinkUtility.IsRangeLink(shape))
                    {
                        linkedShapes.Add(shape);
                    }
                }

                // Find linked shapes in headers/footers
                foreach (Word.Section section in doc.Sections)
                {
                    foreach (Word.HeaderFooter hf in section.Headers)
                    {
                        try
                        {
                            foreach (Word.Shape shape in hf.Shapes)
                            {
                                if (ChartLinkUtility.IsChartLink(shape) || RangeLinkUtility.IsRangeLink(shape))
                                {
                                    linkedShapes.Add(shape);
                                }
                            }
                        }
                        catch { }
                    }

                    foreach (Word.HeaderFooter hf in section.Footers)
                    {
                        try
                        {
                            foreach (Word.Shape shape in hf.Shapes)
                            {
                                if (ChartLinkUtility.IsChartLink(shape) || RangeLinkUtility.IsRangeLink(shape))
                                {
                                    linkedShapes.Add(shape);
                                }
                            }
                        }
                        catch { }
                    }
                }

                // Find text link content controls
                foreach (Word.ContentControl cc in doc.ContentControls)
                {
                    if (TextLinkUtility.IsTextLink(cc))
                    {
                        linkedContentControls.Add(cc);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error collecting links: {ex.Message}");
            }

            int totalLinks = linkedShapes.Count + linkedContentControls.Count;

            if (totalLinks == 0)
            {
                MessageBox.Show("No Excel links found in this document.",
                    "Update Links", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            // Create wait form for progress
            WaitForm waitForm = null;
            if (showProgress)
            {
                waitForm = new WaitForm("Updating Excel Links...");
                waitForm.InitializeProgress(totalLinks);
                waitForm.Show();
                Application.DoEvents();
            }

            var warnings = new WarningCollector();
            Excel.Application excelApp = null;
            bool excelStartedByUs = false;

            try
            {
                // Get or start Excel
                try
                {
                    excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelStartedByUs = true;
                }

                using (var perfScope = new ExcelPerformanceScope(excelApp))
                using (var wbCache = new WorkbookCache(excelApp))
                {
                    int completed = 0;

                    // Update shape links (charts and ranges)
                    // Process in reverse order since refreshing may change shape indices
                    for (int i = linkedShapes.Count - 1; i >= 0; i--)
                    {
                        var shape = linkedShapes[i];
                        completed++;

                        if (waitForm != null)
                        {
                            string linkType = ChartLinkUtility.IsChartLink(shape) ? "chart" : "range";
                            waitForm.SetProgress(completed, $"Updating {linkType} {completed} of {totalLinks}...");
                            Application.DoEvents();

                            if (waitForm.IsCancelled)
                            {
                                warnings.AddWarning("Update cancelled by user");
                                break;
                            }
                        }

                        try
                        {
                            bool success = false;

                            if (ChartLinkUtility.IsChartLink(shape))
                            {
                                success = ChartLinkUtility.RefreshChartLink(shape, excelApp, wbCache, warnings);
                            }
                            else if (RangeLinkUtility.IsRangeLink(shape))
                            {
                                success = RangeLinkUtility.RefreshRangeLink(shape, excelApp, wbCache, warnings);
                            }

                            if (success)
                                warnings.IncrementUpdated();
                            else
                                warnings.IncrementFailed();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error updating shape link: {ex.Message}");
                            warnings.AddWarning($"Shape error: {ex.Message}");
                            warnings.IncrementFailed();
                        }
                    }

                    // Update text links
                    foreach (var cc in linkedContentControls)
                    {
                        completed++;

                        if (waitForm != null)
                        {
                            waitForm.SetProgress(completed, $"Updating text link {completed} of {totalLinks}...");
                            Application.DoEvents();

                            if (waitForm.IsCancelled)
                            {
                                warnings.AddWarning("Update cancelled by user");
                                break;
                            }
                        }

                        try
                        {
                            if (TextLinkUtility.RefreshTextLink(cc, excelApp, wbCache, warnings))
                                warnings.IncrementUpdated();
                            else
                                warnings.IncrementFailed();
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error updating text link: {ex.Message}");
                            warnings.AddWarning($"Text link error: {ex.Message}");
                            warnings.IncrementFailed();
                        }
                    }
                }

                return warnings.UpdateCount > 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in UpdateAllLinks: {ex.Message}");
                MessageBox.Show($"Error updating links: {ex.Message}",
                    "Update Links Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (waitForm != null)
                {
                    waitForm.Close();
                    waitForm.Dispose();
                }

                if (excelStartedByUs && excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                    catch { }
                }

                // Show summary
                if (warnings.UpdateCount > 0 || warnings.FailedCount > 0)
                {
                    warnings.ShowSummary("Update Complete");
                }
            }
        }

        /// <summary>
        /// Updates a single linked shape.
        /// </summary>
        public static bool UpdateSingleLink(Word.Shape shape)
        {
            if (shape == null)
                return false;

            if (!ChartLinkUtility.IsChartLink(shape) && !RangeLinkUtility.IsRangeLink(shape))
            {
                MessageBox.Show("This shape is not an Excel link.",
                    "Update Link", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }

            Excel.Application excelApp = null;
            bool excelStartedByUs = false;

            try
            {
                try
                {
                    excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch
                {
                    excelApp = new Excel.Application();
                    excelApp.Visible = false;
                    excelStartedByUs = true;
                }

                var warnings = new WarningCollector();

                using (var perfScope = new ExcelPerformanceScope(excelApp))
                using (var wbCache = new WorkbookCache(excelApp))
                {
                    bool success = false;

                    if (ChartLinkUtility.IsChartLink(shape))
                    {
                        success = ChartLinkUtility.RefreshChartLink(shape, excelApp, wbCache, warnings);
                    }
                    else if (RangeLinkUtility.IsRangeLink(shape))
                    {
                        success = RangeLinkUtility.RefreshRangeLink(shape, excelApp, wbCache, warnings);
                    }

                    if (success)
                    {
                        // Silent success
                        return true;
                    }
                    else
                    {
                        if (warnings.WarningCount > 0)
                        {
                            warnings.ShowSummary("Update Failed");
                        }
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error updating link: {ex.Message}",
                    "Update Link Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                if (excelStartedByUs && excelApp != null)
                {
                    try
                    {
                        excelApp.Quit();
                        Marshal.ReleaseComObject(excelApp);
                    }
                    catch { }
                }
            }
        }
    }

    /// <summary>
    /// Progress dialog for long-running operations.
    /// </summary>
    internal class WaitForm : Form
    {
        private readonly Label lblStatus;
        private readonly ProgressBar progressBar;
        private readonly Button btnCancel;
        private int totalItems;

        public bool IsCancelled { get; private set; }

        public WaitForm(string title = "Please Wait...")
        {
            // Form setup
            this.Text = title;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;
            this.ShowInTaskbar = false;
            this.Size = new System.Drawing.Size(400, 150);
            this.BackColor = System.Drawing.Color.White;

            // Status label
            lblStatus = new Label
            {
                AutoSize = false,
                Size = new System.Drawing.Size(360, 25),
                Location = new System.Drawing.Point(20, 20),
                Text = "Initializing...",
                Font = new System.Drawing.Font("Segoe UI", 9f)
            };
            this.Controls.Add(lblStatus);

            // Progress bar
            progressBar = new ProgressBar
            {
                Size = new System.Drawing.Size(360, 23),
                Location = new System.Drawing.Point(20, 50),
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };
            this.Controls.Add(progressBar);

            // Cancel button
            btnCancel = new Button
            {
                Text = "Cancel",
                Size = new System.Drawing.Size(80, 30),
                Location = new System.Drawing.Point(300, 85),
                FlatStyle = FlatStyle.Flat,
                BackColor = System.Drawing.Color.FromArgb(240, 240, 240),
                Font = new System.Drawing.Font("Segoe UI", 9f)
            };
            btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.FromArgb(200, 200, 200);
            btnCancel.Click += (s, e) => { IsCancelled = true; btnCancel.Enabled = false; btnCancel.Text = "Cancelling..."; };
            this.Controls.Add(btnCancel);

            // Make cancel button rounded
            btnCancel.Region = new System.Drawing.Region(
                CreateRoundedRectanglePath(new System.Drawing.Rectangle(0, 0, btnCancel.Width, btnCancel.Height), 5));
        }

        public void InitializeProgress(int total)
        {
            totalItems = total;
            progressBar.Maximum = Math.Max(1, total);
            progressBar.Value = 0;
            lblStatus.Text = $"Processing 0 of {total}...";
        }

        public void SetProgress(int completed, string status = null)
        {
            if (completed > progressBar.Maximum)
                completed = progressBar.Maximum;

            progressBar.Value = completed;

            if (status != null)
            {
                lblStatus.Text = status;
            }
            else
            {
                lblStatus.Text = $"Processing {completed} of {totalItems}...";
            }
        }

        private static System.Drawing.Drawing2D.GraphicsPath CreateRoundedRectanglePath(
            System.Drawing.Rectangle rect, int radius)
        {
            var path = new System.Drawing.Drawing2D.GraphicsPath();
            int diameter = radius * 2;

            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            return path;
        }
    }
}