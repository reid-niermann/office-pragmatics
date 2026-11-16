// ChartLinkUtility.cs (.NET Framework 4.8 / C# 7.3)
// Chart linking utilities for Word - copies Excel charts as pictures to Word shapes.

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// Provides chart linking functionality between Excel and Word.
    /// Copies Excel charts as enhanced metafile pictures into Word shapes.
    /// </summary>
    internal static class ChartLinkUtility
    {
        private const string TagWorkbookPath = "Pragmatics_WbPath";
        private const string TagSheetName = "Pragmatics_SheetName";
        private const string TagChartName = "Pragmatics_ChartName";
        private const string TagLinkType = "Pragmatics_LinkType";
        private const string LinkTypeChart = "Chart";

        #region Public Import Methods (Ribbon Entry Points)

        /// <summary>
        /// Imports an Excel chart as a vector (EMF) image with linking.
        /// Shows a dialog to select the chart from an open Excel workbook.
        /// </summary>
        public static Word.Shape ImportExcelChartAsVector(Word.Application wordApp)
        {
            return ImportExcelChart(wordApp, ChartImportFormat.Vector);
        }

        /// <summary>
        /// Imports an Excel chart as a raster (PNG) image with linking.
        /// Shows a dialog to select the chart from an open Excel workbook.
        /// </summary>
        public static Word.Shape ImportExcelChartAsPicture(Word.Application wordApp)
        {
            return ImportExcelChart(wordApp, ChartImportFormat.Picture);
        }

        /// <summary>
        /// Imports an Excel chart as a native embedded chart (editable).
        /// Shows a dialog to select the chart from an open Excel workbook.
        /// </summary>
        public static Word.Shape ImportExcelChartAsNative(Word.Application wordApp)
        {
            return ImportExcelChart(wordApp, ChartImportFormat.Native);
        }

        private enum ChartImportFormat { Vector, Picture, Native }

        private static Word.Shape ImportExcelChart(Word.Application wordApp, ChartImportFormat format)
        {
            if (wordApp == null) return null;

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                MessageBox.Show("No active document.", "Import Chart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (doc == null)
            {
                MessageBox.Show("No active document.", "Import Chart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            // Get Excel application
            Excel.Application excelApp;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                MessageBox.Show("Excel is not running. Please open an Excel workbook with a chart first.",
                    "Import Chart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            // Check for open workbooks
            if (excelApp.Workbooks.Count == 0)
            {
                MessageBox.Show("No Excel workbooks are open. Please open a workbook with a chart.",
                    "Import Chart", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            // Show chart selection dialog
            using (var dialog = new ChartSelectionDialog(excelApp))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return null;

                string wbPath = dialog.SelectedWorkbookPath;
                string sheetName = dialog.SelectedSheetName;
                string chartName = dialog.SelectedChartName;

                if (string.IsNullOrEmpty(wbPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(chartName))
                    return null;

                Word.Range insertRange = wordApp.Selection.Range;

                switch (format)
                {
                    case ChartImportFormat.Vector:
                        return CreateChartLink(doc, excelApp, wbPath, sheetName, chartName, insertRange);

                    case ChartImportFormat.Picture:
                        return CreateChartLinkAsPicture(doc, excelApp, wbPath, sheetName, chartName, insertRange);

                    case ChartImportFormat.Native:
                        return CreateChartAsNative(doc, excelApp, wbPath, sheetName, chartName, insertRange);

                    default:
                        return CreateChartLink(doc, excelApp, wbPath, sheetName, chartName, insertRange);
                }
            }
        }

        #endregion

        /// <summary>
        /// Creates a chart link in Word by copying an Excel chart as a picture.
        /// </summary>
        /// <param name="doc">Target Word document</param>
        /// <param name="excelApp">Excel application instance</param>
        /// <param name="workbookPath">Path to the Excel workbook</param>
        /// <param name="sheetName">Name of the worksheet containing the chart</param>
        /// <param name="chartName">Name of the chart object</param>
        /// <param name="range">Target range in Word for insertion</param>
        /// <returns>The created shape, or null on failure</returns>
        public static Word.Shape CreateChartLink(
            Word.Document doc,
            Excel.Application excelApp,
            string workbookPath,
            string sheetName,
            string chartName,
            Word.Range range)
        {
            if (doc == null || excelApp == null || range == null)
                return null;

            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(chartName))
                return null;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.ChartObject chartObj = null;
            Word.Shape newShape = null;

            using (var perfScope = new ExcelPerformanceScope(excelApp))
            using (var wbCache = new WorkbookCache(excelApp))
            {
                try
                {
                    // Normalize path for cross-user compatibility
                    string normalizedPath = CloudPathHelper.NormalizePath(workbookPath);

                    // Get or open workbook
                    wb = wbCache.GetOrOpen(workbookPath, out bool newlyOpened, out string actualPath);
                    if (wb == null)
                    {
                        MessageBox.Show($"Could not open workbook:\n{workbookPath}",
                            "Chart Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    // Prepare workbook for copy
                    wbCache.PrepareWorkbookForCopy(wb);

                    // Find worksheet
                    try
                    {
                        ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                    }
                    catch
                    {
                        MessageBox.Show($"Could not find worksheet '{sheetName}' in workbook.",
                            "Chart Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    if (ws == null)
                        return null;

                    // Find chart object
                    try
                    {
                        chartObj = ws.ChartObjects(chartName) as Excel.ChartObject;
                    }
                    catch
                    {
                        MessageBox.Show($"Could not find chart '{chartName}' on worksheet '{sheetName}'.",
                            "Chart Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    if (chartObj == null)
                        return null;

                    // Copy chart to clipboard
                    chartObj.Copy();

                    // Wait for clipboard
                    if (!ClipboardHelper.WaitForClipboardReady(300))
                    {
                        System.Diagnostics.Debug.WriteLine("Clipboard not ready after copying chart");
                    }

                    // Paste into Word as picture
                    try
                    {
                        newShape = ClipboardHelper.RetryPaste(() =>
                        {
                            range.Paste();
                            // Get the newly pasted shape
                            if (range.InlineShapes.Count > 0)
                            {
                                var inlineShape = range.InlineShapes[range.InlineShapes.Count];
                                return inlineShape.ConvertToShape();
                            }
                            else if (range.ShapeRange.Count > 0)
                            {
                                return range.ShapeRange[range.ShapeRange.Count];
                            }
                            return null;
                        });
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Paste failed: {ex.Message}");

                        // Fallback: try PasteSpecial
                        try
                        {
                            range.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);
                            if (range.InlineShapes.Count > 0)
                            {
                                var inlineShape = range.InlineShapes[range.InlineShapes.Count];
                                newShape = inlineShape.ConvertToShape();
                            }
                        }
                        catch (Exception ex2)
                        {
                            System.Diagnostics.Debug.WriteLine($"PasteSpecial also failed: {ex2.Message}");
                            return null;
                        }
                    }

                    if (newShape == null)
                        return null;

                    // Store link information in shape tags
                    ShapeTagsHelper.SetTag(newShape, TagLinkType, LinkTypeChart);
                    ShapeTagsHelper.SetTag(newShape, TagWorkbookPath, normalizedPath);
                    ShapeTagsHelper.SetTag(newShape, TagSheetName, sheetName);
                    ShapeTagsHelper.SetTag(newShape, TagChartName, chartName);

                    return newShape;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating chart link: {ex.Message}");
                    return null;
                }
                finally
                {
                    ClipboardHelper.SafeClear();

                    if (chartObj != null)
                    {
                        try { Marshal.ReleaseComObject(chartObj); } catch { }
                    }
                    if (ws != null)
                    {
                        try { Marshal.ReleaseComObject(ws); } catch { }
                    }
                }
            }
        }

        /// <summary>
        /// Refreshes a chart link shape with updated data from Excel.
        /// </summary>
        /// <param name="shape">The linked shape to refresh</param>
        /// <param name="excelApp">Excel application instance</param>
        /// <param name="wbCache">Workbook cache for the refresh operation</param>
        /// <param name="warnings">Warning collector for error reporting</param>
        /// <returns>True if refresh succeeded, false otherwise</returns>
        public static bool RefreshChartLink(
            Word.Shape shape,
            Excel.Application excelApp,
            WorkbookCache wbCache,
            WarningCollector warnings)
        {
            if (shape == null || excelApp == null || wbCache == null)
                return false;

            string linkType = ShapeTagsHelper.GetTag(shape, TagLinkType);
            if (!string.Equals(linkType, LinkTypeChart, StringComparison.OrdinalIgnoreCase))
                return false;

            string wbPath = ShapeTagsHelper.GetTag(shape, TagWorkbookPath);
            string sheetName = ShapeTagsHelper.GetTag(shape, TagSheetName);
            string chartName = ShapeTagsHelper.GetTag(shape, TagChartName);

            if (string.IsNullOrEmpty(wbPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(chartName))
            {
                warnings?.AddWarning($"Chart link missing required tags");
                return false;
            }

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.ChartObject chartObj = null;

            try
            {
                // Get workbook
                wb = wbCache.GetOrOpen(wbPath, out bool newlyOpened, out string actualPath);
                if (wb == null)
                {
                    warnings?.AddWarning($"Could not open workbook: {CloudPathHelper.GetFileName(wbPath)}");
                    return false;
                }

                wbCache.PrepareWorkbookForCopy(wb);

                // Find worksheet
                try
                {
                    ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                }
                catch
                {
                    warnings?.AddWarning($"Worksheet '{sheetName}' not found");
                    return false;
                }

                if (ws == null)
                    return false;

                // Find chart
                try
                {
                    chartObj = ws.ChartObjects(chartName) as Excel.ChartObject;
                }
                catch
                {
                    warnings?.AddWarning($"Chart '{chartName}' not found on '{sheetName}'");
                    return false;
                }

                if (chartObj == null)
                    return false;

                // Handle "Match All" sizing - resize chart to match shape
                bool isMatchAll = SizeModeUtil.IsMatchAll(shape);
                float originalChartWidth = (float)chartObj.Width;
                float originalChartHeight = (float)chartObj.Height;
                bool chartResized = false;

                if (isMatchAll && SizeModeUtil.NeedsResize(chartObj, shape.Width, shape.Height))
                {
                    try
                    {
                        chartObj.Width = shape.Width;
                        chartObj.Height = shape.Height;
                        chartResized = true;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Could not resize chart: {ex.Message}");
                    }
                }

                try
                {
                    // Copy chart
                    chartObj.Copy();

                    if (!ClipboardHelper.WaitForClipboardReady(300))
                    {
                        System.Diagnostics.Debug.WriteLine("Clipboard not ready");
                    }

                    // Get document and range for paste
                    Word.Document doc = shape.Parent as Word.Document;
                    if (doc == null)
                    {
                        try { doc = shape.Anchor?.Document; } catch { }
                    }

                    if (doc == null)
                    {
                        warnings?.AddWarning("Could not determine document for shape");
                        return false;
                    }

                    // Cache position/size info
                    float oldLeft = shape.Left;
                    float oldTop = shape.Top;
                    float oldWidth = shape.Width;
                    float oldHeight = shape.Height;
                    var oldLockAspect = shape.LockAspectRatio;
                    int oldZOrder = shape.ZOrderPosition;
                    string sizeMode = ShapeTagsHelper.GetTag(shape, "Pragmatics_SizeMode");

                    // Get anchor position
                    Word.Range anchor = shape.Anchor;

                    // Delete old shape
                    shape.Delete();

                    // Paste new shape
                    Word.Shape newShape = null;
                    try
                    {
                        anchor.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);

                        // Find the new shape
                        if (anchor.InlineShapes.Count > 0)
                        {
                            var inlineShape = anchor.InlineShapes[anchor.InlineShapes.Count];
                            newShape = inlineShape.ConvertToShape();
                        }
                        else if (anchor.ShapeRange.Count > 0)
                        {
                            newShape = anchor.ShapeRange[anchor.ShapeRange.Count];
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"PasteSpecial failed: {ex.Message}");
                        return false;
                    }

                    if (newShape == null)
                    {
                        warnings?.AddWarning("Failed to paste chart image");
                        return false;
                    }

                    // Restore position and size
                    newShape.Left = oldLeft;
                    newShape.Top = oldTop;
                    newShape.LockAspectRatio = oldLockAspect;

                    if (isMatchAll)
                    {
                        newShape.LockAspectRatio = Office.MsoTriState.msoFalse;
                        newShape.Width = oldWidth;
                        newShape.Height = oldHeight;
                    }
                    else
                    {
                        // Let aspect ratio determine final size based on width
                        newShape.Width = oldWidth;
                    }

                    // Restore tags
                    string normalizedPath = CloudPathHelper.NormalizePath(actualPath);
                    ShapeTagsHelper.SetTag(newShape, TagLinkType, LinkTypeChart);
                    ShapeTagsHelper.SetTag(newShape, TagWorkbookPath, normalizedPath);
                    ShapeTagsHelper.SetTag(newShape, TagSheetName, sheetName);
                    ShapeTagsHelper.SetTag(newShape, TagChartName, chartName);
                    if (!string.IsNullOrEmpty(sizeMode))
                        ShapeTagsHelper.SetTag(newShape, "Pragmatics_SizeMode", sizeMode);

                    return true;
                }
                finally
                {
                    // Restore original chart size if we resized it
                    if (chartResized)
                    {
                        try
                        {
                            chartObj.Width = originalChartWidth;
                            chartObj.Height = originalChartHeight;
                        }
                        catch { }
                    }

                    ClipboardHelper.SafeClear();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error refreshing chart link: {ex.Message}");
                warnings?.AddWarning($"Error refreshing chart: {ex.Message}");
                return false;
            }
            finally
            {
                if (chartObj != null)
                {
                    try { Marshal.ReleaseComObject(chartObj); } catch { }
                }
                if (ws != null)
                {
                    try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }

        /// <summary>
        /// Checks if a shape is a chart link.
        /// </summary>
        public static bool IsChartLink(Word.Shape shape)
        {
            if (shape == null) return false;
            string linkType = ShapeTagsHelper.GetTag(shape, TagLinkType);
            return string.Equals(linkType, LinkTypeChart, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets chart link information from a shape.
        /// </summary>
        public static (string WorkbookPath, string SheetName, string ChartName) GetChartLinkInfo(Word.Shape shape)
        {
            if (shape == null || !IsChartLink(shape))
                return (null, null, null);

            return (
                ShapeTagsHelper.GetTag(shape, TagWorkbookPath),
                ShapeTagsHelper.GetTag(shape, TagSheetName),
                ShapeTagsHelper.GetTag(shape, TagChartName)
            );
        }

        /// <summary>
        /// Creates a chart link as a raster (PNG) image.
        /// </summary>
        private static Word.Shape CreateChartLinkAsPicture(
            Word.Document doc,
            Excel.Application excelApp,
            string workbookPath,
            string sheetName,
            string chartName,
            Word.Range range)
        {
            if (doc == null || excelApp == null || range == null)
                return null;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.ChartObject chartObj = null;
            Word.Shape newShape = null;

            using (var perfScope = new ExcelPerformanceScope(excelApp))
            using (var wbCache = new WorkbookCache(excelApp))
            {
                try
                {
                    string normalizedPath = CloudPathHelper.NormalizePath(workbookPath);

                    wb = wbCache.GetOrOpen(workbookPath, out bool newlyOpened, out string actualPath);
                    if (wb == null) return null;

                    wbCache.PrepareWorkbookForCopy(wb);

                    try { ws = wb.Worksheets[sheetName] as Excel.Worksheet; }
                    catch { return null; }

                    if (ws == null) return null;

                    try { chartObj = ws.ChartObjects(chartName) as Excel.ChartObject; }
                    catch { return null; }

                    if (chartObj == null) return null;

                    // Copy as bitmap/picture format
                    chartObj.Chart.CopyPicture(
                        Appearance: Excel.XlPictureAppearance.xlScreen,
                        Format: Excel.XlCopyPictureFormat.xlBitmap);

                    if (!ClipboardHelper.WaitForClipboardReady(300))
                    {
                        System.Diagnostics.Debug.WriteLine("Clipboard not ready");
                    }

                    try
                    {
                        range.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteBitmap);

                        if (range.InlineShapes.Count > 0)
                        {
                            var inlineShape = range.InlineShapes[range.InlineShapes.Count];
                            newShape = inlineShape.ConvertToShape();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Paste failed: {ex.Message}");
                        return null;
                    }

                    if (newShape == null) return null;

                    ShapeTagsHelper.SetTag(newShape, TagLinkType, LinkTypeChart);
                    ShapeTagsHelper.SetTag(newShape, TagWorkbookPath, normalizedPath);
                    ShapeTagsHelper.SetTag(newShape, TagSheetName, sheetName);
                    ShapeTagsHelper.SetTag(newShape, TagChartName, chartName);

                    return newShape;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating chart link as picture: {ex.Message}");
                    return null;
                }
                finally
                {
                    ClipboardHelper.SafeClear();

                    if (chartObj != null) try { Marshal.ReleaseComObject(chartObj); } catch { }
                    if (ws != null) try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }

        /// <summary>
        /// Creates an embedded native chart (editable in Word).
        /// </summary>
        private static Word.Shape CreateChartAsNative(
            Word.Document doc,
            Excel.Application excelApp,
            string workbookPath,
            string sheetName,
            string chartName,
            Word.Range range)
        {
            if (doc == null || excelApp == null || range == null)
                return null;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.ChartObject chartObj = null;
            Word.Shape newShape = null;

            using (var perfScope = new ExcelPerformanceScope(excelApp))
            using (var wbCache = new WorkbookCache(excelApp))
            {
                try
                {
                    string normalizedPath = CloudPathHelper.NormalizePath(workbookPath);

                    wb = wbCache.GetOrOpen(workbookPath, out bool newlyOpened, out string actualPath);
                    if (wb == null) return null;

                    wbCache.PrepareWorkbookForCopy(wb);

                    try { ws = wb.Worksheets[sheetName] as Excel.Worksheet; }
                    catch { return null; }

                    if (ws == null) return null;

                    try { chartObj = ws.ChartObjects(chartName) as Excel.ChartObject; }
                    catch { return null; }

                    if (chartObj == null) return null;

                    // Copy as OLE object
                    chartObj.Copy();

                    if (!ClipboardHelper.WaitForClipboardReady(300))
                    {
                        System.Diagnostics.Debug.WriteLine("Clipboard not ready");
                    }

                    try
                    {
                        // Paste as embedded object
                        range.PasteSpecial(
                            DataType: Word.WdPasteDataType.wdPasteOLEObject,
                            Placement: Word.WdOLEPlacement.wdInLine);

                        if (range.InlineShapes.Count > 0)
                        {
                            var inlineShape = range.InlineShapes[range.InlineShapes.Count];
                            newShape = inlineShape.ConvertToShape();
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Native paste failed: {ex.Message}");

                        // Fallback to regular paste
                        try
                        {
                            range.Paste();
                            if (range.InlineShapes.Count > 0)
                            {
                                var inlineShape = range.InlineShapes[range.InlineShapes.Count];
                                newShape = inlineShape.ConvertToShape();
                            }
                        }
                        catch
                        {
                            return null;
                        }
                    }

                    if (newShape == null) return null;

                    // Store tags for tracking (though native charts aren't "linked" in the same way)
                    ShapeTagsHelper.SetTag(newShape, TagLinkType, "ChartNative");
                    ShapeTagsHelper.SetTag(newShape, TagWorkbookPath, normalizedPath);
                    ShapeTagsHelper.SetTag(newShape, TagSheetName, sheetName);
                    ShapeTagsHelper.SetTag(newShape, TagChartName, chartName);

                    return newShape;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating native chart: {ex.Message}");
                    return null;
                }
                finally
                {
                    ClipboardHelper.SafeClear();

                    if (chartObj != null) try { Marshal.ReleaseComObject(chartObj); } catch { }
                    if (ws != null) try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }
    }

    /// <summary>
    /// Dialog for selecting an Excel chart to import.
    /// </summary>
    internal class ChartSelectionDialog : Form
    {
        private ComboBox cboWorkbook;
        private ComboBox cboSheet;
        private ComboBox cboChart;
        private Button btnOK;
        private Button btnCancel;
        private readonly Excel.Application _excelApp;

        public string SelectedWorkbookPath { get; private set; }
        public string SelectedSheetName { get; private set; }
        public string SelectedChartName { get; private set; }

        public ChartSelectionDialog(Excel.Application excelApp)
        {
            _excelApp = excelApp;
            InitializeComponent();
            LoadWorkbooks();
        }

        private void InitializeComponent()
        {
            this.Text = "Select Excel Chart";
            this.Size = new System.Drawing.Size(400, 220);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblWorkbook = new Label { Text = "Workbook:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            cboWorkbook = new ComboBox { Location = new System.Drawing.Point(100, 17), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
            cboWorkbook.SelectedIndexChanged += CboWorkbook_SelectedIndexChanged;

            var lblSheet = new Label { Text = "Worksheet:", Location = new System.Drawing.Point(20, 55), AutoSize = true };
            cboSheet = new ComboBox { Location = new System.Drawing.Point(100, 52), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };
            cboSheet.SelectedIndexChanged += CboSheet_SelectedIndexChanged;

            var lblChart = new Label { Text = "Chart:", Location = new System.Drawing.Point(20, 90), AutoSize = true };
            cboChart = new ComboBox { Location = new System.Drawing.Point(100, 87), Width = 260, DropDownStyle = ComboBoxStyle.DropDownList };

            btnOK = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(200, 140), Width = 75 };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(285, 140), Width = 75 };

            this.Controls.AddRange(new Control[] { lblWorkbook, cboWorkbook, lblSheet, cboSheet, lblChart, cboChart, btnOK, btnCancel });
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadWorkbooks()
        {
            cboWorkbook.Items.Clear();
            foreach (Excel.Workbook wb in _excelApp.Workbooks)
            {
                cboWorkbook.Items.Add(new WorkbookItem { Name = wb.Name, Path = wb.FullName });
            }
            if (cboWorkbook.Items.Count > 0)
                cboWorkbook.SelectedIndex = 0;
        }

        private void CboWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboSheet.Items.Clear();
            cboChart.Items.Clear();

            if (cboWorkbook.SelectedItem is WorkbookItem wbItem)
            {
                try
                {
                    Excel.Workbook wb = _excelApp.Workbooks[wbItem.Name];
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        cboSheet.Items.Add(ws.Name);
                    }
                    if (cboSheet.Items.Count > 0)
                        cboSheet.SelectedIndex = 0;
                }
                catch { }
            }
        }

        private void CboSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboChart.Items.Clear();

            if (cboWorkbook.SelectedItem is WorkbookItem wbItem && cboSheet.SelectedItem is string sheetName)
            {
                try
                {
                    Excel.Workbook wb = _excelApp.Workbooks[wbItem.Name];
                    Excel.Worksheet ws = wb.Worksheets[sheetName];
                    Excel.ChartObjects charts = ws.ChartObjects() as Excel.ChartObjects;

                    if (charts != null)
                    {
                        foreach (Excel.ChartObject chart in charts)
                        {
                            cboChart.Items.Add(chart.Name);
                        }
                    }

                    if (cboChart.Items.Count > 0)
                        cboChart.SelectedIndex = 0;
                }
                catch { }
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (cboWorkbook.SelectedItem is WorkbookItem wbItem)
            {
                SelectedWorkbookPath = wbItem.Path;
                SelectedSheetName = cboSheet.SelectedItem as string;
                SelectedChartName = cboChart.SelectedItem as string;
            }
        }

        private class WorkbookItem
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public override string ToString() => Name;
        }
    }
}