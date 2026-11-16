// RangeLinkUtility.cs (.NET Framework 4.8 / C# 7.3)
// Range linking utilities for Word - copies Excel ranges as pictures to Word shapes.

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// Provides Excel range linking functionality for Word documents.
    /// Copies Excel ranges as enhanced metafile pictures.
    /// </summary>
    internal static class RangeLinkUtility
    {
        private const string TagWorkbookPath = "Pragmatics_WbPath";
        private const string TagSheetName = "Pragmatics_SheetName";
        private const string TagRangeAddress = "Pragmatics_RangeAddr";
        private const string TagLinkType = "Pragmatics_LinkType";
        private const string LinkTypeRange = "Range";

        #region Public Import Methods (Ribbon Entry Points)

        /// <summary>
        /// Imports an Excel range as a picture (EMF) with linking.
        /// Shows a dialog to select the range from an open Excel workbook.
        /// </summary>
        public static Word.Shape ImportExcelRangeAsPicture(Word.Application wordApp)
        {
            if (wordApp == null) return null;

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                MessageBox.Show("No active document.", "Import Range", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (doc == null)
            {
                MessageBox.Show("No active document.", "Import Range", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            Excel.Application excelApp;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                MessageBox.Show("Excel is not running. Please open an Excel workbook first.",
                    "Import Range", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (excelApp.Workbooks.Count == 0)
            {
                MessageBox.Show("No Excel workbooks are open.",
                    "Import Range", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            using (var dialog = new RangeSelectionDialog(excelApp))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return null;

                string wbPath = dialog.SelectedWorkbookPath;
                string sheetName = dialog.SelectedSheetName;
                string rangeAddress = dialog.SelectedRangeAddress;

                if (string.IsNullOrEmpty(wbPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(rangeAddress))
                    return null;

                Word.Range insertRange = wordApp.Selection.Range;
                return CreateRangeLink(doc, excelApp, wbPath, sheetName, rangeAddress, insertRange);
            }
        }

        /// <summary>
        /// Imports an Excel range as a native Word table.
        /// Shows a dialog to select the range from an open Excel workbook.
        /// </summary>
        public static Word.Table ImportExcelRangeAsTable(Word.Application wordApp)
        {
            if (wordApp == null) return null;

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                MessageBox.Show("No active document.", "Import Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (doc == null)
            {
                MessageBox.Show("No active document.", "Import Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            Excel.Application excelApp;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                MessageBox.Show("Excel is not running. Please open an Excel workbook first.",
                    "Import Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (excelApp.Workbooks.Count == 0)
            {
                MessageBox.Show("No Excel workbooks are open.",
                    "Import Table", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            using (var dialog = new RangeSelectionDialog(excelApp))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return null;

                string wbPath = dialog.SelectedWorkbookPath;
                string sheetName = dialog.SelectedSheetName;
                string rangeAddress = dialog.SelectedRangeAddress;

                if (string.IsNullOrEmpty(wbPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(rangeAddress))
                    return null;

                Word.Range insertRange = wordApp.Selection.Range;
                return CreateRangeAsTable(doc, excelApp, wbPath, sheetName, rangeAddress, insertRange);
            }
        }

        #endregion

        /// <summary>
        /// Creates a range link in Word by copying an Excel range as a picture.
        /// </summary>
        public static Word.Shape CreateRangeLink(
            Word.Document doc,
            Excel.Application excelApp,
            string workbookPath,
            string sheetName,
            string rangeAddress,
            Word.Range insertRange)
        {
            if (doc == null || excelApp == null || insertRange == null)
                return null;

            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(rangeAddress))
                return null;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;
            Word.Shape newShape = null;

            using (var perfScope = new ExcelPerformanceScope(excelApp))
            using (var wbCache = new WorkbookCache(excelApp))
            {
                try
                {
                    // Normalize path for cross-user compatibility
                    string normalizedPath = CloudPathHelper.NormalizePath(workbookPath);

                    // Get workbook
                    wb = wbCache.GetOrOpen(workbookPath, out bool newlyOpened, out string actualPath);
                    if (wb == null)
                    {
                        MessageBox.Show($"Could not open workbook:\n{workbookPath}",
                            "Range Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    wbCache.PrepareWorkbookForCopy(wb);

                    // Find worksheet
                    try
                    {
                        ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                    }
                    catch
                    {
                        MessageBox.Show($"Could not find worksheet '{sheetName}' in workbook.",
                            "Range Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    if (ws == null)
                        return null;

                    // Get range
                    try
                    {
                        range = ws.Range[rangeAddress];
                    }
                    catch
                    {
                        MessageBox.Show($"Invalid range address: '{rangeAddress}'",
                            "Range Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    if (range == null)
                        return null;

                    // Copy range as picture
                    CopyRangeAsPicture(range);

                    // Wait for clipboard
                    if (!ClipboardHelper.WaitForClipboardReady(300))
                    {
                        System.Diagnostics.Debug.WriteLine("Clipboard not ready after copying range");
                    }

                    // Paste into Word
                    try
                    {
                        newShape = ClipboardHelper.RetryPaste(() =>
                        {
                            insertRange.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);

                            if (insertRange.InlineShapes.Count > 0)
                            {
                                var inlineShape = insertRange.InlineShapes[insertRange.InlineShapes.Count];
                                return inlineShape.ConvertToShape();
                            }
                            else if (insertRange.ShapeRange.Count > 0)
                            {
                                return insertRange.ShapeRange[insertRange.ShapeRange.Count];
                            }
                            return null;
                        });
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Paste failed: {ex.Message}");
                        return null;
                    }

                    if (newShape == null)
                        return null;

                    // Store link information
                    ShapeTagsHelper.SetTag(newShape, TagLinkType, LinkTypeRange);
                    ShapeTagsHelper.SetTag(newShape, TagWorkbookPath, normalizedPath);
                    ShapeTagsHelper.SetTag(newShape, TagSheetName, sheetName);
                    ShapeTagsHelper.SetTag(newShape, TagRangeAddress, rangeAddress);

                    return newShape;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating range link: {ex.Message}");
                    return null;
                }
                finally
                {
                    ClipboardHelper.SafeClear();

                    if (range != null)
                    {
                        try { Marshal.ReleaseComObject(range); } catch { }
                    }
                    if (ws != null)
                    {
                        try { Marshal.ReleaseComObject(ws); } catch { }
                    }
                }
            }
        }

        /// <summary>
        /// Refreshes a range link shape with updated data from Excel.
        /// </summary>
        public static bool RefreshRangeLink(
            Word.Shape shape,
            Excel.Application excelApp,
            WorkbookCache wbCache,
            WarningCollector warnings)
        {
            if (shape == null || excelApp == null || wbCache == null)
                return false;

            string linkType = ShapeTagsHelper.GetTag(shape, TagLinkType);
            if (!string.Equals(linkType, LinkTypeRange, StringComparison.OrdinalIgnoreCase))
                return false;

            string wbPath = ShapeTagsHelper.GetTag(shape, TagWorkbookPath);
            string sheetName = ShapeTagsHelper.GetTag(shape, TagSheetName);
            string rangeAddress = ShapeTagsHelper.GetTag(shape, TagRangeAddress);

            if (string.IsNullOrEmpty(wbPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(rangeAddress))
            {
                warnings?.AddWarning("Range link missing required tags");
                return false;
            }

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

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

                // Get range
                try
                {
                    range = ws.Range[rangeAddress];
                }
                catch
                {
                    warnings?.AddWarning($"Invalid range: '{rangeAddress}'");
                    return false;
                }

                if (range == null)
                    return false;

                // Cache shape properties
                float oldLeft = shape.Left;
                float oldTop = shape.Top;
                float oldWidth = shape.Width;
                float oldHeight = shape.Height;
                var oldLockAspect = shape.LockAspectRatio;
                string sizeMode = ShapeTagsHelper.GetTag(shape, "Pragmatics_SizeMode");
                bool isMatchAll = SizeModeUtil.IsMatchAll(shape);

                // Get document for paste
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

                // Get anchor before deleting shape
                Word.Range anchor = shape.Anchor;

                // Copy range with optional temp resize for "Match All" mode
                if (isMatchAll)
                {
                    CopyRangeWithTempResize(range, oldWidth, oldHeight, ws);
                }
                else
                {
                    CopyRangeAsPicture(range);
                }

                if (!ClipboardHelper.WaitForClipboardReady(300))
                {
                    System.Diagnostics.Debug.WriteLine("Clipboard not ready");
                }

                // Delete old shape
                shape.Delete();

                // Paste new shape
                Word.Shape newShape = null;
                try
                {
                    anchor.PasteSpecial(DataType: Word.WdPasteDataType.wdPasteEnhancedMetafile);

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
                    System.Diagnostics.Debug.WriteLine($"Paste failed: {ex.Message}");
                    warnings?.AddWarning("Failed to paste range image");
                    return false;
                }

                if (newShape == null)
                {
                    warnings?.AddWarning("Failed to paste range image");
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
                    newShape.Width = oldWidth;
                }

                // Restore tags
                string normalizedPath = CloudPathHelper.NormalizePath(actualPath);
                ShapeTagsHelper.SetTag(newShape, TagLinkType, LinkTypeRange);
                ShapeTagsHelper.SetTag(newShape, TagWorkbookPath, normalizedPath);
                ShapeTagsHelper.SetTag(newShape, TagSheetName, sheetName);
                ShapeTagsHelper.SetTag(newShape, TagRangeAddress, rangeAddress);
                if (!string.IsNullOrEmpty(sizeMode))
                    ShapeTagsHelper.SetTag(newShape, "Pragmatics_SizeMode", sizeMode);

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error refreshing range link: {ex.Message}");
                warnings?.AddWarning($"Error refreshing range: {ex.Message}");
                return false;
            }
            finally
            {
                ClipboardHelper.SafeClear();

                if (range != null)
                {
                    try { Marshal.ReleaseComObject(range); } catch { }
                }
                if (ws != null)
                {
                    try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }

        /// <summary>
        /// Copies an Excel range as a picture to the clipboard.
        /// </summary>
        private static void CopyRangeAsPicture(Excel.Range range)
        {
            if (range == null) return;

            try
            {
                // Copy as picture with screen appearance and picture format
                range.CopyPicture(
                    Appearance: Excel.XlPictureAppearance.xlScreen,
                    Format: Excel.XlCopyPictureFormat.xlPicture);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CopyPicture failed: {ex.Message}");

                // Fallback: regular copy
                try
                {
                    range.Copy();
                }
                catch (Exception ex2)
                {
                    System.Diagnostics.Debug.WriteLine($"Regular copy also failed: {ex2.Message}");
                }
            }
        }

        /// <summary>
        /// Copies a range with temporary column width adjustment for "Match All" sizing.
        /// This produces a higher-quality image that matches the target dimensions.
        /// </summary>
        private static void CopyRangeWithTempResize(Excel.Range range, float targetWidth, float targetHeight, Excel.Worksheet ws)
        {
            if (range == null || ws == null)
            {
                CopyRangeAsPicture(range);
                return;
            }

            // Store original column widths
            var originalWidths = new System.Collections.Generic.Dictionary<int, double>();
            var originalRowHeights = new System.Collections.Generic.Dictionary<int, double>();

            try
            {
                // Get range dimensions in points
                double rangeWidth = 0;
                double rangeHeight = 0;

                // Calculate current width (sum of column widths in points)
                for (int col = range.Column; col < range.Column + range.Columns.Count; col++)
                {
                    Excel.Range colRange = ws.Columns[col] as Excel.Range;
                    double colWidth = Convert.ToDouble(colRange.Width); // Width in points
                    originalWidths[col] = Convert.ToDouble(colRange.ColumnWidth); // Width in characters
                    rangeWidth += colWidth;
                    Marshal.ReleaseComObject(colRange);
                }

                // Calculate current height (sum of row heights in points)
                for (int row = range.Row; row < range.Row + range.Rows.Count; row++)
                {
                    Excel.Range rowRange = ws.Rows[row] as Excel.Range;
                    double rowHeight = Convert.ToDouble(rowRange.RowHeight);
                    originalRowHeights[row] = rowHeight;
                    rangeHeight += rowHeight;
                    Marshal.ReleaseComObject(rowRange);
                }

                // Calculate scale factors
                double widthScale = targetWidth / rangeWidth;
                double heightScale = targetHeight / rangeHeight;

                // Only resize if scale is significant (more than 5% difference)
                bool needsWidthResize = Math.Abs(widthScale - 1.0) > 0.05;
                bool needsHeightResize = Math.Abs(heightScale - 1.0) > 0.05;

                if (needsWidthResize)
                {
                    // Adjust column widths proportionally
                    for (int col = range.Column; col < range.Column + range.Columns.Count; col++)
                    {
                        Excel.Range colRange = ws.Columns[col] as Excel.Range;
                        double newWidth = originalWidths[col] * widthScale;
                        // Clamp to Excel limits (0 to 255 characters)
                        newWidth = Math.Max(0.1, Math.Min(255, newWidth));
                        colRange.ColumnWidth = newWidth;
                        Marshal.ReleaseComObject(colRange);
                    }
                }

                if (needsHeightResize)
                {
                    // Adjust row heights proportionally
                    for (int row = range.Row; row < range.Row + range.Rows.Count; row++)
                    {
                        Excel.Range rowRange = ws.Rows[row] as Excel.Range;
                        double newHeight = originalRowHeights[row] * heightScale;
                        // Clamp to Excel limits (0 to 409 points)
                        newHeight = Math.Max(0.5, Math.Min(409, newHeight));
                        rowRange.RowHeight = newHeight;
                        Marshal.ReleaseComObject(rowRange);
                    }
                }

                // Copy the resized range
                CopyRangeAsPicture(range);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in temp resize: {ex.Message}");
                CopyRangeAsPicture(range);
            }
            finally
            {
                // Restore original column widths
                foreach (var kvp in originalWidths)
                {
                    try
                    {
                        Excel.Range colRange = ws.Columns[kvp.Key] as Excel.Range;
                        colRange.ColumnWidth = kvp.Value;
                        Marshal.ReleaseComObject(colRange);
                    }
                    catch { }
                }

                // Restore original row heights
                foreach (var kvp in originalRowHeights)
                {
                    try
                    {
                        Excel.Range rowRange = ws.Rows[kvp.Key] as Excel.Range;
                        rowRange.RowHeight = kvp.Value;
                        Marshal.ReleaseComObject(rowRange);
                    }
                    catch { }
                }
            }
        }

        /// <summary>
        /// Checks if a shape is a range link.
        /// </summary>
        public static bool IsRangeLink(Word.Shape shape)
        {
            if (shape == null) return false;
            string linkType = ShapeTagsHelper.GetTag(shape, TagLinkType);
            return string.Equals(linkType, LinkTypeRange, StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Gets range link information from a shape.
        /// </summary>
        public static (string WorkbookPath, string SheetName, string RangeAddress) GetRangeLinkInfo(Word.Shape shape)
        {
            if (shape == null || !IsRangeLink(shape))
                return (null, null, null);

            return (
                ShapeTagsHelper.GetTag(shape, TagWorkbookPath),
                ShapeTagsHelper.GetTag(shape, TagSheetName),
                ShapeTagsHelper.GetTag(shape, TagRangeAddress)
            );
        }

        /// <summary>
        /// Creates a native Word table from an Excel range.
        /// </summary>
        private static Word.Table CreateRangeAsTable(
            Word.Document doc,
            Excel.Application excelApp,
            string workbookPath,
            string sheetName,
            string rangeAddress,
            Word.Range insertRange)
        {
            if (doc == null || excelApp == null || insertRange == null)
                return null;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range range = null;

            using (var perfScope = new ExcelPerformanceScope(excelApp))
            using (var wbCache = new WorkbookCache(excelApp))
            {
                try
                {
                    wb = wbCache.GetOrOpen(workbookPath, out bool newlyOpened, out string actualPath);
                    if (wb == null) return null;

                    try { ws = wb.Worksheets[sheetName] as Excel.Worksheet; }
                    catch { return null; }

                    if (ws == null) return null;

                    try { range = ws.Range[rangeAddress]; }
                    catch { return null; }

                    if (range == null) return null;

                    // Copy the range
                    range.Copy();

                    if (!ClipboardHelper.WaitForClipboardReady(300))
                    {
                        System.Diagnostics.Debug.WriteLine("Clipboard not ready");
                    }

                    // Paste as formatted text (creates Word table)
                    try
                    {
                        insertRange.PasteExcelTable(
                            LinkedToExcel: false,
                            WordFormatting: false,
                            RTF: true);

                        // Find the newly created table
                        if (insertRange.Tables.Count > 0)
                        {
                            return insertRange.Tables[insertRange.Tables.Count];
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"PasteExcelTable failed: {ex.Message}");

                        // Fallback
                        try
                        {
                            insertRange.Paste();
                            if (insertRange.Tables.Count > 0)
                            {
                                return insertRange.Tables[insertRange.Tables.Count];
                            }
                        }
                        catch
                        {
                            return null;
                        }
                    }

                    return null;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating table: {ex.Message}");
                    return null;
                }
                finally
                {
                    ClipboardHelper.SafeClear();

                    if (range != null) try { Marshal.ReleaseComObject(range); } catch { }
                    if (ws != null) try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }
    }

    /// <summary>
    /// Dialog for selecting an Excel range to import.
    /// </summary>
    internal class RangeSelectionDialog : Form
    {
        private ComboBox cboWorkbook;
        private ComboBox cboSheet;
        private TextBox txtRange;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSelectFromExcel;
        private readonly Excel.Application _excelApp;

        public string SelectedWorkbookPath { get; private set; }
        public string SelectedSheetName { get; private set; }
        public string SelectedRangeAddress { get; private set; }

        public RangeSelectionDialog(Excel.Application excelApp)
        {
            _excelApp = excelApp;
            InitializeComponent();
            LoadWorkbooks();
        }

        private void InitializeComponent()
        {
            this.Text = "Select Excel Range";
            this.Size = new System.Drawing.Size(420, 220);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblWorkbook = new Label { Text = "Workbook:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            cboWorkbook = new ComboBox { Location = new System.Drawing.Point(100, 17), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
            cboWorkbook.SelectedIndexChanged += CboWorkbook_SelectedIndexChanged;

            var lblSheet = new Label { Text = "Worksheet:", Location = new System.Drawing.Point(20, 55), AutoSize = true };
            cboSheet = new ComboBox { Location = new System.Drawing.Point(100, 52), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };

            var lblRange = new Label { Text = "Range:", Location = new System.Drawing.Point(20, 90), AutoSize = true };
            txtRange = new TextBox { Location = new System.Drawing.Point(100, 87), Width = 200 };

            btnSelectFromExcel = new Button { Text = "Select", Location = new System.Drawing.Point(310, 85), Width = 70 };
            btnSelectFromExcel.Click += BtnSelectFromExcel_Click;

            btnOK = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(220, 140), Width = 75 };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(305, 140), Width = 75 };

            this.Controls.AddRange(new Control[] { lblWorkbook, cboWorkbook, lblSheet, cboSheet, lblRange, txtRange, btnSelectFromExcel, btnOK, btnCancel });
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

        private void BtnSelectFromExcel_Click(object sender, EventArgs e)
        {
            try
            {
                this.WindowState = FormWindowState.Minimized;

                MessageBox.Show("Select a range in Excel and click OK.",
                    "Select Range", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // Get current selection from Excel
                Excel.Range selection = _excelApp.Selection as Excel.Range;
                if (selection != null)
                {
                    Excel.Worksheet ws = selection.Worksheet;
                    Excel.Workbook wb = ws.Parent as Excel.Workbook;

                    // Find and select the workbook in dropdown
                    for (int i = 0; i < cboWorkbook.Items.Count; i++)
                    {
                        if (cboWorkbook.Items[i] is WorkbookItem wbItem && wbItem.Name == wb.Name)
                        {
                            cboWorkbook.SelectedIndex = i;
                            break;
                        }
                    }

                    // Select the sheet
                    for (int i = 0; i < cboSheet.Items.Count; i++)
                    {
                        if (cboSheet.Items[i].ToString() == ws.Name)
                        {
                            cboSheet.SelectedIndex = i;
                            break;
                        }
                    }

                    // Set the range address
                    txtRange.Text = selection.Address[false, false];
                }

                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            }
            catch (Exception ex)
            {
                this.WindowState = FormWindowState.Normal;
                MessageBox.Show($"Error getting selection: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (cboWorkbook.SelectedItem is WorkbookItem wbItem)
            {
                SelectedWorkbookPath = wbItem.Path;
                SelectedSheetName = cboSheet.SelectedItem as string;
                SelectedRangeAddress = txtRange.Text.Trim();
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