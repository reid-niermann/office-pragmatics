using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OfficePragmatics_Powerpoint; // Added to access Ribbon1

namespace OfficePragmatics_PowerPoint.ExcelLinking
{
    public static class RangeLinkUtility
    {
        /* ═════════════════════  1. HELPER: GET RANGE DIMENSIONS  ═════════════════════ */
        // (No changes to GetVisibleRowsAndColumns or GetRangeDimensionsWithVisibility)

        public static void GetVisibleRowsAndColumns(Excel.Range range,
            out bool[] visibleRows, out bool[] visibleCols,
            out int visibleRowCount, out int visibleColCount)
        {
            int totalRows = range.Rows.Count;
            int totalCols = range.Columns.Count;

            visibleRows = new bool[totalRows];
            visibleCols = new bool[totalCols];
            visibleRowCount = 0;
            visibleColCount = 0;

            Excel.Range rowsRange = range.Rows;
            for (int r = 1; r <= totalRows; r++)
            {
                Excel.Range row = rowsRange[r];
                bool isVisible = (double)row.Height > 0;
                visibleRows[r - 1] = isVisible;
                if (isVisible) visibleRowCount++;
                Marshal.ReleaseComObject(row);
            }
            Marshal.ReleaseComObject(rowsRange);

            Excel.Range colsRange = range.Columns;
            for (int c = 1; c <= totalCols; c++)
            {
                Excel.Range col = colsRange[c];
                bool isVisible = (double)col.ColumnWidth > 0;
                visibleCols[c - 1] = isVisible;
                if (isVisible) visibleColCount++;
                Marshal.ReleaseComObject(col);
            }
            Marshal.ReleaseComObject(colsRange);
        }

        public static void GetRangeDimensionsWithVisibility(Excel.Range range,
            bool[] visibleRows, bool[] visibleCols,
            out float width, out float height,
            out float[] columnWidths, out float[] rowHeights)
        {
            int totalCols = range.Columns.Count;
            int totalRows = range.Rows.Count;

            int vRows = 0; for (int i = 0; i < totalRows; i++) if (visibleRows[i]) vRows++;
            int vCols = 0; for (int i = 0; i < totalCols; i++) if (visibleCols[i]) vCols++;

            columnWidths = new float[vCols];
            rowHeights = new float[vRows];
            width = 0f;
            height = 0f;

            Excel.Range columnsRange = range.Columns;
            int visColIdx = 0;
            for (int c = 1; c <= totalCols; c++)
            {
                if (visibleCols[c - 1])
                {
                    Excel.Range col = columnsRange[c];
                    float colWidth = (float)(double)col.Width;
                    columnWidths[visColIdx++] = colWidth;
                    width += colWidth;
                    Marshal.ReleaseComObject(col);
                }
            }
            Marshal.ReleaseComObject(columnsRange);

            Excel.Range rowsRange = range.Rows;
            int visRowIdx = 0;
            for (int r = 1; r <= totalRows; r++)
            {
                if (visibleRows[r - 1])
                {
                    Excel.Range row = rowsRange[r];
                    float rowHeight = (float)(double)row.Height;
                    rowHeights[visRowIdx++] = rowHeight;
                    height += rowHeight;
                    Marshal.ReleaseComObject(row);
                }
            }
            Marshal.ReleaseComObject(rowsRange);
        }

        public static void GetRangeDimensions(Excel.Range range, out float width, out float height,
            out float[] columnWidths, out float[] rowHeights)
        {
            GetVisibleRowsAndColumns(range, out bool[] vR, out bool[] vC, out int r, out int c);
            GetRangeDimensionsWithVisibility(range, vR, vC, out width, out height, out columnWidths, out rowHeights);
        }

        /* ═════════════════════  2. TEMP RESIZING (THE FIX)  ═════════════════════ */

        /// <summary>
        /// Temporarily resizes the Excel range columns (and rows if Match All) to match the target PowerPoint width/height.
        /// </summary>
        public static void CopyRangeWithTempResize(Excel.Range range, float targetWidth, float targetHeight, bool scaleHeight, Action copyAction)
        {
            if (range == null || targetWidth <= 0)
            {
                copyAction();
                return;
            }

            Excel.Range cols = range.Columns;
            Excel.Range rows = range.Rows;
            int colCount = cols.Count;
            int rowCount = rows.Count;

            double[] originalColWidths = new double[colCount];
            double[] originalRowHeights = new double[rowCount];
            double currentTotalWidth = 0;
            double currentTotalHeight = 0;

            // 1. Capture original state
            try
            {
                for (int i = 1; i <= colCount; i++)
                {
                    Excel.Range c = cols[i];
                    originalColWidths[i - 1] = (double)c.ColumnWidth;
                    currentTotalWidth += (double)c.Width;
                    Marshal.ReleaseComObject(c);
                }

                if (scaleHeight)
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        Excel.Range r = rows[i];
                        originalRowHeights[i - 1] = (double)r.RowHeight;
                        currentTotalHeight += (double)r.Height;
                        Marshal.ReleaseComObject(r);
                    }
                }
            }
            catch
            {
                if (cols != null) Marshal.ReleaseComObject(cols);
                if (rows != null) Marshal.ReleaseComObject(rows);
                copyAction();
                return;
            }

            // 2. Calculate Scale
            double widthScale = targetWidth / currentTotalWidth;
            double heightScale = scaleHeight && targetHeight > 0 ? targetHeight / currentTotalHeight : 1.0;

            if (Math.Abs(widthScale - 1.0) < 0.01 && Math.Abs(heightScale - 1.0) < 0.01)
            {
                Marshal.ReleaseComObject(cols);
                Marshal.ReleaseComObject(rows);
                copyAction();
                return;
            }

            // 3. Apply Scale & Copy
            try
            {
                for (int i = 1; i <= colCount; i++)
                {
                    Excel.Range c = cols[i];
                    c.ColumnWidth = originalColWidths[i - 1] * widthScale;
                    Marshal.ReleaseComObject(c);
                }

                if (scaleHeight)
                {
                    for (int i = 1; i <= rowCount; i++)
                    {
                        Excel.Range r = rows[i];
                        r.RowHeight = originalRowHeights[i - 1] * heightScale;
                        Marshal.ReleaseComObject(r);
                    }
                }

                System.Windows.Forms.Application.DoEvents();
                copyAction();
            }
            finally
            {
                // 4. Restore Original State
                try
                {
                    for (int i = 1; i <= colCount; i++)
                    {
                        Excel.Range c = cols[i];
                        c.ColumnWidth = originalColWidths[i - 1];
                        Marshal.ReleaseComObject(c);
                    }

                    if (scaleHeight)
                    {
                        for (int i = 1; i <= rowCount; i++)
                        {
                            Excel.Range r = rows[i];
                            r.RowHeight = originalRowHeights[i - 1];
                            Marshal.ReleaseComObject(r);
                        }
                    }
                }
                catch { }

                Marshal.ReleaseComObject(cols);
                Marshal.ReleaseComObject(rows);
            }
        }

        /* ═════════════════════  3. CACHED CELL DATA  ═════════════════════ */
        // (No changes to CachedCellData or ApplyStructuralFixes)

        public class CachedCellData
        {
            public bool IsMerged;
            public int MergeRowSpan;
            public int MergeColSpan;
            public string NumberFormat;
        }

        public static CachedCellData[,] BulkReadCellDataVisible(Excel.Range range,
            bool[] visibleRows, bool[] visibleCols,
            int visibleRowCount, int visibleColCount)
        {
            var data = new CachedCellData[visibleRowCount, visibleColCount];
            if (visibleRowCount == 0 || visibleColCount == 0) return data;

            int totalRows = visibleRows.Length;
            int totalCols = visibleCols.Length;

            int visRowIdx = 0;
            for (int r = 1; r <= totalRows; r++)
            {
                if (!visibleRows[r - 1]) continue;

                int visColIdx = 0;
                for (int c = 1; c <= totalCols; c++)
                {
                    if (!visibleCols[c - 1]) continue;

                    data[visRowIdx, visColIdx] = new CachedCellData();
                    var cellData = data[visRowIdx, visColIdx];
                    Excel.Range cell = range.Cells[r, c];

                    bool isMerged = (bool)cell.MergeCells;
                    if (isMerged)
                    {
                        Excel.Range mergeArea = cell.MergeArea;
                        if (mergeArea.Cells[1, 1].Address == cell.Address)
                        {
                            cellData.IsMerged = true;
                            cellData.MergeRowSpan = mergeArea.Rows.Count;
                            cellData.MergeColSpan = mergeArea.Columns.Count;
                        }
                        Marshal.ReleaseComObject(mergeArea);
                    }

                    try
                    {
                        object fmt = cell.NumberFormat;
                        cellData.NumberFormat = fmt != null ? fmt.ToString() : "";
                    }
                    catch { cellData.NumberFormat = ""; }

                    Marshal.ReleaseComObject(cell);
                    visColIdx++;
                }
                visRowIdx++;
            }
            return data;
        }

        public static void ApplyStructuralFixes(PowerPoint.Cell ppCell, string numberFormat)
        {
            var shape = ppCell.Shape;
            var textFrame = shape.TextFrame;
            var textRange = textFrame.TextRange;

            string rawText = textRange.Text;
            if (!string.IsNullOrEmpty(rawText))
            {
                string trimmed = rawText.Trim();
                bool isAccountingFormat = double.TryParse(trimmed,
                    System.Globalization.NumberStyles.Any,
                    System.Globalization.CultureInfo.CurrentCulture,
                    out _);

                if (rawText.Length != trimmed.Length) textRange.Text = trimmed;
                if (isAccountingFormat) textRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
            }

            const float ACCOUNTING_MARGIN = 3f;
            if (!string.IsNullOrEmpty(numberFormat))
            {
                if (numberFormat.Contains("_)")) textFrame.MarginRight = ACCOUNTING_MARGIN;
                if (numberFormat.Contains("_(")) textFrame.MarginLeft = ACCOUNTING_MARGIN;
            }

            textFrame.WordWrap = Office.MsoTriState.msoFalse;
            textFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
        }

        /* ═════════════════════  4. SHARED PASTE LOGIC  ═════════════════════ */

        // sizeMode decides whether Height is scaled
        private static PowerPoint.Shape PasteAndFixTable(PowerPoint.Application ppApp, PowerPoint.Slide slide, Excel.Range sourceRange, float targetLeft, float targetTop, float targetWidth, float targetHeight, string sizeMode)
        {
            // 1. READ STRUCTURE
            GetVisibleRowsAndColumns(sourceRange, out bool[] vRows, out bool[] vCols, out int rows, out int cols);
            if (rows == 0 || cols == 0) return null;

            var cellData = BulkReadCellDataVisible(sourceRange, vRows, vCols, rows, cols);

            // Calculate native proportions
            GetRangeDimensionsWithVisibility(sourceRange, vRows, vCols, out float sumColW, out float sumRowH, out float[] colWidths, out float[] rowHeights);

            // 2. COPY (With Resize Logic)
            Action copyOp = () => {
                sourceRange.Copy();
                System.Threading.Thread.Sleep(50);
            };

            // Check if we need to resize
            bool useResizeLogic = targetWidth > 0;
            // Only scale height if user specifically requested "Match All"
            bool scaleHeight = useResizeLogic && string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase);

            if (useResizeLogic)
            {
                CopyRangeWithTempResize(sourceRange, targetWidth, targetHeight, scaleHeight, copyOp);
            }
            else
            {
                copyOp();
            }

            // 3. PASTE
            PowerPoint.Shape newShape = null;
            try
            {
                ppApp.CommandBars.ExecuteMso("PasteSourceFormatting");
                System.Windows.Forms.Application.DoEvents();
                if (ppApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                    newShape = ppApp.ActiveWindow.Selection.ShapeRange[1];
            }
            catch
            {
                try { newShape = slide.Shapes.Paste()[1]; } catch { return null; }
            }

            if (newShape == null || newShape.HasTable != Office.MsoTriState.msoTrue) return newShape;

            var table = newShape.Table;

            // 3.5 CLEANUP GHOST COLUMNS
            while (table.Columns.Count > cols)
            {
                try { table.Columns[table.Columns.Count].Delete(); } catch { break; }
            }

            // 4. APPLY FIXES
            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    var data = cellData[r, c];
                    if (data == null) continue;

                    if (data.IsMerged && data.MergeRowSpan > 0)
                    {
                        try { table.Cell(r + 1, c + 1).Merge(table.Cell(r + data.MergeRowSpan, c + data.MergeColSpan)); } catch { }
                    }

                    if (!data.IsMerged || (data.MergeRowSpan > 0))
                    {
                        try { ApplyStructuralFixes(table.Cell(r + 1, c + 1), data.NumberFormat); } catch { }
                    }
                }
            }

            // 5. RESTORE DIMENSIONS
            newShape.Left = targetLeft;
            newShape.Top = targetTop;
            newShape.Width = targetWidth;

            // If Match All, force the height
            if (scaleHeight && targetHeight > 0)
            {
                newShape.Height = targetHeight;
            }

            // Recalculate column widths
            for (int c = 1; c <= cols; c++)
            {
                if (c <= table.Columns.Count)
                {
                    table.Columns[c].Width = (colWidths[c - 1] / sumColW) * targetWidth;
                }
            }

            return newShape;
        }

        /* ═════════════════════  5. PUBLIC METHODS  ═════════════════════ */

        public static void ImportExcelRangeAsTable(PowerPoint.Application ppApp)
        {
            Excel.Application excelApp = null;
            Excel.Range selectedRange = null;

            try
            {
                try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); }
                catch { MessageBox.Show("Excel is not running."); return; }
                selectedRange = excelApp.Selection as Excel.Range;
                if (selectedRange == null) { MessageBox.Show("Select range."); return; }
                PowerPoint.Slide targetSlide = ppApp.ActiveWindow.View.Slide;
                if (targetSlide == null) return;

                // Handle Anchor
                float tLeft = 0, tTop = 0, tWidth = 0, tHeight = 0;
                bool hasAnchor = false;
                PowerPoint.Shape anchor = null;

                if (ppApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ppApp.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    anchor = ppApp.ActiveWindow.Selection.ShapeRange[1];
                    tLeft = anchor.Left; tTop = anchor.Top; tWidth = anchor.Width; tHeight = anchor.Height;
                    hasAnchor = true;
                }
                else
                {
                    GetRangeDimensions(selectedRange, out float w, out float h, out float[] cw, out float[] rh);
                    tWidth = w > 600 ? 600 : w;
                    tHeight = h;
                    tLeft = (ppApp.ActivePresentation.PageSetup.SlideWidth - tWidth) / 2;
                    tTop = (ppApp.ActivePresentation.PageSetup.SlideHeight - h) / 2;
                }

                // Get Size Mode from Ribbon
                Ribbon1 ribbon = Ribbon1.Instance;
                string sizeMode = ribbon?.GetCurrentImportMatchType() ?? "Match Width";

                // If no anchor, fallback to width only to prevent weird floating behavior
                if (!hasAnchor) sizeMode = "Match Width";

                bool screenUpdate = excelApp.ScreenUpdating;
                excelApp.ScreenUpdating = false;

                try
                {
                    // Pass sizeMode to Paste logic
                    PowerPoint.Shape newShape = PasteAndFixTable(ppApp, targetSlide, selectedRange, tLeft, tTop, tWidth, tHeight, sizeMode);

                    if (newShape != null)
                    {
                        if (anchor != null) anchor.Delete();

                        string uniqueId = Guid.NewGuid().ToString().Replace("-", "_");
                        string name = "Pragmatics_Link_" + uniqueId;
                        Excel.Workbook wb = selectedRange.Worksheet.Parent as Excel.Workbook;
                        wb.Names.Add(Name: name, RefersTo: selectedRange, Visible: false);

                        string normalizedWbPath = CloudPathHelper.NormalizePath(wb.FullName);
                        newShape.Tags.Add("Pragmatics_LinkedRange_ID", uniqueId);
                        newShape.Tags.Add("Pragmatics_SourceWorkbookPath", normalizedWbPath);
                        newShape.Tags.Add("Pragmatics_SourceWorksheet", selectedRange.Worksheet.Name);
                        newShape.Tags.Add("Pragmatics_RangeName", name);
                        newShape.Tags.Add("Pragmatics_ImportType", "Table");
                        newShape.Tags.Add("Pragmatics_SizeMode", sizeMode); // Save the mode!

                        newShape.Select();
                    }
                }
                finally { excelApp.ScreenUpdating = screenUpdate; if (selectedRange != null) Marshal.ReleaseComObject(selectedRange); }
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }

        public static void ImportExcelRangeAsPicture(PowerPoint.Application ppApp)
        {
            Excel.Application excelApp = null;
            try { excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application"); } catch { return; }
            Excel.Range rng = excelApp.Selection as Excel.Range;
            if (rng == null) return;
            PowerPoint.Slide sld = ppApp.ActiveWindow.View.Slide;
            if (sld == null) return;

            // Handle Anchor
            float tLeft = 0, tTop = 0, tWidth = 0, tHeight = 0;
            bool hasAnchor = false;
            PowerPoint.Shape anchor = null;

            if (ppApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && ppApp.ActiveWindow.Selection.ShapeRange.Count > 0)
            {
                anchor = ppApp.ActiveWindow.Selection.ShapeRange[1];
                tLeft = anchor.Left; tTop = anchor.Top; tWidth = anchor.Width; tHeight = anchor.Height;
                hasAnchor = true;
            }
            else
            {
                GetRangeDimensions(rng, out float w, out float h, out float[] cw, out float[] rh);
                tWidth = w; tHeight = h;
                tLeft = (ppApp.ActivePresentation.PageSetup.SlideWidth - tWidth) / 2;
                tTop = (ppApp.ActivePresentation.PageSetup.SlideHeight - tHeight) / 2;
            }

            // Get Size Mode from Ribbon
            Ribbon1 ribbon = Ribbon1.Instance;
            string sizeMode = ribbon?.GetCurrentImportMatchType() ?? "Match Width";
            if (!hasAnchor) sizeMode = "Match Width";

            bool originalUpdating = excelApp.ScreenUpdating;
            excelApp.ScreenUpdating = false;

            Excel.Window activeWindow = excelApp.ActiveWindow;
            Excel.XlWindowView originalView = Excel.XlWindowView.xlNormalView;
            bool viewChanged = false;

            try
            {
                try
                {
                    originalView = activeWindow.View;
                    if (originalView != Excel.XlWindowView.xlNormalView)
                    {
                        activeWindow.View = Excel.XlWindowView.xlNormalView;
                        viewChanged = true;
                    }
                }
                catch { }

                Action copyOp = () => {
                    rng.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);
                    System.Threading.Thread.Sleep(50);
                };

                if (hasAnchor)
                {
                    // Pass sizeMode logic to resize helper
                    bool scaleHeight = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase);
                    CopyRangeWithTempResize(rng, tWidth, tHeight, scaleHeight, copyOp);
                }
                else
                {
                    copyOp();
                }

                var shps = sld.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile);
                var shp = shps[1];

                if (anchor != null) anchor.Delete();

                // Position and Size
                shp.LockAspectRatio = Office.MsoTriState.msoFalse;
                shp.Left = tLeft; shp.Top = tTop; shp.Width = tWidth;

                // If Match All, explicitly set height
                if (hasAnchor && string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase))
                {
                    shp.Height = tHeight;
                }
                else if (hasAnchor && string.Equals(sizeMode, "Match Height", StringComparison.OrdinalIgnoreCase))
                {
                    shp.Height = tHeight;
                    // For Match Height, we let width flow, so we don't force width above
                }

                shp.LockAspectRatio = Office.MsoTriState.msoTrue;

                string uniqueId = Guid.NewGuid().ToString().Replace("-", "_");
                string name = "Pragmatics_Link_" + uniqueId;
                Excel.Workbook wb = rng.Worksheet.Parent as Excel.Workbook;
                wb.Names.Add(Name: name, RefersTo: rng, Visible: false);
                string path = CloudPathHelper.NormalizePath(wb.FullName);
                shp.Tags.Add("Pragmatics_LinkedRange_ID", uniqueId);
                shp.Tags.Add("Pragmatics_SourceWorkbookPath", path);
                shp.Tags.Add("Pragmatics_SourceWorksheet", rng.Worksheet.Name);
                shp.Tags.Add("Pragmatics_RangeName", name);
                shp.Tags.Add("Pragmatics_ImportType", "Picture");
                shp.Tags.Add("Pragmatics_SizeMode", sizeMode); // Save mode

                shp.Select();
            }
            finally
            {
                if (viewChanged)
                {
                    try { activeWindow.View = originalView; } catch { }
                }
                excelApp.ScreenUpdating = originalUpdating;
                if (rng != null) Marshal.ReleaseComObject(rng);
            }
        }
    }
}