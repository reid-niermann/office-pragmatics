using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_Excel
{
    public static class CellRangeFormatting
    {
        private static PowerPoint.Application GetPPTApplication()
        {
            try
            {
                return Marshal.GetActiveObject("PowerPoint.Application") as PowerPoint.Application;
            }
            catch (COMException)
            {
                MessageBox.Show("PowerPoint is not running.", "CellRangeFormatting", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return null;
            }
        }

        /// <summary>
        /// Inserts a rectangle in Excel matching the size and outline of the selected PowerPoint shape.
        /// </summary>
        public static void InsertGuide()
        {
            var pptApp = GetPPTApplication();
            if (pptApp == null) return;
            var sel = pptApp.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select a shape in PowerPoint.", "PPTGuide", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var shape = sel.ShapeRange[1];
            float width = shape.Width;
            float height = shape.Height;
            float lineWeight = shape.Line.Weight;

            var excelApp = Globals.ThisAddIn.Application;
            var sheet = excelApp.ActiveSheet as Excel.Worksheet;
            if (sheet == null) return;
            var cell = excelApp.ActiveCell;
            float left = (float)cell.Left;
            float top = (float)cell.Top;

            var excelShape = sheet.Shapes.AddShape(
                MsoAutoShapeType.msoShapeRectangle,
                left, top, width, height);
            excelShape.Fill.Visible = MsoTriState.msoFalse;
            excelShape.Line.Visible = MsoTriState.msoTrue;
            excelShape.Line.Weight = lineWeight;
            excelShape.Line.ForeColor.RGB = ColorTranslator.ToOle(Color.Red);
        }

        /// <summary>
        /// Scales the selected Excel range so its overall width and height match the selected PowerPoint shape.
        /// </summary>
        public static void ResizeToPowerPoint()
        {
            var pptApp = GetPPTApplication();
            if (pptApp == null) return;
            var sel = pptApp.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select a shape in PowerPoint.", "ResizeToPowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            var shape = sel.ShapeRange[1];
            double targetWidth = shape.Width;
            double targetHeight = shape.Height;

            var excelApp = Globals.ThisAddIn.Application;
            var range = excelApp.Selection as Excel.Range;
            if (range == null)
            {
                MessageBox.Show("Please select a range of cells in Excel.", "ResizeToPowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            double currentWidth = (double)range.Width;
            double currentHeight = (double)range.Height;
            if (currentWidth <= 0 || currentHeight <= 0)
            {
                MessageBox.Show("Could not determine current range size.", "ResizeToPowerPoint", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            double widthScale = targetWidth / currentWidth;
            double heightScale = targetHeight / currentHeight;

            // Capture first column metrics for final width correction
            var firstCol = range.Columns[1] as Excel.Range;
            double initialCharWidth = firstCol.ColumnWidth;
            double pointsPerChar = initialCharWidth > 0
                ? ((double)firstCol.Width / initialCharWidth)
                : 1;

            // Scale column widths (character units)
            foreach (Excel.Range col in range.Columns)
            {
                try
                {
                    col.ColumnWidth = col.ColumnWidth * widthScale;
                }
                catch { /* ignore locked/protected columns */ }
            }
            // Scale row heights (points)
            foreach (Excel.Range row in range.Rows)
            {
                try
                {
                    row.RowHeight = row.RowHeight * heightScale;
                }
                catch { /* ignore locked/protected rows */ }
            }

            // Final width correction: adjust first column to eliminate rounding error
            double newWidthPoints = (double)range.Width;
            double errorPoints = targetWidth - newWidthPoints;
            double charAdjust = pointsPerChar > 0
                ? errorPoints / pointsPerChar
                : 0;
            try
            {
                firstCol.ColumnWidth = firstCol.ColumnWidth + charAdjust;
            }
            catch { /* final adjust ignore errors */ }
        }
        public static void mergeToCenterAcrossSelection()
        {
            var excelApp = Globals.ThisAddIn.Application;
            var sel = excelApp.Selection as Excel.Range;
            if (sel == null) return;

            foreach (Excel.Range cell in sel.Cells)
            {
                if (!cell.MergeCells) continue;

                var merged = cell.MergeArea;
                // only operate on single-row merges
                if (merged.Rows.Count == 1)
                {
                    // grab original text from first cell
                    var value = merged.Cells[1, 1].Value2;

                    // unmerge
                    merged.UnMerge();

                    // define the full span for centering
                    var first = merged.Cells[1, 1];
                    var last = merged.Cells[1, merged.Columns.Count];
                    var span = excelApp.Range[first, last];

                    // clear everything, then put text back in first cell only
                    span.ClearContents();
                    first.Value2 = value;

                    // apply center-across-selection on that span
                    span.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenterAcrossSelection;
                    span.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
                }
            }
        }

        /// <summary>
        /// Converts any hidden rows and columns on the current sheet to grouped columns and rows instead.
        /// </summary>
        public static void hideToGroup()
        {
            Excel.Application excelApp = null;
            Excel.Worksheet sheet = null;
            Excel.Range usedRange = null;

            try
            {
                excelApp = Globals.ThisAddIn.Application;
                if (excelApp == null) return;

                sheet = excelApp.ActiveSheet as Excel.Worksheet;
                if (sheet == null)
                {
                    MessageBox.Show("No active worksheet found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                excelApp.ScreenUpdating = false;

                // Clear existing outlines to prevent conflicts or unexpected behavior
                // sheet.Rows.ClearOutline(); // Optional: if you want to start fresh
                // sheet.Columns.ClearOutline(); // Optional: if you want to start fresh

                usedRange = sheet.UsedRange;
                if (usedRange == null || usedRange.Rows.Count == 0 || usedRange.Columns.Count == 0)
                {
                    // Sheet is effectively empty or UsedRange couldn't be determined
                    return;
                }

                // --- Process Rows ---
                int firstUsedRow = usedRange.Row;
                int lastUsedRow = firstUsedRow + usedRange.Rows.Count - 1;
                int hiddenBlockStartRow = 0; // 0 means not in a hidden block

                for (int r = firstUsedRow; r <= lastUsedRow; r++)
                {
                    Excel.Range currentRowRange = null;
                    try
                    {
                        currentRowRange = sheet.Rows[r] as Excel.Range;
                        bool isRowHidden = false;
                        // The Hidden property can return True, False, or DBNull (if mixed).
                        // For an entire row, it should be True or False if user-hidden.
                        object hiddenStatus = currentRowRange.Hidden;
                        if (hiddenStatus is bool && (bool)hiddenStatus)
                        {
                            isRowHidden = true;
                        }

                        if (isRowHidden)
                        {
                            if (hiddenBlockStartRow == 0) // Start of a new hidden block
                            {
                                hiddenBlockStartRow = r;
                            }
                        }
                        else // Current row is visible
                        {
                            if (hiddenBlockStartRow != 0) // A block of hidden rows just ended
                            {
                                Excel.Range groupRows = null;
                                Excel.Range startRowObj = null;
                                Excel.Range endRowObj = null;
                                try
                                {
                                    startRowObj = sheet.Rows[hiddenBlockStartRow];
                                    endRowObj = sheet.Rows[r - 1]; // End of block is previous row
                                    groupRows = sheet.Range[startRowObj, endRowObj];

                                    groupRows.Hidden = false; // Unhide first
                                    groupRows.Group();
                                }
                                finally
                                {
                                    if (groupRows != null) Marshal.ReleaseComObject(groupRows);
                                    if (startRowObj != null) Marshal.ReleaseComObject(startRowObj);
                                    if (endRowObj != null) Marshal.ReleaseComObject(endRowObj);
                                }
                                hiddenBlockStartRow = 0; // Reset
                            }
                        }
                    }
                    finally
                    {
                        if (currentRowRange != null) Marshal.ReleaseComObject(currentRowRange);
                    }
                }
                // After loop, check if a hidden block extends to the end of UsedRange
                if (hiddenBlockStartRow != 0)
                {
                    Excel.Range groupRows = null;
                    Excel.Range startRowObj = null;
                    Excel.Range endRowObj = null;
                    try
                    {
                        startRowObj = sheet.Rows[hiddenBlockStartRow];
                        endRowObj = sheet.Rows[lastUsedRow];
                        groupRows = sheet.Range[startRowObj, endRowObj];

                        groupRows.Hidden = false;
                        groupRows.Group();
                    }
                    finally
                    {
                        if (groupRows != null) Marshal.ReleaseComObject(groupRows);
                        if (startRowObj != null) Marshal.ReleaseComObject(startRowObj);
                        if (endRowObj != null) Marshal.ReleaseComObject(endRowObj);
                    }
                }

                // --- Process Columns ---
                int firstUsedCol = usedRange.Column;
                int lastUsedCol = firstUsedCol + usedRange.Columns.Count - 1;
                int hiddenBlockStartCol = 0; // 0 means not in a hidden block

                for (int c = firstUsedCol; c <= lastUsedCol; c++)
                {
                    Excel.Range currentColRange = null;
                    try
                    {
                        currentColRange = sheet.Columns[c] as Excel.Range;
                        bool isColHidden = false;
                        object hiddenStatus = currentColRange.Hidden;
                        if (hiddenStatus is bool && (bool)hiddenStatus)
                        {
                            isColHidden = true;
                        }

                        if (isColHidden)
                        {
                            if (hiddenBlockStartCol == 0) // Start of a new hidden block
                            {
                                hiddenBlockStartCol = c;
                            }
                        }
                        else // Current column is visible
                        {
                            if (hiddenBlockStartCol != 0) // A block of hidden columns just ended
                            {
                                Excel.Range groupCols = null;
                                Excel.Range startColObj = null;
                                Excel.Range endColObj = null;
                                try
                                {
                                    startColObj = sheet.Columns[hiddenBlockStartCol];
                                    endColObj = sheet.Columns[c - 1]; // End of block is previous column
                                    groupCols = sheet.Range[startColObj, endColObj];

                                    groupCols.Hidden = false; // Unhide first
                                    groupCols.Group();
                                }
                                finally
                                {
                                    if (groupCols != null) Marshal.ReleaseComObject(groupCols);
                                    if (startColObj != null) Marshal.ReleaseComObject(startColObj);
                                    if (endColObj != null) Marshal.ReleaseComObject(endColObj);
                                }
                                hiddenBlockStartCol = 0; // Reset
                            }
                        }
                    }
                    finally
                    {
                        if (currentColRange != null) Marshal.ReleaseComObject(currentColRange);
                    }
                }
                // After loop, check if a hidden block extends to the end of UsedRange
                if (hiddenBlockStartCol != 0)
                {
                    Excel.Range groupCols = null;
                    Excel.Range startColObj = null;
                    Excel.Range endColObj = null;
                    try
                    {
                        startColObj = sheet.Columns[hiddenBlockStartCol];
                        endColObj = sheet.Columns[lastUsedCol];
                        groupCols = sheet.Range[startColObj, endColObj];

                        groupCols.Hidden = false;
                        groupCols.Group();
                    }
                    finally
                    {
                        if (groupCols != null) Marshal.ReleaseComObject(groupCols);
                        if (startColObj != null) Marshal.ReleaseComObject(startColObj);
                        if (endColObj != null) Marshal.ReleaseComObject(endColObj);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in hideToGroup: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (excelApp != null) excelApp.ScreenUpdating = true;
                if (usedRange != null) Marshal.ReleaseComObject(usedRange);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                // excelApp is Globals.ThisAddIn.Application, VSTO manages its lifetime.
            }
        }
        public static void deleteBlankColumns()
        {

        }
        public static void deleteBlankRows()
        {

        }

        // P/Invoke to check for Esc key
        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);
        private const int VK_ESCAPE = 0x1B;

        // Progress indicator fields
        private static Form _progressForm;
        private static int _progressTotal;
        private static int _progressCurrent;

        /// <summary>
        /// Reverses the order of entire columns within the currently selected range
        /// with progress indicator and Esc-to-abort support.
        /// </summary>
        public static void ReverseColumns()
        {
            var app = Globals.ThisAddIn.Application;
            bool aborted = false;
            bool originalScreen = app.ScreenUpdating;

            var sel = app.Selection as Excel.Range;
            if (sel == null)
            {
                MessageBox.Show("Please make a range selection first.", "ReverseColumns",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var ws = (Excel.Worksheet)sel.Worksheet;
            int startCol = sel.Column;
            int colCount = sel.Columns.Count;
            if (colCount < 2) return;

            // Number of swap operations
            int swapCount = colCount / 2;
            ShowProgressIndicator(swapCount);

            // buffer at far right
            int bufferCol = ws.Columns.Count;

            try
            {
                app.ScreenUpdating = false;
                ((Excel.Range)ws.Columns[bufferCol])
                    .Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                for (int i = 0; i < swapCount; i++)
                {
                    // update progress and handle UI
                    _progressCurrent = i + 1;
                    UpdateProgressIndicator();

                    Application.DoEvents();
                    if ((GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0)
                    {
                        aborted = true;
                        break;
                    }

                    int leftIndex = startCol + i;
                    int rightIndex = startCol + colCount - 1 - i;

                    ws.Columns[leftIndex].Cut(ws.Columns[bufferCol]);
                    ws.Columns[rightIndex].Cut(ws.Columns[leftIndex]);
                    ws.Columns[bufferCol].Cut(ws.Columns[rightIndex]);
                }
            }
            finally
            {
                try { ws.Columns[bufferCol].Delete(); } catch { }
                app.ScreenUpdating = originalScreen;
                HideProgressIndicator();
            }

            if (aborted)
                MessageBox.Show("Operation aborted by user.", "ReverseColumns",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>
        /// Reverses the order of entire rows within the currently selected range
        /// with progress indicator and Esc-to-abort support.
        /// </summary>
        public static void ReverseRows()
        {
            var app = Globals.ThisAddIn.Application;
            bool aborted = false;
            bool originalScreen = app.ScreenUpdating;

            var sel = app.Selection as Excel.Range;
            if (sel == null)
            {
                MessageBox.Show("Please make a range selection first.", "ReverseRows",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var ws = (Excel.Worksheet)sel.Worksheet;
            int startRow = sel.Row;
            int rowCount = sel.Rows.Count;
            if (rowCount < 2) return;

            // Number of swap operations
            int swapCount = rowCount / 2;
            ShowProgressIndicator(swapCount);

            // buffer at bottom
            int bufferRow = ws.Rows.Count;

            try
            {
                app.ScreenUpdating = false;
                ((Excel.Range)ws.Rows[bufferRow])
                    .Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                for (int i = 0; i < swapCount; i++)
                {
                    // update progress and handle UI
                    _progressCurrent = i + 1;
                    UpdateProgressIndicator();

                    Application.DoEvents();
                    if ((GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0)
                    {
                        aborted = true;
                        break;
                    }

                    int topIndex = startRow + i;
                    int bottomIndex = startRow + rowCount - 1 - i;

                    ws.Rows[topIndex].Cut(ws.Rows[bufferRow]);
                    ws.Rows[bottomIndex].Cut(ws.Rows[topIndex]);
                    ws.Rows[bufferRow].Cut(ws.Rows[bottomIndex]);
                }
            }
            finally
            {
                try { ws.Rows[bufferRow].Delete(); } catch { }
                app.ScreenUpdating = originalScreen;
                HideProgressIndicator();
            }

            if (aborted)
                MessageBox.Show("Operation aborted by user.", "ReverseRows",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        // ----- Progress Indicator Methods -----
        private static void ShowProgressIndicator(int total)
        {
            _progressTotal = total;
            _progressCurrent = 0;

            _progressForm = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.Manual,
                Size = new System.Drawing.Size(350, 100),
                BackColor = Color.Black,
                Opacity = 0.7,
                TopMost = true,
                ShowInTaskbar = false
            };

            _progressForm.Paint += (sender, e) =>
            {
                string text = $"Processing: {_progressCurrent} / {_progressTotal}  (Hold Esc key to abort)";
                using (var brush = new SolidBrush(Color.White))
                using (var font = new Font("Segoe UI", 10, FontStyle.Bold))
                {
                    e.Graphics.DrawString(text, font, brush, new PointF(10, 35));
                }
            };

            var bounds = Screen.PrimaryScreen.Bounds;
            _progressForm.Location = new Point(
                bounds.Width - _progressForm.Width - 20,
                bounds.Height - _progressForm.Height - 20);

            _progressForm.Show();
            _progressForm.Refresh();
        }

        private static void UpdateProgressIndicator()
        {
            if (_progressForm != null && !_progressForm.IsDisposed)
            {
                _progressForm.Invalidate();
                Application.DoEvents();
            }
        }

        private static void HideProgressIndicator()
        {
            if (_progressForm != null)
            {
                _progressForm.Close();
                _progressForm.Dispose();
                _progressForm = null;
            }
        }
        public static void InsertCagrFormula()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null)
            {
                MessageBox.Show("Please select a range first.", "CAGR", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!TryBuildCagrFormulaFromSelection(sel, out string formula, out string summary))
            {
                MessageBox.Show("Selection must be:\n" +
                                "• 1 row × N (N≥3) or N × 1 (N≥3), OR\n" +
                                "• 2 rows × N (N≥3) or N × 2 (N≥3) with first row/col as dates/years, OR\n" +
                                "• Two disconnected single rows with the same length (N≥3), first is dates/years, second is values.",
                                "CAGR – Invalid Selection",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // NEW: default output = last value cell in the detected values range
            string defaultOut = GetSuggestedOutputTarget(sel);

            object outObj = null;
            try
            {
                // Provide the default address and require a Range (Type:=8)
                outObj = app.InputBox(
                    "Select an output cell for the CAGR result:",
                    "CAGR",
                    Default: defaultOut,
                    Type: 8 /*Range*/);
            }
            catch { /* cancel handled below */ }

            if (outObj is bool canceled && canceled == false) return;

            var outRange = outObj as Excel.Range;
            if (outRange == null || outRange.Cells.Count != 1)
            {
                MessageBox.Show("Please pick a single output cell.", "CAGR", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            outRange.Formula = formula;
            try { outRange.NumberFormat = "0.00%"; } catch { /* ignore */ }

            if (!string.IsNullOrEmpty(summary))
            {
                app.StatusBar = $"CAGR inserted → {summary}";
            }
        }

        // Helper to pick the “end of the values row/column”
        private static string GetSuggestedOutputTarget(Excel.Range sel)
        {
            try
            {
                var ws = (Excel.Worksheet)sel.Worksheet;
                int maxRow = ws.Rows.Count;
                int maxCol = ws.Columns.Count;

                string Addr(int r, int c) => (ws.Cells[r, c] as Excel.Range).Address[true, true];

                // Case: two disconnected single-row areas (Area1=labels, Area2=values)
                if (sel.Areas.Count == 2)
                {
                    var a1 = sel.Areas[1];
                    var a2 = sel.Areas[2];
                    if (a1.Rows.Count == 1 && a2.Rows.Count == 1 && a1.Columns.Count == a2.Columns.Count && a2.Columns.Count >= 3)
                    {
                        int row = a2.Row;
                        int lastCol = a2.Column + a2.Columns.Count - 1;
                        int rightCol = lastCol + 1;

                        if (rightCol <= maxCol) return Addr(row, rightCol);
                        // Edge fallback: try down one row, otherwise stay on last value cell
                        int downRow = row + 1;
                        if (downRow <= maxRow) return Addr(downRow, lastCol);
                        return Addr(row, lastCol);
                    }
                }

                // Contiguous selections
                int rCount = sel.Rows.Count;
                int cCount = sel.Columns.Count;
                int startRow = sel.Row;
                int startCol = sel.Column;

                // 1D: single row of values → suggest one column to the right
                if (rCount == 1 && cCount >= 3)
                {
                    int row = startRow;
                    int lastCol = startCol + cCount - 1;
                    int rightCol = lastCol + 1;

                    if (rightCol <= maxCol) return Addr(row, rightCol);
                    // Edge fallback: try down one row, else stay on last value
                    int downRow = row + 1;
                    if (downRow <= maxRow) return Addr(downRow, lastCol);
                    return Addr(row, lastCol);
                }

                // 1D: single column of values → suggest one row down
                if (cCount == 1 && rCount >= 3)
                {
                    int col = startCol;
                    int lastRow = startRow + rCount - 1;
                    int downRow = lastRow + 1;

                    if (downRow <= maxRow) return Addr(downRow, col);
                    // Edge fallback: try one column to the right, else stay on last value
                    int rightCol = col + 1;
                    if (rightCol <= maxCol) return Addr(lastRow, rightCol);
                    return Addr(lastRow, col);
                }

                // 2×N (first row = labels, second row = values) → one to the right from last value
                if (rCount == 2 && cCount >= 3)
                {
                    int valuesRow = startRow + 1;
                    int lastCol = startCol + cCount - 1;
                    int rightCol = lastCol + 1;

                    if (rightCol <= maxCol) return Addr(valuesRow, rightCol);
                    // Edge fallback
                    int downRow = valuesRow + 1;
                    if (downRow <= maxRow) return Addr(downRow, lastCol);
                    return Addr(valuesRow, lastCol);
                }

                // N×2 (first col = labels, second col = values) → one down from last value
                if (cCount == 2 && rCount >= 3)
                {
                    int valuesCol = startCol + 1;
                    int lastRow = startRow + rCount - 1;
                    int downRow = lastRow + 1;

                    if (downRow <= maxRow) return Addr(downRow, valuesCol);
                    // Edge fallback
                    int rightCol = valuesCol + 1;
                    if (rightCol <= maxCol) return Addr(lastRow, rightCol);
                    return Addr(lastRow, valuesCol);
                }
            }
            catch { /* ignore and fall back below */ }

            // Fallback: active cell (or A1)
            var app = Globals.ThisAddIn.Application;
            return (app.ActiveCell as Excel.Range)?.Address[true, true] ?? "$A$1";
        }

        private static bool TryBuildCagrFormulaFromSelection(Excel.Range sel, out string formula, out string summary)
        {
            formula = null;
            summary = null;

            // Handle two disconnected rows (two areas)
            if (sel.Areas.Count == 2)
            {
                var a1 = sel.Areas[1];
                var a2 = sel.Areas[2];

                if (a1.Rows.Count == 1 && a2.Rows.Count == 1 && a1.Columns.Count == a2.Columns.Count && a1.Columns.Count >= 3)
                {
                    // Area1: labels; Area2: values
                    string startLabel = a1.Cells[1, 1].Address[true, true];
                    string endLabel = a1.Cells[1, a1.Columns.Count].Address[true, true];
                    string startVal = a2.Cells[1, 1].Address[true, true];
                    string endVal = a2.Cells[1, a2.Columns.Count].Address[true, true];

                    bool labelsAreDates = LooksLikeDates(a1);
                    string periods = labelsAreDates
                        ? $"YEARFRAC({startLabel},{endLabel})"
                        : $"({endLabel}-{startLabel})";

                    formula = $"=IFERROR(({endVal}/{startVal})^(1/{periods})-1,\"\")";
                    summary = $"labels: {a1.Address[false, false]}, values: {a2.Address[false, false]}";
                    return true;
                }
                return false;
            }

            // Contiguous region
            int r = sel.Rows.Count;
            int c = sel.Columns.Count;

            // 1D values only (assume consecutive annual data)  ⟶  use COUNTA(range)-1 for period count
            if ((r == 1 && c >= 3) || (c == 1 && r >= 3))
            {
                // First and last value cells
                var firstCell = sel.Cells[1, 1] as Excel.Range;
                var lastCell = (r == 1) ? sel.Cells[1, c] as Excel.Range : sel.Cells[r, 1] as Excel.Range;

                string firstAddr = firstCell.Address[true, true];
                string lastAddr = lastCell.Address[true, true];

                // Periods based on nonblank count in the selection
                string countRange = sel.Address[true, true]; // absolute; change to [false,false] if you prefer relative
                string periods = $"(COUNTA({countRange})-1)";

                formula = $"=IFERROR(({lastAddr}/{firstAddr})^(1/{periods})-1,\"\")";
                summary = $"values: {sel.Address[false, false]} (periods via COUNTA-1)";
                return true;
            }

            // 2xN: first row labels, second row values
            if (r == 2 && c >= 3)
            {
                var labels = sel.Rows[1];
                var vals = sel.Rows[2];

                string startLabel = labels.Cells[1, 1].Address[true, true];
                string endLabel = labels.Cells[1, c].Address[true, true];
                string startVal = vals.Cells[1, 1].Address[true, true];
                string endVal = vals.Cells[1, c].Address[true, true];

                bool labelsAreDates = LooksLikeDates(labels);
                string periods = labelsAreDates
                    ? $"YEARFRAC({startLabel},{endLabel})"
                    : $"({endLabel}-{startLabel})";

                formula = $"=IFERROR(({endVal}/{startVal})^(1/{periods})-1,\"\")";
                summary = $"labels: {labels.Address[false, false]}, values: {vals.Address[false, false]}";
                return true;
            }

            // Nx2: first column labels, second column values
            if (c == 2 && r >= 3)
            {
                var labels = sel.Columns[1];
                var vals = sel.Columns[2];

                string startLabel = labels.Cells[1, 1].Address[true, true];
                string endLabel = labels.Cells[r, 1].Address[true, true];
                string startVal = vals.Cells[1, 1].Address[true, true];
                string endVal = vals.Cells[r, 1].Address[true, true];

                bool labelsAreDates = LooksLikeDates(labels);
                string periods = labelsAreDates
                    ? $"YEARFRAC({startLabel},{endLabel})"
                    : $"({endLabel}-{startLabel})";

                formula = $"=IFERROR(({endVal}/{startVal})^(1/{periods})-1,\"\")";
                summary = $"labels: {labels.Address[false, false]}, values: {vals.Address[false, false]}";
                return true;
            }

            return false;
        }

        // Heuristic: treat as dates if the number format contains year/day tokens or
        // if most cells are date-serial-looking with a date-like number format.
        private static bool LooksLikeDates(Excel.Range labelRange)
        {
            try
            {
                int total = 0, dateLike = 0;
                foreach (Excel.Range cell in labelRange.Cells)
                {
                    total++;
                    object v = cell.Value2;

                    // Quick check: number format hints (common tokens for dates)
                    string nf = "";
                    try { nf = cell.NumberFormat as string ?? ""; } catch { }

                    bool fmtDatey = (nf.IndexOf("y", StringComparison.OrdinalIgnoreCase) >= 0) ||
                                    (nf.IndexOf("d", StringComparison.OrdinalIgnoreCase) >= 0) ||
                                    (nf.IndexOf("m", StringComparison.OrdinalIgnoreCase) >= 0 && nf.IndexOf("yy", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (fmtDatey)
                    {
                        dateLike++;
                        continue;
                    }

                    // If it's a double within Excel date serial range and not a simple 4-digit year
                    if (v is double dv)
                    {
                        // 4-digit year heuristic
                        if (dv >= 1900 && dv < 3000 && Math.Abs(dv - Math.Round(dv)) < 0.0000001)
                        {
                            // looks like a year, not a date-serial (keep as non-date)
                        }
                        else if (dv > 59) // Excel date serials > 59 are safely after 1-Mar-1900
                        {
                            dateLike++;
                        }
                    }
                }
                return total > 0 && dateLike >= Math.Max(1, total / 2);
            }
            catch { return false; }
        }

        // Moved out of the method so C# 7.3 is happy
        private static void Release(object com)
        {
            if (com != null && Marshal.IsComObject(com))
            {
                try { Marshal.FinalReleaseComObject(com); } catch { /* ignore */ }
            }
        }

        public static void HardcodeFormulasInSelectionOrUsedRange()
        {
            var app = Globals.ThisAddIn.Application;
            var ws = app.ActiveSheet as Excel.Worksheet;
            if (ws == null) return;

            Excel.Range target = null;
            Excel.Range formulas = null;

            bool prevScreen = app.ScreenUpdating;
            bool prevEvents = app.EnableEvents;
            bool prevAlerts = app.DisplayAlerts;
            var prevCalc = app.Calculation;
            bool prevPageBreaks = false;

            try
            {
                var sel = app.Selection as Excel.Range;
                target = (sel == null || (sel.Cells != null && sel.Cells.Count == 1))
                         ? ws.UsedRange
                         : sel;

                if (target == null || target.Cells.Count == 0) return;

                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;
                try { prevPageBreaks = ws.DisplayPageBreaks; ws.DisplayPageBreaks = false; } catch { }
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                Excel.Range inUse = null;
                try { inUse = ws.UsedRange; } catch { }
                Excel.Range scope = target;
                if (inUse != null)
                {
                    var inter = app.Intersect(target, inUse);
                    if (inter != null) { Release(scope); scope = inter; }
                }

                try
                {
                    formulas = scope.SpecialCells(Excel.XlCellType.xlCellTypeFormulas, 23);
                }
                catch { return; }
                if (formulas == null) return;

                var seenArrays = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (Excel.Range area in formulas.Areas)
                {
                    foreach (Excel.Range cell in area.Cells)
                    {
                        try
                        {
                            bool isArray = false;
                            try { isArray = cell.HasArray; } catch { }
                            if (isArray)
                            {
                                Excel.Range ca = null;
                                try
                                {
                                    ca = cell.CurrentArray;
                                    string key = ca.Address[true, true, Excel.XlReferenceStyle.xlA1, true, true];
                                    if (seenArrays.Add(key))
                                    {
                                        var v = ca.Value2;
                                        ca.Value2 = v;
                                    }
                                }
                                finally { Release(ca); }
                            }
                        }
                        finally { Release(cell); }
                    }
                    Release(area);
                }

                foreach (Excel.Range area in formulas.Areas)
                {
                    try
                    {
                        var v = area.Value2;
                        area.Value2 = v;
                    }
                    finally { Release(area); }
                }

                app.StatusBar = "Hardcode complete: formulas converted to values.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hardcode failed: " + ex.Message, "Hardcode Formulas",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try { app.ScreenUpdating = prevScreen; } catch { }
                try { app.EnableEvents = prevEvents; } catch { }
                try { app.DisplayAlerts = prevAlerts; } catch { }
                try { app.Calculation = prevCalc; } catch { }
                try { var ws2 = app.ActiveSheet as Excel.Worksheet; if (ws2 != null) ws2.DisplayPageBreaks = prevPageBreaks; } catch { }
                Release(formulas);
                Release(target);
                app.StatusBar = false;
            }
        }
    }
}