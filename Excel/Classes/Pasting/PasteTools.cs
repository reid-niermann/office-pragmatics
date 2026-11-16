using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    internal static partial class PasteTools
    {
        // ───────────────────────── stored metadata ─────────────────────────
        private static string[,] _storedFormulas;          // A1 formula text (loaded on demand)
        private static string[,] _storedFormulasR1C1;      // R1C1 formula text (loaded on demand)
        private static bool[,] _storedHasExternalRef;     // mask for external refs (built on demand)
        private static int _storedRows;
        private static int _storedCols;
        private static int _storedStartRow;
        private static int _storedStartCol;
        private static string _storedSheetName;
        private static string _storedWorkbookName;
        // ───────────────────────────────────────────────────────────────────

        private static readonly Regex A1RefRegex =
            new Regex(@"\$?([A-Za-z]+)\$?(\d+)", RegexOptions.Compiled);

        /// <summary>
        /// Always lazy: on Ctrl+C, execute Excel Copy and store only metadata.
        /// Heavy snapshot (formulas/external mask) is deferred until a special paste is used.
        /// </summary>
        public static void RunCopyAndStore()
        {
            var app = Globals.ThisAddIn.Application;

            // 1) Always run Excel's generic Copy (works on cells, charts, shapes, etc.)
            app.CommandBars.ExecuteMso("Copy");

            // 2) If the selection is a Range, store just the metadata
            if (app.Selection is Excel.Range sel)
            {
                // If the selection is exactly an entire row or entire column, we're done
                if (sel.Address == sel.EntireRow.Address
                 || sel.Address == sel.EntireColumn.Address)
                    return;

                _storedRows = sel.Rows.Count;
                _storedCols = sel.Columns.Count;
                _storedStartRow = sel.Row;
                _storedStartCol = sel.Column;
                _storedSheetName = ((Excel.Worksheet)sel.Worksheet).Name;
                _storedWorkbookName = ((Excel.Workbook)sel.Worksheet.Parent).Name;

                // Clear any previously-loaded heavy snapshots; they'll be reloaded on demand.
                _storedFormulas = null;
                _storedFormulasR1C1 = null;
                _storedHasExternalRef = null;
            }
        }

        // Load formulas + external mask from the source range (single COM calls per property).
        private static void BulkLoadSnapshotFromRange(Excel.Range rng)
        {
            // Sync dimensions in case the source changed shape
            _storedRows = rng.Rows.Count;
            _storedCols = rng.Columns.Count;

            _storedFormulas = ToString2D(GetFormula2(rng), _storedRows, _storedCols);
            _storedFormulasR1C1 = ToString2D(rng.FormulaR1C1, _storedRows, _storedCols);
            _storedHasExternalRef = BuildExternalMask(
                                        _storedFormulas,
                                        _storedStartRow, _storedStartCol,
                                        _storedRows, _storedCols);
        }

        // Ensure we have per-cell arrays; always load lazily from the source using stored metadata.
        private static bool EnsureSnapshotLoaded(Excel.Application app)
        {
            if (_storedRows <= 0 || _storedCols <= 0 || string.IsNullOrEmpty(_storedSheetName))
                return false;

            if (_storedFormulas != null && _storedFormulasR1C1 != null && _storedHasExternalRef != null)
                return true;

            var src = TryGetStoredSourceRange(app);
            if (src == null) return false;

            BulkLoadSnapshotFromRange(src);
            return _storedFormulas != null;
        }

        public static void PasteExact()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null)
            {
                MessageBox.Show("Please select a destination cell first.",
                                "PasteExact",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }

            if (!EnsureSnapshotLoaded(app))
            {
                MessageBox.Show("Nothing stored or source range unavailable. Copy a range first.",
                                "PasteExact", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Determine paste target block (exact shape)
            var topLeft = (Excel.Range)sel.Cells[1, 1];
            var pasteRange = (Excel.Range)topLeft.get_Resize(_storedRows, _storedCols);

            // Fast overwrite detection
            if (RangeHasAnyData(pasteRange))
            {
                var resp = MessageBox.Show("Do you want to overwrite any existing data?",
                                           "PasteExact", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (resp != DialogResult.Yes) return;
            }

            var prevSU = app.ScreenUpdating;
            var prevCal = app.Calculation;
            var prevEvt = app.EnableEvents;
            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Bulk-assign A1 formulas in one call using Formula2 to preserve dynamic arrays
                object[,] arr = new object[_storedRows, _storedCols];
                for (int r = 0; r < _storedRows; r++)
                    for (int c = 0; c < _storedCols; c++)
                        arr[r, c] = _storedFormulas[r, c] ?? string.Empty;

                SetFormula2(pasteRange, arr);

                // Preserve formats using the original source range (robust to clipboard state)
                var srcRangeForFormats = TryGetStoredSourceRange(app);
                CopyFormatsFromSourceToDestination(app, srcRangeForFormats, pasteRange);
            }
            finally
            {
                app.ScreenUpdating = prevSU;
                app.Calculation = prevCal;
                app.EnableEvents = prevEvt;
            }
        }

        public static void PasteDuplicate()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null)
            {
                MessageBox.Show("Please select a destination cell first.",
                                "Smart Exact",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Warning);
                return;
            }
            if (!EnsureSnapshotLoaded(app))
            {
                MessageBox.Show("Nothing stored or source range unavailable. Copy a range first.",
                                "Smart Exact", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Build the paste target block (exact shape)
            var topLeft = (Excel.Range)sel.Cells[1, 1];
            var pasteRange = topLeft.get_Resize(_storedRows, _storedCols);

            // NEW: Large-range confirmation (threshold: > 25,000 cells)
            int targetCells = _storedRows * _storedCols;
            if (targetCells > 25000)
            {
                var confirm = MessageBox.Show(
                    "Large ranges may take a while to execute.\r\nAre you sure you want to proceed?",
                    "Smart Exact – Large Range",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning,
                    MessageBoxDefaultButton.Button2
                );
                if (confirm != DialogResult.Yes) return;
            }

            if (RangeHasAnyData(pasteRange))
            {
                var resp = MessageBox.Show("Overwrite existing data?",
                                           "Smart Exact", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (resp != DialogResult.Yes) return;
            }

            var prevSU = app.ScreenUpdating;
            var prevCal = app.Calculation;
            var prevEvt = app.EnableEvents;
            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // 1) Shift internal refs fast: bulk-assign R1C1
                object[,] r1c1Arr = new object[_storedRows, _storedCols];
                for (int r = 0; r < _storedRows; r++)
                    for (int c = 0; c < _storedCols; c++)
                        r1c1Arr[r, c] = _storedFormulasR1C1[r, c] ?? string.Empty;
                pasteRange.FormulaR1C1 = r1c1Arr;

                // 2) Restore exact A1 text only where the original had external refs
                //    Use Formula2 to preserve dynamic array behavior
                for (int r = 1; r <= _storedRows; r++)
                {
                    for (int c = 1; c <= _storedCols; c++)
                    {
                        if (_storedHasExternalRef[r - 1, c - 1])
                        {
                            var dest = (Excel.Range)pasteRange.Cells[r, c];
                            SetFormula2(dest, _storedFormulas[r - 1, c - 1] ?? string.Empty);
                        }
                    }
                }

                // Preserve formats using the original source range (robust to clipboard state)
                var srcRangeForFormats = TryGetStoredSourceRange(app);
                CopyFormatsFromSourceToDestination(app, srcRangeForFormats, pasteRange);
            }
            finally
            {
                app.ScreenUpdating = prevSU;
                app.Calculation = prevCal;
                app.EnableEvents = prevEvt;
            }
        }

        // helper: column letters → 1-based index
        private static int LetterToNumber(string letters)
        {
            int sum = 0;
            foreach (char c in letters.ToUpper())
            {
                sum = sum * 26 + (c - 'A' + 1);
            }
            return sum;
        }

        // Convert Excel COM return (single value or 2D 1-based array) to string[,]
        private static string[,] ToString2D(object raw, int rows, int cols)
        {
            var dst = new string[rows, cols];

            if (raw is object[,] arr)
            {
                for (int r = 1; r <= rows; r++)
                    for (int c = 1; c <= cols; c++)
                        dst[r - 1, c - 1] = arr[r, c]?.ToString() ?? string.Empty;
            }
            else
            {
                string s = raw?.ToString() ?? string.Empty;
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        dst[r, c] = s;
            }
            return dst;
        }

        private static bool[,] BuildExternalMask(string[,] a1, int startRow, int startCol, int rows, int cols)
        {
            var mask = new bool[rows, cols];
            int rMin = startRow, rMax = startRow + rows - 1;
            int cMin = startCol, cMax = startCol + cols - 1;

            for (int r = 0; r < rows; r++)
            {
                for (int c = 0; c < cols; c++)
                {
                    string f = a1[r, c];
                    if (string.IsNullOrEmpty(f)) { mask[r, c] = false; continue; }

                    bool ext = false;

                    // Sheet-qualified → treat as external for preservation
                    if (f.IndexOf('!') >= 0)
                    {
                        ext = true;
                    }
                    else
                    {
                        foreach (Match m in A1RefRegex.Matches(f))
                        {
                            int col = LetterToNumber(m.Groups[1].Value);
                            int row = int.Parse(m.Groups[2].Value);
                            if (row < rMin || row > rMax || col < cMin || col > cMax) { ext = true; break; }
                        }
                    }
                    mask[r, c] = ext;
                }
            }
            return mask;
        }

        // ───────────────────────── tiling format-only actions ─────────────────────────
        public static void PasteNumberFormat()
        {
            var app = Globals.ThisAddIn.Application;
            if (!EnsureReadyForPaste("Paste Number Format")) return;

            var destSel = app.Selection as Excel.Range;
            var srcRange = TryGetStoredSourceRange(app);
            if (destSel == null || srcRange == null) { ShowSelectionError(); return; }

            var prevScreenUpdating = app.ScreenUpdating;
            var prevCalculation = app.Calculation;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Cache source number formats, detect uniform fast-path
                var nfs = new string[_storedRows, _storedCols];
                string first = null;
                bool uniform = true;

                for (int r = 1; r <= _storedRows; r++)
                {
                    for (int c = 1; c <= _storedCols; c++)
                    {
                        var s = (Excel.Range)srcRange.Cells[r, c];
                        var nfObj = s.NumberFormat;
                        var nf = nfObj == null ? "" : nfObj.ToString();
                        nfs[r - 1, c - 1] = nf;
                        if (first == null) first = nf;
                        else if (!string.Equals(first, nf, StringComparison.Ordinal)) uniform = false;
                    }
                }

                foreach (Excel.Range area in destSel.Areas)
                {
                    int dr = area.Rows.Count, dc = area.Columns.Count;

                    if (uniform)
                    {
                        area.NumberFormat = first; // single COM call
                        continue;
                    }

                    for (int r = 1; r <= dr; r++)
                    {
                        int sr = ((r - 1) % _storedRows) + 1;
                        for (int c = 1; c <= dc; c++)
                        {
                            int sc = ((c - 1) % _storedCols) + 1;
                            var d = (Excel.Range)area.Cells[r, c];
                            d.NumberFormat = nfs[sr - 1, sc - 1];
                        }
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
                app.EnableEvents = prevEnableEvents;
            }
        }

        public static void PasteFillColor()
        {
            var app = Globals.ThisAddIn.Application;
            if (!EnsureReadyForPaste("Paste Fill Color")) return;

            var destSel = app.Selection as Excel.Range;
            var srcRange = TryGetStoredSourceRange(app);
            if (destSel == null || srcRange == null) { ShowSelectionError(); return; }

            var prevScreenUpdating = app.ScreenUpdating;
            var prevCalculation = app.Calculation;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Cache source fill states and colors, detect uniform fast-paths
                var noFill = new bool[_storedRows, _storedCols];
                var colors = new object[_storedRows, _storedCols];

                bool allNoFill = true;
                bool uniformColor = true;
                object firstColor = null;

                for (int r = 1; r <= _storedRows; r++)
                {
                    for (int c = 1; c <= _storedCols; c++)
                    {
                        var s = (Excel.Range)srcRange.Cells[r, c];
                        var si = s.Interior;
                        bool none = (int)si.Pattern == (int)Excel.XlPattern.xlPatternNone;
                        noFill[r - 1, c - 1] = none;

                        if (none)
                        {
                            colors[r - 1, c - 1] = null;
                        }
                        else
                        {
                            var col = si.Color;
                            colors[r - 1, c - 1] = col;
                            allNoFill = false;
                            if (firstColor == null) firstColor = col;
                            else if (!Equals(firstColor, col)) uniformColor = false;
                        }
                    }
                }

                foreach (Excel.Range area in destSel.Areas)
                {
                    int dr = area.Rows.Count, dc = area.Columns.Count;

                    if (allNoFill)
                    {
                        area.Interior.Pattern = Excel.XlPattern.xlPatternNone; // clear in one call
                        continue;
                    }
                    if (uniformColor && firstColor != null)
                    {
                        area.Interior.Color = firstColor; // set in one call
                        continue;
                    }

                    for (int r = 1; r <= dr; r++)
                    {
                        int sr = ((r - 1) % _storedRows) + 1;
                        for (int c = 1; c <= dc; c++)
                        {
                            int sc = ((c - 1) % _storedCols) + 1;
                            var dInt = ((Excel.Range)area.Cells[r, c]).Interior;
                            if (noFill[sr - 1, sc - 1])
                                dInt.Pattern = Excel.XlPattern.xlPatternNone;
                            else
                                dInt.Color = colors[sr - 1, sc - 1];
                        }
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
                app.EnableEvents = prevEnableEvents;
            }
        }

        public static void PasteBorders()
        {
            var app = Globals.ThisAddIn.Application;
            if (!EnsureReadyForPaste("Paste Borders")) return;

            var destSel = app.Selection as Excel.Range;
            var srcRange = TryGetStoredSourceRange(app);
            if (destSel == null || srcRange == null) { ShowSelectionError(); return; }

            var prevScreenUpdating = app.ScreenUpdating;
            var prevCalculation = app.Calculation;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Cache border snapshots once from the source, reuse on tiling
                var snaps = new BorderSnapshot[_storedRows, _storedCols];
                for (int r = 1; r <= _storedRows; r++)
                    for (int c = 1; c <= _storedCols; c++)
                        snaps[r - 1, c - 1] = CaptureBorders((Excel.Range)srcRange.Cells[r, c]);

                foreach (Excel.Range area in destSel.Areas)
                {
                    int dr = area.Rows.Count, dc = area.Columns.Count;

                    for (int r = 1; r <= dr; r++)
                    {
                        int sr = ((r - 1) % _storedRows) + 1;
                        for (int c = 1; c <= dc; c++)
                        {
                            int sc = ((c - 1) % _storedCols) + 1;
                            var d = (Excel.Range)area.Cells[r, c];
                            RestoreBorders(d, snaps[sr - 1, sc - 1]);
                        }
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
                app.EnableEvents = prevEnableEvents;
            }
        }

        public static void pasteCellWidths()
        {
            var app = Globals.ThisAddIn.Application;
            if (!EnsureReadyForPaste("Paste Column Widths")) return;

            var destSel = app.Selection as Excel.Range;
            var srcRange = TryGetStoredSourceRange(app);
            if (destSel == null || srcRange == null) { ShowSelectionError(); return; }

            var prevScreenUpdating = app.ScreenUpdating;
            var prevCalculation = app.Calculation;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Cache source widths
                var widths = new double[_storedCols];
                bool uniform = true;
                double first = 0;

                for (int c = 1; c <= _storedCols; c++)
                {
                    double w = ((Excel.Range)srcRange.Columns[c]).ColumnWidth;
                    widths[c - 1] = w;
                    if (c == 1) first = w;
                    else if (Math.Abs(first - w) > 0.0001) uniform = false;
                }

                foreach (Excel.Range area in destSel.Areas)
                {
                    int dc = area.Columns.Count;
                    if (uniform)
                    {
                        area.Columns.ColumnWidth = first; // one call
                        continue;
                    }
                    for (int c = 1; c <= dc; c++)
                    {
                        int sc = ((c - 1) % _storedCols);
                        ((Excel.Range)area.Columns[c]).ColumnWidth = widths[sc];
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
                app.EnableEvents = prevEnableEvents;
            }
        }

        public static void pasteCellHeights()
        {
            var app = Globals.ThisAddIn.Application;
            if (!EnsureReadyForPaste("Paste Row Heights")) return;

            var destSel = app.Selection as Excel.Range;
            var srcRange = TryGetStoredSourceRange(app);
            if (destSel == null || srcRange == null) { ShowSelectionError(); return; }

            var prevScreenUpdating = app.ScreenUpdating;
            var prevCalculation = app.Calculation;
            var prevEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                // Cache source heights
                var heights = new double[_storedRows];
                bool uniform = true;
                double first = 0;

                for (int r = 1; r <= _storedRows; r++)
                {
                    double h = ((Excel.Range)srcRange.Rows[r]).RowHeight;
                    heights[r - 1] = h;
                    if (r == 1) first = h;
                    else if (Math.Abs(first - h) > 0.0001) uniform = false;
                }

                foreach (Excel.Range area in destSel.Areas)
                {
                    int dr = area.Rows.Count;
                    if (uniform)
                    {
                        area.Rows.RowHeight = first; // one call
                        continue;
                    }
                    for (int r = 1; r <= dr; r++)
                    {
                        int sr = ((r - 1) % _storedRows);
                        ((Excel.Range)area.Rows[r]).RowHeight = heights[sr];
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
                app.EnableEvents = prevEnableEvents;
            }
        }

        // ──────────────────────────── HELPERS ──────────────────────────────────

        // Late-bound Formula2 access to avoid compile-time dependency on newer Excel Interop
        private static object GetFormula2(Excel.Range rng)
        {
            try
            {
                return rng.GetType().InvokeMember("Formula2",
                    System.Reflection.BindingFlags.GetProperty,
                    null, rng, null);
            }
            catch
            {
                return rng.Formula; // Fallback for older Excel
            }
        }

        private static void SetFormula2(Excel.Range rng, object value)
        {
            try
            {
                rng.GetType().InvokeMember("Formula2",
                    System.Reflection.BindingFlags.SetProperty,
                    null, rng, new[] { value });
            }
            catch
            {
                rng.Formula = value; // Fallback for older Excel
            }
        }

        private static void ShowSelectionError() =>
            MessageBox.Show("Please select a destination cell or range first.",
                            "PasteTools", MessageBoxButtons.OK, MessageBoxIcon.Warning);

        // Fast detection of any content (constants or formulas) without iterating all cells
        private static bool RangeHasAnyData(Excel.Range rng)
        {
            try { var _ = rng.SpecialCells(Excel.XlCellType.xlCellTypeConstants); return true; } catch { }
            try { var _ = rng.SpecialCells(Excel.XlCellType.xlCellTypeFormulas); return true; } catch { }
            return false;
        }

        private static bool EnsureReadyForPaste(string title)
        {
            var app = Globals.ThisAddIn.Application;
            if (!(app.Selection is Excel.Range))
            {
                MessageBox.Show("Please select a destination cell or range first.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (_storedRows <= 0 || _storedCols <= 0 || string.IsNullOrEmpty(_storedSheetName))
            {
                MessageBox.Show("Nothing is stored. Use the Copy button that snapshots the source first.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            return true;
        }

        private static Excel.Range TryGetStoredSourceRange(Excel.Application app)
        {
            // Prefer the exact workbook/sheet we copied from.
            Excel.Worksheet srcSheet = null;

            if (!string.IsNullOrEmpty(_storedWorkbookName))
            {
                try
                {
                    var wb = app.Workbooks.Cast<Excel.Workbook>()
                        .FirstOrDefault(w => string.Equals(w.Name, _storedWorkbookName, StringComparison.OrdinalIgnoreCase));
                    if (wb != null)
                        srcSheet = (Excel.Worksheet)wb.Worksheets[_storedSheetName];
                }
                catch { /* ignore */ }
            }

            // Fallback: search any open workbook for a matching sheet name.
            if (srcSheet == null)
            {
                foreach (Excel.Workbook wb in app.Workbooks)
                {
                    try
                    {
                        srcSheet = (Excel.Worksheet)wb.Worksheets[_storedSheetName];
                        if (srcSheet != null) break;
                    }
                    catch { /* keep searching */ }
                }
            }

            if (srcSheet == null)
            {
                // Last resort: if clipboard selection still matches size, use it.
                try
                {
                    var rs = app.ActiveWindow?.RangeSelection as Excel.Range;
                    if (rs != null && rs.Rows.Count == _storedRows && rs.Columns.Count == _storedCols)
                        return rs;
                }
                catch { }
                return null;
            }

            var topLeft = (Excel.Range)srcSheet.Cells[_storedStartRow, _storedStartCol];
            return topLeft.get_Resize(_storedRows, _storedCols);
        }

        // Copy formats without depending on the user's clipboard state.
        // 1) Try src.Copy() + dst.PasteSpecial(formats)
        // 2) If that fails, fall back to a fast manual format transfer.
        private static void CopyFormatsFromSourceToDestination(
            Excel.Application app, Excel.Range src, Excel.Range dst)
        {
            if (src == null || dst == null) return;

            try
            {
                src.Copy();
                dst.PasteSpecial(
                    Excel.XlPasteType.xlPasteFormats,
                    Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone,
                    false, false);
            }
            catch
            {
                ManualFormatCopySameSize(src, dst);
            }
            finally
            {
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }
            }
        }

        // Manual, focused formatting copy (number format, font, fill, alignment, borders).
        // Assumes src and dst are the same size.
        private static void ManualFormatCopySameSize(Excel.Range src, Excel.Range dst)
        {
            int rows = src.Rows.Count;
            int cols = src.Columns.Count;

            for (int r = 1; r <= rows; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    var s = (Excel.Range)src.Cells[r, c];
                    var d = (Excel.Range)dst.Cells[r, c];

                    // Number format
                    d.NumberFormat = s.NumberFormat;

                    // Alignment / text
                    d.HorizontalAlignment = s.HorizontalAlignment;
                    d.VerticalAlignment = s.VerticalAlignment;
                    d.WrapText = s.WrapText;
                    d.Orientation = s.Orientation;
                    d.IndentLevel = s.IndentLevel;

                    // Font (core bits)
                    var sf = s.Font; var df = d.Font;
                    df.Name = sf.Name;
                    df.Size = sf.Size;
                    df.Bold = sf.Bold;
                    df.Italic = sf.Italic;
                    df.Underline = sf.Underline;
                    df.Color = sf.Color;

                    // Fill
                    var si = s.Interior; var di = d.Interior;
                    di.Pattern = si.Pattern;
                    if ((int)si.Pattern != (int)Excel.XlPattern.xlPatternNone)
                        di.Color = si.Color;

                    // Borders (reuse helpers)
                    var snap = CaptureBorders(s);
                    RestoreBorders(d, snap);
                }
            }
        }

        // ─────────── BORDER SNAPSHOT ───────────
        private sealed class BorderSnapshot
        {
            public object[] LineStyles { get; }
            public object[] Weights { get; }
            public object[] Colors { get; }
            public BorderSnapshot(Func<int, object> ls, Func<int, object> wt, Func<int, object> cl)
            {
                LineStyles = Enumerable.Range(1, 12).Select(ls).ToArray();
                Weights = Enumerable.Range(1, 12).Select(wt).ToArray();
                Colors = Enumerable.Range(1, 12).Select(cl).ToArray();
            }
        }

        private static BorderSnapshot CaptureBorders(Excel.Range cell)
        {
            var b = cell.Borders;
            return new BorderSnapshot(
                i => b[(Excel.XlBordersIndex)i].LineStyle,
                i => b[(Excel.XlBordersIndex)i].Weight,
                i => b[(Excel.XlBordersIndex)i].Color);
        }

        private static void RestoreBorders(Excel.Range cell, BorderSnapshot snap)
        {
            var b = cell.Borders;
            for (int i = 1; i <= 12; i++)
            {
                try
                {
                    var idx = (Excel.XlBordersIndex)i;
                    var border = b[idx];

                    if (snap.LineStyles[i - 1] != null)
                    {
                        border.LineStyle = snap.LineStyles[i - 1];

                        // Only set weight and color if line style is not xlLineStyleNone
                        if ((int)snap.LineStyles[i - 1] != (int)Excel.XlLineStyle.xlLineStyleNone)
                        {
                            if (snap.Weights[i - 1] != null)
                                border.Weight = snap.Weights[i - 1];
                            if (snap.Colors[i - 1] != null)
                                border.Color = snap.Colors[i - 1];
                        }
                    }
                }
                catch
                {
                    // Some border indices might not be valid for certain ranges; skip.
                    continue;
                }
            }
        }
    }
}