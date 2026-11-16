// FormulaTools.cs – Smart Fill with Precedent-Aware Pattern Detection (no hard caps)
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes
{
    internal static class AutoFill
    {
        public static void FillRight() => SmartFill(true);
        public static void FillDown() => SmartFill(false);

        private static void SmartFill(bool horizontal)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null) { System.Windows.Forms.MessageBox.Show("Select a source block first."); return; }

            var prevScreenUpdating = app.ScreenUpdating;
            var prevCalculation = app.Calculation;
            var prevEnableEvents = app.EnableEvents;
            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                var ws = sel.Worksheet;

                // 1) Try pattern-aware, precedent-driven extent first
                int extra = horizontal
                    ? TryGetPrecedentDrivenExtentRight(ws, sel)
                    : TryGetPrecedentDrivenExtentDown(ws, sel);

                // 2) Fallback to context scan if no strong signal
                if (extra <= 0)
                {
                    extra = horizontal ? GetSmartFillCountRight(ws, sel)
                                       : GetSmartFillCountDown(ws, sel);
                }

                if (extra <= 0) return;

                Excel.Range dest = horizontal
                    ? sel.Resize[sel.Rows.Count, sel.Columns.Count + extra]
                    : sel.Resize[sel.Rows.Count + extra, sel.Columns.Count];

                sel.AutoFill(dest, Excel.XlAutoFillType.xlFillCopy);
            }
            finally
            {
                app.ScreenUpdating = prevScreenUpdating;
                app.Calculation = prevCalculation;
                app.EnableEvents = prevEnableEvents;
            }
        }

        /* ──────────────── PRECEDENT-AWARE: RIGHTWARD (match row run) ───────────── */
        private static int TryGetPrecedentDrivenExtentRight(Excel.Worksheet ws, Excel.Range sel)
        {
            // Look at the top-left formula cell only; Excel will relative-fill from there.
            var src = (Excel.Range)sel.Cells[1, 1];
            string formula = SafeGetFormula(src);
            if (string.IsNullOrEmpty(formula)) return 0;

            // Prefer using DirectPrecedents (cheap, robust); fallback to regex if unavailable.
            if (TryGetFirstSingleCellPrecedent(src, out int refRow, out int refCol))
            {
                // If the column is anchored ($A) the reference won't slide horizontally; skip.
                if (IsColumnAnchored(formula, refRow, refCol)) return 0;

                int runWidth = ContiguousRunRight(ws, refRow, refCol);
                if (runWidth <= 0) return 0;

                // Desired width equals referenced contiguous width; compute extra beyond current sel width.
                int extra = runWidth - sel.Columns.Count;
                return extra > 0 ? extra : 0;
            }

            // Regex fallback for simple A1 refs in same row
            if (TryParseFirstA1Reference(formula, out bool colAbs, out bool rowAbs, out int r, out int c))
            {
                if (colAbs) return 0; // horizontal fill won’t shift the column if anchored
                int runWidth = ContiguousRunRight(ws, r, c);
                int extra = runWidth - sel.Columns.Count;
                return extra > 0 ? extra : 0;
            }
            return 0;
        }

        /* ──────────────── PRECEDENT-AWARE: DOWNWARD (match column run) ─────────── */
        private static int TryGetPrecedentDrivenExtentDown(Excel.Worksheet ws, Excel.Range sel)
        {
            var src = (Excel.Range)sel.Cells[1, 1];
            string formula = SafeGetFormula(src);
            if (string.IsNullOrEmpty(formula)) return 0;

            if (TryGetFirstSingleCellPrecedent(src, out int refRow, out int refCol))
            {
                // If the row is anchored (A$2) the reference won’t slide vertically; skip.
                if (IsRowAnchored(formula, refRow, refCol)) return 0;

                int runHeight = ContiguousRunDown(ws, refRow, refCol);
                if (runHeight <= 0) return 0;

                int extra = runHeight - sel.Rows.Count;
                return extra > 0 ? extra : 0;
            }

            if (TryParseFirstA1Reference(formula, out bool colAbs, out bool rowAbs, out int r, out int c))
            {
                if (rowAbs) return 0; // vertical fill won’t shift the row if anchored
                int runHeight = ContiguousRunDown(ws, r, c);
                int extra = runHeight - sel.Rows.Count;
                return extra > 0 ? extra : 0;
            }
            return 0;
        }

        /* ───────────────────────── CONTEXT SCAN (STREAMING) ────────────────────── */
        private static int GetSmartFillCountRight(Excel.Worksheet ws, Excel.Range sel)
        {
            int startCol = sel.Column + sel.Columns.Count;
            int lastCol = Math.Max(startCol, LastUsedColumn(ws));

            int bandTop = Math.Max(1, sel.Row - 3);
            int bandBot = Math.Min(LastUsedRow(ws), sel.Row + sel.Rows.Count + 2);
            int nRows = Math.Max(0, bandBot - bandTop + 1);
            if (nRows == 0 || lastCol < startCol) return 0;

            int[] rowRun = new int[nRows];
            bool[] active = new bool[nRows];
            for (int r = 0; r < nRows; r++) active[r] = true;
            int activeCount = nRows;

            int firstEmptyColOffset = -1;
            const int CHUNK = 1024;

            for (int col = startCol; col <= lastCol; col += CHUNK)
            {
                int width = Math.Min(CHUNK, lastCol - col + 1);
                bool[,] mask = ReadNonEmptyMask(ws, bandTop, col, nRows, width);

                for (int j = 0; j < width; j++)
                {
                    bool any = false;
                    for (int r = 0; r < nRows; r++)
                    {
                        bool cellHasData = mask[r, j];
                        if (active[r])
                        {
                            if (cellHasData) { rowRun[r]++; any = true; }
                            else { active[r] = false; activeCount--; }
                        }
                        else if (cellHasData) { any = true; }
                    }
                    if (!any && firstEmptyColOffset < 0)
                    {
                        firstEmptyColOffset = (col - startCol) + j;
                        break;
                    }
                    if (activeCount == 0) break;
                }
                if (firstEmptyColOffset >= 0 || activeCount == 0) break;
            }

            var counts = new List<int>(nRows + 1);
            for (int i = 0; i < nRows; i++) if (rowRun[i] > 0) counts.Add(rowRun[i]);
            if (firstEmptyColOffset >= 0) counts.Add(firstEmptyColOffset);

            return DetermineBestCount(counts);
        }

        private static int GetSmartFillCountDown(Excel.Worksheet ws, Excel.Range sel)
        {
            int startRow = sel.Row + sel.Rows.Count;
            int lastRow = Math.Max(startRow, LastUsedRow(ws));

            int bandLeft = Math.Max(1, sel.Column - 3);
            int bandRight = Math.Min(LastUsedColumn(ws), sel.Column + sel.Columns.Count + 2);
            int nCols = Math.Max(0, bandRight - bandLeft + 1);
            if (nCols == 0 || lastRow < startRow) return 0;

            int[] colRun = new int[nCols];
            bool[] active = new bool[nCols];
            for (int c = 0; c < nCols; c++) active[c] = true;
            int activeCount = nCols;

            int firstEmptyRowOffset = -1;
            const int CHUNK = 4096;

            for (int row = startRow; row <= lastRow; row += CHUNK)
            {
                int height = Math.Min(CHUNK, lastRow - row + 1);
                bool[,] mask = ReadNonEmptyMask(ws, row, bandLeft, height, nCols);

                for (int i = 0; i < height; i++)
                {
                    bool any = false;
                    for (int c = 0; c < nCols; c++)
                    {
                        bool cellHasData = mask[i, c];
                        if (active[c])
                        {
                            if (cellHasData) { colRun[c]++; any = true; }
                            else { active[c] = false; activeCount--; }
                        }
                        else if (cellHasData) { any = true; }
                    }
                    if (!any && firstEmptyRowOffset < 0)
                    {
                        firstEmptyRowOffset = (row - startRow) + i;
                        break;
                    }
                    if (activeCount == 0) break;
                }
                if (firstEmptyRowOffset >= 0 || activeCount == 0) break;
            }

            var counts = new List<int>(nCols + 1);
            for (int i = 0; i < nCols; i++) if (colRun[i] > 0) counts.Add(colRun[i]);
            if (firstEmptyRowOffset >= 0) counts.Add(firstEmptyRowOffset);

            return DetermineBestCount(counts);
        }

        /* ──────────────────────── CONTIGUOUS RUN MEASURES ─────────────────────── */
        private static int ContiguousRunDown(Excel.Worksheet ws, int startRow, int col)
        {
            int lastRow = Math.Max(startRow, LastUsedRow(ws));
            int run = 0;
            const int CHUNK = 4096;
            for (int row = startRow; row <= lastRow; row += CHUNK)
            {
                int height = Math.Min(CHUNK, lastRow - row + 1);
                bool[,] mask = ReadNonEmptyMask(ws, row, col, height, 1);
                for (int i = 0; i < height; i++)
                {
                    if (mask[i, 0]) run++;
                    else return run;
                }
            }
            return run;
        }

        private static int ContiguousRunRight(Excel.Worksheet ws, int row, int startCol)
        {
            int lastCol = Math.Max(startCol, LastUsedColumn(ws));
            int run = 0;
            const int CHUNK = 1024;
            for (int col = startCol; col <= lastCol; col += CHUNK)
            {
                int width = Math.Min(CHUNK, lastCol - col + 1);
                bool[,] mask = ReadNonEmptyMask(ws, row, col, 1, width);
                for (int j = 0; j < width; j++)
                {
                    if (mask[0, j]) run++;
                    else return run;
                }
            }
            return run;
        }

        /* ─────────────────────────  MASK / FIND HELPERS  ──────────────────────── */
        private static bool[,] ReadNonEmptyMask(Excel.Worksheet ws, int r1, int c1, int nRows, int nCols)
        {
            Excel.Range rng = null;
            try
            {
                rng = ws.Range[ws.Cells[r1, c1], ws.Cells[r1 + nRows - 1, c1 + nCols - 1]];
                object value2 = rng.Value2;
                object[,] raw = To2D(value2, nRows, nCols);

                bool[,] mask = new bool[nRows, nCols];
                for (int i = 0; i < nRows; i++)
                {
                    for (int j = 0; j < nCols; j++)
                    {
                        object v = raw[i + 1, j + 1];
                        if (v == null) mask[i, j] = false;
                        else if (v is string s) mask[i, j] = !string.IsNullOrWhiteSpace(s);
                        else mask[i, j] = true;
                    }
                }
                return mask;
            }
            finally
            {
                ReleaseCom(rng);
            }
        }

        private static int LastUsedRow(Excel.Worksheet ws)
        {
            Excel.Range cell = null;
            try
            {
                cell = ws.Cells.Find("*", ws.Cells[1, 1],
                                     Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                     Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                     false, false, false);
                return cell != null ? cell.Row : 1;
            }
            finally { ReleaseCom(cell); }
        }

        private static int LastUsedColumn(Excel.Worksheet ws)
        {
            Excel.Range cell = null;
            try
            {
                cell = ws.Cells.Find("*", ws.Cells[1, 1],
                                     Excel.XlFindLookIn.xlFormulas, Excel.XlLookAt.xlPart,
                                     Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious,
                                     false, false, false);
                return cell != null ? cell.Column : 1;
            }
            finally { ReleaseCom(cell); }
        }

        /* ───────────────────────── PRECEDENT / PARSING ────────────────────────── */
        private static string SafeGetFormula(Excel.Range cell)
        {
            try { return cell.Formula as string; } catch { return null; }
        }

        private static bool TryGetFirstSingleCellPrecedent(Excel.Range cell, out int row, out int col)
        {
            row = col = 0;
            Excel.Range prec = null;
            try
            {
                prec = cell.DirectPrecedents; // may fail for external refs / closed books
                if (prec == null) return false;

                // We only accept a single-cell precedent (strongest signal for ROW/COL sliding).
                if (prec.Areas.Count == 1 && prec.Cells.Count == 1)
                {
                    row = prec.Row;
                    col = prec.Column;
                    return true;
                }

                // If multiple, prefer a single-cell area if present.
                foreach (Excel.Range area in prec.Areas)
                {
                    if (area.Cells.Count == 1)
                    {
                        row = area.Row;
                        col = area.Column;
                        return true;
                    }
                }
                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(prec);
            }
        }

        // Very lightweight A1 ref finder: picks the first cell ref (optionally sheet-qualified).
        private static readonly Regex A1Ref = new Regex(
            @"(?:'[^']+'!|\[[^\]]+\][^!]+!)?\s*(\$?)([A-Za-z]{1,3})(\$?)(\d+)",
            RegexOptions.Compiled);

        private static bool TryParseFirstA1Reference(string formula, out bool colAbs, out bool rowAbs, out int row, out int col)
        {
            colAbs = rowAbs = false; row = col = 0;
            if (string.IsNullOrEmpty(formula)) return false;

            var m = A1Ref.Match(formula);
            if (!m.Success) return false;

            colAbs = m.Groups[1].Value == "$";
            string colLetters = m.Groups[2].Value;
            rowAbs = m.Groups[3].Value == "$";
            row = int.Parse(m.Groups[4].Value);
            col = ColToNum(colLetters);
            return col > 0 && row > 0;
        }

        private static bool IsColumnAnchored(string formula, int refRow, int refCol)
        {
            // Check for $ before the column letters of the first match that points to refCol/refRow.
            // Heuristic: we look at the first A1 ref; good enough for simple patterns.
            if (TryParseFirstA1Reference(formula, out bool colAbs, out _, out int r, out int c))
                if (r == refRow && c == refCol) return colAbs;
            return false;
        }

        private static bool IsRowAnchored(string formula, int refRow, int refCol)
        {
            if (TryParseFirstA1Reference(formula, out _, out bool rowAbs, out int r, out int c))
                if (r == refRow && c == refCol) return rowAbs;
            return false;
        }

        private static int ColToNum(string letters)
        {
            int col = 0;
            for (int i = 0; i < letters.Length; i++)
            {
                col = col * 26 + (char.ToUpperInvariant(letters[i]) - 'A' + 1);
            }
            return col;
        }

        /* ─────────────────────────  DECISION / MODE PICK  ─────────────────────── */
        private static int DetermineBestCount(List<int> candidates)
        {
            if (candidates == null || candidates.Count == 0) return 0;

            var freq = new Dictionary<int, int>();
            for (int i = 0; i < candidates.Count; i++)
            {
                int k = candidates[i];
                if (k <= 0) continue;
                if (freq.TryGetValue(k, out int v)) freq[k] = v + 1; else freq[k] = 1;
            }
            if (freq.Count == 0) return 0;

            int bestKey = 0, bestCount = -1;
            foreach (var kv in freq)
                if (kv.Value > bestCount || (kv.Value == bestCount && kv.Key < bestKey))
                { bestKey = kv.Key; bestCount = kv.Value; }

            if (bestCount == 1 && candidates.Count >= 3)
            {
                candidates.Sort();
                return candidates[candidates.Count / 2];
            }
            return bestKey;
        }

        /* ─────────────────────────────  UTILITIES  ────────────────────────────── */
        private static object[,] To2D(object value2, int rows, int cols)
        {
            var arr = value2 as object[,];
            if (arr != null) return arr;
            var one = (object[,])Array.CreateInstance(typeof(object), new[] { rows, cols }, new[] { 1, 1 });
            one[1, 1] = value2;
            return one;
        }

        private static void ReleaseCom(object comObj)
        {
            if (comObj != null && Marshal.IsComObject(comObj))
            {
                try { Marshal.ReleaseComObject(comObj); } catch { /* ignore */ }
            }
        }
    }
}
