// File: StatBubbles.cs
// C# 7.3 / .NET Framework 4.8 (VSTO)

using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Excel
{
    public static class StatBubbles
    {
        public static void InsertMetricsBubbles()
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return;

            var prevSU = app.ScreenUpdating;
            var prevDA = app.DisplayAlerts;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;

                var chart = ResolveActiveChart(app);
                if (chart == null) { System.Windows.Forms.MessageBox.Show("Select a chart first.", "Metrics Bubbles"); return; }
                if (chart.SeriesCollection().Count < 1) { System.Windows.Forms.MessageBox.Show("The chart has no series.", "Metrics Bubbles"); return; }

                var ws = chart.Parent is Excel.ChartObject co ? (Excel.Worksheet)co.Parent : app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;

                int serCount = chart.SeriesCollection().Count;

                // ---- Try to auto-detect a range from the first series ----
                var sFirst = (Excel.Series)chart.SeriesCollection(1);
                Excel.Range firstY = TryGetSeriesYRange(sFirst, app);   // <- NEW: works even if Values isn't a Range
                Excel.Range lastY = null;

                if (firstY != null && serCount > 1)
                {
                    // find the last series that is also range-bound to maybe treat as a contiguous block
                    for (int s = serCount; s >= 1; s--)
                    {
                        var rng = TryGetSeriesYRange((Excel.Series)chart.SeriesCollection(s), app);
                        if (rng != null) { lastY = rng; break; }
                    }
                }

                bool vertical;
                Excel.Range yHelperRange, labelRange;

                if (firstY != null)
                {
                    // We have a real linked range → no prompt
                    vertical = firstY.Rows.Count >= firstY.Columns.Count;
                    try
                    {
                        if (chart.PlotBy == Excel.XlRowCol.xlRows) vertical = false;
                        if (chart.PlotBy == Excel.XlRowCol.xlColumns) vertical = true;
                    }
                    catch { }

                    int pointCount = vertical ? firstY.Rows.Count : firstY.Columns.Count;

                    // Build contiguous block if lastY aligns with firstY; else just firstY
                    Excel.Range block = firstY;
                    if (lastY != null && SameSheet(firstY, lastY))
                    {
                        bool aligned = vertical
                            ? (firstY.Row == lastY.Row && firstY.Rows.Count == lastY.Rows.Count)
                            : (firstY.Column == lastY.Column && firstY.Columns.Count == lastY.Columns.Count);

                        if (aligned)
                        {
                            block = vertical
                                ? ws.Range[ws.Cells[firstY.Row, firstY.Column], ws.Cells[firstY.Row + firstY.Rows.Count - 1, lastY.Column]]
                                : ws.Range[ws.Cells[firstY.Row, firstY.Column], ws.Cells[lastY.Row, firstY.Column + firstY.Columns.Count - 1]];
                        }
                    }

                    string maxFormula = $"=MAX({AbsA1(block)})*1.25";

                    if (vertical)
                    {
                        int r0 = firstY.Row;
                        int r1 = r0 + pointCount - 1;
                        int insertCol = (block.Column == firstY.Column && block.Columns.Count > 1)
                                        ? (block.Column + block.Columns.Count)           // right of whole block
                                        : (firstY.Column + firstY.Columns.Count);        // right of first series

                        ws.Columns[insertCol].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                        yHelperRange = ws.Range[ws.Cells[r0, insertCol], ws.Cells[r1, insertCol]];
                        yHelperRange.Formula = maxFormula;

                        int labelCol = insertCol + 1;
                        ws.Columns[labelCol].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                        labelRange = ws.Range[ws.Cells[r0, labelCol], ws.Cells[r1, labelCol]];
                        labelRange.Value2 = "TBU";
                    }
                    else
                    {
                        int c0 = firstY.Column;
                        int c1 = c0 + firstY.Columns.Count - 1;
                        int insertRow = (block.Row == firstY.Row && block.Rows.Count > 1)
                                        ? (block.Row + block.Rows.Count)                 // below whole block
                                        : (firstY.Row + firstY.Rows.Count);              // below first series

                        ws.Rows[insertRow].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        yHelperRange = ws.Range[ws.Cells[insertRow, c0], ws.Cells[insertRow, c1]];
                        yHelperRange.Formula = maxFormula;

                        int labelRow = insertRow + 1;
                        ws.Rows[labelRow].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        labelRange = ws.Range[ws.Cells[labelRow, c0], ws.Cells[labelRow, c1]];
                        labelRange.Value2 = "TBU";
                    }

                    AddOverlaySeries(chart, yHelperRange, labelRange);
                    return;
                }

                // ---- No range could be derived from series 1 (literal array). Only now will we ask for the Y range. ----
                // We still need an anchor to know where to insert helpers.
                bool verticalDefault = true;
                try
                {
                    if (chart.PlotBy == Excel.XlRowCol.xlRows) verticalDefault = false;
                    if (chart.PlotBy == Excel.XlRowCol.xlColumns) verticalDefault = true;
                }
                catch { }
                vertical = verticalDefault;

                int pts = 0;
                try { pts = sFirst.Points().Count; } catch { }
                if (pts < 1) { System.Windows.Forms.MessageBox.Show("Could not determine point count.", "Metrics Bubbles"); return; }

                string anchorMsg = vertical
                    ? "Select the TOP cell where the FIRST helper column should be inserted.\nTwo columns (Y then Labels) will be inserted."
                    : "Select the LEFTMOST cell where the FIRST helper row should be inserted.\nTwo rows (Y then Labels) will be inserted.";
                var anchor = app.InputBox(anchorMsg, "Insert helper rows/columns", Type: 8) as Excel.Range;
                if (anchor == null) return;

                // One prompt ONLY when the series truly isn't linked anywhere.
                var maxPick = app.InputBox(
                    "Select the FULL Y data range to MAX over (cells that feed the chart’s Y values).",
                    "Pick range to MAX", Type: 8) as Excel.Range;
                if (maxPick == null) return;

                string maxFormulaArray = $"=MAX({AbsA1(maxPick)})*1.25";

                if (vertical)
                {
                    int r0 = anchor.Row;
                    int yCol = anchor.Column;

                    ws.Columns[yCol].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);
                    ws.Columns[yCol + 1].Insert(Excel.XlInsertShiftDirection.xlShiftToRight);

                    yHelperRange = ws.Range[ws.Cells[r0, yCol], ws.Cells[r0 + pts - 1, yCol]];
                    labelRange = ws.Range[ws.Cells[r0, yCol + 1], ws.Cells[r0 + pts - 1, yCol + 1]];

                    yHelperRange.Formula = maxFormulaArray;
                    labelRange.Value2 = "TBU";
                }
                else
                {
                    int c0 = anchor.Column;
                    int yRow = anchor.Row;

                    ws.Rows[yRow].Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                    ws.Rows[yRow + 1].Insert(Excel.XlInsertShiftDirection.xlShiftDown);

                    yHelperRange = ws.Range[ws.Cells[yRow, c0], ws.Cells[yRow, c0 + pts - 1]];
                    labelRange = ws.Range[ws.Cells[yRow + 1, c0], ws.Cells[yRow + 1, c0 + pts - 1]];

                    yHelperRange.Formula = maxFormulaArray;
                    labelRange.Value2 = "TBU";
                }

                AddOverlaySeries(chart, yHelperRange, labelRange);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Metrics bubble setup failed:\r\n" + ex.Message, "Metrics Bubbles");
            }
            finally
            {
                app.DisplayAlerts = prevDA;
                app.ScreenUpdating = prevSU;
            }
        }

        // ---------- NEW: robust Y-range detection from a Series ----------
        private static Excel.Range TryGetSeriesYRange(Excel.Series s, Excel.Application app)
        {
            if (s == null) return null;

            // 1) Easiest: Values already a Range
            var rng = s.Values as Excel.Range;
            if (rng != null) return rng;

            // 2) Parse =SERIES(name, xRange, yRange, order[, bubbleSize]) from Series.Formula
            try
            {
                string f = s.Formula; // English/invariant, comma separated
                if (string.IsNullOrEmpty(f)) f = s.FormulaR1C1;
                if (!string.IsNullOrEmpty(f) && f.StartsWith("=SERIES(", StringComparison.OrdinalIgnoreCase))
                {
                    var args = SplitSeriesArgs(f);
                    if (args != null && args.Length >= 3)
                    {
                        string yArg = args[2].Trim();
                        // Some charts embed literal arrays {1,2,3}; skip those
                        if (!yArg.StartsWith("{") && !yArg.StartsWith("\"") && yArg.IndexOfAny(new[] { '!', '$', ':', '\'' }) >= 0)
                        {
                            try
                            {
                                var eval = app.Evaluate(yArg);
                                var yRange = eval as Excel.Range;
                                if (yRange != null) return yRange;
                            }
                            catch { /* fall through */ }
                        }
                    }
                }
            }
            catch { }

            return null; // caller may prompt as last resort
        }

        // Split the inside of SERIES(...) into top-level comma-separated args
        private static string[] SplitSeriesArgs(string seriesFormula)
        {
            // Remove "=SERIES(" and trailing ")"
            int start = seriesFormula.IndexOf('(');
            int end = seriesFormula.LastIndexOf(')');
            if (start < 0 || end < 0 || end <= start) return null;
            string inner = seriesFormula.Substring(start + 1, end - start - 1);

            var parts = new System.Collections.Generic.List<string>(5);
            int depth = 0; int last = 0;
            for (int i = 0; i < inner.Length; i++)
            {
                char c = inner[i];
                if (c == '(' || c == '[') depth++;
                else if (c == ')' || c == ']') depth--;
                else if (c == ',' && depth == 0)
                {
                    parts.Add(inner.Substring(last, i - last));
                    last = i + 1;
                }
            }
            parts.Add(inner.Substring(last));
            return parts.Select(p => p.Trim()).ToArray();
        }

        // ---------- Overlay & misc helpers ----------
        private static void AddOverlaySeries(Excel.Chart chart, Excel.Range yHelperRange, Excel.Range labelRange)
        {
            int n = (yHelperRange.Rows.Count >= yHelperRange.Columns.Count)
                    ? yHelperRange.Rows.Count : yHelperRange.Columns.Count;

            var ov = (Excel.Series)chart.SeriesCollection().NewSeries();
            ov.ChartType = Excel.XlChartType.xlXYScatter;
            ov.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
            try { ov.AxisGroup = Excel.XlAxisGroup.xlPrimary; } catch { }
            ov.Format.Line.Visible = Office.MsoTriState.msoFalse;

            ov.XValues = Enumerable.Range(1, n).Select(i => (object)(double)i).ToArray();
            ov.Values = yHelperRange;

            for (int i = 1; i <= n; i++)
            {
                var pt = (Excel.Point)ov.Points(i);
                pt.HasDataLabel = true;
                var dl = (Excel.DataLabel)pt.DataLabel;

                try { dl.Position = Excel.XlDataLabelPosition.xlLabelPositionCenter; } catch { }
                Excel.Range cell = (yHelperRange.Rows.Count >= yHelperRange.Columns.Count)
                    ? (Excel.Range)labelRange.Cells[i, 1]
                    : (Excel.Range)labelRange.Cells[1, i];

                LinkDataLabelToCell(dl, cell);

                dl.Format.Fill.Visible = Office.MsoTriState.msoTrue;
                dl.Format.Fill.Solid();
                dl.Format.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(221, 221, 221));
                dl.Format.Line.Visible = Office.MsoTriState.msoFalse;
            }
        }

        private static bool SameSheet(Excel.Range a, Excel.Range b) => a.Worksheet == b.Worksheet;

        private static string AbsA1(Excel.Range rng)
        {
            try { return rng.get_Address(true, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing); }
            catch
            {
                try { return (string)((dynamic)rng).Address[true, true, Excel.XlReferenceStyle.xlA1, false, Type.Missing]; }
                catch { return rng.Address; }
            }
        }

        private static void LinkDataLabelToCell(Excel.DataLabel dl, Excel.Range cell)
        {
            try
            {
                string addr = AbsA1(cell);
                string sheet = cell.Worksheet.Name.Replace("'", "''");
                ((dynamic)dl).Formula = $"='{sheet}'!{addr}";
            }
            catch
            {
                try { dl.Text = Convert.ToString(cell.Value2) ?? string.Empty; }
                catch
                {
                    Excel.Characters ch = (Excel.Characters)dl.Characters[Type.Missing, Type.Missing];
                    ch.Text = Convert.ToString(cell.Value2) ?? string.Empty;
                }
            }
        }

        private static Excel.Chart ResolveActiveChart(Excel.Application app)
        {
            if (app.ActiveChart != null) return app.ActiveChart;
            var sel = app.Selection;
            if (sel is Excel.ChartObject co) return co.Chart;
            try { if (sel is Excel.Series s) return s.Parent as Excel.Chart; } catch { }
            try
            {
                if (sel is Excel.Shape sh && sh.HasChart == Office.MsoTriState.msoTrue) return sh.Chart;
                if (sel is Excel.ShapeRange sr && sr.Count == 1 && sr.Item(1).HasChart == Office.MsoTriState.msoTrue) return sr.Item(1).Chart;
            }
            catch { }
            return null;
        }
    }
}
