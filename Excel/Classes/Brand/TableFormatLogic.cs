// ──────────────────────────────────────────────────────────────────────────────
// TableFormatLogic.cs  (C# 7.3 / .NET 4.8, VSTO)
// Captures header/body/footer font, fill, borders. Number formats:
//   • Header/Footer: per-column NF
//   • Body: most-frequent NF per column
// Borders: body formatted first, header/footer afterwards so their edges win.
// ──────────────────────────────────────────────────────────────────────────────
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using OfficePragmatics_Excel.Classes.Brand;        // FontFormat / LineFormat
using MsoTriState = Microsoft.Office.Core.MsoTriState;

namespace OfficePragmatics_Excel.Classes.Tables
{

    public class TableFormatSettings
    {
        public bool IsSet { get; set; }
        public string Name { get; set; }

        public RowBlockFormat Header { get; set; } = new RowBlockFormat();
        public RowBlockFormat Body { get; set; } = new RowBlockFormat();
        public RowBlockFormat Footer { get; set; } = new RowBlockFormat();

        public List<BodyColumnFormat> BodyColumnFormats { get; set; } = new List<BodyColumnFormat>();
    }

    public class RowBlockFormat
    {
        public bool IsSet { get; set; }
        public FontFormat Font { get; set; } = new FontFormat();
        public Color? FillColor { get; set; }
        public Color? FontColor { get; set; }
        public BorderSet Borders { get; set; } = new BorderSet();

        // Single-row blocks: per-column NF
        public Dictionary<int, string> ColumnNumberFormats { get; set; } = new Dictionary<int, string>();
        // Multi-row blocks with one uniform NF (rare, but keep)
        public string UniformNumberFormat { get; set; }
    }

    public class BorderSet
    {
        public bool IsSet { get; set; }
        public LineFormat Top { get; set; } = new LineFormat();
        public LineFormat Bottom { get; set; } = new LineFormat();
        public LineFormat Left { get; set; } = new LineFormat();
        public LineFormat Right { get; set; } = new LineFormat();
        public LineFormat InsideH { get; set; } = new LineFormat();
        public LineFormat InsideV { get; set; } = new LineFormat();
    }

    public class BodyColumnFormat
    {
        public int ColumnIndex { get; set; }      // 1-based
        public string NumberFormat { get; set; }  // most common NF in the column (body block)
    }

    public static class TableFormatter
    {
        // ---------------------------------------------------------------------
        // PUBLIC
        // ---------------------------------------------------------------------
        public static TableFormatSettings ImportFromSelection(string name = null)
        {
            Excel.Range rng = GetActiveRange();
            if (rng == null) throw new InvalidOperationException("Select a contiguous range.");

            int rows = rng.Rows.Count;
            int cols = rng.Columns.Count;

            var set = new TableFormatSettings
            {
                IsSet = true,
                Name = string.IsNullOrWhiteSpace(name)
                        ? $"{rng.Worksheet.Name}!{rng.Address[false, false]}"
                        : name
            };

            // Header (first row)
            Excel.Range hdr = rng.Rows[1];
            set.Header = CaptureRowBlock(hdr, cols);

            // Footer (last row, only if at least 3 rows)
            if (rows > 2)
            {
                Excel.Range ftr = rng.Rows[rows];
                set.Footer = CaptureRowBlock(ftr, cols);
            }

            // Body (middle rows)
            if (rows > 2)
            {
                Excel.Range body = rng.Range[rng.Rows[2], rng.Rows[rows - 1]];
                set.Body = CaptureRowBlock(body, cols);
                set.BodyColumnFormats = CaptureBodyColumnFormats(body, cols);
            }

            return set;
        }

        public static void ApplyToSelection(TableFormatSettings set)
        {
            if (set == null || !set.IsSet) return;
            Excel.Range rng = GetActiveRange();
            if (rng == null) return;

            int rows = rng.Rows.Count;

            // 1) BODY FIRST
            if (rows > 2 && set.Body.IsSet)
            {
                Excel.Range body = rng.Range[rng.Rows[2], rng.Rows[rows - 1]];
                ApplyRowBlock(body, set.Body, singleRow: false);
                ApplyBodyColumnFormats(body, set.BodyColumnFormats);
            }

            // 2) HEADER
            if (rows >= 1)
            {
                Excel.Range hdr = rng.Rows[1];
                ApplyRowBlock(hdr, set.Header, singleRow: true);
            }

            // 3) FOOTER
            if (rows > 2)
            {
                Excel.Range ftr = rng.Rows[rows];
                ApplyRowBlock(ftr, set.Footer, singleRow: true);
            }
        }

        // ---------------------------------------------------------------------
        // CAPTURE
        // ---------------------------------------------------------------------
        private static RowBlockFormat CaptureRowBlock(Excel.Range block, int totalCols)
        {
            var rb = new RowBlockFormat { IsSet = true };
            Excel.Range first = block.Cells[1, 1];

            rb.Font = new FontFormat
            {
                IsSet = true,
                Name = (string)first.Font.Name,
                Size = (double?)first.Font.Size,
                Bold = (bool?)first.Font.Bold,
                Color = SafeColor(first.Font.Color)
            };

            rb.FontColor = SafeColor(first.Font.Color);
            rb.FillColor = SafeColor(first.Interior.Color);
            rb.Borders = CaptureBorders(block);

            if (block.Rows.Count == 1)
            {
                // Save per-column NF for header/footer
                for (int c = 1; c <= totalCols; c++)
                {
                    var cell = block.Cells[1, c] as Excel.Range;
                    string nf = cell?.NumberFormat as string;
                    if (!string.IsNullOrEmpty(nf))
                        rb.ColumnNumberFormats[c] = nf;
                }
            }
            else
            {
                rb.UniformNumberFormat = TryGetUniformNumberFormat(block);
            }

            return rb;
        }

        private static BorderSet CaptureBorders(Excel.Range block)
        {
            var bs = new BorderSet { IsSet = true };
            FillBorder(block.Borders[Excel.XlBordersIndex.xlEdgeTop], bs.Top);
            FillBorder(block.Borders[Excel.XlBordersIndex.xlEdgeBottom], bs.Bottom);
            FillBorder(block.Borders[Excel.XlBordersIndex.xlEdgeLeft], bs.Left);
            FillBorder(block.Borders[Excel.XlBordersIndex.xlEdgeRight], bs.Right);
            FillBorder(block.Borders[Excel.XlBordersIndex.xlInsideHorizontal], bs.InsideH);
            FillBorder(block.Borders[Excel.XlBordersIndex.xlInsideVertical], bs.InsideV);
            return bs;
        }

        private static void FillBorder(Excel.Border b, LineFormat lf)
        {
            try
            {
                lf.IsSet = true;
                if (b.LineStyle == Excel.XlLineStyle.xlLineStyleNone)
                {
                    lf.Visible = MsoTriState.msoFalse;
                    lf.Weight = null;
                    lf.ForeColor = null;
                    return;
                }

                lf.Visible = MsoTriState.msoTrue;
                lf.Weight = Convert.ToDouble(b.Weight);
                lf.ForeColor = SafeColor(b.Color);
            }
            catch { lf.IsSet = false; }
        }

        private static List<BodyColumnFormat> CaptureBodyColumnFormats(Excel.Range body, int cols)
        {
            var res = new List<BodyColumnFormat>();

            for (int c = 1; c <= cols; c++)
            {
                var col = body.Columns[c];
                var counts = new Dictionary<string, int>(StringComparer.Ordinal);

                foreach (Excel.Range cell in col.Cells)
                {
                    string nf = cell.NumberFormat as string;
                    if (string.IsNullOrEmpty(nf)) continue;
                    counts[nf] = counts.TryGetValue(nf, out var ct) ? ct + 1 : 1;
                }

                if (counts.Count > 0)
                {
                    string winner = counts.OrderByDescending(kv => kv.Value).First().Key;
                    res.Add(new BodyColumnFormat { ColumnIndex = c, NumberFormat = winner });
                }
            }

            return res;
        }

        // ---------------------------------------------------------------------
        // APPLY
        // ---------------------------------------------------------------------
        private static void ApplyRowBlock(Excel.Range block, RowBlockFormat rb, bool singleRow)
        {
            if (!rb.IsSet) return;

            // Fill
            try
            {
                if (rb.FillColor.HasValue)
                    block.Interior.Color = ColorTranslator.ToOle(rb.FillColor.Value);
            }
            catch { }

            // Font
            try
            {
                if (rb.Font?.IsSet == true)
                {
                    if (!string.IsNullOrEmpty(rb.Font.Name)) block.Font.Name = rb.Font.Name;
                    if (rb.Font.Size.HasValue) block.Font.Size = rb.Font.Size.Value;
                    if (rb.Font.Bold.HasValue) block.Font.Bold = rb.Font.Bold.Value;
                    if (rb.Font.Color.HasValue) block.Font.Color = ColorTranslator.ToOle(rb.Font.Color.Value);
                }
                if (rb.FontColor.HasValue)
                    block.Font.Color = ColorTranslator.ToOle(rb.FontColor.Value);
            }
            catch { }

            // Number formats
            try
            {
                if (singleRow && rb.ColumnNumberFormats.Count > 0)
                {
                    foreach (var kv in rb.ColumnNumberFormats)
                    {
                        var cell = block.Cells[1, kv.Key] as Excel.Range;
                        if (cell != null) cell.NumberFormat = kv.Value;
                    }
                }
                else if (!singleRow && !string.IsNullOrEmpty(rb.UniformNumberFormat))
                {
                    block.NumberFormat = rb.UniformNumberFormat;
                }
            }
            catch { }

            // Borders
            ApplyBorders(block, rb.Borders);
        }

        private static void ApplyBorders(Excel.Range block, BorderSet bs)
        {
            if (bs == null || !bs.IsSet) return;
            SetBorder(block.Borders[Excel.XlBordersIndex.xlEdgeTop], bs.Top);
            SetBorder(block.Borders[Excel.XlBordersIndex.xlEdgeBottom], bs.Bottom);
            SetBorder(block.Borders[Excel.XlBordersIndex.xlEdgeLeft], bs.Left);
            SetBorder(block.Borders[Excel.XlBordersIndex.xlEdgeRight], bs.Right);
            SetBorder(block.Borders[Excel.XlBordersIndex.xlInsideHorizontal], bs.InsideH);
            SetBorder(block.Borders[Excel.XlBordersIndex.xlInsideVertical], bs.InsideV);
        }

        private static void SetBorder(Excel.Border b, LineFormat lf)
        {
            if (lf == null || !lf.IsSet) return;
            try
            {
                if (lf.Visible == MsoTriState.msoFalse)
                {
                    b.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                    return;
                }

                // Make sure style exists before weight/color or Excel ignores it
                b.LineStyle = Excel.XlLineStyle.xlContinuous;

                if (lf.Weight.HasValue) b.Weight = lf.Weight.Value;
                if (lf.ForeColor.HasValue) b.Color = ColorTranslator.ToOle(lf.ForeColor.Value);
            }
            catch { }
        }

        private static void ApplyBodyColumnFormats(Excel.Range body, List<BodyColumnFormat> list)
        {
            if (list == null || list.Count == 0) return;
            int cols = body.Columns.Count;

            foreach (var cf in list)
            {
                if (cf.ColumnIndex < 1 || cf.ColumnIndex > cols) continue;
                try { body.Columns[cf.ColumnIndex].NumberFormat = cf.NumberFormat; } catch { }
            }
        }

        // ---------------------------------------------------------------------
        // UTILS
        // ---------------------------------------------------------------------
        private static string TryGetUniformNumberFormat(Excel.Range block)
        {
            string first = null;
            foreach (Excel.Range c in block.Cells)
            {
                string nf = c.NumberFormat as string;
                if (string.IsNullOrEmpty(nf)) continue;
                if (first == null) first = nf;
                else if (!string.Equals(first, nf, StringComparison.Ordinal)) return null;
            }
            return first;
        }

        private static Color? SafeColor(object ole)
        {
            try { return ColorTranslator.FromOle(Convert.ToInt32(ole)); }
            catch { return null; }
        }

        private static Excel.Range GetActiveRange()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.Selection is Excel.Range r) return r;
            }
            catch { }
            return null;
        }
    }
}
