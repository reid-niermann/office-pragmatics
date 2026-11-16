using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    internal static class OtherFormatting
    {
        // ───────── Underline Cycle ─────────
        private static readonly Excel.XlUnderlineStyle[] _underlineCycle =
        {
            Excel.XlUnderlineStyle.xlUnderlineStyleSingle,
            Excel.XlUnderlineStyle.xlUnderlineStyleSingleAccounting,
            Excel.XlUnderlineStyle.xlUnderlineStyleDouble,
            Excel.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting,
            Excel.XlUnderlineStyle.xlUnderlineStyleNone
        };

        public static void RunUnderlineCycle()
        {
            var sel = GetSelection();
            if (sel == null) return;

            var first = sel.Cells[1, 1] as Excel.Range;
            int current = Array.IndexOf(_underlineCycle, (Excel.XlUnderlineStyle)first.Font.Underline);
            int next = current == -1 ? 0 : (current + 1) % _underlineCycle.Length;

            foreach (Excel.Range c in sel.Cells)
                c.Font.Underline = _underlineCycle[next];
        }

        // ─────────────────────────  Border Cycle  ─────────────────────────

        // ─────────────────────────  Border‑TYPE Cycle  ─────────────────────────
        private static readonly Excel.XlBordersIndex[] _allEdges =
        {
            Excel.XlBordersIndex.xlEdgeLeft,
            Excel.XlBordersIndex.xlEdgeTop,
            Excel.XlBordersIndex.xlEdgeBottom,
            Excel.XlBordersIndex.xlEdgeRight,
            Excel.XlBordersIndex.xlInsideVertical,
            Excel.XlBordersIndex.xlInsideHorizontal
        };

        // order: Bottom → Top → Left → Right → All → None
        private static readonly int[] _borderCodes = { 1, 2, 4, 8, 15, 0 };
        private static readonly Excel.XlBordersIndex[][] _borderTypeCycle =
        {
            new[]{ Excel.XlBordersIndex.xlEdgeBottom },                         // bottom
            new[]{ Excel.XlBordersIndex.xlEdgeTop },                            // top
            new[]{ Excel.XlBordersIndex.xlEdgeLeft },                           // left
            new[]{ Excel.XlBordersIndex.xlEdgeRight },                          // right
            new[]{ Excel.XlBordersIndex.xlEdgeLeft,  Excel.XlBordersIndex.xlEdgeTop,
                    Excel.XlBordersIndex.xlEdgeRight, Excel.XlBordersIndex.xlEdgeBottom,
                    Excel.XlBordersIndex.xlInsideVertical, Excel.XlBordersIndex.xlInsideHorizontal }, // all
            new Excel.XlBordersIndex[0]                                          // none
        };

        // ─────────────────────────  Indicator Form  ─────────────────────────
        private static readonly string[] _borderLabels =
            { "Bottom Border", "Top Border", "Left Border", "Right Border", "All Borders", "No Borders" };
        private static Form _borderForm;
        private static int _borderToHighlight;

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        private struct RECT { public int Left, Top, Right, Bottom; }

        /// <summary>
        /// Cycles through your 6 border‑styles on the current selection,
        /// then pops up a 3‑second “Applied: …” indicator.
        /// </summary>
        public static void RunBorderCycle()
        {
            var sel = GetSelection();        // any range OK
            if (sel == null) return;

            // 1) determine current code
            int code = 0;
            var b = sel.Borders;
            if (IsOn(b, Excel.XlBordersIndex.xlEdgeBottom)) code |= 1;
            if (IsOn(b, Excel.XlBordersIndex.xlEdgeTop)) code |= 2;
            if (IsOn(b, Excel.XlBordersIndex.xlEdgeLeft)) code |= 4;
            if (IsOn(b, Excel.XlBordersIndex.xlEdgeRight)) code |= 8;

            int idx = Array.IndexOf(_borderCodes, code);
            int next = (idx == -1) ? 0 : (idx + 1) % _borderCodes.Length;

            // 2) clear all borders
            foreach (var edge in _allEdges)
                b[edge].LineStyle = Excel.XlLineStyle.xlLineStyleNone;

            // 3) apply next pattern
            foreach (var edge in _borderTypeCycle[next])
            {
                b[edge].LineStyle = Excel.XlLineStyle.xlContinuous;
                b[edge].Weight = Excel.XlBorderWeight.xlThin;
            }

            // 4) show the form indicator
            ShowBorderIndicator(next);
        }

        // helper to check if a given edge is “on”
        private static bool IsOn(Excel.Borders borders, Excel.XlBordersIndex edge) =>
            (Excel.XlLineStyle)borders[edge].LineStyle != Excel.XlLineStyle.xlLineStyleNone;

        // display a small form saying “Applied: …”
        private static void ShowBorderIndicator(int appliedIndex)
        {
            _borderToHighlight = appliedIndex;

            if (_borderForm == null)
            {
                _borderForm = new Form
                {
                    FormBorderStyle = FormBorderStyle.None,
                    StartPosition = FormStartPosition.Manual,
                    Size = new System.Drawing.Size(250, 50),   // 50px tall
                    BackColor = Color.Black,
                    Opacity = 0.7,
                    TopMost = true,
                    ShowInTaskbar = false
                };
                _borderForm.Paint += BorderForm_Paint;
            }

            // position at Excel’s bottom‑right with padding
            IntPtr excelHwnd = new IntPtr(Globals.ThisAddIn.Application.Hwnd);
            if (GetWindowRect(excelHwnd, out RECT rect))
            {
                const int rightPadding = 20;
                const int bottomPadding = 8;
                int x = rect.Right - _borderForm.Width - rightPadding;
                int y = rect.Bottom - _borderForm.Height - bottomPadding;
                _borderForm.Location = new Point(x, y);
            }
            else
            {
                var bounds = Screen.PrimaryScreen.Bounds;
                _borderForm.Location = new Point(
                    bounds.Width - _borderForm.Width - 20,
                    bounds.Height - _borderForm.Height - 20);
            }

            _borderForm.Show();

            // auto‑hide after 3 seconds
            var t = new Timer { Interval = 3000 };
            t.Tick += (s, e) =>
            {
                t.Stop();
                t.Dispose();
                _borderForm.Hide();
            };
            t.Start();
        }

        // draws the “Applied: …” text with C# 7‑style using blocks
        private static void BorderForm_Paint(object sender, PaintEventArgs e)
        {
            using (var font = new Font("Segoe UI", 10, FontStyle.Bold))
            using (var brush = new SolidBrush(Color.White))
            {
                string text = $"Applied: {_borderLabels[_borderToHighlight]}";
                float y = (_borderForm.Height - font.Height) / 2;
                e.Graphics.DrawString(text, font, brush, new PointF(10, y));
            }
        }

        // stub: retrieves the current Excel range selection
        private static Excel.Range GetSelection() =>
            Globals.ThisAddIn.Application.Selection as Excel.Range;

        // ─────────────────────────  Font-Size Cycle  ─────────────────────────
        private static readonly double[] _fontSizes = { 10, 18, 7 };

        public static void RunFontSizeCycle()
        {
            var sel = GetSelection();
            if (sel == null) return;

            double curSize = ((Excel.Range)sel.Cells[1, 1]).Font.Size;
            int idx = Array.IndexOf(_fontSizes, curSize);
            int next = (idx == -1) ? 0 : (idx + 1) % _fontSizes.Length;

            foreach (Excel.Range c in sel.Cells)
                c.Font.Size = _fontSizes[next];
        }

        // ─────────────────────────  Row-Height Cycle  ────────────────────────
        private static readonly double[] _rowHeights = { 5, 20, -1 /*reset*/ };

        public static void RunRowHeightCycle()
        {
            var sel = GetSelection();           // any rows/columns fine
            if (sel == null) return;

            var firstRow = (Excel.Range)sel.Rows[1];
            double cur = firstRow.RowHeight;

            int idx = -1;
            for (int i = 0; i < _rowHeights.Length; i++)
                if (_rowHeights[i] == cur || (_rowHeights[i] == -1 && IsDefaultRowHeight(firstRow)))
                { idx = i; break; }

            int next = (idx == -1) ? 0 : (idx + 1) % _rowHeights.Length;

            foreach (Excel.Range row in sel.Rows)
            {
                if (_rowHeights[next] == -1)
                    row.RowHeight = row.Worksheet.StandardHeight;          // reset
                else
                    row.RowHeight = _rowHeights[next];
            }
        }

        private static bool IsDefaultRowHeight(Excel.Range r) =>
            Math.Abs(r.RowHeight - r.Worksheet.StandardHeight) < 0.1;

        // ─────────────────────────  Column-Width Cycle  ──────────────────────
        private static readonly double[] _colWidths = { 12, 2, 25, -1 /*reset*/ };

        public static void RunColumnWidthCycle()
        {
            var sel = GetSelection();           // any selection okay
            if (sel == null) return;

            var firstCol = (Excel.Range)sel.Columns[1];
            double cur = firstCol.ColumnWidth;

            int idx = -1;
            for (int i = 0; i < _colWidths.Length; i++)
                if (_colWidths[i] == cur || (_colWidths[i] == -1 && IsDefaultColWidth(firstCol)))
                { idx = i; break; }

            int next = (idx == -1) ? 0 : (idx + 1) % _colWidths.Length;

            foreach (Excel.Range col in sel.Columns)
            {
                if (_colWidths[next] == -1)
                    col.ColumnWidth = col.Worksheet.StandardWidth;         // reset
                else
                    col.ColumnWidth = _colWidths[next];
            }
        }

        private static bool IsDefaultColWidth(Excel.Range c) =>
            Math.Abs(c.ColumnWidth - c.Worksheet.StandardWidth) < 0.1;


        // ───────── Footnotes ─────────
        private static readonly Regex _footRx = new Regex(@"\((\d+)\)$", RegexOptions.Compiled);

        public static void RunFootnoteCycle()
        {
            var sel = GetSelection();
            if (sel == null) return;

            foreach (Excel.Range cell in sel.Cells)
            {
                string txt = cell.Value2 as string ?? string.Empty;
                Match m = _footRx.Match(txt);
                int n = m.Success ? int.Parse(m.Groups[1].Value) + 1 : 1;

                if (m.Success) txt = txt.Substring(0, m.Index);     // strip old

                string add = $"({n})";
                cell.Value2 = txt + add;

                int start = txt.Length + 1;                         // Characters is 1-based
                cell.Characters[start, add.Length].Font.Superscript = true;
            }
        }

        public static void RunFootnoteRemove()
        {
            var sel = GetSelection();
            if (sel == null) return;

            foreach (Excel.Range cell in sel.Cells)
            {
                string txt = cell.Value2 as string ?? string.Empty;
                Match m = _footRx.Match(txt);
                if (m.Success) cell.Value2 = txt.Substring(0, m.Index);
            }
        }

        // ───────── List / Bullets (single-column) ─────────
        private enum ListType { Dash, Bullet, Number, UpperLetter, LowerLetter, UpperRoman, LowerRoman, None }
        private const char Bullet = '•';

        public static void RunListBullets()
        {
            var sel = GetSelection(columnsOnly: true);
            if (sel == null) return;

            // Figure out the current cycle state from the first *non-blank* cell
            ListType current = ListType.None;
            foreach (Excel.Range c in sel.Cells)
            {
                string txt = (c.Value2 as string ?? "").TrimStart();
                if (txt.Length == 0) continue;              // skip blanks
                current = DetectListType(c);
                break;
            }
            ListType next = (ListType)(((int)current + 1) %
                                       Enum.GetNames(typeof(ListType)).Length);

            // First strip any existing prefixes (works on all cells)
            foreach (Excel.Range c in sel.Cells)
                c.Value2 = StripPrefix(c.Value2 as string ?? "");

            if (next == ListType.None) return;             // that cycle step is “remove list”

            // Re-apply prefixes, but only to non-blank cells – numbering ignores blanks
            int seq = 0;
            foreach (Excel.Range c in sel.Cells)
            {
                string body = c.Value2 as string ?? "";
                if (body.Trim().Length == 0) continue;     // don’t touch blank rows

                ++seq;
                c.Value2 = GetPrefix(next, seq) + body;
            }
        }

        private static ListType DetectListType(Excel.Range cell)
        {
            string t = (cell.Value2 as string ?? string.Empty).TrimStart();
            if (t.StartsWith("- ")) return ListType.Dash;
            if (t.StartsWith($"{Bullet} ")) return ListType.Bullet;
            if (Regex.IsMatch(t, @"^\d+\. ")) return ListType.Number;
            if (Regex.IsMatch(t, @"^[A-Z]\. ")) return ListType.UpperLetter;
            if (Regex.IsMatch(t, @"^[a-z]\. ")) return ListType.LowerLetter;
            if (Regex.IsMatch(t, @"^[IVXLCDM]+\. ")) return ListType.UpperRoman;
            if (Regex.IsMatch(t, @"^[ivxlcdm]+\. ")) return ListType.LowerRoman;
            return ListType.None;
        }

        private static string StripPrefix(string s) =>
            Regex.Replace(s, @"^(- |• |\d+\. |[A-Z]\. |[a-z]\. |[IVXLCDM]+\. |[ivxlcdm]+\. )", "");

        private static string GetPrefix(ListType t, int idx)
        {
            switch (t)
            {
                case ListType.Dash: return "- ";
                case ListType.Bullet: return $"{Bullet} ";
                case ListType.Number: return $"{idx}. ";
                case ListType.UpperLetter: return $"{(char)('A' + (idx - 1) % 26)}. ";
                case ListType.LowerLetter: return $"{char.ToLower((char)('A' + (idx - 1) % 26))}. ";
                case ListType.UpperRoman: return $"{ToRoman(idx).ToUpper()}. ";
                case ListType.LowerRoman: return $"{ToRoman(idx).ToLower()}. ";
                default: return string.Empty;
            }
        }

        private static string ToRoman(int n)
        {
            if (n < 1 || n > 3999) return n.ToString();
            (int v, string r)[] map = {
                (1000,"M"), (900,"CM"), (500,"D"), (400,"CD"),
                (100,"C"),  (90,"XC"),  (50,"L"),  (40,"XL"),
                (10,"X"),   (9,"IX"),   (5,"V"),   (4,"IV"), (1,"I")
            };

            string res = "";
            foreach (var pair in map)
            {
                int v = pair.v;
                string r = pair.r;
                while (n >= v) { res += r; n -= v; }
            }
            return res;
        }

        // ───────── Accounting Dots (@*.) ─────────
        private const string AccDotsFmt = "@*.";

        public static void RunAccountingDots()
        {
            var sel = GetSelection();
            if (sel == null) return;

            string cur = sel.Cells[1, 1].NumberFormat as string ?? "";
            sel.NumberFormat = cur == AccDotsFmt ? "General" : AccDotsFmt;
        }

        // ───────── Sum-Line Toggle ─────────
        public static void RunSumLine()
        {
            var sel = GetSelection();                 // no row-count restriction
            if (sel == null) return;

            // Detect “active” state from the *first selected row*
            var firstRow = (Excel.Range)sel.Rows[1];
            var firstCell = firstRow.Cells[1, 1] as Excel.Range;

            bool isActive =
                ((Excel.XlUnderlineStyle)firstCell.Font.Underline ==
                     Excel.XlUnderlineStyle.xlUnderlineStyleSingleAccounting) &&
                ((firstCell.Value2 as string) == " ");

            foreach (Excel.Range row in sel.Rows)
            {
                Excel.Range cell = row.Cells[1, 1];
                if (!isActive)
                {
                    cell.Value2 = " ";
                    cell.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingleAccounting;
                    row.RowHeight = 5;                       // thin “sum” line
                }
                else
                {
                    cell.Value2 = "";
                    cell.Font.Underline = Excel.XlUnderlineStyle.xlUnderlineStyleNone;

                    // restore to the row above (or default height if first row)
                    double h = row.Row > 1
                        ? row.Worksheet.Rows[row.Row - 1].RowHeight
                        : row.Worksheet.StandardHeight;
                    row.RowHeight = h;
                }
            }
        }

        // ─────────────────────────  Indent  ─────────────────────────
        private const int MaxIndent = 15;   // Excel hard-limit

        /// <summary>Increases the IndentLevel of every cell in the selection (max 15).</summary>
        public static void RunIndentIncrease()
        {
            var sel = GetSelection();
            if (sel == null) return;

            foreach (Excel.Range c in sel.Cells)
            {
                int lvl = c.IndentLevel;
                if (lvl < MaxIndent) c.IndentLevel = lvl + 1;
            }
        }

        /// <summary>Decreases the IndentLevel of every cell in the selection (down to 0).</summary>
        public static void RunIndentDecrease()
        {
            var sel = GetSelection();
            if (sel == null) return;

            foreach (Excel.Range c in sel.Cells)
            {
                int lvl = c.IndentLevel;
                if (lvl > 0) c.IndentLevel = lvl - 1;
            }
        }


        // ───────── Helpers ─────────
        private static Excel.Range GetSelection(bool columnsOnly = false, bool rowsOnly = false)
        {
            var app = Globals.ThisAddIn.Application;
            if (!(app.Selection is Excel.Range sel))
            {
                System.Windows.Forms.MessageBox.Show(
                    "Please make a range selection first.", "Selection Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return null;
            }

            if (columnsOnly && sel.Columns.Count != 1)
            {
                System.Windows.Forms.MessageBox.Show(
                    "This function works on a single column only.", "Selection Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return null;
            }

            if (rowsOnly && sel.Rows.Count != 1)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Please select exactly one row.", "Selection Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return null;
            }

            return sel;
        }
    }
}