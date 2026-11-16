// File: ChartOffsetDuplicator.cs
// Target: C# 7.3 / .NET 4.8 / VSTO (Excel)
// Ribbon call: OfficePragmatics_Excel.ChartOffsetDuplicator.DuplicateSelectedChartByOffset();

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    public static class ChartOffsetDuplicator
    {
        public static void DuplicateSelectedChartByOffset()
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return;

            var srcCO = GetSelectedEmbeddedChartObject(app);
            if (srcCO == null)
            {
                MessageBox.Show("Select a single embedded chart on a worksheet first.", "Duplicate by Offset",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!(srcCO.Parent is Excel.Worksheet ws))
            {
                MessageBox.Show("Only embedded charts on a worksheet are supported.", "Duplicate by Offset",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            // Ask for inputs
            int colOffset = 0, rowOffset = 1, repetitions = 1;
            bool keepXStatic = true;

            using (var dlg = new OffsetDialog())
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                colOffset = dlg.ColOffset;
                rowOffset = dlg.RowOffset;
                repetitions = dlg.Repetitions;
                keepXStatic = dlg.StaticX;
            }

            if (repetitions <= 0)
            {
                MessageBox.Show("Number of repetitions must be at least 1.", "Duplicate by Offset",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var prevUpdating = app.ScreenUpdating;
            var prevEvents = app.EnableEvents;
            var prevAlerts = app.DisplayAlerts;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.DisplayAlerts = false;

                for (int i = 1; i <= repetitions; i++)
                {
                    // Copy chart formatting/content
                    var newCO = ws.ChartObjects().Add(srcCO.Left, srcCO.Top + i * (srcCO.Height + 12), srcCO.Width, srcCO.Height) as Excel.ChartObject;

                    srcCO.Chart.ChartArea.Copy();
                    newCO.Activate();
                    newCO.Chart.Paste();

                    try { newCO.Placement = srcCO.Placement; } catch { /* best-effort */ }

                    // Relink ranges with offsets
                    ApplySeriesOffsets(app, newCO.Chart, ws, rowOffset * i, colOffset * i, keepXStatic);
                }

                MessageBox.Show($"Created {repetitions} offset chart(s).", "Duplicate by Offset",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to duplicate chart:\r\n" + ex.Message, "Duplicate by Offset",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = prevUpdating;
                app.EnableEvents = prevEvents;
                app.DisplayAlerts = prevAlerts;
            }
        }

        // ---- Core offset logic --------------------------------------------------

        private static void ApplySeriesOffsets(Excel.Application app, Excel.Chart chart, Excel.Worksheet ws, int dRows, int dCols, bool keepXStatic)
        {
            if (chart == null) return;

            Excel.SeriesCollection sc = null;
            try { sc = chart.SeriesCollection(); } catch { }
            if (sc == null) return;

            int n = 0;
            try { n = sc.Count; } catch { n = 0; }

            for (int i = 1; i <= n; i++)
            {
                Excel.Series s = null;
                try { s = sc.Item(i) as Excel.Series; } catch { s = null; }
                if (s == null) continue;

                try
                {
                    // Parse the SERIES() formula to get name / x / y / (optional bubble sizes)
                    if (!TryParseSeriesFormula(s, out var nameExpr, out var xExpr, out var yExpr, out var sizeExpr))
                    {
                        // If we cannot parse, try a simpler path using XValues/Values directly
                        OffsetDirect(app, s, ws, dRows, dCols, keepXStatic);
                        continue;
                    }

                    // Y values (always offset)
                    var yR = TryGetRangeFromRef(app, yExpr);
                    if (yR != null)
                    {
                        var yNew = yR.Offset[dRows, dCols];
                        s.Values = yNew;
                        SafeRelease(yNew); SafeRelease(yR);
                    }

                    // X values (maybe static)
                    if (!keepXStatic && !string.IsNullOrWhiteSpace(xExpr))
                    {
                        var xR = TryGetRangeFromRef(app, xExpr);
                        if (xR != null)
                        {
                            var xNew = xR.Offset[dRows, dCols];
                            s.XValues = xNew;
                            SafeRelease(xNew); SafeRelease(xR);
                        }
                    }

                    // Series name: if it was a cell reference, move it alongside Y
                    if (!string.IsNullOrWhiteSpace(nameExpr) && !IsQuotedLiteral(nameExpr))
                    {
                        var nmR = TryGetRangeFromRef(app, nameExpr);
                        if (nmR != null)
                        {
                            var nmNew = nmR.Offset[dRows, dCols];
                            var addr = nmNew.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
                            s.Name = "=" + addr; // keep as a cell link
                            SafeRelease(nmNew); SafeRelease(nmR);
                        }
                    }

                    // Bubble sizes (if present) move with Y
                    if (!string.IsNullOrWhiteSpace(sizeExpr))
                    {
                        var szR = TryGetRangeFromRef(app, sizeExpr);
                        if (szR != null)
                        {
                            var szNew = szR.Offset[dRows, dCols];
                            try { s.BubbleSizes = szNew; } catch { /* non-bubble types */ }
                            SafeRelease(szNew); SafeRelease(szR);
                        }
                    }
                }
                catch
                {
                    // Fallback to direct method if any part fails
                    try { OffsetDirect(app, s, ws, dRows, dCols, keepXStatic); } catch { }
                }
                finally
                {
                    SafeRelease(s);
                }
            }
        }

        // Fallback: attempt to shift X/Y by using current XValues/Values as ranges if possible.
        private static void OffsetDirect(Excel.Application app, Excel.Series s, Excel.Worksheet ws, int dRows, int dCols, bool keepXStatic)
        {
            // Y first
            try
            {
                var yRef = GetRangeFromSeriesValuesRef(app, s, takeX: false);
                if (yRef != null)
                {
                    var yNew = yRef.Offset[dRows, dCols];
                    s.Values = yNew;
                    SafeRelease(yNew); SafeRelease(yRef);
                }
            }
            catch { }

            // X if not static
            if (!keepXStatic)
            {
                try
                {
                    var xRef = GetRangeFromSeriesValuesRef(app, s, takeX: true);
                    if (xRef != null)
                    {
                        var xNew = xRef.Offset[dRows, dCols];
                        s.XValues = xNew;
                        SafeRelease(xNew); SafeRelease(xRef);
                    }
                }
                catch { }
            }

            // Series name: if it's a ref, move alongside Y
            try
            {
                var nameExpr = ExtractNameExpr(s);
                if (!string.IsNullOrWhiteSpace(nameExpr) && !IsQuotedLiteral(nameExpr))
                {
                    var nmR = TryGetRangeFromRef(app, nameExpr);
                    if (nmR != null)
                    {
                        var nmNew = nmR.Offset[dRows, dCols];
                        var addr = nmNew.Address[true, true, Excel.XlReferenceStyle.xlA1, true];
                        s.Name = "=" + addr;
                        SafeRelease(nmNew); SafeRelease(nmR);
                    }
                }
            }
            catch { }
        }

        // ---- Helpers: selection, parsing, address/offset ------------------------

        private static Excel.ChartObject GetSelectedEmbeddedChartObject(Excel.Application app)
        {
            if (app.Selection is Excel.ChartObject co1) return co1;
            if (app.Selection is Excel.Chart ch && ch.Parent is Excel.ChartObject co2) return co2;
            if (app.ActiveChart?.Parent is Excel.ChartObject co3) return co3;
            return null;
        }

        private static bool TryParseSeriesFormula(Excel.Series s,
            out string nameExpr, out string xExpr, out string yExpr, out string sizeExpr)
        {
            nameExpr = xExpr = yExpr = sizeExpr = null;
            string f;
            try { f = s.Formula; } catch { return false; }
            if (string.IsNullOrWhiteSpace(f)) return false;
            // Expect "=SERIES(name, x, y, order[, sizes])"
            int p0 = f.IndexOf('('), p1 = f.LastIndexOf(')');
            if (p0 < 0 || p1 <= p0) return false;

            var inside = f.Substring(p0 + 1, p1 - p0 - 1).Trim();
            var parts = SplitCsvRespectingQuotes(inside);

            if (parts.Count < 4) return false;

            nameExpr = parts[0].Trim();
            xExpr = parts[1].Trim();
            yExpr = parts[2].Trim();
            // parts[3] is order — ignore
            if (parts.Count >= 5) sizeExpr = parts[4].Trim();

            return true;
        }

        private static List<string> SplitCsvRespectingQuotes(string s)
        {
            var list = new List<string>(5);
            var sb = new StringBuilder();
            bool inQuotes = false;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '"')
                {
                    inQuotes = !inQuotes;
                    sb.Append(c);
                }
                else if (c == ',' && !inQuotes)
                {
                    list.Add(sb.ToString());
                    sb.Clear();
                }
                else
                {
                    sb.Append(c);
                }
            }
            list.Add(sb.ToString());
            return list;
        }

        private static bool IsQuotedLiteral(string expr)
        {
            expr = (expr ?? "").Trim();
            return expr.Length >= 2 && expr[0] == '"' && expr[expr.Length - 1] == '"';
        }

        private static Excel.Range TryGetRangeFromRef(Excel.Application app, string refText)
        {
            if (string.IsNullOrWhiteSpace(refText)) return null;
            // Strip leading '=' if present
            var t = refText.Trim();
            if (t.StartsWith("=")) t = t.Substring(1);
            try
            {
                var o = app.Evaluate(t);
                return o as Excel.Range;
            }
            catch { return null; }
        }

        private static Excel.Range GetRangeFromSeriesValuesRef(Excel.Application app, Excel.Series s, bool takeX)
        {
            try
            {
                // Prefer the parse path: pull the arg we need and Evaluate it
                if (TryParseSeriesFormula(s, out var nameExpr, out var xExpr, out var yExpr, out var sizeExpr))
                {
                    string expr = takeX ? xExpr : yExpr;
                    return TryGetRangeFromRef(app, expr);
                }
            }
            catch { }
            return null;
        }

        private static string ExtractNameExpr(Excel.Series s)
        {
            try
            {
                if (TryParseSeriesFormula(s, out var nameExpr, out _, out _, out _))
                    return nameExpr;
            }
            catch { }
            return null;
        }

        private static void SafeRelease(object comObj)
        {
            try { if (comObj != null && Marshal.IsComObject(comObj)) Marshal.ReleaseComObject(comObj); }
            catch { }
        }

        // ---- Simple input dialog ------------------------------------------------

        // Nested inside ChartOffsetDuplicator (or place in its own file/namespace)
        private sealed class OffsetDialog : Form
        {
            private NumericUpDown numCol, numRow, numRep;
            private CheckBox chkStaticX;
            private Button btnOK, btnCancel;

            public int ColOffset => (int)numCol.Value;
            public int RowOffset => (int)numRow.Value;
            public int Repetitions => (int)numRep.Value;
            public bool StaticX => chkStaticX.Checked;

            public OffsetDialog()
            {
                Text = "Duplicate Chart – Offset Data";
                StartPosition = FormStartPosition.CenterParent;
                FormBorderStyle = FormBorderStyle.FixedDialog;
                MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
                AutoScaleMode = AutoScaleMode.Dpi;

                // Ensure buttons aren’t clipped on high DPI
                ClientSize = new System.Drawing.Size(460, 300);
                MinimumSize = new System.Drawing.Size(460, 300);

                Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);

                var layout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    ColumnCount = 2,
                    RowCount = 5,
                    Padding = new Padding(12)
                };
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
                layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 0: column offset
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 1: row offset
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 2: repetitions
                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // 3: checkbox
                layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 56)); // 4: buttons row (taller)

                // Inputs
                var lblCol = new Label { Text = "Column offset (ΔC):", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 6) };
                numCol = new NumericUpDown
                {
                    Minimum = -100000,
                    Maximum = 100000,
                    Value = 0,
                    DecimalPlaces = 0,
                    Increment = 1,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right
                };

                var lblRow = new Label { Text = "Row offset (ΔR):", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 6) };
                numRow = new NumericUpDown
                {
                    Minimum = -100000,
                    Maximum = 100000,
                    Value = 1,
                    DecimalPlaces = 0,
                    Increment = 1,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right
                };

                var lblRep = new Label { Text = "Repetitions (count):", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 6, 6) };
                numRep = new NumericUpDown
                {
                    Minimum = 1,
                    Maximum = 10000,
                    Value = 3,
                    DecimalPlaces = 0,
                    Increment = 1,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right
                };

                chkStaticX = new CheckBox
                {
                    Text = "Keep X-axis static (don’t offset X)",
                    Checked = true,
                    AutoSize = true,
                    Anchor = AnchorStyles.Left,
                    Margin = new Padding(0, 10, 0, 10)
                };

                layout.Controls.Add(lblCol, 0, 0); layout.Controls.Add(numCol, 1, 0);
                layout.Controls.Add(lblRow, 0, 1); layout.Controls.Add(numRow, 1, 1);
                layout.Controls.Add(lblRep, 0, 2); layout.Controls.Add(numRep, 1, 2);
                layout.Controls.Add(chkStaticX, 0, 3); layout.SetColumnSpan(chkStaticX, 2);

                // Buttons
                var pnlBtns = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft,
                    WrapContents = false
                };
                btnOK = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 96, Height = 28 };
                btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 96, Height = 28, Margin = new Padding(8, 0, 0, 0) };
                pnlBtns.Controls.Add(btnOK);
                pnlBtns.Controls.Add(btnCancel);

                layout.Controls.Add(pnlBtns, 0, 4);
                layout.SetColumnSpan(pnlBtns, 2);

                Controls.Add(layout);

                AcceptButton = btnOK;
                CancelButton = btnCancel;
                KeyPreview = true; // Esc works via CancelButton, Enter via AcceptButton
            }
        }
    }
}
