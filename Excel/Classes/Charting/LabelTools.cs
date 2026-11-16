//  LabelTools.cs
//  OfficePragmatics – Excel VSTO add-in (C# 7.2, .NET Framework)

using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;          //  MsoTriState enum, etc.

namespace OfficePragmatics_Excel.Classes.Charting
{
    internal static class LabelTools
    {
        /*────────────────────────────  PUBLIC RIBBON API  ──────────────────────────*/

        public static void KeyboardControlChartLabels()
        {
            if (_moverForm == null || _moverForm.IsDisposed)
                _moverForm = new ChartLabelMoverForm();

            _moverForm.Show();
            _moverForm.BringToFront();
        }

        public static void ResetLabels()                                    // AutoText on
        {
            var app = Globals.ThisAddIn.Application;
            var chart = app.ActiveChart as Excel.Chart;
            if (chart == null) { MessageBox.Show("Please select a chart first."); return; }

            foreach (Excel.Series ser in chart.SeriesCollection())
            {
                try
                {
                    var dls = ser.DataLabels();
                    if (dls != null) dls.AutoText = true;
                    Marshal.ReleaseComObject(dls);
                }
                catch { /* ignore */ }
                finally { Marshal.ReleaseComObject(ser); }
            }
        }

        public static void SuperscriptFootnotes()
        {
            var app = Globals.ThisAddIn.Application;
            var chart = app.ActiveChart as Excel.Chart;
            if (chart == null) { MessageBox.Show("Please select a chart first."); return; }

            foreach (Excel.Series ser in chart.SeriesCollection())
            {
                foreach (Excel.Point pt in ser.Points())
                {
                    Excel.DataLabel lbl = null;
                    try { lbl = pt.DataLabel; } catch { }

                    if (lbl != null)
                    {
                        string txt = lbl.Text as string;
                        if (!string.IsNullOrEmpty(txt))
                        {
                            for (int j = 1; j <= txt.Length; j++)
                            {
                                char c = txt[j - 1];
                                if ((c >= 'a' && c <= 'w') || c == '(' || c == ')')
                                {
                                    var ch = lbl.Characters[j, 1];
                                    ch.Font.Superscript = true;
                                    Marshal.ReleaseComObject(ch);
                                }
                            }
                        }
                        Marshal.ReleaseComObject(lbl);
                    }
                    Marshal.ReleaseComObject(pt);
                }
                Marshal.ReleaseComObject(ser);
            }
        }

        public static void LabelPoints()                                    // open picker
        {
            if (_pointsForm == null || _pointsForm.IsDisposed)
                _pointsForm = new LabelPointsForm();

            _pointsForm.Show();
            _pointsForm.BringToFront();
        }

        /*─────────────────────────────  PRIVATE CONSTANTS  ─────────────────────────*/

        private const double MoveIncrement = 1.0;           // points per arrow-key nudge

        /*───────────────────────────────  NUDGE ENGINE  ────────────────────────────*/

        private static void NudgeSelectedChartObject(int dTop, int dLeft)
        {
            var app = Globals.ThisAddIn.Application;
            if (app.ActiveChart == null) { MessageBox.Show("Please select a chart."); return; }

            var sel = app.Selection;
            if (sel == null) { MessageBox.Show("Nothing selected."); return; }

            // strongly-typed special-cases
            if (sel is Excel.DataLabel dl) { Move(dl, dTop, dLeft); return; }
            else if (sel is Excel.Legend lg) { Move(lg, dTop, dLeft); return; }
            else if (sel is Excel.AxisTitle at) { Move(at, dTop, dLeft); return; }
            else if (sel is Excel.ChartTitle ct) { Move(ct, dTop, dLeft); return; }
            else if (sel is Excel.PlotArea pa) { Move(pa, dTop, dLeft); return; }

            // DataLabels collection
            if (sel is Excel.DataLabels dls)
            {
                foreach (Excel.DataLabel lbl in dls)
                { Move(lbl, dTop, dLeft); Marshal.ReleaseComObject(lbl); }
                return;
            }

            // ShapeRange
            if (sel is Excel.ShapeRange sr)
            {
                foreach (Excel.Shape shp in sr)
                { Move(shp, dTop, dLeft); Marshal.ReleaseComObject(shp); }
                return;
            }

            // generic COM object fallback
            if (TryNudgeDynamic(sel, dTop, dLeft)) return;

            MessageBox.Show("The selected object cannot be moved with arrow keys.");
        }

        private static void Move(dynamic o, int dT, int dL)
        {
            o.Top = (double)o.Top + dT * MoveIncrement;
            o.Left = (double)o.Left + dL * MoveIncrement;
        }

        private static bool TryNudgeDynamic(object o, int dT, int dL)
        {
            try
            {
                dynamic d = o;
                d.Top = (double)d.Top + dT * MoveIncrement;
                d.Left = (double)d.Left + dL * MoveIncrement;
                return true;
            }
            catch { return false; }
        }

        /*───────────────  ALIGN  |  REMOVE LEADERS  |  RESET POS  ──────────────────*/

        private static void AlignSelectedChartObject(string anchor)
        {
            var app = Globals.ThisAddIn.Application;
            var cht = app.ActiveChart as Excel.Chart;
            if (cht == null) { MessageBox.Show("Select a chart first."); return; }

            var sel = app.Selection;
            if (sel == null) { MessageBox.Show("Nothing selected."); return; }

            // helper to align one object relative to plot area
            void AlignOne(dynamic d)
            {
                var pa = cht.PlotArea;
                double w = d.Width, h = d.Height;

                switch (anchor)
                {
                    case "TopMiddle":
                        d.Top = 0;
                        d.Left = (pa.Width - w) / 2.0;
                        break;
                    case "MiddleRight":
                        d.Top = (pa.Height - h) / 2.0;
                        d.Left = pa.Width - w;
                        break;
                    case "BottomMiddle":
                        d.Top = pa.Height - h;
                        d.Left = (pa.Width - w) / 2.0;
                        break;
                    case "MiddleLeft":
                        d.Top = (pa.Height - h) / 2.0;
                        d.Left = 0;
                        break;
                }
            }

            try
            {
                if (sel is Excel.DataLabels dls)
                {
                    foreach (Excel.DataLabel l in dls) { AlignOne(l); Marshal.ReleaseComObject(l); }
                    return;
                }
                if (sel is Excel.ShapeRange sr)
                {
                    foreach (Excel.Shape s in sr) { AlignOne(s); Marshal.ReleaseComObject(s); }
                    return;
                }
                AlignOne(sel);                    // single object
            }
            catch
            {
                MessageBox.Show("Selected object cannot be aligned by this tool.");
            }
        }

        /*───────────────────────────  LABEL-POINTS ENGINE  ─────────────────────────*/

        private static void ApplyLabelPoints(bool first, bool last, bool max, bool min, string xValue)
        {
            var app = Globals.ThisAddIn.Application;
            var chart = app.ActiveChart as Excel.Chart;
            if (chart == null) { MessageBox.Show("Select a chart first."); return; }

            var allowed = new[]
            {
                Excel.XlChartType.xlLine, Excel.XlChartType.xlLineMarkers,
                Excel.XlChartType.xlLineStacked, Excel.XlChartType.xlLineStacked100,
                Excel.XlChartType.xlLineMarkersStacked, Excel.XlChartType.xlLineMarkersStacked100,
                Excel.XlChartType.xlColumnClustered, Excel.XlChartType.xlColumnStacked,
                Excel.XlChartType.xlBarClustered,    Excel.XlChartType.xlBarStacked
            };
            if (!allowed.Contains(chart.ChartType))
            {
                MessageBox.Show("This tool works only with line or column/bar charts.");
                return;
            }

            foreach (Excel.Series ser in chart.SeriesCollection())
            {
                int ptCount = ser.Points().Count;
                if (ptCount == 0) { Marshal.ReleaseComObject(ser); continue; }

                /* ----- locate max / min ----------------------------------------------------- */
                int maxIdx = 1, minIdx = 1;
                double maxVal = double.MinValue, minVal = double.MaxValue;

                var valsArr = ser.Values as Array;
                if (valsArr != null)
                {
                    int lb = valsArr.GetLowerBound(0);
                    int ub = valsArr.GetUpperBound(0);

                    for (int i = lb; i <= ub; i++)
                    {
                        object vObj = valsArr.GetValue(i);
                        if (vObj == null || !double.TryParse(vObj.ToString(), out double v)) continue;

                        if (v > maxVal) { maxVal = v; maxIdx = i - lb + 1; }   // convert to 1-based point index
                        if (v < minVal) { minVal = v; minIdx = i - lb + 1; }
                    }
                }

                /* ----- locate specific X matches --------------------------------------------- */
                int[] xMatches = Array.Empty<int>();
                bool useX = !string.IsNullOrWhiteSpace(xValue);

                if (useX)
                {
                    var xArr = ser.XValues as Array;
                    if (xArr != null)
                    {
                        int lb = xArr.GetLowerBound(0);
                        int ub = xArr.GetUpperBound(0);

                        xMatches = Enumerable.Range(lb, ub - lb + 1)
                                             .Where(i => IsXMatch(xArr.GetValue(i), xValue))
                                             .Select(i => i - lb + 1)          // convert to 1-based point index
                                             .ToArray();
                    }
                }

                /* ----- helper to label a single point --------------------------------------- */
                void Label(int idx)
                {
                    if (idx < 1 || idx > ptCount) return;
                    try
                    {
                        var pt = ser.Points(idx);

                        // safest: ApplyDataLabels adds + sets .HasDataLabel
                        try { pt.ApplyDataLabels(); }
                        catch { pt.HasDataLabel = true; }

                        // ensure auto-text (value) is shown
                        try { pt.DataLabel.AutoText = true; } catch { }

                        Marshal.ReleaseComObject(pt);
                    }
                    catch { /* ignore */ }
                }

                if (first) Label(1);
                if (last) Label(ptCount);
                if (max) Label(maxIdx);
                if (min) Label(minIdx);
                foreach (var i in xMatches) Label(i);

                Marshal.ReleaseComObject(ser);
            }
        }

        // Helpers
        private static bool TryGetOADate(object o, out double oa)
        {
            oa = 0;
            if (o == null) return false;

            // already a double coming from Excel (dates are serial numbers)
            if (o is double d) { oa = d; return true; }

            // a genuine DateTime instance
            if (o is DateTime dt) { oa = dt.ToOADate(); return true; }

            // fall back to parsing a string such as "1/1/23" or "01-01-2023"
            if (DateTime.TryParse(o.ToString(), out DateTime parsed))
            {
                oa = parsed.ToOADate();
                return true;
            }
            return false;
        }
        private static bool IsXMatch(object xObj, string target)
        {
            // first, try matching as Excel date serials
            if (TryGetOADate(xObj, out double dx) && TryGetOADate(target, out double dt))
                return Math.Abs(dx - dt) < 0.0001;          // dates are whole-day serials

            // numeric compare for non-date numbers
            if (double.TryParse(xObj?.ToString(), out double nx) &&
                double.TryParse(target, out double nt))
                return Math.Abs(nx - nt) < 1e-9;

            // finally, plain case-insensitive text match
            return string.Equals(xObj?.ToString(), target, StringComparison.OrdinalIgnoreCase);
        }

        /*───────────────────────────────  UI FORMS  ────────────────────────────────*/

        private static ChartLabelMoverForm _moverForm;
        private static LabelPointsForm _pointsForm;

        /* ――― 1. CHART-LABEL MOVER  (arrow/align/reset/leaders) ―――――――――――― */

        private sealed class ChartLabelMoverForm : Form
        {
            public ChartLabelMoverForm()
            {
                Text = "Chart-Label Tools";
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                MaximizeBox = MinimizeBox = false;
                Size = new System.Drawing.Size(500, 500);
                KeyPreview = true;

                var lbl = new Label
                {
                    Text = "Arrow keys (↑ ↓ ← →) nudge the\ncurrently selected chart object.",
                    AutoSize = true,
                    Top = 10,
                    Left = 10
                };
                Controls.Add(lbl);

                int y = lbl.Bottom + 12;

                MakeAlignButton("Top", 10, y, "TopMiddle");
                MakeAlignButton("Right", 60, y, "MiddleRight");
                MakeAlignButton("Bottom", 110, y, "BottomMiddle");
                MakeAlignButton("Left", 160, y, "MiddleLeft");
            }

            private void MakeAlignButton(string caption, int l, int t, string anchor)
            {
                var b = new Button { Text = caption, Width = 45, Left = l, Top = t };
                b.Click += (s, e) => AlignSelectedChartObject(anchor);
                Controls.Add(b);
            }

            protected override bool ProcessCmdKey(ref Message m, Keys keyData)
            {
                switch (keyData)
                {
                    case Keys.Up: NudgeSelectedChartObject(-1, 0); return true;
                    case Keys.Down: NudgeSelectedChartObject(1, 0); return true;
                    case Keys.Left: NudgeSelectedChartObject(0, -1); return true;
                    case Keys.Right: NudgeSelectedChartObject(0, 1); return true;
                }
                return base.ProcessCmdKey(ref m, keyData);
            }
        }

        /* ――― 2. LABEL-POINTS PICKER FORM ――――――――――――――――――――――――――――――― */

        private sealed class LabelPointsForm : Form
        {
            private readonly CheckBox _chkFirst = new CheckBox { Text = "First", AutoSize = true };
            private readonly CheckBox _chkLast = new CheckBox { Text = "Last", AutoSize = true };
            private readonly CheckBox _chkMax = new CheckBox { Text = "Maximum", AutoSize = true };
            private readonly CheckBox _chkMin = new CheckBox { Text = "Minimum", AutoSize = true };

            private readonly TextBox _txtX = new TextBox { Width = 200 };
            private readonly Button _btnApply = new Button { Text = "Apply", Width = 200 };

            public LabelPointsForm()
            {
                Text = "Label Points";
                FormBorderStyle = FormBorderStyle.FixedToolWindow;
                MaximizeBox = MinimizeBox = false;
                Width = 320; Height = 260;
                StartPosition = FormStartPosition.CenterScreen;

                int y = 18;
                foreach (var cb in new[] { _chkFirst, _chkLast, _chkMax, _chkMin })
                { cb.Left = 20; cb.Top = y; Controls.Add(cb); y += 28; }

                var lblX = new Label { Text = "At X value:", Left = 20, Top = y + 4, AutoSize = true };
                _txtX.Left = 90; _txtX.Top = y;
                Controls.Add(lblX);
                Controls.Add(_txtX);

                _btnApply.Left = 110;
                _btnApply.Top = y + 40;
                Controls.Add(_btnApply);

                _btnApply.Click += (s, e) =>
                {
                    bool any = _chkFirst.Checked || _chkLast.Checked || _chkMax.Checked || _chkMin.Checked;
                    string xVal = _txtX.Text.Trim();

                    if (!any && string.IsNullOrEmpty(xVal))
                    {
                        MessageBox.Show("Select a point type or enter an X value.");
                        return;
                    }

                    ApplyLabelPoints(_chkFirst.Checked, _chkLast.Checked,
                                     _chkMax.Checked, _chkMin.Checked,
                                     xVal);
                };
            }
        }
    }
}