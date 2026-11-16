// File: ChartDuplicator.cs
// Target: C# 7.3 / .NET 4.8 / VSTO (Excel)
// Call from Ribbon: OfficePragmatics_Excel.ChartDuplicator.DuplicateSelectedChart();

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    public static class ChartDuplicator
    {
        public static void DuplicateSelectedChart()
        {
            var app = Globals.ThisAddIn?.Application;
            if (app == null) return;

            var srcCO = GetSelectedEmbeddedChartObject(app);
            if (srcCO == null)
            {
                MessageBox.Show("Select a single embedded chart on a worksheet first.", "Duplicate Chart",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (!(srcCO.Parent is Excel.Worksheet srcSheet))
            {
                MessageBox.Show("Only embedded charts on a worksheet are supported.", "Duplicate Chart",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var wb = (Excel.Workbook)srcSheet.Parent;
            var wbName = wb.Name;

            List<Excel.Worksheet> targets;
            using (var dlg = new SheetPickerForm(wb, srcSheet))
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;
                targets = dlg.SelectedSheets;
            }
            if (targets == null || targets.Count == 0)
            {
                MessageBox.Show("No destination sheets selected.", "Duplicate Chart",
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

                foreach (var dst in targets)
                    DuplicateOnce(srcCO, srcSheet, dst, wbName);

                MessageBox.Show($"Duplicated chart to {targets.Count} sheet(s).", "Duplicate Chart",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to duplicate chart:\r\n" + ex.Message, "Duplicate Chart",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                app.ScreenUpdating = prevUpdating;
                app.EnableEvents = prevEvents;
                app.DisplayAlerts = prevAlerts;
            }
        }

        private static Excel.ChartObject GetSelectedEmbeddedChartObject(Excel.Application app)
        {
            if (app.Selection is Excel.ChartObject co1) return co1;
            if (app.Selection is Excel.Chart ch && ch.Parent is Excel.ChartObject co2) return co2;
            if (app.ActiveChart?.Parent is Excel.ChartObject co3) return co3;
            return null;
        }

        private static void DuplicateOnce(Excel.ChartObject src, Excel.Worksheet srcSheet, Excel.Worksheet dstSheet, string wbName)
        {
            var dstCO = dstSheet.ChartObjects().Add(src.Left, src.Top, src.Width, src.Height) as Excel.ChartObject;

            src.Chart.ChartArea.Copy();
            dstCO.Activate();
            dstCO.Chart.Paste();

            dstCO.Left = src.Left;
            dstCO.Top = src.Top;
            dstCO.Width = src.Width;
            dstCO.Height = src.Height;
            try { dstCO.Placement = src.Placement; } catch { /* ignore */ }

            RelinkAllReferences(dstCO.Chart, wbName, srcSheet.Name, dstSheet.Name);
        }

        private static void RelinkAllReferences(Excel.Chart chart, string wbName, string fromSheetName, string toSheetName)
        {
            if (chart == null) return;

            var seriesColl = chart.SeriesCollection();
            int count = 0; try { count = seriesColl.Count; } catch { count = 0; }

            for (int i = 1; i <= count; i++)
            {
                var s = seriesColl.Item(i) as Excel.Series;
                if (s == null) continue;

                // SERIES("name", x, y, order)
                try
                {
                    string f = s.Formula;
                    if (!string.IsNullOrEmpty(f))
                    {
                        string f2 = ReplaceSheetRefs(f, wbName, fromSheetName, toSheetName);
                        if (!string.Equals(f2, f, StringComparison.Ordinal))
                            s.Formula = f2;
                    }
                }
                catch { }

                // Fallback: XValues / Values may still be string refs
                try
                {
                    var xv = s.XValues;
                    if (xv is string xs)
                    {
                        var rep = ReplaceSheetRefs(xs, wbName, fromSheetName, toSheetName);
                        if (!string.Equals(rep, xs, StringComparison.Ordinal))
                            s.XValues = rep;
                    }
                }
                catch { }
                try
                {
                    var yv = s.Values;
                    if (yv is string ys)
                    {
                        var rep = ReplaceSheetRefs(ys, wbName, fromSheetName, toSheetName);
                        if (!string.Equals(rep, ys, StringComparison.Ordinal))
                            s.Values = rep;
                    }
                }
                catch { }

                // Point data labels that are cell-linked
                try
                {
                    int ptCount = s.Points().Count;
                    for (int p = 1; p <= ptCount; p++)
                    {
                        var pt = s.Points(p);
                        try
                        {
                            var dl = pt.DataLabel;
                            if (dl != null)
                            {
                                try
                                {
                                    string lf = dl.Formula;
                                    if (!string.IsNullOrEmpty(lf))
                                    {
                                        string lf2 = ReplaceSheetRefs(lf, wbName, fromSheetName, toSheetName);
                                        if (!string.Equals(lf2, lf, StringComparison.Ordinal))
                                            dl.Formula = lf2;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                        finally { SafeRelease(pt); }
                    }
                }
                catch { }
                finally { SafeRelease(s); }
            }

            // Chart title linked to a cell
            try
            {
                if (chart.HasTitle)
                {
                    var ct = chart.ChartTitle;
                    try
                    {
                        string tf = ct.Formula;
                        if (!string.IsNullOrEmpty(tf))
                        {
                            string tf2 = ReplaceSheetRefs(tf, wbName, fromSheetName, toSheetName);
                            if (!string.Equals(tf2, tf, StringComparison.Ordinal))
                                ct.Formula = tf2;
                        }
                    }
                    catch { }
                }
            }
            catch { }

            // Axis titles linked to cells
            TryReplaceAxis(chart, Excel.XlAxisType.xlCategory, wbName, fromSheetName, toSheetName);
            TryReplaceAxis(chart, Excel.XlAxisType.xlValue, wbName, fromSheetName, toSheetName);
            TryReplaceAxis(chart, Excel.XlAxisType.xlSeriesAxis, wbName, fromSheetName, toSheetName); // 3D only
        }

        private static void TryReplaceAxis(Excel.Chart chart, Excel.XlAxisType axisType, string wbName, string fromSheet, string toSheet)
        {
            try
            {
                var axis = chart.Axes(axisType, Excel.XlAxisGroup.xlPrimary);
                if (axis != null && axis.HasTitle)
                {
                    var at = axis.AxisTitle;
                    try
                    {
                        string af = at.Formula;
                        if (!string.IsNullOrEmpty(af))
                        {
                            string af2 = ReplaceSheetRefs(af, wbName, fromSheet, toSheet);
                            if (!string.Equals(af2, af, StringComparison.Ordinal))
                                at.Formula = af2;
                        }
                    }
                    catch { }
                }
                SafeRelease(axis);
            }
            catch { }
        }

        /// <summary>
        /// Replace in-workbook refs from fromSheet → toSheet.
        /// Handles: Sheet!, 'Sheet'!, [ThisBook]Sheet!, '[ThisBook]Sheet'!.
        /// Skips other workbooks. Quotes destination if needed.
        /// </summary>
        private static string ReplaceSheetRefs(string formula, string wbName, string fromSheet, string toSheet)
        {
            if (string.IsNullOrEmpty(formula)) return formula;

            bool toNeedsQuotes = SheetNameNeedsQuotes(toSheet);
            string toQuoted = "'" + toSheet.Replace("'", "''") + "'!";
            string toUnquoted = toSheet + "!";
            string toWrite = toNeedsQuotes ? toQuoted : toUnquoted;

            string fromQuoted = "'" + fromSheet.Replace("'", "''") + "'!";
            string fromUnquoted = fromSheet + "!";

            string book = wbName; // just the file name inside [ ]
            string bFromQuoted = "'[" + book + "]" + fromSheet.Replace("'", "''") + "'!";
            string bFromUnquoted = "[" + book + "]" + fromSheet + "!";
            string bToQuoted = "'[" + book + "]" + toSheet.Replace("'", "''") + "'!";
            string bToUnquoted = "[" + book + "]" + toSheet + "!";
            string bToWrite = toNeedsQuotes ? bToQuoted : bToUnquoted;

            string result = formula;

            // Replace workbook-qualified (this workbook only)
            result = result.Replace(bFromQuoted, bToWrite);
            result = result.Replace(bFromUnquoted, bToWrite);

            // Replace plain quoted
            result = result.Replace(fromQuoted, toQuoted);

            // Replace plain unquoted → write quoted if needed; skip other workbooks
            result = ReplaceUnquotedSkippingOtherBooks(result, fromUnquoted, toWrite);

            return result;
        }

        private static bool SheetNameNeedsQuotes(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            if (char.IsDigit(name[0])) return true;
            for (int i = 0; i < name.Length; i++)
            {
                char c = name[i];
                if (!(char.IsLetterOrDigit(c) || c == '_')) return true; // space or punctuation
            }
            return false;
        }

        private static string ReplaceUnquotedSkippingOtherBooks(string text, string needle, string replacement)
        {
            if (string.IsNullOrEmpty(text) || string.IsNullOrEmpty(needle)) return text;

            var sb = new StringBuilder(text.Length + 16);
            int i = 0;
            while (i < text.Length)
            {
                int idx = text.IndexOf(needle, i, StringComparison.Ordinal);
                if (idx < 0)
                {
                    sb.Append(text, i, text.Length - i);
                    break;
                }

                // If immediately before 'needle' is ']' we are inside [OtherBook]Sheet! -> skip
                bool externalOtherBook = (idx > 0 && text[idx - 1] == ']');

                sb.Append(text, i, idx - i);
                sb.Append(externalOtherBook ? needle : replacement);
                i = idx + needle.Length;
            }
            return sb.ToString();
        }

        private static void SafeRelease(object comObj)
        {
            try { if (comObj != null && Marshal.IsComObject(comObj)) Marshal.ReleaseComObject(comObj); }
            catch { }
        }
    }

    // --- SheetPickerForm (spacious layout, .NET 4.8 safe) ---
    // --- SheetPickerForm (aligned tools row, .NET 4.8 safe) ---
    internal sealed class SheetPickerForm : Form
    {
        private readonly TextBox _filter;
        private readonly CheckedListBox _list;
        private readonly CheckBox _selectAll;
        private readonly Label _subtitle;
        private readonly Label _counter;
        private readonly Button _ok;
        private readonly Button _cancel;

        private bool _suppressSelectAll;
        private readonly List<Excel.Worksheet> _allSheets = new List<Excel.Worksheet>();
        public List<Excel.Worksheet> SelectedSheets { get; private set; } = new List<Excel.Worksheet>();

        public SheetPickerForm(Excel.Workbook wb, Excel.Worksheet exclude)
        {
            var srcName = exclude?.Name ?? "";
            Text = "Duplicate Chart – Pick Sheets";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false; ShowInTaskbar = false;
            AutoScaleMode = AutoScaleMode.Dpi;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
            MinimumSize = new System.Drawing.Size(480, 420);
            Size = new System.Drawing.Size(520, 560);

            // Root grid (adds real vertical spacing)
            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 7,
                Padding = new Padding(12)
            };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 0: subtitle
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 8));     // 1: spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));    // 2: tools row (fixed height so we can center reliably)
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 10));    // 3: spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));    // 4: list/card
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 12));    // 5: spacer
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));    // 6: buttons

            // Subtitle
            _subtitle = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Text = $"Copy the selected chart from \"{srcName}\" to:"
            };
            layout.Controls.Add(_subtitle, 0, 0);

            // Tools row: filter + (Select all + counter) — with runtime vertical centering
            var tools = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2
            };
            tools.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); // filter
            tools.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));     // right stack

            _filter = new TextBox
            {
                Dock = DockStyle.Fill,
                Margin = new Padding(0, 0, 8, 0)
            };
            Win32Edit.SetCue(_filter, "Filter sheets…");
            _filter.TextChanged += (s, e) => ApplyFilter();

            // Right panel holds checkbox + counter; we center them vertically to the filter at layout-time.
            var right = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0)
            };

            _selectAll = new CheckBox
            {
                Text = "Select all",
                ThreeState = true,
                AutoSize = true,
                Margin = new Padding(0, 0, 12, 0) // top margin will be adjusted dynamically
            };
            _selectAll.CheckStateChanged += SelectAll_CheckStateChanged;

            _counter = new Label
            {
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
                Margin = new Padding(0, 0, 0, 0)  // top margin will be adjusted dynamically
            };

            right.Controls.Add(_selectAll);
            right.Controls.Add(_counter);

            tools.Controls.Add(_filter, 0, 0);
            tools.Controls.Add(right, 1, 0);

            // On layout/resize, vertically center the right-side controls to the filter height
            void AlignTools(object _1 = null, EventArgs _2 = null)
            {
                // Actual heights after layout
                int rowH = tools.Height;                       // fixed 34 from RowStyle
                int selH = _selectAll.Height;
                int cntH = _counter.Height;
                int selTop = Math.Max(0, (rowH - selH) / 2);
                int cntTop = Math.Max(0, (rowH - cntH) / 2);

                // Apply margins without breaking horizontal spacing
                _selectAll.Margin = new Padding(_selectAll.Margin.Left, selTop, _selectAll.Margin.Right, 0);
                _counter.Margin = new Padding(_counter.Margin.Left, cntTop, _counter.Margin.Right, 0);
            }
            tools.Layout += (s, e) => AlignTools();
            tools.Resize += (s, e) => AlignTools();

            layout.Controls.Add(new Panel { Height = 8, Dock = DockStyle.Fill }, 0, 1);
            layout.Controls.Add(tools, 0, 2);
            layout.Controls.Add(new Panel { Height = 10, Dock = DockStyle.Fill }, 0, 3);

            // List inside a simple border "card"
            var card = new Panel { Dock = DockStyle.Fill, Padding = new Padding(8), BackColor = Color.White };
            card.Paint += (s, e) =>
            {
                var r = new Rectangle(0, 0, card.Width - 1, card.Height - 1);
                using (var p = new Pen(Color.FromArgb(210, 210, 210))) e.Graphics.DrawRectangle(p, r);
            };

            _list = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                IntegralHeight = false
            };
            try { _list.ItemHeight = Math.Max(_list.ItemHeight, Font.Height + 6); } catch { }
            _list.ItemCheck += (s, e) => BeginInvoke((Action)UpdateCounterAndSelectAll);
            _list.DoubleClick += (s, e) =>
            {
                var idx = _list.IndexFromPoint(_list.PointToClient(MousePosition));
                if (idx >= 0) _list.SetItemChecked(idx, !_list.GetItemChecked(idx));
            };

            card.Controls.Add(_list);
            layout.Controls.Add(card, 0, 4);
            layout.Controls.Add(new Panel { Height = 12, Dock = DockStyle.Fill }, 0, 5);

            // Buttons (right-aligned with vertical centering)
            var btns = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.RightToLeft,
                WrapContents = false,
                AutoSize = false,
                Height = 48,
                Padding = new Padding(0, 10, 0, 10)  // Center buttons vertically in the 48px row
            };
            _ok = new Button { Text = "OK", DialogResult = DialogResult.OK, Width = 96, Height = 28, Margin = new Padding(0) };
            _cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Width = 96, Height = 28, Margin = new Padding(8, 0, 0, 0) };
            btns.Controls.Add(_ok);
            btns.Controls.Add(_cancel);
            layout.Controls.Add(btns, 0, 6);

            Controls.Add(layout);
            AcceptButton = _ok;
            CancelButton = _cancel;

            // Data
            LoadSheets(wb, exclude);
            UpdateCounterAndSelectAll();
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == (Keys.Control | Keys.F)) { _filter.Focus(); _filter.SelectAll(); return true; }
            if (keyData == (Keys.Control | Keys.A))
            {
                _selectAll.CheckState = _selectAll.CheckState == CheckState.Checked ? CheckState.Unchecked : CheckState.Checked;
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void LoadSheets(Excel.Workbook wb, Excel.Worksheet exclude)
        {
            _allSheets.Clear();
            int count = wb.Worksheets.Count;
            for (int i = 1; i <= count; i++)
            {
                var ws = wb.Worksheets[i] as Excel.Worksheet;
                if (ws == null || ws == exclude) continue;
                try { if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible) continue; } catch { }
                _allSheets.Add(ws);
            }
            ApplyFilter();
        }

        private void ApplyFilter()
        {
            string q = (_filter.Text ?? "").Trim();
            var items = string.IsNullOrEmpty(q)
                ? _allSheets.ToList()
                : _allSheets.Where(s => s.Name.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            var prevChecked = new HashSet<string>(EnumerateCheckedNames(), StringComparer.OrdinalIgnoreCase);

            _list.BeginUpdate();
            _list.Items.Clear();
            foreach (var ws in items.OrderBy(w => w.Index))
            {
                int idx = _list.Items.Add(new SheetItem(ws));
                if (prevChecked.Contains(ws.Name)) _list.SetItemChecked(idx, true);
            }
            _list.EndUpdate();

            UpdateCounterAndSelectAll();
        }

        private IEnumerable<string> EnumerateCheckedNames()
        {
            foreach (var obj in _list.CheckedItems)
                yield return (obj is SheetItem si) ? si.Name : (obj?.ToString() ?? "");
        }

        private void UpdateCounterAndSelectAll()
        {
            int total = _list.Items.Count;
            int sel = _list.CheckedItems.Count;
            _counter.Text = $"{sel} selected";

            _suppressSelectAll = true;
            if (sel == 0) _selectAll.CheckState = CheckState.Unchecked;
            else if (sel == total) _selectAll.CheckState = CheckState.Checked;
            else _selectAll.CheckState = CheckState.Indeterminate;
            _suppressSelectAll = false;
        }

        private void SelectAll_CheckStateChanged(object sender, EventArgs e)
        {
            if (_suppressSelectAll) return;
            bool checkAll = _selectAll.CheckState != CheckState.Unchecked;
            if (_selectAll.CheckState == CheckState.Indeterminate) checkAll = true;
            for (int i = 0; i < _list.Items.Count; i++) _list.SetItemChecked(i, checkAll);
            UpdateCounterAndSelectAll();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (DialogResult != DialogResult.OK) return;

            var picked = new List<Excel.Worksheet>();
            for (int i = 0; i < _list.Items.Count; i++)
                if (_list.GetItemChecked(i) && _list.Items[i] is SheetItem si)
                    picked.Add(si.Worksheet);

            if (picked.Count == 0)
            {
                MessageBox.Show("Pick at least one destination sheet, or press Cancel.", "Duplicate Chart",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                e.Cancel = true; return;
            }
            SelectedSheets = picked;
        }

        private sealed class SheetItem
        {
            public Excel.Worksheet Worksheet { get; }
            public string Name { get; }
            public int Index { get; }
            public SheetItem(Excel.Worksheet ws) { Worksheet = ws; Name = ws?.Name ?? ""; Index = ws?.Index ?? 0; }
            public override string ToString() => Name;
        }
    }

    // ---- native helper for TextBox cue banner (placeholder) on .NET 4.8 ----
    internal static class Win32Edit
    {
        private const int EM_SETCUEBANNER = 0x1501;

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);

        public static void SetCue(TextBox tb, string text)
        {
            if (tb == null) return;
            if (tb.IsHandleCreated)
                SendMessage(tb.Handle, EM_SETCUEBANNER, (IntPtr)1, text);
            else
                tb.HandleCreated += (s, e) => SendMessage(tb.Handle, EM_SETCUEBANNER, (IntPtr)1, text);
        }
    }
}