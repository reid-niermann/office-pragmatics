// TraceDependent.cs – Trace Dependents with colour “breadcrumb”
// C# 7.2-compatible – paste into your VSTO project as-is

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    public static class TraceDependent
    {
        public static void Run(Excel.Application app)
        {
            var activeCell = app.ActiveCell as Excel.Range;

            // Save the current cell for RevertBackwards functionality
            RevertBackwards.SaveCurrentCell(activeCell);

            if (activeCell == null || activeCell.CountLarge > 1)
            {
                MessageBox.Show("Please select a single cell to trace.", "Trace Dependents",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Cursor.Current = Cursors.WaitCursor;
            bool scrn = app.ScreenUpdating;
            app.ScreenUpdating = false;

            List<DependentInfo> dependents = null;
            try
            {
                dependents = FindDependentsForCell(activeCell);
            }
            finally
            {
                app.ScreenUpdating = scrn;
                Cursor.Current = Cursors.Default;
            }

            if (dependents.Count == 0)
            {
                MessageBox.Show("The selected cell has no dependents in the active workbook.",
                                "Trace Dependents",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string formTitle = $"Dependents of {activeCell.Address[true, true, Excel.XlReferenceStyle.xlA1, true]}";
            var form = new DependentsForm(formTitle, dependents, activeCell);
            form.Show();
        }

        private static List<DependentInfo> FindDependentsForCell(Excel.Range sourceCell)
        {
            var dependents = new List<DependentInfo>();
            var foundCells = new HashSet<string>();

            Excel.Workbook workbook = sourceCell.Worksheet.Parent as Excel.Workbook;
            Excel.Worksheet sourceSheet = (Excel.Worksheet)sourceCell.Worksheet;
            string sourceSheetName = sourceSheet.Name;

            var sourceAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            sourceAddresses.Add(sourceCell.Address[true, true, Excel.XlReferenceStyle.xlA1, false]);
            sourceAddresses.Add(sourceCell.Address[true, false, Excel.XlReferenceStyle.xlA1, false]);
            sourceAddresses.Add(sourceCell.Address[false, true, Excel.XlReferenceStyle.xlA1, false]);
            sourceAddresses.Add(sourceCell.Address[false, false, Excel.XlReferenceStyle.xlA1, false]);

            string quotedSheetName = sourceSheetName;
            if (sourceSheetName.Contains(' ') || sourceSheetName.IndexOfAny(new[] { '!', '[', ']', '\'', '-' }) >= 0)
            {
                quotedSheetName = $"'{sourceSheetName}'";
            }
            foreach (var addr in sourceAddresses.ToList())
            {
                sourceAddresses.Add($"{quotedSheetName}!{addr}");
                sourceAddresses.Add($"{sourceSheetName}!{addr}");
            }

            Excel.Range dependentCells = null;
            try
            {
                dependentCells = sourceCell.Dependents;
                foreach (Excel.Range area in dependentCells.Areas)
                {
                    foreach (Excel.Range cell in area.Cells)
                    {
                        bool isDirectDependent = false;
                        try
                        {
                            object formulaObj = cell.Formula;
                            if (formulaObj != null && formulaObj.ToString().StartsWith("="))
                            {
                                string formula = formulaObj.ToString();
                                foreach (var sourceAddr in sourceAddresses)
                                {
                                    if (ContainsWholeWordReference(formula, sourceAddr))
                                    {
                                        isDirectDependent = true;
                                        break;
                                    }
                                }
                            }
                        }
                        catch (COMException) { continue; }

                        if (isDirectDependent)
                        {
                            string sheetName = ((Excel.Worksheet)cell.Worksheet).Name;
                            string address = cell.Address[false, false];
                            string displayText = $"'{sheetName}'!{address}";
                            string key = $"{sheetName}!{address}";
                            if (foundCells.Add(key))
                            {
                                bool hasSubDependents = CheckForDependents(cell);
                                dependents.Add(new DependentInfo(cell, displayText, hasSubDependents));
                            }
                        }
                    }
                }
            }
            catch (COMException) { /* No dependents found via built-in method */ }
            finally
            {
                if (dependentCells != null) Marshal.ReleaseComObject(dependentCells);
            }

            foreach (Excel.Worksheet worksheet in workbook.Worksheets)
            {
                try
                {
                    Excel.Range formulaCells = null;
                    try
                    {
                        formulaCells = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeFormulas);
                    }
                    catch (COMException) { continue; }

                    if (formulaCells != null)
                    {
                        foreach (Excel.Range area in formulaCells.Areas)
                        {
                            foreach (Excel.Range cell in area.Cells)
                            {
                                try
                                {
                                    object formulaObj = cell.Formula;
                                    if (formulaObj != null)
                                    {
                                        string formula = formulaObj.ToString();
                                        bool isDependent = false;
                                        if (worksheet != sourceSheet)
                                        {
                                            foreach (var addr in sourceAddresses)
                                            {
                                                if (addr.Contains("!") && ContainsWholeWordReference(formula, addr))
                                                {
                                                    isDependent = true;
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            foreach (var addr in sourceAddresses)
                                            {
                                                if (!addr.Contains("!") && ContainsWholeWordReference(formula, addr))
                                                {
                                                    isDependent = true;
                                                    break;
                                                }
                                            }
                                        }
                                        if (isDependent)
                                        {
                                            string sheetName = worksheet.Name;
                                            string address = cell.Address[false, false];
                                            string displayText = $"'{sheetName}'!{address}";
                                            string key = $"{sheetName}!{address}";
                                            if (foundCells.Add(key))
                                            {
                                                bool hasSubDependents = CheckForDependents(cell);
                                                dependents.Add(new DependentInfo(cell, displayText, hasSubDependents));
                                            }
                                        }
                                    }
                                }
                                catch (COMException) { /* Skip cells we can't read */ }
                            }
                        }
                        Marshal.ReleaseComObject(formulaCells);
                    }
                }
                catch (Exception) { /* Skip sheets we can't access */ }
            }

            Marshal.ReleaseComObject(sourceSheet);
            Marshal.ReleaseComObject(workbook);
            dependents.Sort((a, b) => a.DisplayText.CompareTo(b.DisplayText));
            return dependents;
        }

        private static bool ContainsWholeWordReference(string formula, string reference)
        {
            int index = 0;
            while ((index = formula.IndexOf(reference, index, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                bool preOK = (index == 0) || !char.IsLetterOrDigit(formula[index - 1]);
                int endIndex = index + reference.Length;
                bool postOK = (endIndex == formula.Length) || !char.IsLetterOrDigit(formula[endIndex]);
                if (preOK && postOK) return true;
                index++;
            }
            return false;
        }

        private static bool CheckForDependents(Excel.Range cell)
        {
            Excel.Range deps = null;
            try
            {
                deps = cell.Dependents;
                return true;
            }
            catch (COMException) { return false; }
            finally
            {
                if (deps != null) Marshal.ReleaseComObject(deps);
            }
        }

        public class DependentInfo
        {
            public Excel.Range Range { get; }
            public string DisplayText { get; }
            public bool HasDependents { get; }
            public bool IsSource { get; set; } = false;
            public int Level { get; set; } = 0;

            public DependentInfo(Excel.Range rng, string disp, bool hasDependents)
            {
                Range = rng;
                DisplayText = disp;
                HasDependents = hasDependents;
            }
        }

        public class DependentsForm : Form
        {
            [DllImport("user32.dll")] public static extern bool ReleaseCapture();
            [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            private struct UIColors
            {
                public static readonly Color Surface = Color.FromArgb(255, 255, 255);
                public static readonly Color SurfaceAlt = Color.FromArgb(248, 248, 248);
                public static readonly Color TitleBg = Color.FromArgb(32, 42, 102);
                public static readonly Color TitleText = Color.White;
                public static readonly Color Border = Color.FromArgb(220, 220, 220);
                public static readonly Color TextPrimary = Color.FromArgb(32, 31, 30);
                public static readonly Color Accent = Color.FromArgb(110, 125, 220);
                public static readonly Color Danger = Color.FromArgb(232, 17, 35);
            }

            private class DrillDownState
            {
                public List<DependentInfo> Items { get; }
                public Excel.Range OriginCell { get; }
                public int SelectedIndex { get; }
                public string Title { get; }
                public DrillDownState(List<DependentInfo> items, Excel.Range origin, int selectedIndex, string title)
                {
                    Items = items; OriginCell = origin; SelectedIndex = selectedIndex; Title = title;
                }
            }

            private readonly Stack<DrillDownState> _drillDownStack = new Stack<DrillDownState>();
            private readonly Panel _formulaPanel;
            private RichTextBox _currentFormulaLabel;
            private readonly Excel.Range _initialOriginCell;
            private List<DependentInfo> _dependents;
            private readonly List<DependentInfo> _shown = new List<DependentInfo>();
            private Excel.Range _origin;
            private int _cur = -1;
            private readonly ListView _list;
            // Changed INDENT to 10 for tighter spacing
            private const int INDENT = 10, FXCOL = 30;
            private readonly HashSet<Excel.Range> _unhiddenRows = new HashSet<Excel.Range>();
            private readonly HashSet<Excel.Range> _unhiddenCols = new HashSet<Excel.Range>();
            private bool _keepExposed = false;
            private Point _originalPosition = Point.Empty;
            private bool _hasMovedAway = false;
            private readonly Dictionary<Excel.Range, string> _valueCache = new Dictionary<Excel.Range, string>();
            private bool _updating = false;
            private bool _isClosing = false;
            private string _currentTitle;

            private readonly Color[] _colors = { Color.FromArgb(0, 112, 192), Color.FromArgb(0, 176, 80), Color.FromArgb(112, 48, 160), Color.FromArgb(192, 0, 0) };
            private static Color Lighten(Color c, double f = 0.7) { return Color.FromArgb(c.A, c.R + (int)((255 - c.R) * f), c.G + (int)((255 - c.G) * f), c.B + (int)((255 - c.B) * f)); }

            private Font _italicFont;
            private Font _titleFont;
            private Font _closeButtonFont;

            public DependentsForm(string title, List<DependentInfo> items, Excel.Range origin)
            {
                this._initialOriginCell = origin;
                this._dependents = items;
                this._origin = origin;
                this._currentTitle = title;

                this.FormBorderStyle = FormBorderStyle.None;
                this.DoubleBuffered = true;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.BackColor = UIColors.Border;
                this.Padding = new Padding(1);
                this.ShowIcon = false;
                this.Text = "Trace Dependents";
                this.TopMost = true;
                this.ResizeRedraw = true;

                float scaleFactor = 1.0f;
                using (Graphics g = CreateGraphics())
                {
                    float dpiX = g.DpiX;
                    float dpiY = g.DpiY;
                    scaleFactor = dpiX / 96f;
                    Size = new System.Drawing.Size((int)(3.5f * dpiX), (int)(3.5f * dpiY));
                }

                var mainPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = UIColors.Surface
                };
                mainPanel.Paint += (s, e) => { using (var p = new Pen(UIColors.Border)) e.Graphics.DrawRectangle(p, 0, 0, mainPanel.Width - 1, mainPanel.Height - 1); };
                this.Controls.Add(mainPanel);

                _titleFont = new Font("Segoe UI Semibold", 10f);
                _closeButtonFont = new Font("Segoe UI", 11f);

                int titleHeight = (int)(32 * scaleFactor);
                var pnlTitle = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = titleHeight,
                    BackColor = UIColors.TitleBg
                };
                pnlTitle.MouseDown += (s, e) =>
                {
                    if (e.Button == MouseButtons.Left) { ReleaseCapture(); SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); }
                };
                pnlTitle.Paint += (s, e) =>
                {
                    e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    using (var brush = new SolidBrush(UIColors.TitleText))
                    {
                        float y = (pnlTitle.Height - _titleFont.Height) / 2f;
                        e.Graphics.DrawString("Trace Dependents", _titleFont, brush, 10, y);
                    }
                };

                var btnClose = new Label
                {
                    Text = "✕",
                    Font = _closeButtonFont,
                    ForeColor = UIColors.TitleText,
                    Dock = DockStyle.Right,
                    Width = (int)(40 * scaleFactor),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = UIColors.TitleBg
                };
                btnClose.MouseEnter += (s, e) => { btnClose.BackColor = UIColors.Danger; btnClose.ForeColor = Color.White; };
                btnClose.MouseLeave += (s, e) => { btnClose.BackColor = UIColors.TitleBg; btnClose.ForeColor = UIColors.TitleText; };
                btnClose.Click += (s, e) => { Close(); };
                pnlTitle.Controls.Add(btnClose);
                mainPanel.Controls.Add(pnlTitle);

                KeyPreview = true;
                KeyDown += OnKeyDown;
                Resize += (s, e) => LayoutControls();
                this.FormClosing += (s, e) => { _isClosing = true; };
                FormClosed += OnFormClosed;

                _formulaPanel = new Panel
                {
                    Location = new Point(10, pnlTitle.Bottom + 10),
                    Width = mainPanel.Width - 20
                };
                _formulaPanel.Padding = new Padding(1);
                mainPanel.Controls.Add(_formulaPanel);

                _currentFormulaLabel = new RichTextBox
                {
                    Dock = DockStyle.Top,
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    Font = new Font("Segoe UI", 10),
                    Text = title,
                    HideSelection = true,
                    TabStop = false,
                    WordWrap = true,
                    ScrollBars = RichTextBoxScrollBars.None
                };
                _formulaPanel.Controls.Add(_currentFormulaLabel);

                _list = new ListView
                {
                    View = View.Details,
                    FullRowSelect = true,
                    HideSelection = false,
                    MultiSelect = false,
                    Font = new Font("Segoe UI", 10),
                    OwnerDraw = true,
                    BorderStyle = BorderStyle.FixedSingle
                };
                var doubleBufferProperty = typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
                doubleBufferProperty?.SetValue(_list, true, null);

                _italicFont = new Font(_list.Font, FontStyle.Italic);

                _list.Columns.Add("Dependent Cell");
                _list.Columns.Add("*");
                _list.Columns.Add("Value");

                _list.DrawColumnHeader += (s, e) =>
                {
                    e.Graphics.FillRectangle(new SolidBrush(UIColors.SurfaceAlt), e.Bounds);
                    TextRenderer.DrawText(e.Graphics, e.Header.Text, e.Font, e.Bounds, UIColors.TextPrimary, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
                    using (var p = new Pen(UIColors.Border)) e.Graphics.DrawLine(p, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom);
                    using (var p = new Pen(UIColors.Border)) e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
                };
                _list.DrawItem += OnDrawItem;
                _list.DrawSubItem += OnDrawSub;
                _list.SelectedIndexChanged += OnSelectedIndexChanged;

                mainPanel.Controls.Add(_list);

                BeginInvoke(new Action(() => { PopulateList(); LayoutControls(); if (_list.Items.Count > 0) { _list.Items[0].Selected = true; _cur = 0; Highlight(); } _list.Focus(); }));
            }

            private void OnDrawItem(object s, DrawListViewItemEventArgs e)
            {
                Color backColor = e.Item.BackColor;
                if (e.Item.Selected) backColor = UIColors.Accent;
                using (var brush = new SolidBrush(backColor)) e.Graphics.FillRectangle(brush, e.Bounds);
            }

            private void OnDrawSub(object s, DrawListViewSubItemEventArgs e)
            {
                var info = (DependentInfo)e.Item.Tag;
                if (info == null) { e.DrawDefault = true; return; }
                Color txt = e.SubItem.ForeColor;
                if (e.Item.Selected) txt = Color.White;

                switch (e.ColumnIndex)
                {
                    case 0: // Dependent Cell
                        int startLevel = info.IsSource ? 1 : 0;
                        using (var pen = new Pen(Color.Silver))
                        {
                            for (int l = startLevel; l <= info.Level; l++)
                            {
                                int x = e.Bounds.Left + 6 + l * INDENT;
                                e.Graphics.DrawLine(pen, x, e.Bounds.Top, x, e.Bounds.Bottom - 1);
                            }
                        }

                        int textOffset = 6 + info.Level * INDENT;
                        var rect = new Rectangle(e.Bounds.Left + textOffset, e.Bounds.Top, e.Bounds.Width - textOffset, e.Bounds.Height);
                        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.Item.Font, rect, txt, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                        break;
                    case 1: // Arrow
                        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, e.Bounds, txt, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                        break;
                    case 2: // Value
                        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, e.Bounds, txt, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                        break;
                    default:
                        e.DrawDefault = true;
                        break;
                }
            }

            private void OnKeyDown(object s, KeyEventArgs e)
            {
                if (_shown.Count == 0 && e.KeyCode != Keys.Escape) return;
                switch (e.KeyCode)
                {
                    case Keys.Escape: _keepExposed = false; Close(); break;
                    case Keys.Enter: DrillDown(true); break;
                    case Keys.Down: _cur = (_cur + 1) % _shown.Count; Highlight(); break;
                    case Keys.Up: _cur = (_cur - 1 + _shown.Count) % _shown.Count; Highlight(); break;
                    case Keys.Right: DrillDown(false); break;
                    case Keys.Left: DrillUp(); break;
                    default: return;
                }
                e.Handled = true;
            }

            private void DrillDown(bool closeOnNoDependents)
            {
                if (_cur < 0 || _cur >= _shown.Count) return;
                var itemToDrill = _shown[_cur];

                if (itemToDrill.IsSource) return;

                if (!itemToDrill.HasDependents)
                {
                    if (closeOnNoDependents) { _keepExposed = true; Close(); }
                    else { System.Media.SystemSounds.Asterisk.Play(); }
                    return;
                }
                var targetCell = itemToDrill.Range;
                var app = (Excel.Application)_origin.Application;
                Cursor.Current = Cursors.WaitCursor;
                bool scrn = app.ScreenUpdating;
                app.ScreenUpdating = false;
                List<DependentInfo> newDependents = null;
                try { newDependents = FindDependentsForCell(targetCell); }
                finally { app.ScreenUpdating = scrn; Cursor.Current = Cursors.Default; }

                _drillDownStack.Push(new DrillDownState(_dependents, _origin, _cur, _currentTitle));
                _dependents = newDependents;
                _origin = targetCell;
                _currentTitle = $"Dependents of {targetCell.Address[true, true, Excel.XlReferenceStyle.xlA1, true]}";

                _currentFormulaLabel.Text = _currentTitle;
                PopulateList();
                LayoutControls();
                _cur = 0;
                Highlight();
                _list.Focus();
            }

            private void DrillUp()
            {
                if (_drillDownStack.Count == 0) return;
                var prevState = _drillDownStack.Pop();
                _dependents = prevState.Items;
                _origin = prevState.OriginCell;
                _cur = prevState.SelectedIndex;
                _currentTitle = prevState.Title;

                _currentFormulaLabel.Text = _currentTitle;
                PopulateList();
                LayoutControls();
                Highlight();
                _list.Focus();
            }

            private void OnFormClosed(object sender, FormClosedEventArgs e)
            {
                _italicFont?.Dispose();
                _titleFont?.Dispose();
                _closeButtonFont?.Dispose();
                if (!_keepExposed) { RestoreHidden(); try { SafeSelect(_initialOriginCell); } catch { } }
            }

            private void OnSelectedIndexChanged(object sender, EventArgs e) { if (_updating) return; if (_list.SelectedIndices.Count > 0) { _cur = _list.SelectedIndices[0]; Highlight(); } }

            private void LayoutControls()
            {
                int rtbHeight = Math.Max(26, Math.Min(_currentFormulaLabel.GetPositionFromCharIndex(_currentFormulaLabel.TextLength - 1).Y + _currentFormulaLabel.Font.Height + 10, 150));
                _currentFormulaLabel.Height = rtbHeight;
                _formulaPanel.Height = rtbHeight;

                _list.Location = new Point(10, _formulaPanel.Bottom + 10);
                _list.Size = new System.Drawing.Size(ClientSize.Width - 20, ClientSize.Height - _formulaPanel.Bottom - 20);

                int total = _list.ClientSize.Width; _list.Columns[1].Width = FXCOL;
                _list.Columns[0].Width = (int)(total * 0.65);
                _list.Columns[2].Width = total - _list.Columns[0].Width - FXCOL;
            }

            private void PopulateList()
            {
                _list.BeginUpdate();
                _list.Items.Clear();
                _shown.Clear();

                // 1. Add Source Item
                var sourceInfo = new DependentInfo(_origin, _origin.Address[false, false], false) { IsSource = true, Level = 0 };
                var sourceItem = new ListViewItem(sourceInfo.DisplayText) { Tag = sourceInfo, UseItemStyleForSubItems = false, ForeColor = UIColors.TextPrimary };
                sourceItem.SubItems.Add(new ListViewItem.ListViewSubItem(sourceItem, "") { ForeColor = Color.Gray });
                sourceItem.SubItems.Add(new ListViewItem.ListViewSubItem(sourceItem, GetFormattedValue(_origin)));
                _list.Items.Add(sourceItem);
                _shown.Add(sourceInfo);

                // 2. Add Dependents
                foreach (var dep in _dependents)
                {
                    dep.Level = 1;
                    var item = new ListViewItem(dep.DisplayText) { Tag = dep, UseItemStyleForSubItems = false };
                    string arrow = dep.HasDependents ? "↪" : "";
                    var arrowSub = new ListViewItem.ListViewSubItem(item, arrow) { ForeColor = Color.Gray };
                    item.SubItems.Add(arrowSub);
                    var valSub = new ListViewItem.ListViewSubItem(item, GetFormattedValue(dep.Range));
                    item.SubItems.Add(valSub);
                    _list.Items.Add(item);
                    _shown.Add(dep);
                }
                _list.EndUpdate();
            }

            private void RestoreHidden()
            {
                foreach (var r in _unhiddenRows) { try { r.Hidden = true; } catch { } }
                foreach (var c in _unhiddenCols) { try { c.Hidden = true; } catch { } }
                _unhiddenRows.Clear();
                _unhiddenCols.Clear();
            }

            private string GetFormattedValue(Excel.Range rng)
            {
                if (rng == null) return ""; if (_valueCache.TryGetValue(rng, out string cached)) return cached;
                Excel.Range first = null; object rawValue = null; string rawFmt = "";
                try { first = (Excel.Range)rng.Cells[1, 1]; rawFmt = first.NumberFormat as string ?? ""; rawValue = first.Value; }
                catch { }
                finally { if (first != null) Marshal.ReleaseComObject(first); }
                string cleanFmt = rawFmt.Replace("*", ""); if (rawValue == null) { _valueCache[rng] = ""; return ""; }
                try { var app = (Excel.Application)rng.Application; string result = app.WorksheetFunction.Text(rawValue, cleanFmt).Trim(); if (_valueCache.Count < 500) _valueCache[rng] = result; return result; }
                catch { try { string result = rawValue.ToString().Trim(); if (_valueCache.Count < 500) _valueCache[rng] = result; return result; } catch { return ""; } }
            }

            private void Highlight()
            {
                if (_cur < 0 || _cur >= _shown.Count) return;
                _updating = true;
                RestoreHidden();

                _list.Items[_cur].Selected = true;
                _list.EnsureVisible(_cur);
                _updating = false;

                var sel = _shown[_cur];

                if (sel.IsSource)
                {
                    _currentFormulaLabel.Text = _currentTitle;
                    _currentFormulaLabel.SelectAll();
                    _currentFormulaLabel.SelectionColor = Color.Black;
                    _currentFormulaLabel.SelectionBackColor = Color.White;
                    _currentFormulaLabel.SelectionFont = new Font(_currentFormulaLabel.Font, FontStyle.Regular);
                    _currentFormulaLabel.DeselectAll();
                    LayoutControls();

                    try { UnhideForRange(sel.Range); } catch { }
                    BeginInvoke((MethodInvoker)(() => { if (_isClosing) return; try { SafeSelect(sel.Range); AutoMoveFormAwayFromRange(sel.Range); this.Activate(); } catch { } }));
                    return;
                }

                if (sel.Range != null)
                {
                    string formula = "";
                    try { var formulaObj = sel.Range.Formula; formula = formulaObj?.ToString() ?? ""; if (!formula.StartsWith("=")) { formula = sel.Range.Text?.ToString() ?? ""; } }
                    catch (COMException) { try { formula = sel.Range.Text?.ToString() ?? "Error reading cell"; } catch { formula = "Error reading cell"; } }

                    string cellAddress = sel.Range.Address[false, false, Excel.XlReferenceStyle.xlA1, false];
                    string fullText = formula.StartsWith("=") ? $"{cellAddress}{formula}" : formula;

                    _currentFormulaLabel.Text = fullText;
                    LayoutControls();

                    _currentFormulaLabel.SelectAll();
                    _currentFormulaLabel.SelectionColor = Color.Black;
                    _currentFormulaLabel.SelectionBackColor = Color.White;

                    if (formula.StartsWith("="))
                    {
                        _currentFormulaLabel.Select(0, cellAddress.Length);
                        _currentFormulaLabel.SelectionFont = new Font(_currentFormulaLabel.Font, FontStyle.Underline);

                        Excel.Worksheet dependentSheet = null;
                        Excel.Worksheet originSheet = null;
                        try
                        {
                            dependentSheet = (Excel.Worksheet)sel.Range.Worksheet;
                            originSheet = (Excel.Worksheet)_origin.Worksheet;
                            string[] addressesToSearch;
                            if (!dependentSheet.Equals(originSheet))
                            {
                                string originSheetName = originSheet.Name;
                                if (originSheetName.Contains(' ') || originSheetName.IndexOfAny(new[] { '!', '[', ']', '\'', '-' }) >= 0) { originSheetName = $"'{originSheetName}'"; }
                                string prefix = $"{originSheetName}!";
                                addressesToSearch = new[] { prefix + _origin.Address[true, true], prefix + _origin.Address[true, false], prefix + _origin.Address[false, true], prefix + _origin.Address[false, false] };
                            }
                            else { addressesToSearch = new[] { _origin.Address[true, true], _origin.Address[true, false], _origin.Address[false, true], _origin.Address[false, false] }; }

                            Color baseColor = _colors[_drillDownStack.Count % _colors.Length];
                            Color highlightColor = Lighten(baseColor, 0.85);

                            foreach (var term in addressesToSearch.Distinct().OrderByDescending(s => s.Length))
                            {
                                if (string.IsNullOrEmpty(term)) continue;
                                int startIndex = cellAddress.Length;
                                while ((startIndex = fullText.IndexOf(term, startIndex, StringComparison.InvariantCultureIgnoreCase)) != -1)
                                {
                                    bool preOK = (startIndex == 0) || !char.IsLetterOrDigit(fullText[startIndex - 1]);
                                    int endIndex = startIndex + term.Length;
                                    bool postOK = (endIndex == fullText.Length) || !char.IsLetterOrDigit(fullText[endIndex]);
                                    if (preOK && postOK)
                                    {
                                        _currentFormulaLabel.Select(startIndex, term.Length);
                                        _currentFormulaLabel.SelectionColor = baseColor;
                                        _currentFormulaLabel.SelectionBackColor = highlightColor;
                                        break;
                                    }
                                    startIndex += term.Length;
                                }
                            }
                        }
                        finally { if (dependentSheet != null) Marshal.ReleaseComObject(dependentSheet); if (originSheet != null) Marshal.ReleaseComObject(originSheet); }
                    }
                    _currentFormulaLabel.DeselectAll();

                    try { UnhideForRange(sel.Range); } catch { }
                    BeginInvoke((MethodInvoker)(() => { if (_isClosing) return; try { SafeSelect(sel.Range); AutoMoveFormAwayFromRange(sel.Range); this.Activate(); } catch { } }));
                }
            }

            private void UnhideForRange(Excel.Range rng) { if (rng == null) return; if (rng.Address == rng.EntireColumn.Address) { var ec = rng.EntireColumn; if (ec.Hidden) { ec.Hidden = false; _unhiddenCols.Add(ec); } return; } if (rng.Address == rng.EntireRow.Address) { var er = rng.EntireRow; if (er.Hidden) { er.Hidden = false; _unhiddenRows.Add(er); } return; } foreach (Excel.Range row in rng.Rows) { var er = row.EntireRow; if (!er.Hidden && !_unhiddenRows.Contains(er)) continue; try { er.Hidden = false; _unhiddenRows.Add(er); } catch { } } foreach (Excel.Range col in rng.Columns) { var ec = col.EntireColumn; if (!ec.Hidden && !_unhiddenCols.Contains(ec)) continue; try { ec.Hidden = false; _unhiddenCols.Add(ec); } catch { } } }

            private void AutoMoveFormAwayFromRange(Excel.Range selectedRange)
            {
                Rectangle formBounds = this.Bounds;
                Rectangle rangeScreenCoords = GetRangeScreenCoordinates(selectedRange);
                if (!formBounds.IntersectsWith(rangeScreenCoords))
                {
                    if (_hasMovedAway && _originalPosition != Point.Empty)
                    {
                        Rectangle originalBounds = new Rectangle(_originalPosition, this.Size);
                        if (!originalBounds.IntersectsWith(rangeScreenCoords)) { this.Location = _originalPosition; _hasMovedAway = false; }
                    }
                    return;
                }
                if (!_hasMovedAway) _originalPosition = this.Location;
                Rectangle workingArea = Screen.FromControl(this).WorkingArea;
                Point[] candidatePositions = new Point[] { new Point(rangeScreenCoords.Right + 10, this.Location.Y), new Point(rangeScreenCoords.Left - this.Width - 10, this.Location.Y), new Point(this.Location.X, rangeScreenCoords.Bottom + 10), new Point(this.Location.X, rangeScreenCoords.Top - this.Height - 10), new Point(rangeScreenCoords.Right + 10, rangeScreenCoords.Top), new Point(rangeScreenCoords.Left - this.Width - 10, rangeScreenCoords.Top) };
                Point bestPosition = this.Location; double minDistance = double.MaxValue;
                foreach (Point candidatePos in candidatePositions)
                {
                    Rectangle potentialBounds = new Rectangle(candidatePos, this.Size);
                    if (!workingArea.Contains(potentialBounds) || potentialBounds.IntersectsWith(rangeScreenCoords)) continue;
                    double distance = Math.Sqrt(Math.Pow(candidatePos.X - this.Location.X, 2) + Math.Pow(candidatePos.Y - this.Location.Y, 2));
                    if (distance < minDistance) { minDistance = distance; bestPosition = candidatePos; }
                }
                if (bestPosition != this.Location) { this.Location = bestPosition; _hasMovedAway = true; }
            }

            private Rectangle GetRangeScreenCoordinates(Excel.Range range) { if (range == null) return Rectangle.Empty; try { var app = (Excel.Application)range.Application; var window = app.ActiveWindow; double left = (double)range.Left, top = (double)range.Top, width = (double)range.Width, height = (double)range.Height; int screenLeft = window.ActivePane.PointsToScreenPixelsX((int)left); int screenTop = window.ActivePane.PointsToScreenPixelsY((int)top); int screenRight = window.ActivePane.PointsToScreenPixelsX((int)(left + width)); int screenBottom = window.ActivePane.PointsToScreenPixelsY((int)(top + height)); return new Rectangle(screenLeft, screenTop, screenRight - screenLeft, screenBottom - screenTop); } catch { return Rectangle.Empty; } }

            private static void SafeSelect(Excel.Range rng)
            {
                var app = (Excel.Application)rng.Application;
                bool scrn = app.ScreenUpdating;
                bool evts = app.EnableEvents;
                try
                {
                    app.ScreenUpdating = false;
                    app.EnableEvents = false;
                    Excel.Range target = rng;
                    target.Worksheet.Activate();
                    app.Goto(target, false);
                }
                finally
                {
                    app.EnableEvents = evts;
                    app.ScreenUpdating = scrn;
                }
            }
        }
    }
}