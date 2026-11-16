using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    internal static class SheetSelector
    {
        public static void Show(Excel.Application app)
        {
            if (app == null)
            {
                MessageBox.Show("Excel application not available.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            using (var picker = new SheetSelectForm(app))
            {
                var result = picker.ShowDialog();
                try { app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }

                if (result == DialogResult.OK)
                {
                    var selectedNames = picker.SelectedSheetNames;

                    if (selectedNames.Count == 0) return;

                    try
                    {
                        if (selectedNames.Count == 1)
                        {
                            var sheet = app.Worksheets[selectedNames[0]] as Excel.Worksheet;
                            if (sheet != null)
                            {
                                if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                                sheet.Select(true);
                            }
                        }
                        else
                        {
                            bool isFirst = true;
                            foreach (string name in selectedNames)
                            {
                                var sheet = app.Worksheets[name] as Excel.Worksheet;
                                if (sheet != null)
                                {
                                    if (sheet.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                                        sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                                    sheet.Select(isFirst ? true : false);
                                    isFirst = false;
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Could not activate sheet(s): {ex.Message}",
                            "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
    }

    public class SheetSelectForm : Form
    {
        // ── P/Invoke ─────────────────────────────────────────────────────────────
        [DllImport("user32.dll")] public static extern bool ReleaseCapture();
        [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")] private static extern int GetDpiForWindow(IntPtr hWnd);

        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;
        private const int WM_DPICHANGED = 0x02E0;

        // ── Metrics ──────────────────────────────────────────────────────────────
        private const double FORM_WIDTH_IN = 3.2;
        private const double FORM_HEIGHT_IN = 4.0;
        private const double HEADER_H_IN = 0.45;
        private const double FOOTER_H_IN = 0.36;
        private const double LIST_ITEM_H_IN = 0.28;
        private const double TITLE_BAR_H_IN = 0.32;

        // Fonts
        private const double FONT_BODY_IN = 9.0 / 72.0;

        // MATCHED FONT SIZE (Now same as Body)
        private const double FONT_SEARCH_IN = 9.0 / 72.0;

        private const double FONT_SMALL_IN = 8.5 / 72.0;

        // ── Colors ───────────────────────────────────────────────────────────────
        private static readonly Color TitleBgColor = Color.FromArgb(32, 42, 102);
        private static readonly Color TitleTextColor = Color.White;
        private static readonly Color DangerColor = Color.FromArgb(232, 17, 35);

        private readonly Color BackgroundColor = Color.FromArgb(255, 255, 255);
        private readonly Color ControlBgColor = Color.FromArgb(243, 243, 243);
        private readonly Color BorderColor = Color.FromArgb(180, 180, 180);

        private readonly Color TextColor = Color.FromArgb(30, 30, 30);
        private readonly Color SubTextColor = Color.FromArgb(100, 100, 100);
        private readonly Color HintColor = Color.FromArgb(160, 160, 160);

        private readonly Color SelectionBgColor = Color.FromArgb(230, 242, 250);

        private readonly Color IconExcelGreen = Color.FromArgb(33, 115, 70);
        private readonly Color IconHiddenGray = Color.FromArgb(160, 160, 160);

        // ── State ────────────────────────────────────────────────────────────────
        private int _lastAppliedDpiX = 0;
        private int _lastAppliedDpiY = 0;
        private int _dragSourceIndex = -1;
        private Point _dragStart;
        private bool _isDragging;
        private bool _isRenaming = false;

        private readonly Excel.Application _app;
        private readonly string _activeSheetName;
        private readonly HashSet<string> _preSelectedSheets = new HashSet<string>();
        private List<SheetItem> _allItems;

        // ── Controls ─────────────────────────────────────────────────────────────
        private Panel _mainPanel;
        private Panel _titleBarPanel;
        private Label _closeButton;
        private Panel _headerPanel;
        private Panel _searchContainer;
        private TextBox _searchBox;
        private Panel _listContainerPanel;
        private ListBox _listBox;
        private Panel _footerPanel;
        private CheckBox _showHiddenToggle;
        private Label _tipsLabel;
        private ToolTip _toolTip;
        private TextBox _editBox;

        private class SheetItem
        {
            public string Name;
            public bool Hidden;
            public override string ToString() => Name;
        }

        public List<string> SelectedSheetNames
        {
            get
            {
                var list = new List<string>();
                foreach (SheetItem item in _listBox.SelectedItems)
                {
                    if (item != null) list.Add(item.Name);
                }
                return list;
            }
        }

        public SheetSelectForm(Excel.Application app)
        {
            if (app == null) throw new ArgumentNullException(nameof(app));
            _app = app;
            _activeSheetName = app.ActiveSheet?.Name ?? "";

            try
            {
                if (app.ActiveWindow != null && app.ActiveWindow.SelectedSheets != null)
                {
                    foreach (object s in app.ActiveWindow.SelectedSheets)
                    {
                        if (s is Excel.Worksheet ws) _preSelectedSheets.Add(ws.Name);
                    }
                }
            }
            catch { }
            if (_preSelectedSheets.Count == 0 && !string.IsNullOrEmpty(_activeSheetName))
                _preSelectedSheets.Add(_activeSheetName);

            InitializeForm();
            CreateControls();
            PopulateSheets();
            RefreshList();
            SelectInitialSheets();

            this.Load += (s, e) => { _searchBox?.Focus(); };
        }

        private void InitializeForm()
        {
            AutoScaleMode = AutoScaleMode.None;
            ShowIcon = false;
            Text = "Activate Worksheet";
            KeyPreview = true;
            BackColor = Color.White;
            FormBorderStyle = FormBorderStyle.None;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
        }

        private void CreateControls()
        {
            // 1. Outer Frame
            _mainPanel = new Panel { Dock = DockStyle.Fill, BackColor = BackgroundColor, Padding = new Padding(1) };
            _mainPanel.Paint += (s, e) => {
                using (var p = new Pen(BorderColor))
                    e.Graphics.DrawRectangle(p, 0, 0, _mainPanel.Width - 1, _mainPanel.Height - 1);
            };
            this.Controls.Add(_mainPanel);

            // 2. Title Bar
            _titleBarPanel = new Panel { Dock = DockStyle.Top, BackColor = TitleBgColor, Height = 32 };
            _titleBarPanel.MouseDown += (s, e) => {
                if (e.Button == MouseButtons.Left) { ReleaseCapture(); SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); }
            };
            _titleBarPanel.Paint += (s, e) => {
                e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                using (var titleFont = new Font("Segoe UI Semibold", 10f))
                using (var brush = new SolidBrush(TitleTextColor))
                {
                    float y = (_titleBarPanel.Height - titleFont.Height) / 2f;
                    e.Graphics.DrawString("Activate Worksheet", titleFont, brush, 12, y);
                }
            };

            _closeButton = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 10f),
                ForeColor = TitleTextColor,
                Dock = DockStyle.Right,
                Width = 45,
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                BackColor = TitleBgColor
            };
            _closeButton.MouseEnter += (s, e) => _closeButton.BackColor = DangerColor;
            _closeButton.MouseLeave += (s, e) => _closeButton.BackColor = TitleBgColor;
            _closeButton.Click += (s, e) => { DialogResult = DialogResult.Cancel; Close(); };
            _titleBarPanel.Controls.Add(_closeButton);

            // 3. Header
            _headerPanel = new Panel { Dock = DockStyle.Top, BackColor = ControlBgColor };
            _searchContainer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                Padding = new Padding(5, 3, 2, 2)
            };
            _searchContainer.Paint += (s, e) => {
                ControlPaint.DrawBorder(e.Graphics, _searchContainer.ClientRectangle, BorderColor, ButtonBorderStyle.Solid);
            };
            _searchBox = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                ForeColor = HintColor,
                Text = "Search...",
                BackColor = Color.White
            };
            _searchBox.GotFocus += (s, e) => {
                if (_searchBox.ForeColor == HintColor) { _searchBox.Text = ""; _searchBox.ForeColor = TextColor; }
            };
            _searchBox.LostFocus += (s, e) => {
                if (string.IsNullOrWhiteSpace(_searchBox.Text)) { _searchBox.Text = "Search..."; _searchBox.ForeColor = HintColor; }
            };
            _searchBox.TextChanged += (s, e) => RefreshList();
            _searchBox.KeyDown += CommonKeyDown;
            _searchContainer.Controls.Add(_searchBox);
            _headerPanel.Controls.Add(_searchContainer);

            // 4. Footer
            _footerPanel = new Panel { Dock = DockStyle.Bottom, BackColor = ControlBgColor };

            _showHiddenToggle = new CheckBox
            {
                Text = "Show Hidden Sheets",
                ForeColor = TextColor,
                Checked = true,
                AutoSize = true,
                FlatStyle = FlatStyle.Standard,
                Cursor = Cursors.Hand,
                UseVisualStyleBackColor = true
            };
            _showHiddenToggle.CheckedChanged += (s, e) => RefreshList();

            _tipsLabel = new Label
            {
                Text = "Shortcut Tips",
                AutoSize = true,
                Cursor = Cursors.Hand,
                ForeColor = Color.FromArgb(0, 100, 200),
                BackColor = ControlBgColor,
                Font = new Font("Segoe UI Semibold", 9f)
            };

            _toolTip = new ToolTip
            {
                AutoPopDelay = 10000,
                InitialDelay = 100,
                ReshowDelay = 500,
                ShowAlways = true,
                ToolTipTitle = "Keyboard Shortcuts"
            };
            _toolTip.SetToolTip(_tipsLabel,
                "F2\t\t Rename Sheet\n" +
                "Ctrl+Arrow\t Reorder Sheet(s)\n" +
                "Shift+Click\t Multi-Select\n" +
                "Enter\t\t Select Sheet\n" +
                "Esc\t\t Cancel");

            _footerPanel.Controls.Add(_tipsLabel);
            _footerPanel.Controls.Add(_showHiddenToggle);

            // 5. List
            _listContainerPanel = new Panel { Dock = DockStyle.Fill, BackColor = BackgroundColor };
            _listBox = new ListBox
            {
                Dock = DockStyle.Fill,
                DrawMode = DrawMode.OwnerDrawFixed,
                BorderStyle = BorderStyle.None,
                BackColor = BackgroundColor,
                IntegralHeight = true,
                AllowDrop = true,
                SelectionMode = SelectionMode.MultiExtended
            };
            _listBox.DrawItem += ListBox_DrawItem;
            _listBox.KeyDown += CommonKeyDown;
            _listBox.DoubleClick += (s, e) => AcceptSelection();
            _listBox.MouseDown += ListBox_MouseDown;
            _listBox.MouseMove += ListBox_MouseMove;
            _listBox.DragOver += ListBox_DragOver;
            _listBox.DragDrop += ListBox_DragDrop;

            _listContainerPanel.Controls.Add(_listBox);

            // Assemble
            _mainPanel.Controls.Add(_listContainerPanel);
            _mainPanel.Controls.Add(_footerPanel);
            _mainPanel.Controls.Add(_headerPanel);
            _mainPanel.Controls.Add(_titleBarPanel);

            KeyDown += Form_KeyDown;
            KeyPress += Form_KeyPress;
        }

        private void ListBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _listBox.Items.Count) return;

            var item = _listBox.Items[e.Index] as SheetItem;
            if (item == null) return;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            // Background
            if (isSelected)
            {
                using (var brush = new SolidBrush(SelectionBgColor)) g.FillRectangle(brush, e.Bounds);
            }
            else
            {
                using (var brush = new SolidBrush(BackgroundColor)) g.FillRectangle(brush, e.Bounds);
            }

            int dpiX = _lastAppliedDpiX > 0 ? _lastAppliedDpiX : 96;
            int padding = InchesToPxX(0.1, dpiX);
            int iconSize = InchesToPxX(0.14, dpiX);

            int iconY = e.Bounds.Y + (e.Bounds.Height - iconSize) / 2;
            int iconX = e.Bounds.X + padding + 6;

            Color iconColor = item.Hidden ? IconHiddenGray : IconExcelGreen;

            // Icon Drawing
            using (var p = new Pen(iconColor, 1.5f))
            {
                g.DrawRectangle(p, iconX, iconY, iconSize, iconSize);
                g.DrawLine(p, iconX, iconY + iconSize / 3, iconX + iconSize, iconY + iconSize / 3);
                g.DrawLine(p, iconX + iconSize / 2, iconY + iconSize / 3, iconX + iconSize / 2, iconY + iconSize);
            }

            int textX = iconX + iconSize + padding;
            Color textColor = item.Hidden ? SubTextColor : TextColor;

            using (var font = new Font("Segoe UI", PointsFromInches(FONT_BODY_IN), FontStyle.Regular))
            using (var brush = new SolidBrush(textColor))
            {
                float textY = e.Bounds.Y + (e.Bounds.Height - font.Height) / 2f;
                g.DrawString(item.Name, font, brush, textX, textY);
                if (item.Hidden)
                {
                    float nameWidth = g.MeasureString(item.Name, font).Width;
                    using (var italicFont = new Font("Segoe UI", PointsFromInches(FONT_SMALL_IN), FontStyle.Italic))
                    using (var tagBrush = new SolidBrush(SubTextColor))
                    {
                        g.DrawString("(hidden)", italicFont, tagBrush, textX + nameWidth + 5, textY + 1);
                    }
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_DPICHANGED)
            {
                int dpiX = (int)(m.WParam.ToInt64() & 0xFFFF);
                int dpiY = (int)((m.WParam.ToInt64() >> 16) & 0xFFFF);
                ApplyPhysicalLayout(dpiX, dpiY);
            }
            base.WndProc(ref m);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            int d = 96;
            try { d = GetDpiForWindow(Handle); } catch { using (var g = CreateGraphics()) d = (int)g.DpiX; }
            if (d == 0) d = 96;
            ApplyPhysicalLayout(d, d);
        }

        private void ApplyPhysicalLayout(int dpiX, int dpiY)
        {
            if (dpiX == _lastAppliedDpiX && dpiY == _lastAppliedDpiY) return;
            _lastAppliedDpiX = dpiX; _lastAppliedDpiY = dpiY;

            this.Size = new System.Drawing.Size(InchesToPxX(FORM_WIDTH_IN, dpiX), InchesToPxY(FORM_HEIGHT_IN, dpiY));

            _titleBarPanel.Height = InchesToPxY(TITLE_BAR_H_IN, dpiY);
            _headerPanel.Height = InchesToPxY(HEADER_H_IN, dpiY);
            _footerPanel.Height = InchesToPxY(FOOTER_H_IN, dpiY);

            int hPad = InchesToPxX(0.1, dpiX);
            int vPad = InchesToPxY(0.08, dpiY);
            _headerPanel.Padding = new Padding(hPad, vPad, hPad, vPad);
            _searchBox.Font = new Font("Segoe UI", PointsFromInches(FONT_SEARCH_IN));

            // Vertically center logic
            int containerHeight = _headerPanel.Height - (vPad * 2);
            int textBoxHeight = _searchBox.PreferredHeight;
            int centerPad = Math.Max(0, (containerHeight - textBoxHeight) / 2);
            _searchContainer.Padding = new Padding(InchesToPxX(0.05, dpiX), centerPad, 0, 0);

            _listBox.ItemHeight = InchesToPxY(LIST_ITEM_H_IN, dpiY);
            _listBox.Font = new Font("Segoe UI", PointsFromInches(FONT_BODY_IN));

            _showHiddenToggle.Font = new Font("Segoe UI", PointsFromInches(FONT_BODY_IN));
            _tipsLabel.Font = new Font("Segoe UI", PointsFromInches(FONT_BODY_IN), FontStyle.Underline);

            int midY = _footerPanel.Height / 2;
            _showHiddenToggle.Top = midY - (_showHiddenToggle.PreferredSize.Height / 2);
            _showHiddenToggle.Left = hPad;

            _tipsLabel.Top = midY - (_tipsLabel.PreferredHeight / 2);
            _tipsLabel.Left = _footerPanel.Width - _tipsLabel.PreferredWidth - hPad;
        }

        private static int InchesToPxX(double inches, int dpiX) => (int)Math.Round(inches * dpiX);
        private static int InchesToPxY(double inches, int dpiY) => (int)Math.Round(inches * dpiY);
        private static float PointsFromInches(double inches) => (float)(inches * 72.0);

        // ── Logic ────────────────────────────────────────────────────────────────
        private void PopulateSheets()
        {
            _allItems = new List<SheetItem>();
            if (_app?.Worksheets != null)
            {
                foreach (Excel.Worksheet ws in _app.Worksheets)
                {
                    if (ws != null) _allItems.Add(new SheetItem { Name = ws.Name, Hidden = ws.Visible != Excel.XlSheetVisibility.xlSheetVisible });
                }
            }
        }

        private void RefreshList()
        {
            if (_searchBox == null || _listBox == null || _allItems == null) return;
            string query = (_searchBox.ForeColor == HintColor) ? "" : _searchBox.Text?.Trim() ?? "";

            _listBox.BeginUpdate();
            _listBox.Items.Clear();
            var filtered = _allItems.Where(si =>
                (string.IsNullOrEmpty(query) || si.Name.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) &&
                (_showHiddenToggle.Checked || !si.Hidden));
            foreach (var item in filtered) _listBox.Items.Add(item);
            _listBox.EndUpdate();
        }

        private void SelectInitialSheets()
        {
            _listBox.ClearSelected();
            if (_listBox.Items.Count > 0 && (_preSelectedSheets == null || _preSelectedSheets.Count == 0))
            {
                _listBox.SelectedIndex = 0;
                return;
            }
            bool scrolled = false;
            for (int i = 0; i < _listBox.Items.Count; i++)
            {
                var item = _listBox.Items[i] as SheetItem;
                if (item != null && _preSelectedSheets.Contains(item.Name))
                {
                    _listBox.SetSelected(i, true);
                    if (!scrolled && item.Name == _activeSheetName)
                    {
                        _listBox.TopIndex = Math.Max(0, i - 2);
                        scrolled = true;
                    }
                }
            }
        }

        private void AcceptSelection() { DialogResult = DialogResult.OK; Close(); }

        private void Form_KeyDown(object sender, KeyEventArgs e)
        {
            if (_editBox != null) return;

            // Reorder (Ctrl+Arrow)
            if (e.Control && (e.KeyCode == Keys.Up || e.KeyCode == Keys.Down))
            {
                int dir = (e.KeyCode == Keys.Up) ? -1 : 1;
                MoveSelectedSheets(dir);
                e.Handled = true; return;
            }

            switch (e.KeyCode)
            {
                case Keys.Up:
                    if (Control.ModifierKeys == Keys.None)
                    {
                        MoveListSelection(-1); e.Handled = true;
                    }
                    else
                    {
                        e.Handled = false;
                    }
                    break;

                case Keys.Down:
                    if (Control.ModifierKeys == Keys.None)
                    {
                        MoveListSelection(1); e.Handled = true;
                    }
                    else
                    {
                        e.Handled = false;
                    }
                    break;

                case Keys.Enter: AcceptSelection(); e.Handled = true; break;
                case Keys.Escape: DialogResult = DialogResult.Cancel; Close(); e.Handled = true; break;
                case Keys.F2: BeginInlineRename(); e.Handled = true; break;
            }
        }

        private void Form_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (_editBox != null || char.IsControl(e.KeyChar)) return;
            if (_searchBox.ForeColor == HintColor) { _searchBox.Text = ""; _searchBox.ForeColor = TextColor; }
            _searchBox.Focus(); _searchBox.AppendText(e.KeyChar.ToString()); e.Handled = true;
        }

        private void CommonKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (_editBox != null) EndInlineRename(true); else AcceptSelection();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                if (_editBox != null) EndInlineRename(false); else { DialogResult = DialogResult.Cancel; Close(); }
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F2)
            {
                BeginInlineRename(); e.Handled = true;
            }
            else if (sender == _searchBox && (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up))
            {
                MoveListSelection(e.KeyCode == Keys.Down ? 1 : -1); e.Handled = true;
            }
        }

        private void MoveListSelection(int delta)
        {
            if (_listBox.Items.Count == 0) return;
            int current = (_listBox.SelectedIndex < 0) ? 0 : _listBox.SelectedIndex;
            int next = Math.Max(0, Math.Min(current + delta, _listBox.Items.Count - 1));

            _listBox.ClearSelected();
            _listBox.SelectedIndex = next;
        }

        private void BeginInlineRename()
        {
            if (_editBox != null || _listBox.SelectedIndex < 0 || _listBox.SelectedItems.Count > 1) return;

            int idx = _listBox.SelectedIndex;
            var rect = _listBox.GetItemRectangle(idx);
            var item = _listBox.Items[idx] as SheetItem;
            int dpiX = _lastAppliedDpiX > 0 ? _lastAppliedDpiX : 96;
            int iconAreaWidth = InchesToPxX(0.35, dpiX);

            _editBox = new TextBox
            {
                Bounds = new Rectangle(rect.X + iconAreaWidth, rect.Y + _listContainerPanel.Top + 2, rect.Width - iconAreaWidth - InchesToPxX(0.1, dpiX), rect.Height),
                Text = item.Name,
                BorderStyle = BorderStyle.FixedSingle,
                Font = _listBox.Font,
                BackColor = Color.White
            };
            _editBox.KeyDown += CommonKeyDown;
            _editBox.LostFocus += (s, e) => EndInlineRename(true);
            _mainPanel.Controls.Add(_editBox); _editBox.BringToFront(); _editBox.Focus(); _editBox.SelectAll();
        }

        private void EndInlineRename(bool commit)
        {
            if (_editBox == null || _isRenaming) return;
            _isRenaming = true;
            string oldName = "";
            string newName = "";
            var item = _listBox.SelectedItem as SheetItem;
            if (item != null) oldName = item.Name;

            try
            {
                newName = _editBox.Text?.Trim();
                if (commit && !string.IsNullOrEmpty(newName) && item != null && newName != item.Name)
                {
                    try
                    {
                        var ws = _app.Worksheets[item.Name] as Excel.Worksheet;
                        if (ws != null) { ws.Name = newName; item.Name = newName; }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Rename failed: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                }
                _mainPanel.Controls.Remove(_editBox); _editBox.Dispose(); _editBox = null;
                RefreshList();
                Reselect(new List<string> { string.IsNullOrEmpty(newName) ? oldName : newName });

            }
            finally
            {
                _isRenaming = false;
                Cursor.Current = Cursors.Default;
                try { _app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }
            }
        }

        private void ListBox_MouseDown(object s, MouseEventArgs e)
        {
            if (_editBox != null) return;
            _dragSourceIndex = _listBox.IndexFromPoint(e.Location);
            _dragStart = e.Location;
            _isDragging = false;
        }

        private void ListBox_MouseMove(object s, MouseEventArgs e)
        {
            if (_listBox.SelectedItems.Count > 1) return;
            if (e.Button == MouseButtons.Left && _dragSourceIndex >= 0 && !_isDragging)
            {
                if (Math.Abs(e.X - _dragStart.X) > 5 || Math.Abs(e.Y - _dragStart.Y) > 5)
                {
                    _isDragging = true;
                    _listBox.DoDragDrop(_dragSourceIndex, DragDropEffects.Move);
                }
            }
        }

        private void ListBox_DragOver(object s, DragEventArgs e) { e.Effect = DragDropEffects.Move; Point pt = _listBox.PointToClient(new Point(e.X, e.Y)); int idx = _listBox.IndexFromPoint(pt); if (idx >= 0) _listBox.SelectedIndex = idx; }
        private void ListBox_DragDrop(object s, DragEventArgs e) { if (!e.Data.GetDataPresent(typeof(int))) return; int from = (int)e.Data.GetData(typeof(int)); Point pt = _listBox.PointToClient(new Point(e.X, e.Y)); int to = _listBox.IndexFromPoint(pt); if (to == ListBox.NoMatches) to = _listBox.Items.Count - 1; if (from != to) MoveSheetByVisibleIndices(from, to); _isDragging = false; _dragSourceIndex = -1; }

        private void MoveSheetByVisibleIndices(int from, int to)
        {
            if (from == to) return;
            var srcItem = _listBox.Items[from] as SheetItem; var dstItem = _listBox.Items[to] as SheetItem;
            if (srcItem == null || dstItem == null) return;
            try
            {
                var src = _app.Worksheets[srcItem.Name] as Excel.Worksheet; var dst = _app.Worksheets[dstItem.Name] as Excel.Worksheet;
                if (src != null && dst != null) { if (to > from) src.Move(Type.Missing, dst); else src.Move(dst); }
                PopulateSheets(); RefreshList(); Reselect(new List<string> { srcItem.Name });
            }
            catch { }
            finally
            {
                Cursor.Current = Cursors.Default;
                try { _app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }
            }
        }

        private void MoveSelectedSheets(int dir)
        {
            var selectedNames = SelectedSheetNames;
            if (selectedNames.Count == 0) return;

            try
            {
                bool isFirst = true;
                var sheetsToMove = new List<Excel.Worksheet>();

                foreach (var name in selectedNames)
                {
                    var ws = _app.Worksheets[name] as Excel.Worksheet;
                    if (ws != null)
                    {
                        if (ws.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                            ws.Visible = Excel.XlSheetVisibility.xlSheetVisible;

                        sheetsToMove.Add(ws);
                        ws.Select(isFirst ? true : false);
                        isFirst = false;
                    }
                }

                if (sheetsToMove.Count == 0) return;

                // Sort by index
                sheetsToMove.Sort((a, b) => a.Index.CompareTo(b.Index));

                if (dir < 0) // UP
                {
                    int topIndex = sheetsToMove[0].Index;
                    if (topIndex > 1)
                    {
                        int targetIndex = topIndex - 1;
                        var neighbor = _app.Worksheets[targetIndex] as Excel.Worksheet;

                        if (neighbor != null && neighbor.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            // Hidden? Scan UP to find the START of this hidden block.
                            while (targetIndex > 1)
                            {
                                var prev = _app.Worksheets[targetIndex - 1] as Excel.Worksheet;
                                if (prev != null && prev.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                                    targetIndex--;
                                else
                                    break;
                            }
                        }

                        var target = _app.Worksheets[targetIndex];
                        ((Excel.Sheets)_app.ActiveWindow.SelectedSheets).Move(Before: target);
                    }
                }
                else // DOWN
                {
                    int botIndex = sheetsToMove[sheetsToMove.Count - 1].Index;
                    if (botIndex < _app.Worksheets.Count)
                    {
                        int targetIndex = botIndex + 1;
                        var neighbor = _app.Worksheets[targetIndex] as Excel.Worksheet;

                        if (neighbor != null && neighbor.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                        {
                            // Hidden? Scan DOWN to find the END of this hidden block.
                            while (targetIndex < _app.Worksheets.Count)
                            {
                                var next = _app.Worksheets[targetIndex + 1] as Excel.Worksheet;
                                if (next != null && next.Visible != Excel.XlSheetVisibility.xlSheetVisible)
                                    targetIndex++;
                                else
                                    break;
                            }
                        }

                        var target = _app.Worksheets[targetIndex];
                        ((Excel.Sheets)_app.ActiveWindow.SelectedSheets).Move(After: target);
                    }
                }

                PopulateSheets();
                RefreshList();
                Reselect(selectedNames);
            }
            catch { }
            finally
            {
                Cursor.Current = Cursors.Default;
                try { _app.Cursor = Excel.XlMousePointer.xlDefault; } catch { }
            }
        }

        private void Reselect(List<string> names)
        {
            _listBox.ClearSelected();
            if (names == null || names.Count == 0) return;

            var set = new HashSet<string>(names);
            for (int i = 0; i < _listBox.Items.Count; i++)
            {
                if (_listBox.Items[i] is SheetItem item && set.Contains(item.Name))
                {
                    _listBox.SetSelected(i, true);
                }
            }
            if (_listBox.SelectedIndices.Count > 0)
                _listBox.TopIndex = Math.Max(0, _listBox.SelectedIndices[0] - 2);
        }
    }
}