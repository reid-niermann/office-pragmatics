// SuperFind.cs – OfficePragmatics_Excel  (targets C# 7.3 / .NET 4.8, VSTO Excel)
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;                       // StringBuilder
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;    // for MsoTriState
using Timer = System.Windows.Forms.Timer;                 // disambiguate Timer
using DockStyle = System.Windows.Forms.DockStyle;         // ensure WinForms DockStyle

namespace OfficePragmatics_Excel
{
    /*───────────────────────────── ENTRY POINT ─────────────────────────────*/
    internal static class SuperFind
    {
        private static CustomTaskPane _pane;
        private static SuperFindControl _ui;

        public static void Show()
        {
            var addIn = Globals.ThisAddIn; // VSTO host
            if (_pane == null)
            {
                _ui = new SuperFindControl(addIn.Application);
                _ui.EscPressed += (s, e) =>
                {
                    if (_pane != null) _pane.Visible = false;
                };

                _pane = addIn.CustomTaskPanes.Add(_ui, "Super Find");
                _pane.Width = 450; // Slightly adjusted default for better initial fit
                _pane.VisibleChanged += (s, e) =>
                {
                    if (_pane.Visible) _ui.FocusSearchBox();
                };
            }
            _pane.Visible = true;
            _ui.FocusSearchBox();
        }
    }

    /*────────────────────────────── UI CONTROL ─────────────────────────────*/
    internal sealed class SuperFindControl : UserControl
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern IntPtr SetFocus(IntPtr hWnd);

        private readonly Excel.Application _app;

        #region UI elements
        // New root/top/bottom containers to control layout precisely
        private Panel _root;     // fills the task pane
        private Panel _top;      // scrollable area with find/types/options/scope/buttons
        private Panel _bottom;   // compact results+status block, docked to bottom

        // top content layout
        private TableLayoutPanel _main;
        private ComboBox _cboScope;
        private FlowLayoutPanel _pnlTypes, _pnlOpts, _pnlButtons;
        private TextBox _txtTerm;
        private CheckBox _chkMultiTerm, _chkMatchCase, _chkWholeWord, _chkIncludeHidden;
        private Button _btnSearch, _btnCancel, _btnSelectAll;

        // bottom content (results + status)
        private SmoothListView _lstHits;
        private Label _lblStatus;
        private ProgressBar _progress;

        private ToolTip _tips;
        #endregion

        private List<SearchResult> _hits = new List<SearchResult>();
        private CancellationTokenSource _cts;

        // Notify host to close the task pane on Esc when idle
        public event EventHandler EscPressed;

        internal SuperFindControl(Excel.Application app)
        {
            _app = app;
            BuildUI();
            WireGlobalKeyHandlers(this);
        }

        /* focus helper ---------------------------------------------------- */
        internal void FocusSearchBox()
        {
            if (!IsHandleCreated || _txtTerm == null) return;

            void TryFocus()
            {
                try
                {
                    SetForegroundWindow(this.Handle);   // bring pane's host to front
                    SetFocus(_txtTerm.Handle);          // shove focus into the box
                    _txtTerm.SelectAll();
                }
                catch { /* ignore */ }
            }

            BeginInvoke((Action)(() =>
            {
                TryFocus();
                var t = new Timer { Interval = 120 };
                t.Tick += (s, e) => { t.Stop(); t.Dispose(); TryFocus(); };
                t.Start();
            }));
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            FocusSearchBox(); // safety
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                HandleEsc();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        /* --------------------------- Build UI ----------------------------- */
        private void BuildUI()
        {
            // CRITICAL: Handle 4K/High DPI properly
            this.AutoScaleMode = AutoScaleMode.Font;

            Dock = DockStyle.Fill;
            BackColor = Color.FromArgb(250, 250, 250);
            Font = new Font("Segoe UI", 9.5F);

            _tips = new ToolTip { ShowAlways = true, AutomaticDelay = 150 };

            // ── Root contains two children:
            //    1) _top  (Dock=Fill, AutoScroll)    – find/types/options/scope/buttons
            //    2) _bottom (Dock=Bottom, fixed-ish) – results + status
            _root = new Panel { Dock = DockStyle.Fill, Padding = new Padding(12, 12, 12, 8), BackColor = BackColor };
            Controls.Add(_root);

            _top = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,           // never clip checkboxes; scroll if very short
                BackColor = BackColor
            };
            _root.Controls.Add(_top);

            _bottom = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 140,                // compact default; will adjust on resize
                BackColor = BackColor,
                Padding = new Padding(0, 6, 0, 0)
            };
            _root.Controls.Add(_bottom);

            // ── TOP: Table for controls
            _main = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                BackColor = BackColor
            };

            // OPTIMIZATION: Use AutoSize for labels instead of fixed pixels.
            // This recovers space on low resolutions (1280x800) and scales correctly on 4K.
            _main.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            _top.Controls.Add(_main);

            /* Row 0 – Find */
            var lblFind = new Label
            {
                Text = "&Find what:",
                UseMnemonic = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(70, 70, 70),
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 10, 8, 8),
                TabIndex = 0
            };
            _main.Controls.Add(lblFind, 0, 0);

            _txtTerm = CreateModernTextBox();
            _txtTerm.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    e.SuppressKeyPress = true;
                    _btnSearch.PerformClick();
                }
            };
            _main.Controls.Add(_txtTerm, 1, 0);

            /* Row 1 – Content Types */
            var grpTypes = CreateModernGroupBox("Content &Types");
            _pnlTypes = CreateModernCheckboxPanel();

            // Add resize handler to ensure proper width calculation
            grpTypes.Resize += (s, e) =>
            {
                if (_pnlTypes != null && grpTypes.ClientSize.Width > 0)
                {
                    _pnlTypes.MaximumSize = new System.Drawing.Size(
                        grpTypes.ClientSize.Width - _pnlTypes.Padding.Horizontal - 4, // Tighter tolerance
                        0);
                }
            };

            _pnlTypes.Controls.Add(MkCheck("&Values", true, "Alt+V"));
            _pnlTypes.Controls.Add(MkCheck("&Formulas", true, "Alt+F"));
            _pnlTypes.Controls.Add(MkCheck("&Notes", true, "Alt+N"));
            _pnlTypes.Controls.Add(MkCheck("&Hyperlinks", true, "Alt+H"));
            _pnlTypes.Controls.Add(MkCheck("Cha&rts", true, "Alt+R"));
            _pnlTypes.Controls.Add(MkCheck("&Shapes", true, "Alt+S"));
            grpTypes.Controls.Add(_pnlTypes);
            _main.SetColumnSpan(grpTypes, 2);
            _main.Controls.Add(grpTypes, 0, 1);

            /* Row 2 – Options */
            var grpOpts = CreateModernGroupBox("&Options");
            _pnlOpts = CreateModernCheckboxPanel();

            // Add resize handler for options panel too
            grpOpts.Resize += (s, e) =>
            {
                if (_pnlOpts != null && grpOpts.ClientSize.Width > 0)
                {
                    _pnlOpts.MaximumSize = new System.Drawing.Size(
                        grpOpts.ClientSize.Width - _pnlOpts.Padding.Horizontal - 4,
                        0);
                }
            };

            _chkMultiTerm = MkCheck("Multi-&term (;)", false, "Alt+T");
            _chkMatchCase = MkCheck("Match &case", false, "Alt+C");
            _chkWholeWord = MkCheck("&Whole word", false, "Alt+W");
            _chkIncludeHidden = MkCheck("Include hi&dden", false, "Alt+D"); // avoid Alt+H clash
            _pnlOpts.Controls.AddRange(new Control[] { _chkMultiTerm, _chkMatchCase, _chkWholeWord, _chkIncludeHidden });
            grpOpts.Controls.Add(_pnlOpts);
            _main.SetColumnSpan(grpOpts, 2);
            _main.Controls.Add(grpOpts, 0, 2);

            /* Row 3 – Scope */
            var lblScope = new Label
            {
                Text = "Searc&h in:",
                UseMnemonic = true,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(70, 70, 70),
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                Margin = new Padding(0, 10, 8, 8),
                TabIndex = 2
            };
            _main.Controls.Add(lblScope, 0, 3);

            _cboScope = CreateModernComboBox();
            _cboScope.Items.AddRange(new[] { "Selection", "Active Sheet(s)", "Active Workbook", "All Workbooks" });
            _cboScope.SelectedIndex = 1;
            _main.Controls.Add(_cboScope, 1, 3);

            /* Row 4 – Buttons */
            _pnlButtons = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Margin = new Padding(0, 8, 0, 0)
            };

            _btnSearch = CreateCmdButton("S&earch"); Tip(_btnSearch, "Alt+E");
            _btnCancel = CreateCmdButton("Cance&l"); Tip(_btnCancel, "Alt+L"); _btnCancel.Visible = false;
            _btnSelectAll = CreateCmdButton("Select &All"); Tip(_btnSelectAll, "Alt+A"); _btnSelectAll.Enabled = false;

            _btnSearch.Click += BtnSearch_Click;
            _btnCancel.Click += (s, e) => _cts?.Cancel();
            _btnSelectAll.Click += BtnSelectAll_Click;

            _pnlButtons.Controls.AddRange(new Control[] { _btnSearch, _btnCancel, _btnSelectAll });
            _main.SetColumnSpan(_pnlButtons, 2);
            _main.Controls.Add(_pnlButtons, 0, 4);

            // ── BOTTOM: results list (fills) + tiny status bar under it
            var bottomLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2,
                BackColor = BackColor,
                Padding = new Padding(0)
            };
            bottomLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100F)); // list
            bottomLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // status
            _bottom.Controls.Add(bottomLayout);

            _lstHits = new SmoothListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                MultiSelect = false,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9F),
                Margin = new Padding(0, 0, 0, 6),
                TabIndex = 4
            };
            _lstHits.Columns.Add("Location", 260);
            _lstHits.DoubleClick += (s, e) =>
            {
                if (_lstHits.SelectedItems.Count > 0)
                    JumpTo(_lstHits.SelectedItems[0].Index);
            };
            _lstHits.KeyDown += LstHits_KeyDown;
            _lstHits.Resize += (s, e) =>
            {
                if (_lstHits.Columns.Count > 0)
                    _lstHits.Columns[0].Width = Math.Max(50, _lstHits.ClientSize.Width - 4);
            };
            bottomLayout.Controls.Add(_lstHits, 0, 0);

            var statusPanel = new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                Margin = new Padding(0),
                Padding = new Padding(0),
                BackColor = BackColor
            };

            _lblStatus = new Label
            {
                AutoSize = true,
                ForeColor = Color.FromArgb(100, 100, 100),
                Font = new Font("Segoe UI", 8.5F, FontStyle.Regular),
                Margin = new Padding(0, 0, 8, 0)
            };

            _progress = new ProgressBar
            {
                Width = 160,
                Height = 6,
                Visible = false,
                Margin = new Padding(0, 2, 0, 0),
                Style = ProgressBarStyle.Continuous
            };

            statusPanel.Controls.AddRange(new Control[] { _lblStatus, _progress });
            bottomLayout.Controls.Add(statusPanel, 0, 1);

            // Tab order: labels hand focus to the next control
            lblFind.TabIndex = 0; _txtTerm.TabIndex = 1;
            lblScope.TabIndex = 2; _cboScope.TabIndex = 3;

            // responsive bottom height: compact but never dominates
            SizeChanged += (s, e) => AdjustBottomHeight();
            AdjustBottomHeight();
        }

        // Responsive compact results block; top panel always shows full groups (scrolls if needed)
        private void AdjustBottomHeight()
        {
            int h = this.ClientSize.Height;
            int min = 110;                     // never smaller than this
            int max = 170;                     // never larger than this
            int target = Math.Max(min, Math.Min(max, (int)(h * 0.33))); // ~33% of pane
            if (_bottom.Height != target) _bottom.Height = target;
        }

        // Modern control creation helpers
        private TextBox CreateModernTextBox()
        {
            return new TextBox
            {
                Dock = DockStyle.Top,
                Font = new Font("Segoe UI", 10F),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Margin = new Padding(0, 6, 0, 8),
                Padding = new Padding(8, 6, 8, 6),
                TabIndex = 1
            };
        }

        private ComboBox CreateModernComboBox()
        {
            return new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5F),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Standard,
                IntegralHeight = false,
                Dock = DockStyle.Fill,
                MinimumSize = new System.Drawing.Size(0, 28),
                Margin = new Padding(0, 6, 0, 8),
                TabIndex = 3
            };
        }

        private GroupBox CreateModernGroupBox(string title)
        {
            return new GroupBox
            {
                Text = title,
                Dock = DockStyle.Top,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
                ForeColor = Color.FromArgb(50, 50, 50),
                Padding = new Padding(8, 8, 8, 8), // Reduced from 12 to 8 to save vertical space
                Margin = new Padding(0, 8, 0, 4),
                BackColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                TabStop = false
            };
        }

        private FlowLayoutPanel CreateModernCheckboxPanel()
        {
            return new FlowLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill, // Changed from DockStyle.Top
                WrapContents = true,
                FlowDirection = FlowDirection.LeftToRight,
                Padding = new Padding(2), // Reduced from 4
                Margin = new Padding(0)
            };
        }

        private CheckBox MkCheck(string textWithAmpersand, bool @checked, string tip)
        {
            var cb = new CheckBox
            {
                Text = textWithAmpersand,
                Checked = @checked,
                AutoSize = true,
                UseMnemonic = true,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(70, 70, 70),
                FlatStyle = FlatStyle.Flat,
                // OPTIMIZATION: Drastically reduced Right margin (20->8) to prevent premature line wrapping
                Margin = new Padding(4, 3, 8, 4),
                TabStop = true,
                UseVisualStyleBackColor = true
            };
            Tip(cb, tip);
            return cb;
        }

        private void Tip(Control c, string text) { if (_tips != null) _tips.SetToolTip(c, text); }

        private static Button CreateCmdButton(string text) => new Button
        {
            Text = text,
            AutoSize = true,
            Font = new Font("Segoe UI", 9.5F, FontStyle.Regular),
            BackColor = text.Contains("Search") ? Color.FromArgb(0, 120, 212) : Color.FromArgb(240, 240, 240),
            ForeColor = text.Contains("Search") ? Color.White : Color.FromArgb(70, 70, 70),
            FlatStyle = FlatStyle.Flat,
            Padding = new Padding(12, 1, 12, 1),
            Margin = new Padding(0, 0, 8, 0),
            UseVisualStyleBackColor = false
        };

        /* Catch Esc/Enter even when focus is deep inside child controls --- */
        private void WireGlobalKeyHandlers(Control root)
        {
            root.KeyDown += AnyKeyDown;
            foreach (Control c in root.Controls)
                WireGlobalKeyHandlers(c);
        }
        private void AnyKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                e.SuppressKeyPress = true;
                HandleEsc();
            }
            else if (e.KeyCode == Keys.Enter && (sender is TextBox || sender is ComboBox))
            {
                e.SuppressKeyPress = true;
                _btnSearch.PerformClick();
            }
        }
        private void HandleEsc()
        {
            if (_cts != null)
                _cts.Cancel();
            else
                EscPressed?.Invoke(this, EventArgs.Empty);
        }

        #region Event handlers
        private async void BtnSearch_Click(object sender, EventArgs e)
        {
            var opts = CollectOptions();
            if (opts.Terms.Length == 0 || string.IsNullOrWhiteSpace(opts.Terms[0]))
            {
                _lblStatus.Text = "Please enter a search term.";
                return;
            }

            SetUiState(isSearching: true);
            _cts = new CancellationTokenSource();
            var progress = new Progress<ProgressReport>(ReportProgress);
            var sw = Stopwatch.StartNew();

            try
            {
                _hits = await Task.Run(() =>
                    SearchEngine.Run(_app, opts, progress, _cts.Token));
                sw.Stop();

                PopulateResults();
                _lblStatus.Text = _cts.IsCancellationRequested
                    ? "Search cancelled."
                    : $"Found {_hits.Count} match(es) in {sw.Elapsed.TotalSeconds:F2}s.";

                if (_hits.Count > 0) JumpTo(0);
            }
            catch (Exception ex)
            {
                _lblStatus.Text = "Search failed.";
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetUiState(isSearching: false);
                _cts.Dispose();
                _cts = null;
            }
        }

        private void BtnSelectAll_Click(object sender, EventArgs e)
        {
            if (_hits.Count == 0) return;

            var groups = _hits.Where(h => h.Type != SearchResultType.Object)
                              .GroupBy(h => new { h.WorkbookName, h.WorksheetName });

            _app.ScreenUpdating = false;
            try
            {
                foreach (var g in groups)
                {
                    Excel.Workbook wb = null;
                    Excel.Worksheet ws = null;
                    Excel.Range union = null;
                    try
                    {
                        wb = _app.Workbooks[g.Key.WorkbookName];
                        ws = wb.Worksheets[g.Key.WorksheetName];
                        ws.Activate();

                        foreach (var hit in g)
                        {
                            var cell = ws.Range[hit.Address];
                            union = union == null ? cell : _app.Union(union, cell);
                        }
                        union?.Select();
                    }
                    finally
                    {
                        if (union != null) Marshal.ReleaseComObject(union);
                        if (ws != null) Marshal.ReleaseComObject(ws);
                        if (wb != null) Marshal.ReleaseComObject(wb);
                    }
                }
            }
            finally { _app.ScreenUpdating = true; }
        }

        private void LstHits_KeyDown(object sender, KeyEventArgs e)
        {
            if (_lstHits.SelectedItems.Count == 0) return;
            int i = _lstHits.SelectedItems[0].Index;

            if (e.KeyCode == Keys.Up || e.KeyCode == Keys.Left)
            {
                e.Handled = true;
                JumpTo(Math.Max(0, i - 1));
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Right)
            {
                e.Handled = true;
                JumpTo(Math.Min(_hits.Count - 1, i + 1));
            }
        }
        #endregion

        #region Helper methods
        private SearchOptions CollectOptions()
        {
            var o = new SearchOptions
            {
                Scope = (SearchScope)_cboScope.SelectedIndex,
                CaseSensitive = _chkMatchCase.Checked,
                MatchWholeWord = _chkWholeWord.Checked,
                IncludeHidden = _chkIncludeHidden.Checked
            };

            var checks = _pnlTypes.Controls.OfType<CheckBox>().ToList();
            ContentMask mask = 0;
            if (checks[0].Checked) mask |= ContentMask.Values;
            if (checks[1].Checked) mask |= ContentMask.Formulas;
            if (checks[2].Checked) mask |= ContentMask.Notes;
            if (checks[3].Checked) mask |= ContentMask.Hyperlinks;
            if (checks[4].Checked) mask |= ContentMask.Charts;
            if (checks[5].Checked) mask |= ContentMask.Shapes;
            o.Mask = mask;

            string raw = _txtTerm.Text;
            o.Terms = _chkMultiTerm.Checked
                ? raw.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                     .Select(t => t.Trim()).Where(t => t.Length > 0).ToArray()
                : new[] { raw };

            return o;
        }

        private void SetUiState(bool isSearching)
        {
            _btnSearch.Visible = !isSearching;
            _btnCancel.Visible = isSearching;
            _btnSelectAll.Enabled = !isSearching && _hits.Count > 0;

            _cboScope.Enabled = !isSearching;
            _pnlTypes.Enabled = !isSearching;
            _txtTerm.Enabled = !isSearching;
            _pnlOpts.Enabled = !isSearching;
            _lstHits.Enabled = !isSearching;

            _progress.Visible = isSearching;
            if (!isSearching) { _progress.Value = 0; _lblStatus.Text = ""; }
        }

        private void ReportProgress(ProgressReport r)
        {
            _lblStatus.Text = r.Message;
            if (r.Total > 0)
            {
                _progress.Maximum = r.Total;
                _progress.Value = Math.Min(r.Current, _progress.Maximum);
            }
        }

        private void PopulateResults()
        {
            _lstHits.BeginUpdate();
            _lstHits.Items.Clear();
            foreach (var h in _hits)
                _lstHits.Items.Add(new ListViewItem(new[] { h.FormattedLocation }));
            _lstHits.EndUpdate();
            _btnSelectAll.Enabled = _hits.Any(h => h.Type != SearchResultType.Object);

            if (_lstHits.Columns.Count > 0)
                _lstHits.Columns[0].Width = Math.Max(50, _lstHits.ClientSize.Width - 4);
        }

        private void JumpTo(int index)
        {
            if (index < 0 || index >= _hits.Count) return;

            _lstHits.Items[index].Selected = true;
            _lstHits.EnsureVisible(index);

            var hit = _hits[index];
            try
            {
                _app.ScreenUpdating = false;

                var wb = _app.Workbooks[hit.WorkbookName];
                wb.Activate();
                dynamic sheet = wb.Sheets[hit.WorksheetName];
                sheet.Activate();

                if (hit.Type == SearchResultType.Object)
                    sheet.Shapes.Range(new[] { hit.Address }).Select();
                else
                    sheet.Range(hit.Address).Select();
            }
            finally { _app.ScreenUpdating = true; }
        }

        /* Small derived ListView with double-buffering to reduce flicker --- */
        private sealed class SmoothListView : ListView
        {
            public SmoothListView()
            {
                DoubleBuffered = true;
            }
        }
        #endregion
    }

    /*────────────────────────── SEARCH ENGINE ────────────────────────────*/
    #region DTOs
    [Flags]
    internal enum ContentMask
    {
        Values = 1, Formulas = 2, Notes = 4,
        Hyperlinks = 8, Charts = 16, Shapes = 32
    }
    internal enum SearchScope { Selection, ActiveSheets, ActiveWorkbook, AllWorkbooks }
    internal enum SearchResultType { Cell, Object }

    internal sealed class SearchOptions
    {
        public SearchScope Scope { get; set; }
        public ContentMask Mask { get; set; }
        public string[] Terms { get; set; } = Array.Empty<string>();
        public bool CaseSensitive { get; set; }
        public bool MatchWholeWord { get; set; }
        public bool IncludeHidden { get; set; }

        public StringComparison ComparisonType =>
            CaseSensitive ? StringComparison.CurrentCulture
                          : StringComparison.CurrentCultureIgnoreCase;
    }

    internal sealed class SearchResult
    {
        public string WorkbookName { get; set; }
        public string WorksheetName { get; set; }
        public string Address { get; set; }   // always relative – no $
        public string Content { get; set; }
        public SearchResultType Type { get; set; }

        public string FormattedLocation
        {
            get
            {
                string safeSheet = WorksheetName.Replace("'", "''");
                return $"'{safeSheet}'!{Address}";
            }
        }
    }

    internal sealed class ProgressReport
    {
        public string Message { get; set; }
        public int Current { get; set; }
        public int Total { get; set; }
    }
    #endregion

    internal static class SearchEngine
    {
        public static List<SearchResult> Run(
            Excel.Application app, SearchOptions o,
            IProgress<ProgressReport> progress, CancellationToken token)
        {
            var hits = new List<SearchResult>();

            var wbs = TargetWorkbooks(app, o).ToList();
            int totalSheets = wbs.SelectMany(wb => TargetSheets(wb, app, o)).Count();
            int processed = 0;

            foreach (var wb in wbs)
            {
                foreach (dynamic ws in TargetSheets(wb, app, o))
                {
                    token.ThrowIfCancellationRequested();
                    processed++;
                    progress.Report(new ProgressReport
                    {
                        Message = $"Searching '{ws.Name}'...",
                        Current = processed,
                        Total = totalSheets
                    });

                    ScanSheet(app, wb, ws, o, hits, token);
                    Marshal.ReleaseComObject(ws);
                }
                Marshal.ReleaseComObject(wb);
            }
            return hits;
        }

        private static IEnumerable<Excel.Workbook> TargetWorkbooks(
            Excel.Application app, SearchOptions o) =>
            o.Scope == SearchScope.AllWorkbooks
                ? app.Workbooks.Cast<Excel.Workbook>()
                : new[] { app.ActiveWorkbook };

        private static IEnumerable<object> TargetSheets(
            Excel.Workbook wb, Excel.Application app, SearchOptions o)
        {
            switch (o.Scope)
            {
                case SearchScope.ActiveSheets:
                    if (app.ActiveWindow.SelectedSheets is Excel.Sheets sel)
                        return sel.Cast<object>();
                    return new[] { wb.ActiveSheet };

                case SearchScope.Selection:
                    return new[] { app.ActiveSheet };

                default: // ActiveWorkbook or AllWorkbooks
                    return wb.Sheets.Cast<object>()
                                     .Where(sh => o.IncludeHidden ||
                                                  ((dynamic)sh).Visible ==
                                                  Excel.XlSheetVisibility.xlSheetVisible);
            }
        }

        private static void ScanSheet(Excel.Application app, Excel.Workbook wb,
            dynamic ws, SearchOptions o, List<SearchResult> hits, CancellationToken token)
        {
            try
            {
                if (ws is Excel.Worksheet sheet)
                {
                    Excel.Range rng = null;
                    try
                    {
                        rng = (o.Scope == SearchScope.Selection)
                            ? app.Selection as Excel.Range
                            : sheet.UsedRange;

                        if (rng == null ||
                            (rng.Count == 1 && string.IsNullOrEmpty(Convert.ToString(rng.Value2))))
                            return;

                        if ((o.Mask & (ContentMask.Values | ContentMask.Formulas)) != 0)
                            ScanCellContents(rng, wb.Name, sheet.Name, o, hits, token);

                        if ((o.Mask & ContentMask.Notes) != 0)
                            ScanCollection<Excel.Comment>(
                                sheet.Comments, wb.Name, sheet.Name, o,
                                new Func<Excel.Comment, string>(c => c.Text()),
                                new Func<Excel.Comment, string>(c => c.Parent.Address),
                                hits);

                        if ((o.Mask & ContentMask.Hyperlinks) != 0)
                            ScanCollection<Excel.Hyperlink>(
                                sheet.Hyperlinks, wb.Name, sheet.Name, o,
                                new Func<Excel.Hyperlink, string>(
                                    h => $"{h.Address} {h.TextToDisplay}"),
                                new Func<Excel.Hyperlink, string>(h => h.Range.Address),
                                hits);

                        if ((o.Mask & ContentMask.Charts) != 0)
                            ScanCollection<Excel.ChartObject>(
                                sheet.ChartObjects(), wb.Name, sheet.Name, o,
                                new Func<Excel.ChartObject, string>(co =>
                                    co.Name + (co.Chart.HasTitle
                                        ? " " + co.Chart.ChartTitle.Text
                                        : "")),
                                new Func<Excel.ChartObject, string>(co => co.Name),
                                hits, isObject: true);
                    }
                    finally { if (rng != null) Marshal.ReleaseComObject(rng); }
                }

                if ((o.Mask & ContentMask.Shapes) != 0)
                {
                    ScanCollection<Excel.Shape>(
                        ws.Shapes, wb.Name, ws.Name, o,
                        new Func<Excel.Shape, string>(sh =>
                        {
                            if (!o.IncludeHidden &&
                                sh.Visible != Office.MsoTriState.msoTrue)
                                return null;
                            return GetShapeText(sh);
                        }),
                        new Func<Excel.Shape, string>(sh => sh.Name),
                        hits, isObject: true);
                }
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"COM error scanning '{ws.Name}': {ex.Message}");
            }
        }

        private static void ScanCellContents(
            Excel.Range rng, string wb, string ws, SearchOptions o,
            List<SearchResult> hits, CancellationToken token)
        {
            object[,] vals = null, frms = null;

            if (rng.CountLarge == 1)
            {
                vals = new object[1, 1] { { rng.Value2 } };
                frms = new object[1, 1] { { rng.Formula } };
            }
            else
            {
                if ((o.Mask & ContentMask.Values) != 0) vals = rng.Value2;
                if ((o.Mask & ContentMask.Formulas) != 0) frms = rng.Formula;
            }

            int startRow = rng.Row, startCol = rng.Column;
            int rows = (vals ?? frms).GetUpperBound(0);
            int cols = (vals ?? frms).GetUpperBound(1);

            for (int r = 1; r <= rows; r++)
            {
                token.ThrowIfCancellationRequested();
                for (int c = 1; c <= cols; c++)
                {
                    if ((o.Mask & ContentMask.Values) != 0 && vals?[r, c] != null)
                    {
                        string v = Convert.ToString(vals[r, c]);
                        if (MatchAny(v, o))
                            hits.Add(CreateHit(wb, ws, startRow + r - 1, startCol + c - 1, v));
                    }
                    if ((o.Mask & ContentMask.Formulas) != 0 && frms?[r, c] != null)
                    {
                        string f = Convert.ToString(frms[r, c]);
                        if (f.StartsWith("=") && MatchAny(f, o))
                            hits.Add(CreateHit(wb, ws, startRow + r - 1, startCol + c - 1, f));
                    }
                }
            }
        }

        private static void ScanCollection<T>(
            dynamic coll, string wb, string ws, SearchOptions o,
            Func<T, string> getText, Func<T, string> getAddr,
            List<SearchResult> hits, bool isObject = false)
        {
            try
            {
                foreach (T item in coll)
                {
                    if (item == null) continue;
                    string txt = null, addr = null;
                    try
                    {
                        txt = getText(item);
                        if (string.IsNullOrEmpty(txt)) continue;
                        if (MatchAny(txt, o))
                        {
                            addr = getAddr(item);
                            hits.Add(new SearchResult
                            {
                                WorkbookName = wb,
                                WorksheetName = ws,
                                Address = addr,
                                Content = txt.Length > 100 ? txt.Substring(0, 100) + "..." : txt,
                                Type = isObject ? SearchResultType.Object : SearchResultType.Cell
                            });
                        }
                    }
                    finally
                    {
                        if (item != null && Marshal.IsComObject(item))
                            Marshal.ReleaseComObject(item);
                    }
                }
            }
            catch (COMException ex)
            {
                Debug.WriteLine($"Collection scan error: {ex.Message}");
            }
            finally
            {
                if (coll != null && Marshal.IsComObject(coll))
                    Marshal.ReleaseComObject(coll);
            }
        }

        private static string GetShapeText(Excel.Shape sh)
        {
            var sb = new StringBuilder(sh.Name);

            try
            {
                if (sh.TextFrame2 != null &&
                    sh.TextFrame2.HasText == Office.MsoTriState.msoTrue)
                {
                    sb.Append(" ").Append(sh.TextFrame2.TextRange.Text);
                }
                else if (sh.TextFrame != null)
                {
                    string legacy = sh.TextFrame.Characters().Text;
                    if (!string.IsNullOrEmpty(legacy))
                        sb.Append(" ").Append(legacy);
                }
            }
            catch { /* some shape types throw – ignore */ }

            if (!string.IsNullOrEmpty(sh.AlternativeText))
                sb.Append(" ").Append(sh.AlternativeText);

            return sb.ToString();
        }

        private static SearchResult CreateHit(
            string wb, string ws, int row, int col, string content) =>
            new SearchResult
            {
                WorkbookName = wb,
                WorksheetName = ws,
                Address = $"{ColumnIndexToName(col)}{row}",
                Content = content.Length > 100 ? content.Substring(0, 100) + "..." : content,
                Type = SearchResultType.Cell
            };

        private static string ColumnIndexToName(int ci)
        {
            int div = ci;
            string name = "";
            while (div > 0)
            {
                int mod = (div - 1) % 26;
                name = Convert.ToChar(65 + mod) + name;
                div = (div - mod) / 26;
            }
            return name;
        }

        private static bool MatchAny(string hay, SearchOptions o)
        {
            if (string.IsNullOrEmpty(hay)) return false;
            foreach (string term in o.Terms)
            {
                if (o.MatchWholeWord)
                {
                    if (Regex.IsMatch(hay,
                        @"\b" + Regex.Escape(term) + @"\b",
                        o.CaseSensitive ? RegexOptions.None : RegexOptions.IgnoreCase))
                        return true;
                }
                else if (hay.IndexOf(term, o.ComparisonType) >= 0)
                    return true;
            }
            return false;
        }
    }
}