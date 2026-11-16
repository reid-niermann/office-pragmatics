// AllSettings.cs (Excel) (C# 7.3 / .NET 4.8, VSTO)
// Requires: PragmaticsUI.cs (Dpi, Theme, PragmaticsForm), Newtonsoft.Json

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MsoTriState = Microsoft.Office.Core.MsoTriState;
using Newtonsoft.Json;

using OfficePragmatics_Excel.Classes.Brand;
using OfficePragmatics_Excel.Classes.Tables;

// Avoid naming clashes with File/Directory types
using IOFile = System.IO.File;
using IODir = System.IO.Directory;

namespace OfficePragmatics_Excel.Settings
{
    // ------------------------------
    // Persisted settings (JSON)
    // ------------------------------
    public class AllSettings
    {
        // Schema metadata (future-proofing)
        public int SchemaVersion { get; set; } = 1;
        public DateTime LastSavedUtc { get; set; } = DateTime.UtcNow;

        // General
        public bool AlignToBorder { get; set; } = true;
        public bool AutoFontColorEnabled { get; set; } = false;
        public bool DisableF1 { get; set; } = false; // true = disable Excel F1

        // Shared Library (now lives in JSON)
        public string SharedLibraryBasePath { get; set; } = string.Empty;

        // Charts / Tables (existing types)
        public BrandChartSettings BrandCharts { get; set; } = new BrandChartSettings();
        public TableFormatSettings TableFormat { get; set; } = new TableFormatSettings();

        // Shortcuts persisted as Method ↔ OnKey
        public List<ShortcutBinding> Shortcuts { get; set; } = new List<ShortcutBinding>();
    }

    public class ShortcutBinding
    {
        public string Method { get; set; }
        public string OnKey { get; set; }
    }

    // ------------------------------
    // Central manager for JSON settings
    // ------------------------------
    internal static class SettingsManager
    {
        private const string FolderName = "OfficePragmatics";
        private const string JsonFileName = "settings.json";

        private static readonly object _lock = new object();

        public static string AppDataDir =>
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), FolderName);

        public static string JsonPath => Path.Combine(AppDataDir, JsonFileName);

        public static AllSettings Current { get; private set; } = new AllSettings();

        public static AllSettings Load()
        {
            lock (_lock)
            {
                EnsureDirectory();

                if (IOFile.Exists(JsonPath))
                {
                    try
                    {
                        var json = IOFile.ReadAllText(JsonPath, Encoding.UTF8);
                        var loaded = JsonConvert.DeserializeObject<AllSettings>(json) ?? new AllSettings();
                        Current = loaded;
                        if (Current.BrandCharts == null) Current.BrandCharts = new BrandChartSettings();
                        if (Current.TableFormat == null) Current.TableFormat = new TableFormatSettings();
                        if (Current.Shortcuts == null) Current.Shortcuts = new List<ShortcutBinding>();
                        return Current;
                    }
                    catch
                    {
                        // Corrupt file → back up and reset
                        TryMoveBad(JsonPath);
                        Current = new AllSettings();
                        Save(Current);
                        return Current;
                    }
                }

                Current = new AllSettings();
                Save(Current);
                return Current;
            }
        }

        public static void Save(AllSettings settings)
        {
            lock (_lock)
            {
                EnsureDirectory();
                settings.LastSavedUtc = DateTime.UtcNow;

                var json = JsonConvert.SerializeObject(settings, Formatting.Indented);
                var tmp = JsonPath + ".tmp";

                // Atomic-ish write
                IOFile.WriteAllText(tmp, json, Encoding.UTF8);
                if (IOFile.Exists(JsonPath))
                    IOFile.Replace(tmp, JsonPath, null);
                else
                    IOFile.Move(tmp, JsonPath);

                Current = settings;
            }
        }

        private static void EnsureDirectory()
        {
            if (!IODir.Exists(AppDataDir))
            {
                try { IODir.CreateDirectory(AppDataDir); } catch { /* ignore */ }
            }
        }

        private static void TryMoveBad(string path)
        {
            try
            {
                var bad = path + ".bad_" + DateTime.UtcNow.ToString("yyyyMMdd_HHmmss");
                IOFile.Move(path, bad);
            }
            catch { /* ignore */ }
        }
    }

    public static class SettingsRuntime
    {
        /// <summary>
        /// Call this once in ThisAddIn_Startup to load settings and apply them:
        /// - F1 disable/enable
        /// - Shortcuts (registers)
        /// - Auto Font Color sheet change handler
        /// </summary>
        public static void InitializeOnAddInStartup()
        {
            var s = SettingsManager.Load();

            // Apply F1 mapping
            try
            {
                ApplyF1Setting(s.DisableF1);
            }
            catch { /* ignore */ }

            // Apply Shortcuts: overlay saved keys onto known methods
            try
            {
                var utils = new AddInUtilities();
                utils.UnregisterShortcuts();

                if (s.Shortcuts != null && s.Shortcuts.Count > 0)
                {
                    var arr = AddInUtilities._shortcuts;
                    var map = s.Shortcuts.Where(b => !string.IsNullOrWhiteSpace(b.Method))
                                         .ToDictionary(b => b.Method, b => NormalizeOnKey(b.OnKey ?? ""),
                                                       StringComparer.OrdinalIgnoreCase);

                    for (int i = 0; i < arr.Length; i++)
                    {
                        var (key, method) = arr[i];
                        if (map.TryGetValue(method, out var savedOnKey))
                        {
                            arr[i] = (savedOnKey, method);
                        }
                    }
                    AddInUtilities._shortcuts = arr;
                }

                utils.RegisterShortcuts();
            }
            catch { /* ignore */ }

            // Hook/unhook Auto Font Color
            try { Globals.ThisAddIn?.UpdateSheetChangeHandler(); } catch { /* ignore */ }
        }

        /// <summary>
        /// Apply F1 key registration/unregistration based on disable setting
        /// </summary>
        internal static void ApplyF1Setting(bool disableF1)
        {
            try
            {
                var app = Globals.ThisAddIn?.Application;
                if (app == null) return;

                if (disableF1)
                {
                    // Disable F1 by mapping it to nothing
                    app.OnKey("{F1}", "");
                }
                else
                {
                    // Re-enable F1 by removing the override (null = restore default)
                    app.OnKey("{F1}");
                }
            }
            catch { /* ignore */ }
        }

        internal static string NormalizeOnKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            bool ctrl = s.IndexOf('^') >= 0;
            bool shift = s.IndexOf('+') >= 0 || s.Any(ch => char.IsLetter(ch) && char.IsUpper(ch));
            bool alt = s.IndexOf('%') >= 0;

            var raw = s.Replace("^", "").Replace("+", "").Replace("%", "");
            string key;
            if (raw.StartsWith("{") && raw.EndsWith("}"))
            {
                key = "{" + raw.Substring(1, raw.Length - 2).ToUpperInvariant() + "}";
            }
            else if (raw.Length == 1 && char.IsLetter(raw[0]))
            {
                key = shift ? raw.ToUpperInvariant() : raw.ToLowerInvariant();
            }
            else
            {
                key = raw.ToUpperInvariant();
            }

            var sb = new StringBuilder();
            if (ctrl) sb.Append('^');
            if (shift) sb.Append('+');
            if (alt) sb.Append('%');
            sb.Append(key);
            return sb.ToString();
        }
    }

    // ------------------------------
    // Settings Form (UI) - Modern Style
    // ------------------------------
    public sealed class SettingsForm : Form
    {
        #region Native Interop for Draggable Title Bar
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;
        #endregion

        #region Modern Theme Colors (matching LogoSearchForm)
        private struct UIColors
        {
            public static readonly Color Surface = Color.FromArgb(255, 255, 255);
            public static readonly Color SurfaceAlt = Color.FromArgb(248, 248, 248);

            // BRAND COLOR: Deep Navy from "Office Pragmatics" Logo Text
            public static readonly Color TitleBg = Color.FromArgb(32, 42, 102);
            public static readonly Color TitleText = Color.White;

            public static readonly Color Border = Color.FromArgb(220, 220, 220);
            public static readonly Color TextPrimary = Color.FromArgb(32, 31, 30);
            public static readonly Color TextSecondary = Color.FromArgb(96, 94, 92);

            // BRAND COLOR: Periwinkle Blue from "Microsoft Office" Text / "14 Day Trial" Button
            public static readonly Color Accent = Color.FromArgb(110, 125, 220);
            public static readonly Color AccentHover = Color.FromArgb(85, 95, 180);

            public static readonly Color Danger = Color.FromArgb(232, 17, 35);

            // Navigation
            public static readonly Color NavBg = Color.FromArgb(245, 245, 250);
            public static readonly Color NavText = Color.FromArgb(50, 50, 50);
            public static readonly Color NavSelected = Color.White;
            public static readonly Color HeaderBg = Color.FromArgb(250, 250, 252);
        }
        #endregion

        // ✅ Now backed by JSON
        public static AllSettings CurrentSettings { get; private set; } = SettingsManager.Load();

        private float _scale;

        public static void ShowSettingsDialog()
        {
            using (Dpi.UsePerMonitorV2())
            {
                var owner = new WindowWrapper(new IntPtr(Globals.ThisAddIn.Application.Hwnd));
                using (var dlg = new SettingsForm(CurrentSettings))
                {
                    if (dlg.ShowDialog(owner) == DialogResult.OK)
                        Globals.ThisAddIn.UpdateSheetChangeHandler();
                }
            }
        }

        private sealed class WindowWrapper : IWin32Window { public WindowWrapper(IntPtr h) { Handle = h; } public IntPtr Handle { get; } }

        #region Fields
        private ListBox _nav;
        private Panel _contentHost, _header, _pageContainer;
        private Panel _titleBar;
        private Label _btnClose;

        private Label _headerDesc;

        private Button _btnOK, _btnCancel;


        // General
        private Panel _generalPanel;
        private CheckBox _alignChk, _autoFontColorChk, _disableF1Chk;

        // Shared library UI
        private Label _sharedHeader;
        private TextBox _sharedPath;
        private Button _btnChangeShared;
        private Button _btnClearShared;


        // Brand Chart
        private Panel _brandChartPanel, _chartCard;
        private Button _btnImportChart, _btnRemoveSecondary;
        private ListView _lvSecondaryCharts;
        private ComboBox _cmbSameAs;
        private readonly List<Excel.XlChartType> _allTypes;
        private int _hitItemIndex = -1, _hitSubIndex = -1;
        private const int COL_SAMEAS = 2;

        // Tables
        private Panel _tablePanel, _tableCard;
        private Button _btnImportTable, _btnClearTable;
        private ListView _lvTableProps;

        // Shortcuts
        private Panel _shortcutsPanel, _shortcutsCard;
        private ListView _lvShortcuts;
        private Button _btnResetShortcuts;
        private int _editItemIndex = -1;

        // ✅ Single declaration only
        private static readonly (string Key, string Method)[] _defaultShortcuts;
        #endregion

        #region Init
        static SettingsForm()
        {
            _defaultShortcuts = new (string Key, string Method)[AddInUtilities._shortcuts.Length];
            Array.Copy(AddInUtilities._shortcuts, _defaultShortcuts, AddInUtilities._shortcuts.Length);
        }

        private int Px(int logical) => (int)(logical * _scale);

        public SettingsForm(AllSettings initial)
        {
            // DPI Scaling
            using (var g = CreateGraphics()) { _scale = g.DpiX / 96f; }

            // Form Settings - Borderless Modern Style
            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.StartPosition = FormStartPosition.CenterParent;
            this.BackColor = UIColors.Border;
            this.Padding = new Padding(1);
            this.KeyPreview = true;

            int w = Px(820);
            int h = Px(560);
            this.Size = new System.Drawing.Size(w, h);

            // Always operate on the manager's live instance
            CurrentSettings = initial ?? SettingsManager.Load();
            _allTypes = new List<Excel.XlChartType>
            {
                Excel.XlChartType.xlColumnClustered, Excel.XlChartType.xlColumnStacked, Excel.XlChartType.xlColumnStacked100,
                Excel.XlChartType.xlBarClustered, Excel.XlChartType.xlBarStacked, Excel.XlChartType.xlBarStacked100,
                Excel.XlChartType.xlLine, Excel.XlChartType.xlLineMarkers,
                Excel.XlChartType.xlArea, Excel.XlChartType.xlAreaStacked, Excel.XlChartType.xlAreaStacked100,
                Excel.XlChartType.xlPie, Excel.XlChartType.xlDoughnut,
                Excel.XlChartType.xlXYScatter, Excel.XlChartType.xlXYScatterLines, Excel.XlChartType.xlXYScatterLinesNoMarkers
            };

            InitializeFormChrome();

            _cmbSameAs = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Visible = false };
            _cmbSameAs.SelectionChangeCommitted += SameAsCommitted;
            _cmbSameAs.KeyDown += (s, e) => { if (e.KeyCode == Keys.Escape) { _cmbSameAs.Visible = false; e.Handled = true; } };
            _brandChartPanel.Controls.Add(_cmbSameAs);

            // Populate current values from settings
            _alignChk.Checked = CurrentSettings.AlignToBorder;
            _autoFontColorChk.Checked = CurrentSettings.AutoFontColorEnabled;
            _disableF1Chk.Checked = CurrentSettings.DisableF1;

            PopulateBrandChartPanel();
            PopulateTableProps();
            PopulateShortcutsPanelFromRuntimeOrSaved();

            // Events
            _nav.SelectedIndexChanged += (s, e) => ShowSelectedPanel();

            _lvSecondaryCharts.ItemCheck += ListView_ItemCheck;
            _lvSecondaryCharts.MouseDown += Lv_MouseDownRemember;
            _lvSecondaryCharts.MouseUp += Lv_MouseUpShowCombo;

            _btnImportChart.Click += OnImportChartClick;
            _btnRemoveSecondary.Click += RemoveSecondary_Click;

            _btnImportTable.Click += OnImportTableClick;
            _btnClearTable.Click += OnClearTableClick;

            _btnChangeShared.Click += OnChangeSharedClick;
            _btnClearShared.Click += OnClearSharedClick;

            _lvShortcuts.ItemSelectionChanged += (s, e) => { if (e.IsSelected) _editItemIndex = e.ItemIndex; };
            _lvShortcuts.GotFocus += (s, e) => { if (_lvShortcuts.SelectedItems.Count > 0) _editItemIndex = _lvShortcuts.SelectedItems[0].Index; };
            _lvShortcuts.LostFocus += (s, e) => { _editItemIndex = -1; };

            _btnResetShortcuts.Click += OnResetShortcutsClick;

            _lvTableProps.Resize += (s, e) => FitTwoColumns(_lvTableProps, 260);
            _lvShortcuts.Resize += (s, e) => FitTwoColumnsEqual(_lvShortcuts);
            _lvSecondaryCharts.Resize += (s, e) => FitThreeColumnsFillLast(_lvSecondaryCharts, 80, 260);

            _nav.SelectedIndex = 0;
            ShowSelectedPanel();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            RefreshSharedPathDisplay();
        }

        private void InitializeFormChrome()
        {
            SuspendLayout();

            // Main surface panel (inside the 1px border)
            var mainPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.Surface
            };
            this.Controls.Add(mainPanel);

            // ═══════════════════════════════════════════════════════════════
            // TITLE BAR (Dock=Top) - Deep Navy with custom paint
            // ═══════════════════════════════════════════════════════════════
            _titleBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = Px(50),
                BackColor = UIColors.TitleBg
            };
            _titleBar.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            };
            _titleBar.Paint += (s, e) =>
            {
                e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
                e.Graphics.CompositingQuality = CompositingQuality.HighQuality;

                using (var font = new Font("Segoe UI", 13f, FontStyle.Bold))
                using (var brush = new SolidBrush(UIColors.TitleText))
                {
                    float y = (_titleBar.Height - font.Height) / 2f;
                    e.Graphics.DrawString("Settings - Office Pragmatics", font, brush, Px(20), y);
                }
            };

            // Close button (X)
            _btnClose = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 11f),
                ForeColor = UIColors.TitleText,
                Dock = DockStyle.Right,
                Width = Px(46),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                BackColor = UIColors.TitleBg
            };
            _btnClose.MouseEnter += (s, e) => { _btnClose.BackColor = UIColors.Danger; _btnClose.ForeColor = Color.White; };
            _btnClose.MouseLeave += (s, e) => { _btnClose.BackColor = UIColors.TitleBg; _btnClose.ForeColor = UIColors.TitleText; };
            _btnClose.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            _titleBar.Controls.Add(_btnClose);

            // ═══════════════════════════════════════════════════════════════
            // BOTTOM BAR (Dock=Bottom) - OK/Cancel
            // ═══════════════════════════════════════════════════════════════
            var bottomBar = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = Px(54),
                BackColor = UIColors.Surface,
                Padding = new Padding(Px(16), Px(0), Px(16), Px(0))
            };
            bottomBar.Paint += (s, e) =>
            {
                using (var pen = new Pen(UIColors.Border))
                    e.Graphics.DrawLine(pen, 0, 0, bottomBar.Width, 0);
            };

            // TableLayoutPanel for bottom bar layout
            var bottomLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 3,
                RowCount = 1,
                BackColor = Color.Transparent,
                Margin = new Padding(0)
            };
            bottomLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // Right: OK / Cancel buttons in a TableLayoutPanel for proper centering
            var buttonsTable = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent,
                Margin = new Padding(0),
                Anchor = AnchorStyles.None
            };
            buttonsTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            buttonsTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttonsTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            BuildButtons(out _btnOK, out _btnCancel);
            StylePrimaryButton(_btnOK);
            StyleGhostButton(_btnCancel);
            _btnOK.Anchor = AnchorStyles.None;
            _btnCancel.Anchor = AnchorStyles.None;
            _btnCancel.Margin = new Padding(0, 0, Px(8), 0);
            buttonsTable.Controls.Add(_btnCancel, 0, 0);
            buttonsTable.Controls.Add(_btnOK, 1, 0);

            bottomLayout.Controls.Add(buttonsTable, 2, 0);
            bottomBar.Controls.Add(bottomLayout);

            // ═══════════════════════════════════════════════════════════════
            // MAIN CONTENT AREA
            // ═══════════════════════════════════════════════════════════════
            var mainContent = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 1,
                ColumnCount = 2,
                BackColor = UIColors.SurfaceAlt
            };
            mainContent.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, Px(180)));
            mainContent.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

            // Navigation List
            _nav = BuildNavList();
            mainContent.Controls.Add(_nav, 0, 0);

            // Content Host
            _contentHost = BuildContentHost(out _header, out _headerDesc, out _pageContainer);
            mainContent.Controls.Add(_contentHost, 1, 0);

            // Build panels
            _generalPanel = BuildGeneralPanel(out _alignChk, out _autoFontColorChk, out _disableF1Chk);
            _brandChartPanel = BuildBrandChartPanel();
            _tablePanel = BuildTablePanel();
            _shortcutsPanel = BuildShortcutsPanel();

            _pageContainer.Controls.AddRange(new Control[]
            {
                _generalPanel, _brandChartPanel, _tablePanel, _shortcutsPanel
            });

            mainPanel.Controls.Add(mainContent);
            mainPanel.Controls.Add(bottomBar);
            mainPanel.Controls.Add(_titleBar);

            this.AcceptButton = _btnOK;
            this.CancelButton = _btnCancel;

            ResumeLayout(false);
        }

        private void StylePrimaryButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = UIColors.TitleBg;
            btn.ForeColor = UIColors.TitleText;
            btn.Font = new Font("Segoe UI Semibold", 10f);
            btn.Cursor = Cursors.Hand;
            btn.FlatAppearance.BorderSize = 0;
            btn.FlatAppearance.MouseOverBackColor = UIColors.Accent;
        }

        private void StyleGhostButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = UIColors.Surface;
            btn.ForeColor = UIColors.TextPrimary;
            btn.Font = new Font("Segoe UI", 9f);
            btn.Cursor = Cursors.Hand;
            btn.FlatAppearance.BorderColor = UIColors.Border;
            btn.FlatAppearance.MouseOverBackColor = UIColors.SurfaceAlt;
        }

        private void StyleActionButton(Button btn)
        {
            btn.FlatStyle = FlatStyle.Flat;
            btn.BackColor = UIColors.Surface;
            btn.ForeColor = UIColors.TextPrimary;
            btn.Font = new Font("Segoe UI", 9f);
            btn.Cursor = Cursors.Hand;
            btn.AutoSize = true;
            btn.Padding = new Padding(Px(4), Px(1), Px(4), Px(1));
            btn.FlatAppearance.BorderColor = UIColors.Border;
            btn.FlatAppearance.MouseOverBackColor = UIColors.SurfaceAlt;
        }
        #endregion

        #region Builders (nav/content)
        private ListBox BuildNavList()
        {
            var lb = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawVariable,
                IntegralHeight = false,
                BackColor = UIColors.NavBg,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
            };
            lb.MeasureItem += (s, e) => { e.ItemHeight = Px(36); };
            lb.Items.AddRange(new object[] { "General", "Charts", "Tables", "Shortcuts" });
            lb.DrawItem += Nav_DrawItem;
            DoubleBuffer(lb);
            return lb;
        }

        private Panel BuildContentHost(out Panel header, out Label desc, out Panel pageContainer)
        {
            var host = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.SurfaceAlt,
                Padding = new Padding(Px(16), Px(12), Px(16), Px(8))
            };

            header = new Panel
            {
                Dock = DockStyle.Top,
                Height = Px(40),
                BackColor = UIColors.SurfaceAlt,
                Padding = new Padding(0, 0, 0, Px(8))
            };
            desc = new Label
            {
                AutoSize = false,
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = UIColors.TextSecondary,
                TextAlign = ContentAlignment.MiddleLeft
            };
            header.Controls.Add(desc);

            pageContainer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.Surface
            };
            Cardize(pageContainer);
            DoubleBuffer(pageContainer);

            host.Controls.Add(pageContainer);
            host.Controls.Add(header);
            _headerDesc = desc;
            return host;
        }

        private Panel BuildGeneralPanel(out CheckBox alignOut, out CheckBox autoFontColorOut, out CheckBox disableF1Out)
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface, Padding = new Padding(Px(16)) };

            var inner = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true
            };

            var chkAlign = new CheckBox
            {
                Text = "  Align shapes to border (including outline)",
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = UIColors.TextPrimary,
                Margin = new Padding(0, 0, 0, Px(8))
            };
            var chkFont = new CheckBox
            {
                Text = "  Enable automatic font color on cell edit",
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = UIColors.TextPrimary,
                Margin = new Padding(0, 0, 0, Px(8))
            };
            var chkF1 = new CheckBox
            {
                Text = "  Disable F1 help key in Excel",
                AutoSize = true,
                Font = new Font("Segoe UI", 9.5f),
                ForeColor = UIColors.TextPrimary,
                Margin = new Padding(0, 0, 0, Px(20))
            };
            inner.Controls.Add(chkAlign);
            inner.Controls.Add(chkFont);
            inner.Controls.Add(chkF1);

            // Shared library section
            _sharedHeader = new Label
            {
                Text = "Shared Library",
                Font = new Font("Segoe UI Semibold", 10f),
                ForeColor = UIColors.TextPrimary,
                AutoSize = true,
                Margin = new Padding(0, Px(8), 0, Px(4))
            };
            inner.Controls.Add(_sharedHeader);

            var sharedDesc = new Label
            {
                Text = "Configure a shared folder for Objects, Slides, and Decks accessible by all users.",
                Font = new Font("Segoe UI", 9f),
                ForeColor = UIColors.TextSecondary,
                AutoSize = true,
                MaximumSize = new System.Drawing.Size(Px(500), 0),
                Margin = new Padding(0, 0, 0, Px(8))
            };
            inner.Controls.Add(sharedDesc);

            _sharedPath = new TextBox
            {
                ReadOnly = true,
                Width = Px(400),
                Font = new Font("Segoe UI", 9f),
                BackColor = UIColors.SurfaceAlt,
                BorderStyle = BorderStyle.FixedSingle,
                Margin = new Padding(0, 0, 0, Px(8))
            };
            inner.Controls.Add(_sharedPath);

            var btnFlow = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                WrapContents = false,
                Margin = new Padding(0, 0, 0, Px(20))
            };

            _btnChangeShared = new Button { Text = "Change…", Margin = new Padding(0, 0, Px(8), 0) };
            StyleActionButton(_btnChangeShared);

            _btnClearShared = new Button { Text = "Clear", Margin = new Padding(0) };
            StyleActionButton(_btnClearShared);

            btnFlow.Controls.Add(_btnChangeShared);
            btnFlow.Controls.Add(_btnClearShared);
            inner.Controls.Add(btnFlow);

            p.Controls.Add(inner);
            alignOut = chkAlign;
            autoFontColorOut = chkFont;
            disableF1Out = chkF1;
            return p;
        }

        private FlowLayoutPanel BuildActionBar()
        {
            return new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(Px(12), Px(8), Px(12), Px(8)),
                Margin = new Padding(0),
                BackColor = UIColors.Surface
            };
        }

        private void StyleListHeader(ListView lv)
        {
            lv.OwnerDraw = true;
            lv.DrawColumnHeader += (s, e) =>
            {
                using (var bg = new SolidBrush(UIColors.HeaderBg))
                    e.Graphics.FillRectangle(bg, e.Bounds);
                TextRenderer.DrawText(e.Graphics, e.Header.Text, lv.Font,
                    new Rectangle(e.Bounds.Left + Px(8), e.Bounds.Top, e.Bounds.Width - Px(8), e.Bounds.Height),
                    UIColors.TextSecondary, TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
                using (var pen = new Pen(UIColors.Border))
                    e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
            };
            lv.DrawItem += (s, e) => e.DrawDefault = true;
            lv.DrawSubItem += (s, e) => e.DrawDefault = true;
        }

        private Panel BuildBrandChartPanel()
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface };

            var actions = BuildActionBar();
            _btnImportChart = new Button { Text = "Import" }; StyleActionButton(_btnImportChart);
            _btnRemoveSecondary = new Button { Text = "Clear" }; StyleActionButton(_btnRemoveSecondary);
            actions.Controls.Add(_btnImportChart);
            actions.Controls.Add(_btnRemoveSecondary);

            _chartCard = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface, Padding = new Padding(Px(12), 0, Px(12), Px(12)) };

            _lvSecondaryCharts = new ListView
            {
                View = View.Details,
                Dock = DockStyle.Fill,
                FullRowSelect = true,
                CheckBoxes = true,
                BorderStyle = BorderStyle.None,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                Font = new Font("Segoe UI", 9f)
            };
            _lvSecondaryCharts.Columns.Add("Base", Px(80), HorizontalAlignment.Center);
            _lvSecondaryCharts.Columns.Add("Chart Type", Px(260));
            _lvSecondaryCharts.Columns.Add("Same As", Px(260));
            DoubleBuffer(_lvSecondaryCharts);
            StyleListHeader(_lvSecondaryCharts);

            _chartCard.Controls.Add(_lvSecondaryCharts);
            p.Controls.Add(_chartCard);
            p.Controls.Add(actions);
            return p;
        }

        private Panel BuildTablePanel()
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface };

            var actions = BuildActionBar();
            _btnImportTable = new Button { Text = "Import" }; StyleActionButton(_btnImportTable);
            _btnClearTable = new Button { Text = "Clear" }; StyleActionButton(_btnClearTable);
            actions.Controls.Add(_btnImportTable);
            actions.Controls.Add(_btnClearTable);

            _tableCard = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface, Padding = new Padding(Px(12), 0, Px(12), Px(12)) };

            _lvTableProps = new ListView
            {
                View = View.Details,
                Dock = DockStyle.Fill,
                FullRowSelect = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9f)
            };
            _lvTableProps.Columns.Add("Property", Px(260));
            _lvTableProps.Columns.Add("Value", Px(520));
            DoubleBuffer(_lvTableProps);
            StyleListHeader(_lvTableProps);

            _tableCard.Controls.Add(_lvTableProps);
            p.Controls.Add(_tableCard);
            p.Controls.Add(actions);
            return p;
        }

        private Panel BuildShortcutsPanel()
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface };

            var bar = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                ColumnCount = 2,
                Height = Px(50),
                Padding = new Padding(Px(12), Px(8), Px(12), Px(8)),
                BackColor = UIColors.Surface
            };
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var help = new Label
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9f),
                ForeColor = UIColors.TextSecondary,
                TextAlign = ContentAlignment.MiddleLeft,
                Text = "Select a row, then press the new keys to change shortcuts."
            };
            _btnResetShortcuts = new Button { Text = "Reset Defaults", Anchor = AnchorStyles.Right };
            StyleActionButton(_btnResetShortcuts);

            bar.Controls.Add(help, 0, 0);
            bar.Controls.Add(_btnResetShortcuts, 1, 0);

            _shortcutsCard = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface, Padding = new Padding(Px(12), 0, Px(12), Px(12)) };

            _lvShortcuts = new ListView
            {
                View = View.Details,
                Dock = DockStyle.Fill,
                FullRowSelect = true,
                BorderStyle = BorderStyle.None,
                GridLines = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                HideSelection = false,
                MultiSelect = false,
                Font = new Font("Segoe UI", 9f)
            };
            _lvShortcuts.Columns.Add("Command", Px(360));
            _lvShortcuts.Columns.Add("Shortcut", Px(360));
            DoubleBuffer(_lvShortcuts);
            StyleListHeader(_lvShortcuts);

            _shortcutsCard.Controls.Add(_lvShortcuts);
            p.Controls.Add(_shortcutsCard);
            p.Controls.Add(bar);
            return p;
        }

        private void Cardize(Panel p)
        {
            p.Paint += (s, e) =>
            {
                using (var pen = new Pen(UIColors.Border))
                {
                    e.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1);
                }
            };
        }

        private void DoubleBuffer(Control c)
        {
            typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                ?.SetValue(c, true, null);
        }
        #endregion

        #region Column fit helpers
        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        private const int GWL_STYLE = -16;
        private const int WS_VSCROLL = 0x00200000;
        private static bool HasVScrollBar(Control c) => (GetWindowLong(c.Handle, GWL_STYLE) & WS_VSCROLL) != 0;

        private void FitTwoColumns(ListView lv, int leftColWidthLogical)
        {
            if (lv.Columns.Count < 2) return;
            int scroll = HasVScrollBar(lv) ? SystemInformation.VerticalScrollBarWidth : 0;
            int left = Px(leftColWidthLogical);
            lv.Columns[0].Width = left;
            lv.Columns[1].Width = Math.Max(Px(80), lv.ClientSize.Width - left - scroll - 1);
        }

        private void FitTwoColumnsEqual(ListView lv, int minEachLogical = 120)
        {
            if (lv.Columns.Count < 2) return;
            int scroll = HasVScrollBar(lv) ? SystemInformation.VerticalScrollBarWidth : 0;
            int available = Math.Max(0, lv.ClientSize.Width - scroll - 1);
            int half = available / 2;
            int minEach = Px(minEachLogical);
            if (half < minEach) half = minEach;
            lv.Columns[0].Width = half;
            lv.Columns[1].Width = available - half;
        }

        private void FitThreeColumnsFillLast(ListView lv, int col0Logical, int col1Logical)
        {
            if (lv.Columns.Count < 3) return;
            int scroll = HasVScrollBar(lv) ? SystemInformation.VerticalScrollBarWidth : 0;
            int c0 = Px(col0Logical);
            int c1 = Px(col1Logical);
            lv.Columns[0].Width = c0;
            lv.Columns[1].Width = c1;
            lv.Columns[2].Width = Math.Max(Px(80), lv.ClientSize.Width - c0 - c1 - scroll - 1);
        }
        #endregion

        #region Shortcuts population & editing
        private void PopulateShortcutsPanelFromRuntimeOrSaved()
        {
            var saved = SettingsManager.Current.Shortcuts?.ToDictionary(s => s.Method, s => s.OnKey, StringComparer.OrdinalIgnoreCase)
                       ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            _lvShortcuts.BeginUpdate();
            _lvShortcuts.Items.Clear();

            for (int i = 0; i < AddInUtilities._shortcuts.Length; i++)
            {
                var sc = AddInUtilities._shortcuts[i];
                if (sc.Method == nameof(AddInUtilities.RunCopyAndStore)) continue;

                var lvi = new ListViewItem(FormatMethodName(sc.Method)) { Tag = i };

                string onKey = sc.Key ?? "";
                if (saved.TryGetValue(sc.Method, out var savedKey) && !string.IsNullOrWhiteSpace(savedKey))
                    onKey = NormalizeOnKey(savedKey);

                var display = string.IsNullOrEmpty(onKey) ? "(None)" : AddInUtilities.FormatOnKeyForDisplay(onKey);
                var sub = new ListViewItem.ListViewSubItem(lvi, display) { Tag = onKey };
                lvi.SubItems.Add(sub);
                _lvShortcuts.Items.Add(lvi);
            }
            _lvShortcuts.EndUpdate();
            FitTwoColumnsEqual(_lvShortcuts);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (_lvShortcuts.ContainsFocus && _lvShortcuts.SelectedItems.Count > 0)
            {
                var keyCode = keyData & Keys.KeyCode;

                // Allow navigation
                if (keyCode == Keys.Up || keyCode == Keys.Down || keyCode == Keys.Left || keyCode == Keys.Right ||
                    keyCode == Keys.PageUp || keyCode == Keys.PageDown || keyCode == Keys.Home || keyCode == Keys.End ||
                    keyCode == Keys.Tab)
                {
                    return base.ProcessCmdKey(ref msg, keyData);
                }

                // Cancel
                if (keyCode == Keys.Escape) { _editItemIndex = -1; return true; }

                // Clear
                if (keyCode == Keys.Delete || keyCode == Keys.Back)
                {
                    ApplyShortcutEdit("");
                    return true;
                }

                // Ignore pure modifiers
                if (keyCode == Keys.ControlKey || keyCode == Keys.ShiftKey || keyCode == Keys.Menu)
                    return true;

                // Ensure edit index tracks selection
                _editItemIndex = _lvShortcuts.SelectedItems[0].Index;

                bool ctrl = (keyData & Keys.Control) == Keys.Control;
                bool shift = (keyData & Keys.Shift) == Keys.Shift;
                bool alt = (keyData & Keys.Alt) == Keys.Alt;

                var mods = new StringBuilder();
                if (ctrl) mods.Append("^");
                if (shift) mods.Append("+");
                if (alt) mods.Append("%");

                string onKey = mods + ConvertKeyCodeToOnKeyString(keyCode);
                onKey = NormalizeOnKey(onKey);
                ApplyShortcutEdit(onKey);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ApplyShortcutEdit(string onKey)
        {
            onKey = NormalizeOnKey(onKey);
            if (_editItemIndex < 0 && _lvShortcuts.SelectedItems.Count > 0)
                _editItemIndex = _lvShortcuts.SelectedItems[0].Index;

            int idx = _editItemIndex;
            if (idx < 0) return;

            // Live duplicate prevention
            if (!string.IsNullOrWhiteSpace(onKey))
            {
                for (int i = 0; i < _lvShortcuts.Items.Count; i++)
                {
                    if (i == idx) continue;
                    var other = _lvShortcuts.Items[i].SubItems[1].Tag as string;
                    if (!string.IsNullOrEmpty(other) && string.Equals(other, onKey, StringComparison.OrdinalIgnoreCase))
                    {
                        string otherCmd = _lvShortcuts.Items[i].SubItems[0].Text;
                        MessageBox.Show(
                            $"The shortcut '{AddInUtilities.FormatOnKeyForDisplay(onKey)}' is already assigned to \"{otherCmd}\".\nPlease choose a different shortcut.",
                            "Duplicate Shortcut", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        System.Media.SystemSounds.Beep.Play();
                        return;
                    }
                }
            }

            var sub = _lvShortcuts.Items[idx].SubItems[1];
            sub.Tag = onKey;
            sub.Text = string.IsNullOrEmpty(onKey) ? "(None)" : AddInUtilities.FormatOnKeyForDisplay(onKey);

            _editItemIndex = -1;
        }

        private void OnResetShortcutsClick(object sender, EventArgs e)
        {
            if (MessageBox.Show("Reset all shortcuts to their default values?\nThis won't be saved until you click OK.",
                "Confirm Reset", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;

            _lvShortcuts.BeginUpdate();
            foreach (ListViewItem item in _lvShortcuts.Items)
            {
                int originalIndex = (int)item.Tag;
                var def = _defaultShortcuts[originalIndex];
                var norm = NormalizeOnKey(def.Key ?? "");
                var sub = item.SubItems[1];
                sub.Text = string.IsNullOrEmpty(norm) ? "(None)" : AddInUtilities.FormatOnKeyForDisplay(norm);
                sub.Tag = norm;
            }
            _lvShortcuts.EndUpdate();
            FitTwoColumnsEqual(_lvShortcuts);
        }

        private string ConvertKeyCodeToOnKeyString(Keys keyCode)
        {
            if (keyCode >= Keys.A && keyCode <= Keys.Z) return keyCode.ToString().ToLower();
            if (keyCode >= Keys.D0 && keyCode <= Keys.D9) return ((char)keyCode).ToString();
            if (keyCode >= Keys.NumPad0 && keyCode <= Keys.NumPad9) return (keyCode - Keys.NumPad0).ToString();
            switch (keyCode)
            {
                case Keys.OemOpenBrackets: return "{[}";
                case Keys.Oem6: return "{]}";
                case Keys.Oemcomma: return ",";
                case Keys.OemPeriod: return ".";
                case Keys.OemSemicolon: return ";";
                case Keys.OemQuotes: return "'";
                case Keys.OemQuestion: return "/";
                case Keys.Space: return "{SPACE}";
                case Keys.Tab: return "{TAB}";
                default: return "{" + keyCode.ToString().ToUpper() + "}";
            }
        }

        private string FormatMethodName(string methodName)
        {
            if (string.IsNullOrEmpty(methodName)) return "";
            if (methodName.StartsWith("Run")) methodName = methodName.Substring(3);
            var sb = new StringBuilder();
            if (methodName.Length > 0)
            {
                sb.Append(methodName[0]);
                for (int i = 1; i < methodName.Length; i++)
                {
                    if (char.IsUpper(methodName[i]) && !char.IsUpper(methodName[i - 1])) sb.Append(' ');
                    sb.Append(methodName[i]);
                }
            }
            return sb.ToString();
        }
        #endregion

        #region Charts / Table logic
        private void PopulateBrandChartPanel()
        {
            var list = CurrentSettings.BrandCharts.SecondaryCharts;
            foreach (var t in _allTypes)
                if (!list.Any(x => x.ChartType == t))
                    list.Add(new ChartFormatSettings { ChartType = t });

            _lvSecondaryCharts.BeginUpdate();
            _lvSecondaryCharts.Items.Clear();

            var ordered = list.OrderByDescending(x => x.IsSet).ThenBy(x => x.ChartType.ToString()).ToList();
            foreach (var s in ordered)
            {
                var lvi = new ListViewItem("", 0) { Checked = s.IsBaseChart, Tag = s, UseItemStyleForSubItems = true };
                lvi.SubItems.Add(s.ChartType.ToString());
                lvi.SubItems.Add(s.SameAsChartType?.ToString() ?? "");
                if (!s.IsSet) lvi.ForeColor = UIColors.TextSecondary;
                _lvSecondaryCharts.Items.Add(lvi);
            }
            _lvSecondaryCharts.EndUpdate();
            FitThreeColumnsFillLast(_lvSecondaryCharts, 80, 260);
        }

        private void OnImportChartClick(object sender, EventArgs e)
        {
            Excel.Chart chartToImport = ChartFormat.GetActiveChart();
            if (chartToImport == null)
            {
                MessageBox.Show("Select a chart in Excel first.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var imported = ChartFormat.ImportSettingsFromChart(chartToImport);
            var list = CurrentSettings.BrandCharts.SecondaryCharts;
            var existing = list.FirstOrDefault(x => x.ChartType == imported.ChartType);

            if (existing != null)
            {
                imported.IsBaseChart = existing.IsBaseChart;
                imported.SameAsChartType = existing.SameAsChartType;
                list[list.IndexOf(existing)] = imported;
            }
            else list.Add(imported);

            PopulateBrandChartPanel();
            MessageBox.Show("Chart format imported.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RemoveSecondary_Click(object sender, EventArgs e)
        {
            if (_lvSecondaryCharts.SelectedItems.Count == 0) return;
            if (!(_lvSecondaryCharts.SelectedItems[0].Tag is ChartFormatSettings s)) return;
            s.IsSet = false; s.IsBaseChart = false; s.SameAsChartType = null; s.SeriesFormats.Clear();
            PopulateBrandChartPanel();
        }

        private void ListView_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            var item = _lvSecondaryCharts.Items[e.Index];
            if (!(item.Tag is ChartFormatSettings s)) return;

            if (!s.IsSet) { e.NewValue = e.CurrentValue; return; }

            if (e.NewValue == CheckState.Checked)
            {
                for (int i = 0; i < _lvSecondaryCharts.Items.Count; i++)
                {
                    if (i == e.Index) continue;
                    _lvSecondaryCharts.Items[i].Checked = false;
                    if (_lvSecondaryCharts.Items[i].Tag is ChartFormatSettings s2) s2.IsBaseChart = false;
                }
                s.IsBaseChart = true;
            }
            else s.IsBaseChart = false;
        }

        private void PopulateTableProps()
        {
            _lvTableProps.BeginUpdate();
            _lvTableProps.Items.Clear();
            var tf = CurrentSettings.TableFormat;

            if (tf == null || !tf.IsSet)
            {
                _lvTableProps.Items.Add(new ListViewItem(new[] { "Status", "No table format saved" }));
                _lvTableProps.EndUpdate();
                FitTwoColumns(_lvTableProps, 260);
                return;
            }

            _lvTableProps.Items.Add(new ListViewItem(new[] { "Name", tf.Name }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Header Fill", ColorHex(tf.Header.FillColor) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Header Font", FontDesc(tf.Header.Font) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Header Borders", BordersDesc(tf.Header.Borders) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Body Fill", ColorHex(tf.Body.FillColor) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Body Font", FontDesc(tf.Body.Font) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Body Borders", BordersDesc(tf.Body.Borders) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Footer Fill", ColorHex(tf.Footer.FillColor) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Footer Font", FontDesc(tf.Footer.Font) }));
            _lvTableProps.Items.Add(new ListViewItem(new[] { "Footer Borders", BordersDesc(tf.Footer.Borders) }));
            _lvTableProps.EndUpdate();
            FitTwoColumns(_lvTableProps, 260);
        }

        private static string ColorHex(Color? c) => c.HasValue ? $"#{c.Value.R:X2}{c.Value.G:X2}{c.Value.B:X2}" : "";
        private static string FontDesc(FontFormat f) => (f == null || !f.IsSet) ? "" : $"{f.Name} {(f.Size ?? 0):0}{(f.Bold == true ? " Bold" : "")}";
        private static string BordersDesc(BorderSet b)
        {
            if (b == null || !b.IsSet) return "";
            string Edge(LineFormat lf) => lf != null && lf.IsSet && lf.Visible != MsoTriState.msoFalse ? "Y" : "-";
            return $"T:{Edge(b.Top)} B:{Edge(b.Bottom)} L:{Edge(b.Left)} R:{Edge(b.Right)} IH:{Edge(b.InsideH)} IV:{Edge(b.InsideV)}";
        }
        #endregion

        #region Shared Library Handlers (JSON-backed)
        private void OnChangeSharedClick(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog()
            {
                Description = "Select a Shared Library folder accessible by all users.\n" +
                              "Subfolders (Objects, Slides, Decks) will be created if needed."
            })
            {
                if (fbd.ShowDialog(this) == DialogResult.OK && IODir.Exists(fbd.SelectedPath))
                {
                    CurrentSettings.SharedLibraryBasePath = fbd.SelectedPath;
                    SettingsManager.Save(CurrentSettings);
                    CreateTypeSubfoldersIfMissing(fbd.SelectedPath);
                    RefreshSharedPathDisplay();
                }
            }
        }

        private void OnClearSharedClick(object sender, EventArgs e)
        {
            CurrentSettings.SharedLibraryBasePath = string.Empty;
            SettingsManager.Save(CurrentSettings);
            RefreshSharedPathDisplay();
        }

        private void RefreshSharedPathDisplay()
        {
            var path = CurrentSettings.SharedLibraryBasePath;
            _sharedPath.Text = string.IsNullOrEmpty(path) ? "(not set)" : path;
        }

        private static void CreateTypeSubfoldersIfMissing(string basePath)
        {
            try
            {
                foreach (var name in new[] { "Objects", "Slides", "Decks" })
                {
                    var p = Path.Combine(basePath, name);
                    if (!IODir.Exists(p)) IODir.CreateDirectory(p);
                }
            }
            catch { /* ignore */ }
        }
        #endregion

        #region Navigation + close
        private void ShowSelectedPanel()
        {
            int idx = _nav.SelectedIndex;
            if (idx < 0 || idx >= _nav.Items.Count) return;

            var name = _nav.Items[idx].ToString();

            _generalPanel.Visible = name == "General";
            _brandChartPanel.Visible = name == "Charts";
            _tablePanel.Visible = name == "Tables";
            _shortcutsPanel.Visible = name == "Shortcuts";

            _cmbSameAs.Visible = false;

            _headerDesc.Text = name == "General" ? "General preferences for alignment, cell font coloring, and shared library configuration."
                           : name == "Charts" ? "Save a base chart and optional secondary styles by chart type."
                           : name == "Tables" ? "Save a table format and review imported properties."
                           : name == "Shortcuts" ? "Customize keyboard shortcuts used by Pragmatics."
                           : "";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                // Final duplicate check
                var shortcutMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (ListViewItem item in _lvShortcuts.Items)
                {
                    string commandName = item.SubItems[0].Text;
                    string onKey = NormalizeOnKey(item.SubItems[1].Tag as string ?? "");
                    if (string.IsNullOrWhiteSpace(onKey)) continue;

                    if (shortcutMap.ContainsKey(onKey))
                    {
                        string other = shortcutMap[onKey];
                        MessageBox.Show(
                            $"The shortcut '{AddInUtilities.FormatOnKeyForDisplay(onKey)}' is used by two commands:\n\n- {other}\n- {commandName}\n\nPlease ensure all shortcuts are unique.",
                            "Duplicate Shortcut", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true; return;
                    }
                    shortcutMap.Add(onKey, commandName);
                }

                // Apply to runtime first (so session has new behavior), then persist
                var utils = new AddInUtilities();
                utils.UnregisterShortcuts();

                CurrentSettings.AlignToBorder = _alignChk.Checked;
                CurrentSettings.AutoFontColorEnabled = _autoFontColorChk.Checked;
                CurrentSettings.DisableF1 = _disableF1Chk.Checked;

                // Update runtime F1 mapping
                try
                {
                    SettingsRuntime.ApplyF1Setting(CurrentSettings.DisableF1);
                }
                catch { /* ignore */ }

                // Rewrite runtime shortcuts and capture to JSON
                var saved = new List<ShortcutBinding>();
                foreach (ListViewItem item in _lvShortcuts.Items)
                {
                    int originalIndex = (int)item.Tag;
                    string newOnKey = NormalizeOnKey(item.SubItems[1].Tag as string ?? "");
                    var originalMethod = AddInUtilities._shortcuts[originalIndex].Method;

                    AddInUtilities._shortcuts[originalIndex] = (newOnKey, originalMethod);
                    saved.Add(new ShortcutBinding { Method = originalMethod, OnKey = newOnKey });
                }
                CurrentSettings.Shortcuts = saved;

                utils.RegisterShortcuts();

                // Persist everything
                SettingsManager.Save(CurrentSettings);
            }
            base.OnFormClosing(e);
        }

        private void Nav_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;
            var lb = (ListBox)sender;
            bool selected = (e.State & DrawItemState.Selected) != 0;

            // Background
            using (var bg = new SolidBrush(UIColors.NavBg))
                e.Graphics.FillRectangle(bg, e.Bounds);

            if (selected)
            {
                var selRect = new Rectangle(e.Bounds.Left, e.Bounds.Top, e.Bounds.Width, e.Bounds.Height);
                using (var white = new SolidBrush(UIColors.NavSelected))
                    e.Graphics.FillRectangle(white, selRect);

                // Accent bar on left
                using (var accent = new SolidBrush(UIColors.TitleBg))
                    e.Graphics.FillRectangle(accent, new Rectangle(selRect.Left, selRect.Top, Px(4), selRect.Height));

                // Border
                using (var pen = new Pen(UIColors.Border))
                    e.Graphics.DrawRectangle(pen, new Rectangle(selRect.Left, selRect.Top, selRect.Width - 1, selRect.Height - 1));
            }

            // Text
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            var textRect = new Rectangle(e.Bounds.Left + Px(16), e.Bounds.Top, e.Bounds.Width - Px(20), e.Bounds.Height);
            TextRenderer.DrawText(e.Graphics, lb.Items[e.Index].ToString(), lb.Font, textRect,
                selected ? UIColors.TextPrimary : UIColors.NavText,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter);
        }

        private static string NormalizeOnKey(string s) => SettingsRuntime.NormalizeOnKey(s);
        #endregion

        #region Helper Methods
        private void BuildButtons(out Button ok, out Button cancel)
        {
            ok = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Size = new System.Drawing.Size(Px(90), Px(30)),
                MinimumSize = new System.Drawing.Size(Px(90), Px(30)),
                Margin = new Padding(0)
            };
            cancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Size = new System.Drawing.Size(Px(90), Px(30)),
                MinimumSize = new System.Drawing.Size(Px(90), Px(30)),
                Margin = new Padding(0)
            };
        }

        private void Lv_MouseDownRemember(object sender, MouseEventArgs e)
        {
            var hit = _lvSecondaryCharts.HitTest(e.Location);
            _hitItemIndex = hit.Item?.Index ?? -1;
            _hitSubIndex = (hit.Item != null && hit.SubItem != null)
                            ? hit.Item.SubItems.IndexOf(hit.SubItem)
                            : -1;
        }

        private void Lv_MouseUpShowCombo(object sender, MouseEventArgs e)
        {
            if (_hitItemIndex < 0 || _hitSubIndex != COL_SAMEAS) return;

            var item = _lvSecondaryCharts.Items[_hitItemIndex];
            if (!(item.Tag is ChartFormatSettings setting)) return;

            var avail = CurrentSettings.BrandCharts.SecondaryCharts
                .Where(x => x.IsSet && x.ChartType != setting.ChartType)
                .Select(x => x.ChartType).Distinct().ToList();

            _cmbSameAs.Items.Clear();
            _cmbSameAs.Items.Add("(None)");
            foreach (var t in avail) _cmbSameAs.Items.Add(t.ToString());

            if (setting.SameAsChartType.HasValue)
            {
                var sel = setting.SameAsChartType.Value.ToString();
                _cmbSameAs.SelectedItem = _cmbSameAs.Items.Contains(sel) ? sel : "(None)";
            }
            else _cmbSameAs.SelectedIndex = 0;

            var subR = item.SubItems[_hitSubIndex].Bounds;
            var lvTopLeft = _lvSecondaryCharts.PointToScreen(Point.Empty);
            var contTopLeft = _chartCard.PointToScreen(Point.Empty);
            int dx = lvTopLeft.X - contTopLeft.X;
            int dy = lvTopLeft.Y - contTopLeft.Y;

            _cmbSameAs.Parent = _chartCard;
            _cmbSameAs.SetBounds(subR.Left + dx, subR.Top + dy, subR.Width, subR.Height);
            _cmbSameAs.Tag = item;
            _cmbSameAs.Visible = true;
            _cmbSameAs.BringToFront();
            _cmbSameAs.Focus();
            _cmbSameAs.DroppedDown = true;
        }

        private void SameAsCommitted(object sender, EventArgs e)
        {
            var item = _cmbSameAs.Tag as ListViewItem;
            if (item == null) return;
            if (!(item.Tag is ChartFormatSettings setting)) return;

            string val = _cmbSameAs.SelectedItem?.ToString() ?? "(None)";
            if (val == "(None)")
            {
                setting.SameAsChartType = null;
                item.SubItems[COL_SAMEAS].Text = "";
            }
            else
            {
                var match = CurrentSettings.BrandCharts.SecondaryCharts
                    .FirstOrDefault(x => x.ChartType.ToString() == val);
                if (match != null)
                {
                    setting.SameAsChartType = match.ChartType;
                    item.SubItems[COL_SAMEAS].Text = val;
                }
            }
            _cmbSameAs.Visible = false;
        }

        private void OnImportTableClick(object sender, EventArgs e)
        {
            try
            {
                var tf = TableFormatter.ImportFromSelection();
                CurrentSettings.TableFormat = tf;
                PopulateTableProps();
                MessageBox.Show("Table format imported.", "Done", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OnClearTableClick(object sender, EventArgs e)
        {
            CurrentSettings.TableFormat = new TableFormatSettings();
            PopulateTableProps();
        }
        #endregion
    }

    /// <summary>
    /// Backwards-compatible façade that now reads/writes the JSON-backed settings.
    /// Keep this so any other code that uses it won't break.
    /// </summary>
    public static class SharedLibrarySettings
    {
        public static string GetSharedBasePath()
        {
            try { return SettingsManager.Current?.SharedLibraryBasePath ?? SettingsManager.Load().SharedLibraryBasePath; }
            catch { return string.Empty; }
        }

        public static void SetSharedBasePath(string path)
        {
            var s = SettingsManager.Current ?? SettingsManager.Load();
            s.SharedLibraryBasePath = path ?? string.Empty;
            SettingsManager.Save(s);
        }
    }
}