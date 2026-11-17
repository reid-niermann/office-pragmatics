// AllSettings.cs (PowerPoint version - C# 7.3 / .NET 4.8, VSTO)
// Requires: Newtonsoft.Json

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
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
using IODir = System.IO.Directory;
using IOFile = System.IO.File;
using OfficePragmatics_Powerpoint;

namespace OfficePragmatics_PowerPoint.Settings
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

        // Shared Library (now lives in JSON)
        public string SharedLibraryBasePath { get; set; } = string.Empty;

        // ------------------------------
        // Footnote Settings
        // ------------------------------
        public string FootnoteFontName { get; set; } = "Arial";
        public float FootnoteFontSize { get; set; } = 8f;

        // Default to "1." (ppBulletArabicPeriod = 1)
        public int FootnoteBulletStyle { get; set; } = 1;

        // Position & Size Defaults (Bottom Leftish)
        public float FootnoteLeft { get; set; } = 50f;
        public float FootnoteTop { get; set; } = 500f;
        public float FootnoteWidth { get; set; } = 400f;
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

        // Current is backed by a field to support lazy loading
        private static AllSettings _current;
        public static AllSettings Current
        {
            get
            {
                lock (_lock)
                {
                    if (_current == null)
                    {
                        // If we haven't loaded yet, try to load from disk
                        Load();
                    }
                    return _current;
                }
            }
            private set => _current = value;
        }

        // Store the full JSON object to preserve properties from other add-ins
        private static JObject _fullJson = null;

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
                        _fullJson = JObject.Parse(json);

                        // Deserialize only the properties we care about
                        var loaded = _fullJson.ToObject<AllSettings>() ?? new AllSettings();
                        _current = loaded; // Direct assignment to backing field
                        return _current;
                    }
                    catch
                    {
                        // Corrupt file → back up and reset
                        TryMoveBad(JsonPath);
                        _fullJson = null;
                        _current = new AllSettings();
                        Save(_current);
                        return _current;
                    }
                }

                // Migrate from old text file if it exists
                MigrateFromLegacyTextFile();

                _fullJson = null;

                // If _current is still null after migration attempts, create fresh defaults
                if (_current == null) _current = new AllSettings();

                Save(_current);
                return _current;
            }
        }

        public static void Save(AllSettings settings)
        {
            lock (_lock)
            {
                EnsureDirectory();
                settings.LastSavedUtc = DateTime.UtcNow;

                JObject jsonToSave;

                if (_fullJson != null)
                {
                    // Merge our settings into the existing JSON
                    jsonToSave = _fullJson;

                    jsonToSave["SchemaVersion"] = settings.SchemaVersion;
                    jsonToSave["LastSavedUtc"] = settings.LastSavedUtc;
                    jsonToSave["AlignToBorder"] = settings.AlignToBorder;
                    jsonToSave["SharedLibraryBasePath"] = settings.SharedLibraryBasePath;

                    // Footnotes
                    jsonToSave["FootnoteFontName"] = settings.FootnoteFontName;
                    jsonToSave["FootnoteFontSize"] = settings.FootnoteFontSize;
                    jsonToSave["FootnoteBulletStyle"] = settings.FootnoteBulletStyle;
                    jsonToSave["FootnoteLeft"] = settings.FootnoteLeft;
                    jsonToSave["FootnoteTop"] = settings.FootnoteTop;
                    jsonToSave["FootnoteWidth"] = settings.FootnoteWidth;
                }
                else
                {
                    jsonToSave = JObject.FromObject(settings);
                }

                var json = jsonToSave.ToString(Formatting.Indented);
                var tmp = JsonPath + ".tmp";

                IOFile.WriteAllText(tmp, json, Encoding.UTF8);
                if (IOFile.Exists(JsonPath))
                    IOFile.Replace(tmp, JsonPath, null);
                else
                    IOFile.Move(tmp, JsonPath);

                _fullJson = jsonToSave;
                _current = settings;
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

        private static void MigrateFromLegacyTextFile()
        {
            try
            {
                var legacyFile = Path.Combine(AppDataDir, "SharedLibraryPath.txt");
                if (IOFile.Exists(legacyFile))
                {
                    var path = IOFile.ReadAllText(legacyFile, Encoding.UTF8).Trim();
                    if (!string.IsNullOrEmpty(path))
                    {
                        // Temporarily create a current instance if null to perform migration
                        if (_current == null) _current = new AllSettings();
                        _current.SharedLibraryBasePath = path;
                        Save(_current);
                    }
                    try { IOFile.Delete(legacyFile); } catch { /* ignore */ }
                }
            }
            catch { /* ignore migration errors */ }
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
            public static readonly System.Drawing.Color Surface = System.Drawing.Color.FromArgb(255, 255, 255);
            public static readonly System.Drawing.Color SurfaceAlt = System.Drawing.Color.FromArgb(248, 248, 248);

            // BRAND COLOR: Deep Navy from "Office Pragmatics" Logo Text
            public static readonly System.Drawing.Color TitleBg = System.Drawing.Color.FromArgb(32, 42, 102);
            public static readonly System.Drawing.Color TitleText = System.Drawing.Color.White;

            public static readonly System.Drawing.Color Border = System.Drawing.Color.FromArgb(220, 220, 220);
            public static readonly System.Drawing.Color TextPrimary = System.Drawing.Color.FromArgb(32, 31, 30);
            public static readonly System.Drawing.Color TextSecondary = System.Drawing.Color.FromArgb(96, 94, 92);

            // BRAND COLOR: Periwinkle Blue
            public static readonly System.Drawing.Color Accent = System.Drawing.Color.FromArgb(110, 125, 220);
            public static readonly System.Drawing.Color AccentHover = System.Drawing.Color.FromArgb(85, 95, 180);

            public static readonly System.Drawing.Color Danger = System.Drawing.Color.FromArgb(232, 17, 35);

            // Navigation
            public static readonly System.Drawing.Color NavBg = System.Drawing.Color.FromArgb(245, 245, 250);
            public static readonly System.Drawing.Color NavText = System.Drawing.Color.FromArgb(50, 50, 50);
            public static readonly System.Drawing.Color NavSelected = System.Drawing.Color.White;
        }
        #endregion

        // ✅ Now backed by JSON
        public static AllSettings CurrentSettings { get; private set; } = SettingsManager.Load();

        private float _scale;

        public const string ApiKeysPageName = "API Keys";

        public static void ShowSettingsDialog(string startPage = null)
        {
            using (var dlg = new SettingsForm(CurrentSettings))
            {
                if (!string.IsNullOrEmpty(startPage))
                {
                    int idx = dlg._nav.Items.IndexOf(startPage);
                    if (idx >= 0) dlg._nav.SelectedIndex = idx;
                }
                dlg.ShowDialog();
            }
        }

        #region Fields
        private ListBox _nav;
        private Panel _contentHost, _header, _pageContainer;
        private Panel _titleBar;
        private Label _btnClose;

        private Label _headerDesc;

        private Button _btnOK, _btnCancel;


        // General
        private Panel _generalPanel;
        private CheckBox _alignChk;
        private TextBox _sharedPath;
        private Button _btnChangeShared;
        private Button _btnClearShared;


        // Footnotes UI
        private Panel _footnotesPanel;
        private Label _lblFootnoteFont;
        private Button _btnChangeFootnoteFont;
        private NumericUpDown _numFootnoteSize;
        private ComboBox _cmbFootnoteStyle;
        private Label _lblFootnotePos;
        private Button _btnCaptureShape;

        // API Keys UI
        private Panel _apiKeysPanel;
        private TextBox _txtBraveApiKey;
        private TextBox _txtFreepikApiKey;
        private TextBox _txtMapboxToken;
        private CheckBox _chkShowKeys;
        #endregion

        #region Init
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

            InitializeFormChrome();

            // Populate General
            _alignChk.Checked = CurrentSettings.AlignToBorder;

            // Populate Footnotes
            RefreshFootnoteFontLabel();
            InitializeBulletCombo();
            RefreshFootnotePosLabel();
            _numFootnoteSize.Value = (decimal)CurrentSettings.FootnoteFontSize;

            // Populate API Keys
            _txtBraveApiKey.Text = ApiKeys.BraveSearchApiKey;
            _txtFreepikApiKey.Text = ApiKeys.FreepikApiKey;
            _txtMapboxToken.Text = ApiKeys.MapboxAccessToken;

            // Events
            _nav.SelectedIndexChanged += (s, e) => ShowSelectedPanel();

            _btnChangeShared.Click += OnChangeSharedClick;
            _btnClearShared.Click += OnClearSharedClick;

            // Footnote Events
            _btnChangeFootnoteFont.Click += OnChangeFootnoteFontClick;
            _btnCaptureShape.Click += OnCaptureShapeClick;
            _numFootnoteSize.ValueChanged += (s, e) => {
                CurrentSettings.FootnoteFontSize = (float)_numFootnoteSize.Value;
                RefreshFootnoteFontLabel();
            };

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
            _btnClose.MouseEnter += (s, e) => { _btnClose.BackColor = UIColors.Danger; _btnClose.ForeColor = System.Drawing.Color.White; };
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
                BackColor = System.Drawing.Color.Transparent,
                Margin = new Padding(0)
            };
            bottomLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            bottomLayout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            // Right: OK / Cancel buttons
            var buttonsTable = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = System.Drawing.Color.Transparent,
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
            _generalPanel = BuildGeneralPanel(out _alignChk);
            _footnotesPanel = BuildFootnotesPanel();
            _apiKeysPanel = BuildApiKeysPanel();

            _pageContainer.Controls.Add(_apiKeysPanel);
            _pageContainer.Controls.Add(_footnotesPanel);
            _pageContainer.Controls.Add(_generalPanel);

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

        #region Builders
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
            lb.Items.Add("General");
            lb.Items.Add("Footnotes");
            lb.Items.Add(ApiKeysPageName);
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

        private Panel BuildGeneralPanel(out CheckBox alignOut)
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
                Margin = new Padding(0, 0, 0, Px(20))
            };
            inner.Controls.Add(chkAlign);

            // Shared library section
            AddSectionHeader(inner, "Shared Library");
            AddDescLabel(inner, "Configure a shared folder for Objects, Slides, and Decks.");

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
            return p;
        }

        private Panel BuildFootnotesPanel()
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface, Padding = new Padding(Px(16)) };
            var inner = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };

            // 1. Font Settings
            AddSectionHeader(inner, "Default Font");
            _lblFootnoteFont = AddDescLabel(inner, "Current: Arial, 8pt");
            _lblFootnoteFont.Margin = new Padding(0, 0, 0, Px(10));

            // Use a FlowLayoutPanel for horizontal alignment
            var fontFlow = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                Margin = new Padding(0, 0, 0, Px(20)),
                WrapContents = false
            };

            _btnChangeFootnoteFont = new Button { Text = "Select Font Family..." };
            StyleActionButton(_btnChangeFootnoteFont);
            // Add right margin to separate from the label
            _btnChangeFootnoteFont.Margin = new Padding(0, 0, Px(10), 0);
            fontFlow.Controls.Add(_btnChangeFootnoteFont);

            var lblSize = new Label
            {
                Text = "Size (pt):",
                AutoSize = true,
                Font = new Font("Segoe UI", 9f),
                TextAlign = ContentAlignment.MiddleLeft,
                // Add top padding to align text baseline with the button
                Margin = new Padding(0, Px(5), Px(5), 0)
            };
            fontFlow.Controls.Add(lblSize);

            _numFootnoteSize = new NumericUpDown
            {
                Minimum = 5,
                Maximum = 24,
                DecimalPlaces = 1,
                Width = Px(60),
                Font = new Font("Segoe UI", 9f)
            };
            // Add slight top margin to align the box with the button
            _numFootnoteSize.Margin = new Padding(0, Px(3), 0, 0);
            fontFlow.Controls.Add(_numFootnoteSize);

            inner.Controls.Add(fontFlow);

            // 2. List Format
            AddSectionHeader(inner, "List Format");
            AddDescLabel(inner, "Choose the numbering style for footnotes.");

            _cmbFootnoteStyle = new ComboBox
            {
                Width = Px(250),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9.5f),
                Margin = new Padding(0, 0, 0, Px(20))
            };
            inner.Controls.Add(_cmbFootnoteStyle);

            // 3. Location & Size
            AddSectionHeader(inner, "Location & Size");
            AddDescLabel(inner, "Select a shape in PowerPoint and click Capture to save its location and width as the default for new footnotes.");

            _lblFootnotePos = new Label
            {
                AutoSize = true,
                Font = new Font("Consolas", 9f),
                ForeColor = UIColors.TextSecondary,
                Margin = new Padding(0, Px(5), 0, Px(10))
            };
            inner.Controls.Add(_lblFootnotePos);

            _btnCaptureShape = new Button { Text = "Capture from Selected Shape" };
            StylePrimaryButton(_btnCaptureShape);
            _btnCaptureShape.AutoSize = true;
            _btnCaptureShape.Padding = new Padding(Px(10), Px(5), Px(10), Px(5));
            inner.Controls.Add(_btnCaptureShape);

            p.Controls.Add(inner);
            return p;
        }

        private Panel BuildApiKeysPanel()
        {
            var p = new Panel { Dock = DockStyle.Fill, BackColor = UIColors.Surface, Padding = new Padding(Px(16)) };
            var inner = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoScroll = true };

            AddDescLabel(inner,
                "Online features call these services directly using your own keys. " +
                "Keys are stored encrypted for your Windows user account. Leave a field blank to turn that feature off.");

            // Logo search (Brave Search)
            AddSectionHeader(inner, "Logo Search (Brave Search)");
            AddLinkLabel(inner,
                "Used by Insert Logo. ", null,
                "Get a Brave Search API key", ApiKeys.BraveSearchApiKeyUrl,
                " (the Search plan includes image search).", null);
            _txtBraveApiKey = AddKeyField(inner, "API key", secret: true);

            // Icon search (Freepik)
            AddSectionHeader(inner, "Icon Search (Freepik)");
            AddLinkLabel(inner,
                "Used by Insert Icon. ", null,
                "Get a Freepik API key", ApiKeys.FreepikApiKeyUrl,
                ".", null);
            _txtFreepikApiKey = AddKeyField(inner, "API key", secret: true);

            // Maps (Mapbox)
            AddSectionHeader(inner, "Maps (Mapbox)");
            AddLinkLabel(inner,
                "Used by Plot Map to geocode addresses. ", null,
                "Get a Mapbox access token", ApiKeys.MapboxTokenUrl,
                ".", null);
            _txtMapboxToken = AddKeyField(inner, "Access token", secret: true);

            _chkShowKeys = new CheckBox
            {
                Text = "  Show keys",
                AutoSize = true,
                Font = new Font("Segoe UI", 9f),
                ForeColor = UIColors.TextPrimary,
                Margin = new Padding(0, Px(12), 0, 0)
            };
            _chkShowKeys.CheckedChanged += (s, e) =>
            {
                bool hide = !_chkShowKeys.Checked;
                _txtBraveApiKey.UseSystemPasswordChar = hide;
                _txtFreepikApiKey.UseSystemPasswordChar = hide;
                _txtMapboxToken.UseSystemPasswordChar = hide;
            };
            inner.Controls.Add(_chkShowKeys);

            p.Controls.Add(inner);
            return p;
        }

        private TextBox AddKeyField(Control parent, string label, bool secret)
        {
            parent.Controls.Add(new Label
            {
                Text = label,
                Font = new Font("Segoe UI", 9f),
                ForeColor = UIColors.TextPrimary,
                AutoSize = true,
                Margin = new Padding(0, Px(4), 0, Px(2))
            });

            var tb = new TextBox
            {
                Width = Px(400),
                Font = new Font("Segoe UI", 9f),
                BorderStyle = BorderStyle.FixedSingle,
                UseSystemPasswordChar = secret,
                Margin = new Padding(0, 0, 0, Px(6))
            };
            parent.Controls.Add(tb);
            return tb;
        }

        /// <summary>
        /// Adds a description line made of (text, url) pairs. A pair with a null url is
        /// plain text; a pair with a url becomes a clickable link that opens the browser.
        /// </summary>
        private void AddLinkLabel(Control parent, params string[] textUrlPairs)
        {
            var link = new LinkLabel
            {
                Font = new Font("Segoe UI", 9f),
                ForeColor = UIColors.TextSecondary,
                LinkColor = UIColors.Accent,
                ActiveLinkColor = UIColors.AccentHover,
                AutoSize = true,
                MaximumSize = new System.Drawing.Size(Px(500), 0),
                Margin = new Padding(0, 0, 0, Px(8))
            };

            var sb = new StringBuilder();
            var links = new List<Tuple<int, int, string>>();
            for (int i = 0; i + 1 < textUrlPairs.Length; i += 2)
            {
                string text = textUrlPairs[i] ?? string.Empty;
                string url = textUrlPairs[i + 1];
                if (!string.IsNullOrEmpty(url)) links.Add(Tuple.Create(sb.Length, text.Length, url));
                sb.Append(text);
            }

            link.Text = sb.ToString();
            link.Links.Clear();
            foreach (var l in links) link.Links.Add(l.Item1, l.Item2, l.Item3);

            link.LinkClicked += (s, e) =>
            {
                try { System.Diagnostics.Process.Start(e.Link.LinkData as string); } catch { /* no browser */ }
            };
            parent.Controls.Add(link);
        }

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

        private void AddSectionHeader(Control parent, string text)
        {
            var l = new Label
            {
                Text = text,
                Font = new Font("Segoe UI Semibold", 10f),
                ForeColor = UIColors.TextPrimary,
                AutoSize = true,
                Margin = new Padding(0, Px(20), 0, Px(4))
            };
            parent.Controls.Add(l);
        }

        private Label AddDescLabel(Control parent, string text)
        {
            var l = new Label
            {
                Text = text,
                Font = new Font("Segoe UI", 9f),
                ForeColor = UIColors.TextSecondary,
                AutoSize = true,
                MaximumSize = new System.Drawing.Size(Px(500), 0),
                Margin = new Padding(0, 0, 0, Px(8))
            };
            parent.Controls.Add(l);
            return l;
        }
        #endregion

        #region Footnote Logic & Events
        private class BulletOption
        {
            public string Display { get; set; }
            public Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle Value { get; set; }
            public override string ToString() => Display;
        }

        private void InitializeBulletCombo()
        {
            var opts = new List<BulletOption>
            {
                new BulletOption { Display = "1. 2. 3.", Value = Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle.ppBulletArabicPeriod },
                new BulletOption { Display = "(1) (2) (3)", Value = Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle.ppBulletArabicParenBoth },
                new BulletOption { Display = "a) b) c)", Value = Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle.ppBulletAlphaLCParenRight },
                new BulletOption { Display = "(a) (b) (c)", Value = Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle.ppBulletAlphaLCParenBoth },
                new BulletOption { Display = "A) B) C)", Value = Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle.ppBulletAlphaUCParenRight },
                new BulletOption { Display = "(A) (B) (C)", Value = Microsoft.Office.Interop.PowerPoint.PpNumberedBulletStyle.ppBulletAlphaUCParenBoth },
            };

            _cmbFootnoteStyle.DataSource = opts;

            // Select current setting
            var current = opts.FirstOrDefault(x => (int)x.Value == CurrentSettings.FootnoteBulletStyle);
            if (current != null) _cmbFootnoteStyle.SelectedItem = current;
            else if (opts.Count > 0) _cmbFootnoteStyle.SelectedIndex = 0;
        }

        private void OnChangeFootnoteFontClick(object sender, EventArgs e)
        {
            using (var fd = new FontDialog())
            {
                fd.ShowColor = false;
                fd.ShowEffects = false;
                try { fd.Font = new Font(CurrentSettings.FootnoteFontName, CurrentSettings.FootnoteFontSize); }
                catch { fd.Font = new Font("Arial", 8f); }

                if (fd.ShowDialog(this) == DialogResult.OK)
                {
                    CurrentSettings.FootnoteFontName = fd.Font.Name;
                    RefreshFootnoteFontLabel();
                }
            }
        }

        private void RefreshFootnoteFontLabel()
        {
            if (_lblFootnoteFont != null)
                _lblFootnoteFont.Text = $"Current: {CurrentSettings.FootnoteFontName}, {CurrentSettings.FootnoteFontSize}pt";
        }

        private void RefreshFootnotePosLabel()
        {
            if (_lblFootnotePos != null)
            {
                _lblFootnotePos.Text = $"Left: {CurrentSettings.FootnoteLeft:F1}, Top: {CurrentSettings.FootnoteTop:F1}, Width: {CurrentSettings.FootnoteWidth:F1}";
            }
        }

        private void OnCaptureShapeClick(object sender, EventArgs e)
        {
            try
            {
                // Access PowerPoint via Globals
                var app = OfficePragmatics_Powerpoint.Globals.ThisAddIn.Application;
                var win = app.ActiveWindow;

                if (win != null && win.Selection.Type == Microsoft.Office.Interop.PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    if (win.Selection.ShapeRange.Count > 0)
                    {
                        var shp = win.Selection.ShapeRange[1]; // 1-based index

                        CurrentSettings.FootnoteLeft = shp.Left;
                        CurrentSettings.FootnoteTop = shp.Top;
                        CurrentSettings.FootnoteWidth = shp.Width;

                        RefreshFootnotePosLabel();

                        MessageBox.Show(this, "Location and size captured successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                MessageBox.Show(this, "Please select a Shape in PowerPoint to capture its position.", "No Shape Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, $"Error capturing shape: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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

            _generalPanel.Visible = false;
            _footnotesPanel.Visible = false;
            _apiKeysPanel.Visible = false;

            if (name == "General")
            {
                _generalPanel.Visible = true;
                _headerDesc.Text = "General preferences for alignment and shared library configuration.";
            }
            else if (name == "Footnotes")
            {
                _footnotesPanel.Visible = true;
                _headerDesc.Text = "Configure default styles, font, and positioning for auto-inserted footnotes.";
            }
            else if (name == ApiKeysPageName)
            {
                _apiKeysPanel.Visible = true;
                _headerDesc.Text = "Your own API keys for logo search, icon search, and maps.";
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                CurrentSettings.AlignToBorder = _alignChk.Checked;

                // Save Combo Selection
                if (_cmbFootnoteStyle.SelectedItem is BulletOption opt)
                {
                    CurrentSettings.FootnoteBulletStyle = (int)opt.Value;
                }

                // Font Size is updated in the ValueChanged event, 
                // so we don't need to force set it here unless we want to be safe.
                CurrentSettings.FootnoteFontSize = (float)_numFootnoteSize.Value;

                // Persist everything
                SettingsManager.Save(CurrentSettings);

                try
                {
                    ApiKeys.BraveSearchApiKey = _txtBraveApiKey.Text;
                    ApiKeys.FreepikApiKey = _txtFreepikApiKey.Text;
                    ApiKeys.MapboxAccessToken = _txtMapboxToken.Text;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, $"Could not save API keys:\n{ex.Message}", "Settings",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
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