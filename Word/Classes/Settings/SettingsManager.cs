// AllSettings.cs (Word version - C# 7.3 / .NET 4.8, VSTO)
// Requires: Newtonsoft.Json

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using IODir = System.IO.Directory;
using IOFile = System.IO.File;

namespace OfficePragmatics_Word.Settings
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
                        Current = loaded;
                        return Current;
                    }
                    catch
                    {
                        // Corrupt file → back up and reset
                        TryMoveBad(JsonPath);
                        _fullJson = null;
                        Current = new AllSettings();
                        Save(Current);
                        return Current;
                    }
                }

                // Migrate from old text file if it exists
                MigrateFromLegacyTextFile();

                _fullJson = null;
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

                JObject jsonToSave;

                if (_fullJson != null)
                {
                    // Merge our settings into the existing JSON, preserving other add-ins' properties
                    jsonToSave = _fullJson;

                    // Update only our properties
                    jsonToSave["SchemaVersion"] = settings.SchemaVersion;
                    jsonToSave["LastSavedUtc"] = settings.LastSavedUtc;
                    jsonToSave["AlignToBorder"] = settings.AlignToBorder;
                    jsonToSave["SharedLibraryBasePath"] = settings.SharedLibraryBasePath;
                }
                else
                {
                    // No existing JSON, create fresh
                    jsonToSave = JObject.FromObject(settings);
                }

                var json = jsonToSave.ToString(Formatting.Indented);
                var tmp = JsonPath + ".tmp";

                // Atomic-ish write
                IOFile.WriteAllText(tmp, json, Encoding.UTF8);
                if (IOFile.Exists(JsonPath))
                    IOFile.Replace(tmp, JsonPath, null);
                else
                    IOFile.Move(tmp, JsonPath);

                _fullJson = jsonToSave;
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
                        Current.SharedLibraryBasePath = path;
                        Save(Current);
                    }
                    // Optionally delete the old file after migration
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
            public static readonly Color Surface = Color.FromArgb(255, 255, 255);
            public static readonly Color SurfaceAlt = Color.FromArgb(248, 248, 248);

            // BRAND COLOR: Deep Navy from "Office Pragmatics" Logo Text
            public static readonly Color TitleBg = Color.FromArgb(32, 42, 102);
            public static readonly Color TitleText = Color.White;

            public static readonly Color Border = Color.FromArgb(220, 220, 220);
            public static readonly Color TextPrimary = Color.FromArgb(32, 31, 30);
            public static readonly Color TextSecondary = Color.FromArgb(96, 94, 92);

            // BRAND COLOR: Periwinkle Blue
            public static readonly Color Accent = Color.FromArgb(110, 125, 220);
            public static readonly Color AccentHover = Color.FromArgb(85, 95, 180);

            public static readonly Color Danger = Color.FromArgb(232, 17, 35);

            // Navigation
            public static readonly Color NavBg = Color.FromArgb(245, 245, 250);
            public static readonly Color NavText = Color.FromArgb(50, 50, 50);
            public static readonly Color NavSelected = Color.White;
        }
        #endregion

        // ✅ Now backed by JSON
        public static AllSettings CurrentSettings { get; private set; } = SettingsManager.Load();

        private float _scale;

        public static void ShowSettingsDialog()
        {
            using (var dlg = new SettingsForm(CurrentSettings))
            {
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

        // Shared library UI
        private Label _sharedHeader;
        private TextBox _sharedPath;
        private Button _btnChangeShared;
        private Button _btnClearShared;

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

            // Populate current values from settings
            _alignChk.Checked = CurrentSettings.AlignToBorder;

            // Events
            _nav.SelectedIndexChanged += (s, e) => ShowSelectedPanel();

            _btnChangeShared.Click += OnChangeSharedClick;
            _btnClearShared.Click += OnClearSharedClick;

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

            // Right: OK / Cancel buttons
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
            _generalPanel = BuildGeneralPanel(out _alignChk);

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
            return p;
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

            _headerDesc.Text = name == "General"
                ? "General preferences for alignment and shared library configuration."
                : "";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                CurrentSettings.AlignToBorder = _alignChk.Checked;

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