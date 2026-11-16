using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Runtime.InteropServices;
using OfficePragmatics_Excel.Settings;
using IOFile = System.IO.File;
using IODir = System.IO.Directory;
using SD = System.Drawing;

namespace Office_Pragmatics_Word
{
    public enum LibraryAssetType
    {
        Object,
        Page,
        Document
    }

    public class LibraryAsset
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public string ThumbnailPath { get; set; }
        public LibraryAssetType AssetType { get; set; }
        public string UniqueID { get; set; }

        public override string ToString()
        {
            return Name;
        }
    }

    internal static class SharedLibrarySettings
    {
        private static readonly string SettingsDir =
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficePragmatics");

        private static readonly string SettingsFile =
            Path.Combine(SettingsDir, "SharedLibraryPath.txt");

        public static string GetSharedBasePath()
        {
            try
            {
                if (IOFile.Exists(SettingsFile))
                {
                    var path = (IOFile.ReadAllText(SettingsFile) ?? "").Trim();
                    if (!string.IsNullOrEmpty(path) && IODir.Exists(path)) return path;
                }
            }
            catch { /* ignore */ }
            return null;
        }

        public static void SetSharedBasePath(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            try
            {
                if (!IODir.Exists(SettingsDir)) IODir.CreateDirectory(SettingsDir);
                IOFile.WriteAllText(SettingsFile, path.Trim());
            }
            catch { /* ignore write errors */ }
        }
    }

    public partial class LibraryForm : Form
    {
        // ── NATIVE INTEROP FOR DRAGGABLE TITLE BAR ──
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HT_CAPTION = 0x2;

        // ── MODERN THEME COLORS (matching PowerPoint LibraryForm) ──
        private static class UIColors
        {
            public static readonly SD.Color Surface = SD.Color.FromArgb(255, 255, 255);
            public static readonly SD.Color SurfaceAlt = SD.Color.FromArgb(248, 248, 248);
            public static readonly SD.Color SurfaceDark = SD.Color.FromArgb(240, 240, 240);

            // Deep Navy from brand
            public static readonly SD.Color TitleBg = SD.Color.FromArgb(32, 42, 102);
            public static readonly SD.Color TitleText = SD.Color.White;

            public static readonly SD.Color Border = SD.Color.FromArgb(220, 220, 220);
            public static readonly SD.Color TextPrimary = SD.Color.FromArgb(32, 31, 30);
            public static readonly SD.Color TextSecondary = SD.Color.FromArgb(96, 94, 92);

            // Periwinkle Accent
            public static readonly SD.Color Accent = SD.Color.FromArgb(110, 125, 220);
            public static readonly SD.Color AccentHover = SD.Color.FromArgb(85, 95, 180);
            public static readonly SD.Color AccentLight = SD.Color.FromArgb(235, 238, 250);

            public static readonly SD.Color Danger = SD.Color.FromArgb(232, 17, 35);

            // List selection
            public static readonly SD.Color ListHover = SD.Color.FromArgb(245, 246, 250);
            public static readonly SD.Color ListSelected = SD.Color.FromArgb(232, 235, 248);
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                this.DialogResult = DialogResult.Cancel;
                this.Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private LibraryAssetType _currentAssetType;
        private List<LibraryAsset> _allAssets;
        private List<LibraryAsset> _filteredAssets;
        private Word.Application _wordApplication;

        // Controls
        private Panel pnlTitleBar;
        private Panel pnlMain;
        private Panel pnlTabs;
        private ListBox listBoxAssets;
        private PictureBox pictureBoxThumbnail;
        private Label lblAssetName;
        private TextBox txtSearch;
        private Button btnSaveNew;
        private Button btnReplaceSelected;
        private Button btnInsert;
        private Button btnDelete;
        private Button btnRename;
        private Label lblStatus;

        private bool _useSharedLibrary = false;
        private CheckBox chkSharedLibrary;
        private int _hoveredIndex = -1;
        private int _selectedTabIndex = 0;

        // DPI scaling
        private float _scale = 1f;

        public LibraryForm(Word.Application wordApp, LibraryAssetType? initialAssetType = null)
        {
            _wordApplication = wordApp;
            _currentAssetType = initialAssetType ?? LibraryAssetType.Object;

            // DPI scaling
            using (var g = CreateGraphics()) { _scale = g.DpiX / 96f; }

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = UIColors.Border;
            this.Padding = new Padding(1);

            InitializeComponentManual();
            InitializeFormSpecifics();

            if (initialAssetType.HasValue)
            {
                _selectedTabIndex = (int)initialAssetType.Value;
                _currentAssetType = initialAssetType.Value;
            }

            LoadAssets();
            this.Shown += LibraryForm_Shown;
        }

        private int Scale(int value) => (int)(value * _scale);

        private void LibraryForm_Shown(object sender, EventArgs e)
        {
            this.CenterToScreen();
            pnlTabs?.Invalidate();
        }

        private void InitializeComponentManual()
        {
            this.SuspendLayout();

            this.ClientSize = new SD.Size(Scale(780), Scale(560));
            this.MinimumSize = new SD.Size(Scale(600), Scale(450));

            // Main container (inside border)
            pnlMain = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.Surface
            };
            this.Controls.Add(pnlMain);

            // ══════════════════════════════════════════════════════════
            // LAYOUT ORDER (add in reverse dock order)
            // ══════════════════════════════════════════════════════════

            // ── STATUS BAR (Dock=Bottom) ──
            var pnlStatus = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = Scale(32),
                BackColor = UIColors.Surface,
                Padding = new Padding(Scale(15), 0, Scale(15), 0)
            };
            pnlStatus.Paint += (s, e) =>
            {
                using (var p = new SD.Pen(UIColors.Border))
                    e.Graphics.DrawLine(p, 0, 0, pnlStatus.Width, 0);
            };

            lblStatus = new Label
            {
                Dock = DockStyle.Fill,
                TextAlign = SD.ContentAlignment.MiddleLeft,
                Font = new SD.Font("Segoe UI", 9f),
                ForeColor = UIColors.TextSecondary
            };
            pnlStatus.Controls.Add(lblStatus);
            pnlMain.Controls.Add(pnlStatus);

            // ── BUTTON BAR (Dock=Bottom) ──
            var pnlButtons = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = Scale(65),
                BackColor = UIColors.SurfaceAlt,
                Padding = new Padding(Scale(15), Scale(12), Scale(15), Scale(12))
            };
            pnlButtons.Paint += (s, e) =>
            {
                using (var p = new SD.Pen(UIColors.Border))
                    e.Graphics.DrawLine(p, 0, 0, pnlButtons.Width, 0);
            };

            var flowButtons = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = SD.Color.Transparent
            };

            var btnSize = new SD.Size(Scale(100), Scale(36));

            btnSaveNew = CreateModernButton("New", btnSize, true);
            btnSaveNew.Click += btnSaveNew_Click;

            btnInsert = CreateModernButton("Insert", btnSize, true);
            btnInsert.Click += btnInsert_Click;

            btnDelete = CreateModernButton("Delete", btnSize, false);
            btnDelete.Click += btnDelete_Click;

            btnReplaceSelected = CreateModernButton("Replace", btnSize, false);
            btnReplaceSelected.Click += btnReplaceSelected_Click;

            btnRename = CreateModernButton("Rename", btnSize, false);
            btnRename.Click += btnRename_Click;

            flowButtons.Controls.Add(btnSaveNew);
            flowButtons.Controls.Add(btnInsert);
            flowButtons.Controls.Add(btnDelete);
            flowButtons.Controls.Add(btnReplaceSelected);
            flowButtons.Controls.Add(btnRename);

            pnlButtons.Controls.Add(flowButtons);
            pnlMain.Controls.Add(pnlButtons);

            // ── CONTENT AREA ──
            var pnlContent = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.SurfaceAlt,
                Padding = new Padding(Scale(10), Scale(10), Scale(10), Scale(5))
            };

            // Split: Left list, Right preview
            var pnlListContainer = new Panel
            {
                Dock = DockStyle.Left,
                Width = Scale(240),
                BackColor = UIColors.Surface,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            pnlListContainer.Paint += (s, e) =>
            {
                using (var p = new SD.Pen(UIColors.Border))
                    e.Graphics.DrawRectangle(p, 0, 0, pnlListContainer.Width - 1, pnlListContainer.Height - 1);
            };

            listBoxAssets = new ListBox
            {
                Dock = DockStyle.Fill,
                Font = new SD.Font("Segoe UI", 9.5f),
                BackColor = UIColors.Surface,
                ForeColor = UIColors.TextPrimary,
                BorderStyle = BorderStyle.None,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = Scale(38),
                IntegralHeight = false,
                Margin = new Padding(1)
            };
            listBoxAssets.DrawItem += ListBoxAssets_DrawItem;
            listBoxAssets.SelectedIndexChanged += listBoxAssets_SelectedIndexChanged;
            listBoxAssets.MouseMove += ListBoxAssets_MouseMove;
            listBoxAssets.MouseLeave += (s, e) => { _hoveredIndex = -1; listBoxAssets.Invalidate(); };

            pnlListContainer.Controls.Add(listBoxAssets);

            // Splitter
            var splitter = new Splitter
            {
                Dock = DockStyle.Left,
                Width = Scale(8),
                MinSize = Scale(180),
                BackColor = UIColors.SurfaceAlt
            };

            // Preview panel
            var pnlPreviewOuter = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.Surface
            };
            pnlPreviewOuter.Paint += (s, e) =>
            {
                using (var p = new SD.Pen(UIColors.Border))
                    e.Graphics.DrawRectangle(p, 0, 0, pnlPreviewOuter.Width - 1, pnlPreviewOuter.Height - 1);
            };

            lblAssetName = new Label
            {
                Dock = DockStyle.Bottom,
                Height = Scale(28),
                Font = new SD.Font("Segoe UI Semibold", 10f),
                ForeColor = UIColors.TextPrimary,
                TextAlign = SD.ContentAlignment.MiddleLeft,
                AutoEllipsis = true,
                BackColor = UIColors.Surface,
                Padding = new Padding(Scale(8), 0, Scale(8), 0)
            };

            var pnlPictureContainer = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = UIColors.SurfaceAlt,
                Padding = new Padding(Scale(8))
            };

            pictureBoxThumbnail = new PictureBox
            {
                Dock = DockStyle.Fill,
                SizeMode = PictureBoxSizeMode.Zoom,
                BackColor = UIColors.SurfaceAlt
            };

            pnlPictureContainer.Controls.Add(pictureBoxThumbnail);
            pnlPreviewOuter.Controls.Add(pnlPictureContainer);

            pnlContent.Controls.Add(pnlPreviewOuter);
            pnlContent.Controls.Add(splitter);
            pnlContent.Controls.Add(pnlListContainer);

            pnlMain.Controls.Add(pnlContent);

            // ── SEARCH BAR (Dock=Top) ──
            var pnlSearch = new Panel
            {
                Dock = DockStyle.Top,
                Height = Scale(50),
                BackColor = UIColors.Surface,
                Padding = new Padding(Scale(15), Scale(8), Scale(15), Scale(8))
            };
            pnlSearch.Paint += (s, e) =>
            {
                using (var p = new SD.Pen(UIColors.Border))
                    e.Graphics.DrawLine(p, 0, pnlSearch.Height - 1, pnlSearch.Width, pnlSearch.Height - 1);
            };

            var lblSearchLabel = new Label
            {
                Dock = DockStyle.Left,
                Width = Scale(55),
                Text = "Search:",
                Font = new SD.Font("Segoe UI", 9.5f),
                TextAlign = SD.ContentAlignment.MiddleLeft,
                ForeColor = UIColors.TextSecondary
            };

            var pnlSearchBorder = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(1),
                BackColor = UIColors.Border,
                Margin = new Padding(0)
            };
            var pnlSearchInner = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = SD.Color.White,
                Padding = new Padding(Scale(8), 0, Scale(8), 0)
            };

            txtSearch = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new SD.Font("Segoe UI", 10f),
                ForeColor = UIColors.TextPrimary
            };
            txtSearch.TextChanged += txtSearch_TextChanged;

            pnlSearchInner.Resize += (s, e) =>
            {
                int topPadding = (pnlSearchInner.Height - txtSearch.Height) / 2;
                if (topPadding < 0) topPadding = 0;
                pnlSearchInner.Padding = new Padding(Scale(8), topPadding, Scale(8), 0);
            };

            pnlSearchInner.Controls.Add(txtSearch);
            pnlSearchBorder.Controls.Add(pnlSearchInner);

            chkSharedLibrary = new CheckBox
            {
                Dock = DockStyle.Right,
                AutoSize = true,
                Text = "Shared Library",
                Font = new SD.Font("Segoe UI", 9f),
                ForeColor = UIColors.TextSecondary,
                Checked = false,
                Padding = new Padding(Scale(10), 0, 0, 0)
            };
            chkSharedLibrary.CheckedChanged += chkSharedLibrary_CheckedChanged;

            pnlSearch.Controls.Add(pnlSearchBorder);
            pnlSearch.Controls.Add(lblSearchLabel);
            pnlSearch.Controls.Add(chkSharedLibrary);

            pnlMain.Controls.Add(pnlSearch);

            // ── TAB BAR (Dock=Top) ──
            pnlTabs = new Panel
            {
                Dock = DockStyle.Top,
                Height = Scale(48),
                BackColor = UIColors.Surface
            };
            pnlTabs.Paint += PnlTabs_Paint;
            pnlTabs.MouseClick += PnlTabs_MouseClick;
            pnlTabs.MouseMove += PnlTabs_MouseMove;
            pnlTabs.MouseLeave += (s, e) => { _hoveredTabIndex = -1; pnlTabs.Invalidate(); };

            pnlMain.Controls.Add(pnlTabs);

            // ── TITLE BAR (Dock=Top) ──
            pnlTitleBar = new Panel
            {
                Dock = DockStyle.Top,
                Height = Scale(50),
                BackColor = UIColors.TitleBg
            };
            pnlTitleBar.MouseDown += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            };
            pnlTitleBar.Paint += (s, e) =>
            {
                e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                e.Graphics.SmoothingMode = SmoothingMode.HighQuality;

                using (var font = new SD.Font("Segoe UI", 13f, SD.FontStyle.Bold))
                using (var brush = new SD.SolidBrush(UIColors.TitleText))
                {
                    float y = (pnlTitleBar.Height - font.Height) / 2f;
                    e.Graphics.DrawString("Asset Library", font, brush, Scale(20), y);
                }
            };

            // Close button
            var btnClose = new Label
            {
                Text = "✕",
                Font = new SD.Font("Segoe UI", 11f),
                ForeColor = UIColors.TitleText,
                Dock = DockStyle.Right,
                Width = Scale(50),
                TextAlign = SD.ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand,
                BackColor = UIColors.TitleBg
            };
            btnClose.MouseEnter += (s, e) => { btnClose.BackColor = UIColors.Danger; };
            btnClose.MouseLeave += (s, e) => { btnClose.BackColor = UIColors.TitleBg; };
            btnClose.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

            pnlTitleBar.Controls.Add(btnClose);
            pnlMain.Controls.Add(pnlTitleBar);
            pnlContent.BringToFront();

            this.ResumeLayout(false);
        }

        private Button CreateModernButton(string text, SD.Size size, bool isPrimary)
        {
            var btn = new Button
            {
                Text = text,
                Size = size,
                Font = new SD.Font("Segoe UI Semibold", 9.5f),
                FlatStyle = FlatStyle.Flat,
                Cursor = Cursors.Hand,
                Margin = new Padding(0, 0, Scale(8), 0)
            };

            if (isPrimary)
            {
                btn.BackColor = UIColors.TitleBg;
                btn.ForeColor = UIColors.TitleText;
                btn.FlatAppearance.BorderSize = 0;
                btn.FlatAppearance.MouseOverBackColor = UIColors.Accent;
                btn.FlatAppearance.MouseDownBackColor = UIColors.AccentHover;
            }
            else
            {
                btn.BackColor = UIColors.Surface;
                btn.ForeColor = UIColors.TextPrimary;
                btn.FlatAppearance.BorderColor = UIColors.Border;
                btn.FlatAppearance.BorderSize = 1;
                btn.FlatAppearance.MouseOverBackColor = UIColors.SurfaceAlt;
                btn.FlatAppearance.MouseDownBackColor = UIColors.SurfaceDark;
            }

            return btn;
        }

        // ── TAB DRAWING ──
        private readonly string[] _tabNames = { "Objects", "Pages", "Documents" };
        private int _hoveredTabIndex = -1;

        private void PnlTabs_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            int tabWidth = Scale(100);
            int startX = Scale(15);
            int tabHeight = pnlTabs.Height - Scale(4);

            for (int i = 0; i < _tabNames.Length; i++)
            {
                var rect = new SD.Rectangle(startX + i * tabWidth, Scale(2), tabWidth, tabHeight);
                bool isSelected = (i == _selectedTabIndex);
                bool isHovered = (i == _hoveredTabIndex && !isSelected);

                if (isSelected)
                {
                    using (var brush = new SD.SolidBrush(UIColors.AccentLight))
                        g.FillRectangle(brush, rect);
                }
                else if (isHovered)
                {
                    using (var brush = new SD.SolidBrush(UIColors.SurfaceAlt))
                        g.FillRectangle(brush, rect);
                }

                var textColor = isSelected ? UIColors.TitleBg : UIColors.TextSecondary;
                var fontStyle = isSelected ? SD.FontStyle.Bold : SD.FontStyle.Regular;
                using (var font = new SD.Font("Segoe UI", 9.5f, fontStyle))
                using (var brush = new SD.SolidBrush(textColor))
                {
                    var sf = new SD.StringFormat { Alignment = SD.StringAlignment.Center, LineAlignment = SD.StringAlignment.Center };
                    g.DrawString(_tabNames[i], font, brush, rect, sf);
                }

                if (isSelected)
                {
                    using (var pen = new SD.Pen(UIColors.TitleBg, Scale(2)))
                        g.DrawLine(pen, rect.X + Scale(10), rect.Bottom - Scale(2), rect.Right - Scale(10), rect.Bottom - Scale(2));
                }
            }

            using (var pen = new SD.Pen(UIColors.Border))
                g.DrawLine(pen, 0, pnlTabs.Height - 1, pnlTabs.Width, pnlTabs.Height - 1);
        }

        private void PnlTabs_MouseClick(object sender, MouseEventArgs e)
        {
            int tabWidth = Scale(100);
            int startX = Scale(15);

            for (int i = 0; i < _tabNames.Length; i++)
            {
                var rect = new SD.Rectangle(startX + i * tabWidth, 0, tabWidth, pnlTabs.Height);
                if (rect.Contains(e.Location))
                {
                    if (_selectedTabIndex != i)
                    {
                        _selectedTabIndex = i;
                        _currentAssetType = (LibraryAssetType)i;
                        txtSearch.Text = "";
                        LoadAssets();
                        InitializeFormSpecifics();
                        pnlTabs.Invalidate();
                    }
                    break;
                }
            }
        }

        private void PnlTabs_MouseMove(object sender, MouseEventArgs e)
        {
            int tabWidth = Scale(100);
            int startX = Scale(15);
            int newHover = -1;

            for (int i = 0; i < _tabNames.Length; i++)
            {
                var rect = new SD.Rectangle(startX + i * tabWidth, 0, tabWidth, pnlTabs.Height);
                if (rect.Contains(e.Location))
                {
                    newHover = i;
                    break;
                }
            }

            if (newHover != _hoveredTabIndex)
            {
                _hoveredTabIndex = newHover;
                pnlTabs.Invalidate();
            }
        }

        // ── LIST DRAWING ──
        private void ListBoxAssets_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            bool isHovered = (e.Index == _hoveredIndex && !isSelected);

            SD.Color backColor = isSelected ? UIColors.ListSelected :
                              isHovered ? UIColors.ListHover : UIColors.Surface;

            using (var brush = new SD.SolidBrush(backColor))
                g.FillRectangle(brush, e.Bounds);

            if (isSelected)
            {
                using (var brush = new SD.SolidBrush(UIColors.TitleBg))
                    g.FillRectangle(brush, e.Bounds.X, e.Bounds.Y, Scale(3), e.Bounds.Height);
            }

            var textRect = new SD.Rectangle(
                e.Bounds.X + Scale(isSelected ? 15 : 12),
                e.Bounds.Y,
                e.Bounds.Width - Scale(20),
                e.Bounds.Height);

            var textColor = isSelected ? UIColors.TitleBg : UIColors.TextPrimary;
            using (var font = new SD.Font("Segoe UI", 9.5f, isSelected ? SD.FontStyle.Bold : SD.FontStyle.Regular))
            using (var brush = new SD.SolidBrush(textColor))
            {
                var sf = new SD.StringFormat
                {
                    Alignment = SD.StringAlignment.Near,
                    LineAlignment = SD.StringAlignment.Center,
                    Trimming = SD.StringTrimming.EllipsisCharacter
                };
                g.DrawString(listBoxAssets.Items[e.Index].ToString(), font, brush, textRect, sf);
            }
        }

        private void ListBoxAssets_MouseMove(object sender, MouseEventArgs e)
        {
            int newHoverIndex = listBoxAssets.IndexFromPoint(e.Location);
            if (newHoverIndex != _hoveredIndex)
            {
                _hoveredIndex = newHoverIndex;
                listBoxAssets.Invalidate();
            }
        }

        private void InitializeFormSpecifics()
        {
            btnInsert.Text = (_currentAssetType == LibraryAssetType.Document) ? "Open" : "Insert";

            bool canSaveFromSelection = false;
            try
            {
                bool objectSavable = _currentAssetType == LibraryAssetType.Object &&
                                     _wordApplication?.Selection != null &&
                                     _wordApplication.Selection.Type != Word.WdSelectionType.wdSelectionIP;

                bool pageSavable = _currentAssetType == LibraryAssetType.Page &&
                                   _wordApplication?.ActiveDocument != null;

                bool documentSavable = _currentAssetType == LibraryAssetType.Document &&
                                       _wordApplication?.Documents?.Count > 0 &&
                                       _wordApplication.ActiveDocument != null;

                canSaveFromSelection = objectSavable || pageSavable || documentSavable;
            }
            catch (COMException ex)
            {
                Console.WriteLine($"COMException in InitializeFormSpecifics: {ex.Message}");
            }
            catch (Exception exGen)
            {
                Console.WriteLine($"General Exception in InitializeFormSpecifics: {exGen.Message}");
            }

            btnSaveNew.Visible = canSaveFromSelection;
            btnReplaceSelected.Visible = canSaveFromSelection;
            btnReplaceSelected.Enabled = false;
            btnRename.Enabled = false;
            btnInsert.Visible = (listBoxAssets.SelectedItem != null);
        }

        private void chkSharedLibrary_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSharedLibrary.Checked)
            {
                if (!EnsureSharedBasePathChosen())
                {
                    chkSharedLibrary.CheckedChanged -= chkSharedLibrary_CheckedChanged;
                    chkSharedLibrary.Checked = false;
                    chkSharedLibrary.CheckedChanged += chkSharedLibrary_CheckedChanged;
                    _useSharedLibrary = false;
                    return;
                }
                _useSharedLibrary = true;
            }
            else
            {
                _useSharedLibrary = false;
            }

            LoadAssets();
        }

        private bool EnsureSharedBasePathChosen()
        {
            var existing = SharedLibrarySettings.GetSharedBasePath();
            if (!string.IsNullOrEmpty(existing) && IODir.Exists(existing))
            {
                CreateTypeSubfoldersIfMissing(existing);
                return true;
            }

            using (var fbd = new FolderBrowserDialog()
            {
                Description = "Select a Shared Library folder accessible by all users.\n" +
                              "Subfolders (Objects, Pages, Documents) will be created if needed."
            })
            {
                if (fbd.ShowDialog(this) == DialogResult.OK && IODir.Exists(fbd.SelectedPath))
                {
                    SharedLibrarySettings.SetSharedBasePath(fbd.SelectedPath);
                    CreateTypeSubfoldersIfMissing(fbd.SelectedPath);
                    return true;
                }
            }
            return false;
        }

        private void CreateTypeSubfoldersIfMissing(string basePath)
        {
            try
            {
                var wordRoot = Path.Combine(basePath, "Word");
                if (!IODir.Exists(wordRoot)) IODir.CreateDirectory(wordRoot);

                foreach (var name in new[] { "Objects", "Pages", "Documents" })
                {
                    var p = Path.Combine(wordRoot, name);
                    if (!IODir.Exists(p)) IODir.CreateDirectory(p);
                }
            }
            catch { /* ignore */ }
        }

        private string GetLibraryBasePath()
        {
            if (_useSharedLibrary)
            {
                var shared = SharedLibrarySettings.GetSharedBasePath();
                if (!string.IsNullOrEmpty(shared) && IODir.Exists(shared))
                    return Path.Combine(shared, "Word");

                _useSharedLibrary = false;
                if (chkSharedLibrary != null)
                {
                    chkSharedLibrary.CheckedChanged -= chkSharedLibrary_CheckedChanged;
                    chkSharedLibrary.Checked = false;
                    chkSharedLibrary.CheckedChanged += chkSharedLibrary_CheckedChanged;
                }
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "OfficePragmatics_Library",
                "Word");
        }

        private string GetLibraryPath()
        {
            string basePath = GetLibraryBasePath();
            string specificPath = Path.Combine(basePath, _currentAssetType.ToString() + "s");
            if (!IODir.Exists(specificPath))
            {
                try { IODir.CreateDirectory(specificPath); } catch { /* ignore */ }
            }
            return specificPath;
        }

        private void LoadAssets()
        {
            _allAssets = new List<LibraryAsset>();
            string libPath = GetLibraryPath();
            string[] assetFiles = Directory.GetFiles(libPath, "*.docx");

            foreach (string filePath in assetFiles)
            {
                string fileNameNoExt = Path.GetFileNameWithoutExtension(filePath);
                string assetName = fileNameNoExt;
                string thumbnailPath = Path.ChangeExtension(filePath, ".png");
                string uniqueID = "";

                int lastUnderscore = fileNameNoExt.LastIndexOf('_');
                if ((_currentAssetType == LibraryAssetType.Object || _currentAssetType == LibraryAssetType.Page) &&
                    lastUnderscore > 0 && fileNameNoExt.Length - lastUnderscore - 1 == 36)
                {
                    string potentialGuid = fileNameNoExt.Substring(lastUnderscore + 1);
                    if (Guid.TryParse(potentialGuid, out _))
                    {
                        uniqueID = potentialGuid;
                        assetName = fileNameNoExt.Substring(0, lastUnderscore);
                    }
                }

                if (!IOFile.Exists(thumbnailPath))
                {
                    try
                    {
                        SD.Bitmap bmp = new SD.Bitmap(120, 90);
                        using (SD.Graphics g = SD.Graphics.FromImage(bmp))
                        {
                            g.Clear(UIColors.SurfaceAlt);
                            using (var font = new SD.Font("Segoe UI", 9f))
                            using (var brush = new SD.SolidBrush(UIColors.TextSecondary))
                                g.DrawString("No Preview", font, brush, 10, 38);
                        }
                        bmp.Save(thumbnailPath, System.Drawing.Imaging.ImageFormat.Png);
                        bmp.Dispose();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error creating dummy thumbnail for {thumbnailPath}: {ex.Message}");
                    }
                }

                _allAssets.Add(new LibraryAsset
                {
                    Name = assetName,
                    FilePath = filePath,
                    ThumbnailPath = thumbnailPath,
                    AssetType = _currentAssetType,
                    UniqueID = uniqueID
                });
            }

            ApplyFilter();
            UpdateStatus();
        }

        private void UpdateStatus()
        {
            int total = _allAssets?.Count ?? 0;
            int showing = _filteredAssets?.Count ?? 0;
            string typeLabel = _currentAssetType.ToString().ToLower() + "s";

            if (total == showing)
                lblStatus.Text = $"{total} {typeLabel} in library";
            else
                lblStatus.Text = $"Showing {showing} of {total} {typeLabel}";
        }

        private void ApplyFilter()
        {
            string searchTerm = txtSearch.Text.ToLowerInvariant();
            _filteredAssets = string.IsNullOrWhiteSpace(searchTerm) ?
                              new List<LibraryAsset>(_allAssets) :
                              _allAssets.Where(a => a.Name.ToLowerInvariant().Contains(searchTerm)).ToList();

            listBoxAssets.BeginUpdate();
            listBoxAssets.DataSource = null;
            listBoxAssets.DataSource = _filteredAssets;
            listBoxAssets.DisplayMember = "Name";
            listBoxAssets.EndUpdate();

            if (listBoxAssets.Items.Count > 0)
                listBoxAssets.SelectedIndex = 0;
            else
                ClearPreview();

            UpdateStatus();
        }

        private void ClearPreview()
        {
            if (pictureBoxThumbnail.Image != null)
            {
                pictureBoxThumbnail.Image.Dispose();
                pictureBoxThumbnail.Image = null;
            }
            lblAssetName.Text = "";
            btnReplaceSelected.Enabled = false;
            btnInsert.Visible = false;
            btnDelete.Enabled = false;
            btnRename.Enabled = false;
        }

        private void listBoxAssets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxAssets.SelectedItem is LibraryAsset selectedAsset)
            {
                lblAssetName.Text = selectedAsset.Name;
                if (IOFile.Exists(selectedAsset.ThumbnailPath))
                {
                    try
                    {
                        if (pictureBoxThumbnail.Image != null) pictureBoxThumbnail.Image.Dispose();
                        using (FileStream fs = new FileStream(selectedAsset.ThumbnailPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                        {
                            pictureBoxThumbnail.Image = SD.Image.FromStream(fs);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error loading thumbnail {selectedAsset.ThumbnailPath}: {ex.Message}");
                        if (pictureBoxThumbnail.Image != null) pictureBoxThumbnail.Image.Dispose();
                        pictureBoxThumbnail.Image = null;
                    }
                }
                else
                {
                    if (pictureBoxThumbnail.Image != null) pictureBoxThumbnail.Image.Dispose();
                    pictureBoxThumbnail.Image = null;
                }
                btnReplaceSelected.Enabled = true;
                btnInsert.Visible = true;
                btnDelete.Enabled = true;
                btnRename.Enabled = true;
            }
            else
            {
                ClearPreview();
            }
        }

        private void txtSearch_TextChanged(object sender, EventArgs e)
        {
            ApplyFilter();
        }

        private void btnSaveNew_Click(object sender, EventArgs e)
        {
            string assetName = PromptForAssetName("Enter a name for the new asset:", lblAssetName.Text);
            if (string.IsNullOrWhiteSpace(assetName)) return;

            string uniqueID = Guid.NewGuid().ToString();
            string cleanAssetName = Path.GetInvalidFileNameChars().Aggregate(assetName, (current, c) => current.Replace(c.ToString(), string.Empty));
            string baseFileName = (_currentAssetType == LibraryAssetType.Object || _currentAssetType == LibraryAssetType.Page)
                                  ? $"{cleanAssetName}_{uniqueID}"
                                  : cleanAssetName;

            string filePath = Path.Combine(GetLibraryPath(), baseFileName + ".docx");
            string thumbPath = Path.ChangeExtension(filePath, ".png");

            if (IOFile.Exists(filePath))
            {
                MessageBox.Show("An asset with this file name already exists. Please choose a different name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                switch (_currentAssetType)
                {
                    case LibraryAssetType.Object:
                        SaveSelectedObjects(filePath, thumbPath, uniqueID);
                        break;
                    case LibraryAssetType.Page:
                        SaveActivePage(filePath, thumbPath, uniqueID);
                        break;
                    case LibraryAssetType.Document:
                        SaveActiveDocument(filePath, thumbPath);
                        break;
                }
                MessageBox.Show($"{_currentAssetType} '{assetName}' saved successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnReplaceSelected_Click(object sender, EventArgs e)
        {
            if (listBoxAssets.SelectedItem is LibraryAsset selectedAsset)
            {
                string newAssetNameInput = lblAssetName.Text;
                if (string.IsNullOrWhiteSpace(newAssetNameInput))
                {
                    MessageBox.Show("Asset name cannot be empty.", "Invalid Name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                string cleanNewAssetName = Path.GetInvalidFileNameChars().Aggregate(newAssetNameInput, (current, c) => current.Replace(c.ToString(), string.Empty));

                DialogResult confirm = MessageBox.Show($"Replace '{selectedAsset.Name}' with current selection?\nNew name: '{cleanNewAssetName}'",
                                                      "Confirm Replace", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm == DialogResult.No) return;

                string uniqueID = selectedAsset.UniqueID;
                if (string.IsNullOrEmpty(uniqueID) && (_currentAssetType == LibraryAssetType.Object || _currentAssetType == LibraryAssetType.Page))
                {
                    uniqueID = Guid.NewGuid().ToString();
                }

                try
                {
                    if (pictureBoxThumbnail.Image != null)
                    {
                        pictureBoxThumbnail.Image.Dispose();
                        pictureBoxThumbnail.Image = null;
                    }
                    if (IOFile.Exists(selectedAsset.FilePath)) IOFile.Delete(selectedAsset.FilePath);
                    if (IOFile.Exists(selectedAsset.ThumbnailPath)) IOFile.Delete(selectedAsset.ThumbnailPath);
                }
                catch (IOException ioEx)
                {
                    MessageBox.Show($"Error deleting old asset files: {ioEx.Message}", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                string baseFileName = (_currentAssetType == LibraryAssetType.Object || _currentAssetType == LibraryAssetType.Page) && !string.IsNullOrEmpty(uniqueID)
                                  ? $"{cleanNewAssetName}_{uniqueID}"
                                  : cleanNewAssetName;
                string newFilePath = Path.Combine(GetLibraryPath(), baseFileName + ".docx");
                string newThumbPath = Path.ChangeExtension(newFilePath, ".png");

                try
                {
                    switch (_currentAssetType)
                    {
                        case LibraryAssetType.Object:
                            SaveSelectedObjects(newFilePath, newThumbPath, uniqueID);
                            break;
                        case LibraryAssetType.Page:
                            SaveActivePage(newFilePath, newThumbPath, uniqueID);
                            break;
                        case LibraryAssetType.Document:
                            SaveActiveDocument(newFilePath, newThumbPath);
                            break;
                    }
                    MessageBox.Show($"Asset replaced successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadAssets();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error replacing asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LoadAssets();
                }
            }
        }

        private void btnInsert_Click(object sender, EventArgs e)
        {
            if (listBoxAssets.SelectedItem is LibraryAsset selectedAsset)
            {
                try
                {
                    switch (_currentAssetType)
                    {
                        case LibraryAssetType.Object:
                            InsertObjectsFromLibrary(selectedAsset);
                            break;
                        case LibraryAssetType.Page:
                            InsertPageFromLibrary(selectedAsset);
                            break;
                        case LibraryAssetType.Document:
                            OpenDocumentFromLibrary(selectedAsset);
                            break;
                    }
                    if (_currentAssetType != LibraryAssetType.Document)
                        MessageBox.Show($"'{selectedAsset.Name}' inserted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error inserting/opening asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnRename_Click(object sender, EventArgs e)
        {
            if (listBoxAssets.SelectedItem is LibraryAsset selectedAsset)
            {
                string newName = PromptForAssetName("Enter a new name for this asset:", selectedAsset.Name);
                if (string.IsNullOrWhiteSpace(newName)) return;

                string cleanNewName = Path.GetInvalidFileNameChars().Aggregate(newName, (current, c) => current.Replace(c.ToString(), string.Empty));

                string newBaseFileName = (_currentAssetType == LibraryAssetType.Object || _currentAssetType == LibraryAssetType.Page) && !string.IsNullOrEmpty(selectedAsset.UniqueID)
                                      ? $"{cleanNewName}_{selectedAsset.UniqueID}"
                                      : cleanNewName;

                string newFilePath = Path.Combine(GetLibraryPath(), newBaseFileName + ".docx");
                string newThumbPath = Path.ChangeExtension(newFilePath, ".png");

                if (IOFile.Exists(newFilePath) && newFilePath != selectedAsset.FilePath)
                {
                    MessageBox.Show("An asset with this name already exists.", "Name Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                try
                {
                    if (pictureBoxThumbnail.Image != null)
                    {
                        pictureBoxThumbnail.Image.Dispose();
                        pictureBoxThumbnail.Image = null;
                    }

                    if (IOFile.Exists(selectedAsset.FilePath))
                        IOFile.Move(selectedAsset.FilePath, newFilePath);
                    if (IOFile.Exists(selectedAsset.ThumbnailPath))
                        IOFile.Move(selectedAsset.ThumbnailPath, newThumbPath);

                    MessageBox.Show($"Renamed to '{cleanNewName}'.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadAssets();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error renaming asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (listBoxAssets.SelectedItem is LibraryAsset selectedAsset)
            {
                DialogResult confirm = MessageBox.Show($"Delete '{selectedAsset.Name}' from the library?",
                                                      "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (confirm == DialogResult.Yes)
                {
                    try
                    {
                        if (pictureBoxThumbnail.Image != null)
                        {
                            pictureBoxThumbnail.Image.Dispose();
                            pictureBoxThumbnail.Image = null;
                        }
                        if (IOFile.Exists(selectedAsset.FilePath)) IOFile.Delete(selectedAsset.FilePath);
                        if (IOFile.Exists(selectedAsset.ThumbnailPath)) IOFile.Delete(selectedAsset.ThumbnailPath);
                        MessageBox.Show($"'{selectedAsset.Name}' deleted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        LoadAssets();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error deleting asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private string PromptForAssetName(string promptMessage, string defaultName = "")
        {
            string name = "";
            using (Form prompt = new Form()
            {
                Width = Scale(380),
                Height = Scale(160),
                Text = "Asset Name",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterParent,
                MinimizeBox = false,
                MaximizeBox = false,
                BackColor = UIColors.Surface
            })
            {
                Label textLabel = new Label()
                {
                    Left = Scale(20),
                    Top = Scale(20),
                    Text = promptMessage,
                    AutoSize = true,
                    Font = new SD.Font("Segoe UI", 9.5f),
                    ForeColor = UIColors.TextPrimary
                };

                TextBox textBox = new TextBox()
                {
                    Left = Scale(20),
                    Top = Scale(50),
                    Width = prompt.ClientSize.Width - Scale(40),
                    Text = defaultName,
                    Font = new SD.Font("Segoe UI", 10f),
                    BorderStyle = BorderStyle.FixedSingle
                };

                Button confirmation = CreateModernButton("OK", new SD.Size(Scale(80), Scale(32)), true);
                confirmation.Left = prompt.ClientSize.Width - Scale(180);
                confirmation.Top = Scale(90);
                confirmation.DialogResult = DialogResult.OK;
                confirmation.Click += (s, ev) => { prompt.Close(); };

                Button cancel = CreateModernButton("Cancel", new SD.Size(Scale(80), Scale(32)), false);
                cancel.Left = prompt.ClientSize.Width - Scale(90);
                cancel.Top = Scale(90);
                cancel.DialogResult = DialogResult.Cancel;
                cancel.Click += (s, ev) => { prompt.Close(); };

                prompt.Controls.Add(textBox);
                prompt.Controls.Add(confirmation);
                prompt.Controls.Add(cancel);
                prompt.Controls.Add(textLabel);
                prompt.AcceptButton = confirmation;
                prompt.CancelButton = cancel;

                if (prompt.ShowDialog(this) == DialogResult.OK)
                {
                    name = textBox.Text;
                }
                return name;
            }
        }

        // --- Word Interaction Methods ---
        private void SaveSelectedObjects(string docxFilePath, string pngThumbPath, string uniqueID)
        {
            Word.Selection sel = _wordApplication.Selection;
            if (sel.Type == Word.WdSelectionType.wdSelectionIP)
            {
                throw new InvalidOperationException("No content selected for saving.");
            }

            Word.Document tempDoc = null;

            try
            {
                tempDoc = _wordApplication.Documents.Add();
                sel.Copy();
                tempDoc.Content.Paste();

                Office.CustomXMLPart xmlPart = tempDoc.CustomXMLParts.Add(
                    $"<LibraryMetadata><ID>{uniqueID}</ID></LibraryMetadata>");

                tempDoc.SaveAs2(docxFilePath);

                CreateDocumentThumbnail(tempDoc, pngThumbPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving object asset: {ex.Message}", ex);
            }
            finally
            {
                if (tempDoc != null)
                {
                    tempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(tempDoc);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void SaveActivePage(string docxFilePath, string pngThumbPath, string uniqueID)
        {
            Word.Document activeDoc = null;
            Word.Document tempDoc = null;
            Word.Range pageRange = null;

            try
            {
                activeDoc = _wordApplication.ActiveDocument;
                if (activeDoc == null)
                {
                    throw new InvalidOperationException("No active document.");
                }

                int currentPage = (int)_wordApplication.Selection.Information[Word.WdInformation.wdActiveEndPageNumber];

                object whatPage = Word.WdGoToItem.wdGoToPage;
                object whichPage = Word.WdGoToDirection.wdGoToAbsolute;
                object pageNum = currentPage;

                Word.Range startRange = activeDoc.GoTo(whatPage, whichPage, pageNum);
                int startPos = startRange.Start;

                pageNum = currentPage + 1;
                Word.Range endRange;
                int endPos;

                try
                {
                    endRange = activeDoc.GoTo(whatPage, whichPage, pageNum);
                    endPos = endRange.Start - 1;
                }
                catch
                {
                    endPos = activeDoc.Content.End;
                }

                pageRange = activeDoc.Range(startPos, endPos);

                tempDoc = _wordApplication.Documents.Add();
                pageRange.Copy();
                tempDoc.Content.Paste();

                Office.CustomXMLPart xmlPart = tempDoc.CustomXMLParts.Add(
                    $"<LibraryMetadata><ID>{uniqueID}</ID></LibraryMetadata>");

                tempDoc.SaveAs2(docxFilePath);

                CreateDocumentThumbnail(tempDoc, pngThumbPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error processing page for library: {ex.Message}", ex);
            }
            finally
            {
                if (pageRange != null) Marshal.ReleaseComObject(pageRange);
                if (tempDoc != null)
                {
                    tempDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(tempDoc);
                }
                if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void SaveActiveDocument(string docxFilePath, string pngThumbPath)
        {
            Word.Document activeDoc = null;
            try
            {
                activeDoc = _wordApplication.ActiveDocument;
                if (activeDoc == null)
                {
                    throw new InvalidOperationException("No active document to save.");
                }

                activeDoc.SaveAs2(docxFilePath);

                CreateDocumentThumbnail(activeDoc, pngThumbPath);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error saving active document: {ex.Message}", ex);
            }
            finally
            {
                if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
            }
        }

        private void CreateDocumentThumbnail(Word.Document doc, string pngPath)
        {
            try
            {
                string tempHtmlPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".html");
                doc.SaveAs2(tempHtmlPath, Word.WdSaveFormat.wdFormatFilteredHTML);

                SD.Bitmap bmp = new SD.Bitmap(120, 90);
                using (SD.Graphics g = SD.Graphics.FromImage(bmp))
                {
                    g.Clear(SD.Color.White);
                    using (var font = new SD.Font("Segoe UI", 9f))
                    using (var brush = new SD.SolidBrush(UIColors.TextSecondary))
                        g.DrawString(doc.Name, font, brush, 10, 40);
                }
                bmp.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
                bmp.Dispose();

                try { IOFile.Delete(tempHtmlPath); } catch { }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Thumbnail generation error: {ex.Message}");
                SD.Bitmap bmp = new SD.Bitmap(120, 90);
                using (SD.Graphics g = SD.Graphics.FromImage(bmp))
                {
                    g.Clear(UIColors.SurfaceAlt);
                    using (var font = new SD.Font("Segoe UI", 9f))
                    using (var brush = new SD.SolidBrush(UIColors.TextSecondary))
                        g.DrawString("No Preview", font, brush, 10, 40);
                }
                bmp.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
                bmp.Dispose();
            }
        }

        private void InsertObjectsFromLibrary(LibraryAsset asset)
        {
            Word.Document sourceDoc = null;
            Word.Range sourceRange = null;

            try
            {
                if (_wordApplication.ActiveDocument == null)
                {
                    throw new InvalidOperationException("Please open a document to insert objects into.");
                }

                sourceDoc = _wordApplication.Documents.Open(asset.FilePath, ReadOnly: true);
                sourceRange = sourceDoc.Content;
                sourceRange.Copy();

                _wordApplication.Selection.Paste();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error inserting objects: {ex.Message}", ex);
            }
            finally
            {
                if (sourceRange != null) Marshal.ReleaseComObject(sourceRange);
                if (sourceDoc != null)
                {
                    sourceDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(sourceDoc);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void InsertPageFromLibrary(LibraryAsset asset)
        {
            Word.Document activeDoc = null;
            Word.Document sourceDoc = null;
            Word.Range insertRange = null;

            try
            {
                activeDoc = _wordApplication.ActiveDocument;
                if (activeDoc == null)
                {
                    throw new InvalidOperationException("No active document to insert page into.");
                }

                _wordApplication.Selection.InsertBreak(Word.WdBreakType.wdPageBreak);

                sourceDoc = _wordApplication.Documents.Open(asset.FilePath, ReadOnly: true);
                sourceDoc.Content.Copy();

                _wordApplication.Selection.Paste();
            }
            catch (Exception ex)
            {
                throw new Exception($"Error inserting page: {ex.Message}", ex);
            }
            finally
            {
                if (insertRange != null) Marshal.ReleaseComObject(insertRange);
                if (sourceDoc != null)
                {
                    sourceDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                    Marshal.ReleaseComObject(sourceDoc);
                }
                if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
            }
        }

        private void OpenDocumentFromLibrary(LibraryAsset asset)
        {
            Word.Document newDoc = null;
            try
            {
                newDoc = _wordApplication.Documents.Open(
                    FileName: asset.FilePath,
                    ReadOnly: false);
            }
            catch (Exception ex)
            {
                throw new Exception($"Error opening document: {ex.Message}", ex);
            }
            finally
            {
                if (newDoc != null) Marshal.ReleaseComObject(newDoc);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (pictureBoxThumbnail?.Image != null)
                {
                    pictureBoxThumbnail.Image.Dispose();
                    pictureBoxThumbnail.Image = null;
                }
            }
            base.Dispose(disposing);
        }
    }

    public static class LibraryUtility
    {
        private static Word.Application WordApplication => Globals.ThisAddIn.Application;

        public static void OpenLibrary(LibraryAssetType? initialTab = null)
        {
            try
            {
                using (var form = new LibraryForm(WordApplication, initialTab))
                {
                    form.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening Library: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void UpdateLinks()
        {
            Word.Document activeDoc = null;
            try
            {
                if (WordApplication.Documents.Count == 0)
                {
                    MessageBox.Show("No active document to update links in.", "Update Links", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                activeDoc = WordApplication.ActiveDocument;
            }
            catch (COMException)
            {
                MessageBox.Show("Could not get the active document.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            int updatedCount = 0;

            var basePaths = new List<string>
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                             "OfficePragmatics_Library", "Word")
            };
            var sharedBase = SharedLibrarySettings.GetSharedBasePath();
            if (!string.IsNullOrEmpty(sharedBase) && IODir.Exists(sharedBase))
                basePaths.Add(Path.Combine(sharedBase, "Word"));

            try
            {
                foreach (Office.CustomXMLPart xmlPart in activeDoc.CustomXMLParts)
                {
                    try
                    {
                        string xml = xmlPart.XML;
                        if (xml.Contains("<LibraryMetadata>") && xml.Contains("<ID>"))
                        {
                            int idStart = xml.IndexOf("<ID>") + 4;
                            int idEnd = xml.IndexOf("</ID>");
                            if (idStart > 3 && idEnd > idStart)
                            {
                                string id = xml.Substring(idStart, idEnd - idStart);
                                updatedCount++;
                            }
                        }
                    }
                    catch { }
                }

                MessageBox.Show($"Update complete.\nItems updated: {updatedCount}", "Update Links", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error during link update: {ex.Message}", "Update Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}