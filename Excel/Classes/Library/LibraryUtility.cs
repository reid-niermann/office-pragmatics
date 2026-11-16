// ExcelLibrary.cs  (C# 7.3 / .NET 4.8)

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using Excel = Microsoft.Office.Interop.Excel;
using OfficePragmatics_Excel.Settings;
using OfficePragmatics_Excel.Classes;
using IOFile = System.IO.File;
using IODir = System.IO.Directory;
using SD = System.Drawing;

namespace OfficePragmatics_Excel.Library
{
    public enum ExcelLibraryAssetType
    {
        Sheet,
        Workbook
    }

    public class LibraryAsset
    {
        public string Name { get; set; }
        public string FilePath { get; set; }
        public string ThumbnailPath { get; set; }
        public ExcelLibraryAssetType AssetType { get; set; }
        public string UniqueID { get; set; }

        public override string ToString() => Name;
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
            catch { /* ignore */ }
        }
    }

    public sealed class ExcelLibraryForm : Form
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

        private const string AppFolderName = "XLS";
        private ExcelLibraryAssetType _currentAssetType;
        private List<LibraryAsset> _allAssets;
        private List<LibraryAsset> _filteredAssets;
        private readonly Excel.Application _xl;

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
        private CheckBox chkSharedLibrary;

        private bool _useSharedLibrary = false;
        private int _hoveredIndex = -1;
        private int _selectedTabIndex = 0;

        // DPI scaling
        private float _scale = 1f;

        public ExcelLibraryForm(Excel.Application xlApp, ExcelLibraryAssetType? initialType = null)
        {
            _xl = xlApp;
            _currentAssetType = initialType ?? ExcelLibraryAssetType.Sheet;

            // DPI scaling
            using (var g = CreateGraphics()) { _scale = g.DpiX / 96f; }

            this.FormBorderStyle = FormBorderStyle.None;
            this.DoubleBuffered = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = UIColors.Border;
            this.Padding = new Padding(1);
            this.ShowIcon = false;

            InitializeComponentManual();
            InitializeFormSpecifics();

            if (initialType.HasValue)
            {
                _selectedTabIndex = (int)initialType.Value;
                _currentAssetType = initialType.Value;
            }

            LoadAssets();
            this.Shown += ExcelLibraryForm_Shown;
        }

        private int Scale(int value) => (int)(value * _scale);

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ExcelLibraryForm_Shown(object sender, EventArgs e)
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
            btnSaveNew.Click += (_, __) => SaveNew();

            btnInsert = CreateModernButton("Insert", btnSize, true);
            btnInsert.Click += (_, __) => InsertSelected();

            btnDelete = CreateModernButton("Delete", btnSize, false);
            btnDelete.Click += (_, __) => DeleteSelected();

            btnReplaceSelected = CreateModernButton("Replace", btnSize, false);
            btnReplaceSelected.Click += (_, __) => ReplaceSelected();

            btnRename = CreateModernButton("Rename", btnSize, false);
            btnRename.Click += (_, __) => RenameSelected();

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
            listBoxAssets.SelectedIndexChanged += ListBoxAssets_SelectedIndexChanged;
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
            txtSearch.TextChanged += (_, __) => ApplyFilter();

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
            chkSharedLibrary.CheckedChanged += ChkSharedLibrary_CheckedChanged;

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

            this.Name = "ExcelLibraryForm";
            this.Text = "Asset Library";

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
        private readonly string[] _tabNames = { "Sheets", "Workbooks" };
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
                        _currentAssetType = (ExcelLibraryAssetType)i;
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
            btnInsert.Text = (_currentAssetType == ExcelLibraryAssetType.Workbook) ? "Open" : "Insert";

            bool canSave = false;
            try
            {
                if (_currentAssetType == ExcelLibraryAssetType.Sheet)
                {
                    var selSheets = _xl?.ActiveWindow?.SelectedSheets as Excel.Sheets;
                    canSave = selSheets != null && selSheets.Count > 0;
                    if (selSheets != null) Marshal.ReleaseComObject(selSheets);
                }
                else if (_currentAssetType == ExcelLibraryAssetType.Workbook)
                {
                    canSave = _xl?.ActiveWorkbook != null;
                }
            }
            catch { /* ignore */ }

            btnSaveNew.Visible = canSave;
            btnReplaceSelected.Visible = canSave;
            btnReplaceSelected.Enabled = false;
            btnRename.Enabled = false;
            btnInsert.Visible = (listBoxAssets.SelectedItem != null);
        }

        private void ChkSharedLibrary_CheckedChanged(object sender, EventArgs e)
        {
            if (chkSharedLibrary.Checked)
            {
                if (!EnsureSharedBasePathChosen())
                {
                    chkSharedLibrary.CheckedChanged -= ChkSharedLibrary_CheckedChanged;
                    chkSharedLibrary.Checked = false;
                    chkSharedLibrary.CheckedChanged += ChkSharedLibrary_CheckedChanged;
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

            using (var fbd = new FolderBrowserDialog
            {
                Description = "Select a Shared Library folder accessible by all users.\n" +
                              "Subfolders (XLS/Sheets, XLS/Workbooks) will be created if needed."
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
                var xlRoot = Path.Combine(basePath, AppFolderName);
                if (!IODir.Exists(xlRoot)) IODir.CreateDirectory(xlRoot);

                foreach (var name in new[] { "Sheets", "Workbooks" })
                {
                    var p = Path.Combine(xlRoot, name);
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
                    return Path.Combine(shared, AppFolderName);

                _useSharedLibrary = false;
                if (chkSharedLibrary != null)
                {
                    chkSharedLibrary.CheckedChanged -= ChkSharedLibrary_CheckedChanged;
                    chkSharedLibrary.Checked = false;
                    chkSharedLibrary.CheckedChanged += ChkSharedLibrary_CheckedChanged;
                }
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "OfficePragmatics_Library",
                AppFolderName);
        }

        private string GetLibraryPath()
        {
            string basePath = GetLibraryBasePath();
            string specific = Path.Combine(basePath, _currentAssetType == ExcelLibraryAssetType.Sheet ? "Sheets" : "Workbooks");
            if (!IODir.Exists(specific))
            {
                try { IODir.CreateDirectory(specific); } catch { /* ignore */ }
            }
            return specific;
        }

        private void LoadAssets()
        {
            _allAssets = new List<LibraryAsset>();
            string libPath = GetLibraryPath();
            string[] files = Directory.Exists(libPath) ? Directory.GetFiles(libPath, "*.xlsx") : Array.Empty<string>();

            foreach (var filePath in files)
            {
                var fileNameNoExt = Path.GetFileNameWithoutExtension(filePath);
                var assetName = fileNameNoExt;
                var thumbnailPath = Path.ChangeExtension(filePath, ".png");

                if (!IOFile.Exists(thumbnailPath))
                {
                    try
                    {
                        using (var bmp = new SD.Bitmap(120, 90))
                        using (var g = SD.Graphics.FromImage(bmp))
                        {
                            g.Clear(UIColors.SurfaceAlt);
                            using (var font = new SD.Font("Segoe UI", 9f))
                            using (var brush = new SD.SolidBrush(UIColors.TextSecondary))
                                g.DrawString("No Preview", font, brush, 10, 38);
                            bmp.Save(thumbnailPath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                    catch { /* ignore */ }
                }

                _allAssets.Add(new LibraryAsset
                {
                    Name = assetName,
                    FilePath = filePath,
                    ThumbnailPath = thumbnailPath,
                    AssetType = _currentAssetType,
                    UniqueID = ""
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
            string term = txtSearch.Text.ToLowerInvariant();
            _filteredAssets = string.IsNullOrWhiteSpace(term)
                ? new List<LibraryAsset>(_allAssets)
                : _allAssets.Where(a => a.Name.ToLowerInvariant().Contains(term)).ToList();

            listBoxAssets.BeginUpdate();
            listBoxAssets.DataSource = null;
            listBoxAssets.DataSource = _filteredAssets;
            listBoxAssets.DisplayMember = "Name";
            listBoxAssets.EndUpdate();

            if (listBoxAssets.Items.Count > 0) listBoxAssets.SelectedIndex = 0;
            else ClearPreview();

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

        private void ListBoxAssets_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxAssets.SelectedItem is LibraryAsset a)
            {
                lblAssetName.Text = a.Name;
                if (IOFile.Exists(a.ThumbnailPath))
                {
                    try
                    {
                        if (pictureBoxThumbnail.Image != null) pictureBoxThumbnail.Image.Dispose();
                        using (var fs = new FileStream(a.ThumbnailPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            pictureBoxThumbnail.Image = SD.Image.FromStream(fs);
                    }
                    catch
                    {
                        if (pictureBoxThumbnail.Image != null) { pictureBoxThumbnail.Image.Dispose(); pictureBoxThumbnail.Image = null; }
                    }
                }
                else
                {
                    if (pictureBoxThumbnail.Image != null) { pictureBoxThumbnail.Image.Dispose(); pictureBoxThumbnail.Image = null; }
                }

                btnReplaceSelected.Enabled = true;
                btnInsert.Visible = true;
                btnDelete.Enabled = true;
                btnRename.Enabled = true;
            }
            else ClearPreview();
        }

        private void SaveNew()
        {
            var defaultName = lblAssetName.Text;
            string assetName = PromptForAssetName("Enter a name for the new asset:", defaultName);
            if (string.IsNullOrWhiteSpace(assetName)) return;

            string clean = Path.GetInvalidFileNameChars().Aggregate(assetName, (curr, c) => curr.Replace(c.ToString(), ""));
            string filePath = Path.Combine(GetLibraryPath(), clean + ".xlsx");
            string thumbPath = Path.ChangeExtension(filePath, ".png");

            if (IOFile.Exists(filePath))
            {
                MessageBox.Show("An asset with this name already exists. Choose another name.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            try
            {
                if (_currentAssetType == ExcelLibraryAssetType.Sheet)
                {
                    SaveSelectedSheets(filePath, thumbPath);
                }
                else
                {
                    SaveActiveWorkbook(filePath, thumbPath);
                }
                MessageBox.Show($"{_currentAssetType} '{assetName}' saved successfully.", "Success",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saving asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReplaceSelected()
        {
            if (!(listBoxAssets.SelectedItem is LibraryAsset selected)) return;

            string newName = lblAssetName.Text;
            if (string.IsNullOrWhiteSpace(newName))
            {
                MessageBox.Show("Asset name cannot be empty.", "Invalid Name", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string clean = Path.GetInvalidFileNameChars().Aggregate(newName, (curr, c) => curr.Replace(c.ToString(), ""));

            var confirm = MessageBox.Show($"Replace '{selected.Name}' with current {(_currentAssetType == ExcelLibraryAssetType.Sheet ? "selected sheets" : "workbook")}?\nNew name will be '{clean}'.",
                                          "Confirm Replace", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            try
            {
                try
                {
                    if (pictureBoxThumbnail.Image != null) { pictureBoxThumbnail.Image.Dispose(); pictureBoxThumbnail.Image = null; }
                    if (IOFile.Exists(selected.FilePath)) IOFile.Delete(selected.FilePath);
                    if (IOFile.Exists(selected.ThumbnailPath)) IOFile.Delete(selected.ThumbnailPath);
                }
                catch (IOException ioex)
                {
                    MessageBox.Show($"File in use: {ioex.Message}", "File Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                string newFile = Path.Combine(GetLibraryPath(), clean + ".xlsx");
                string newThumb = Path.ChangeExtension(newFile, ".png");

                if (_currentAssetType == ExcelLibraryAssetType.Sheet)
                    SaveSelectedSheets(newFile, newThumb);
                else
                    SaveActiveWorkbook(newFile, newThumb);

                MessageBox.Show($"Asset '{clean}' replaced successfully.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error replacing asset: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                LoadAssets();
            }
        }

        private void InsertSelected()
        {
            if (!(listBoxAssets.SelectedItem is LibraryAsset a)) return;

            try
            {
                if (_currentAssetType == ExcelLibraryAssetType.Sheet)
                    InsertSheetsFromLibrary(a);
                else
                    OpenWorkbookFromLibrary(a);

                if (_currentAssetType != ExcelLibraryAssetType.Workbook)
                    MessageBox.Show($"'{a.Name}' inserted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Insert/Open error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RenameSelected()
        {
            if (!(listBoxAssets.SelectedItem is LibraryAsset a)) return;

            string newName = PromptForAssetName("Enter a new name for this asset:", a.Name);
            if (string.IsNullOrWhiteSpace(newName)) return;
            string clean = Path.GetInvalidFileNameChars().Aggregate(newName, (curr, c) => curr.Replace(c.ToString(), ""));

            string newFile = Path.Combine(GetLibraryPath(), clean + ".xlsx");
            string newThumb = Path.ChangeExtension(newFile, ".png");

            if (IOFile.Exists(newFile) && !newFile.Equals(a.FilePath, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("An asset with this name already exists.", "Name Conflict", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            try
            {
                if (pictureBoxThumbnail.Image != null) { pictureBoxThumbnail.Image.Dispose(); pictureBoxThumbnail.Image = null; }
                if (IOFile.Exists(a.FilePath)) IOFile.Move(a.FilePath, newFile);
                if (IOFile.Exists(a.ThumbnailPath)) IOFile.Move(a.ThumbnailPath, newThumb);

                MessageBox.Show($"Renamed to '{clean}'.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Rename error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DeleteSelected()
        {
            if (!(listBoxAssets.SelectedItem is LibraryAsset a)) return;

            var confirm = MessageBox.Show($"Delete '{a.Name}' from the library?",
                                          "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (confirm != DialogResult.Yes) return;

            try
            {
                if (pictureBoxThumbnail.Image != null) { pictureBoxThumbnail.Image.Dispose(); pictureBoxThumbnail.Image = null; }
                if (IOFile.Exists(a.FilePath)) IOFile.Delete(a.FilePath);
                if (IOFile.Exists(a.ThumbnailPath)) IOFile.Delete(a.ThumbnailPath);
                MessageBox.Show($"'{a.Name}' deleted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                LoadAssets();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Delete error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // -------- Core Excel Operations --------

        private void SaveSelectedSheets(string xlsxFilePath, string pngThumbPath)
        {
            Excel.Workbook srcWb = _xl?.ActiveWorkbook;
            if (srcWb == null)
                throw new InvalidOperationException("No active workbook.");

            if (string.IsNullOrEmpty(srcWb.FullName) || !IOFile.Exists(srcWb.FullName))
                throw new InvalidOperationException("Please save the workbook locally before exporting sheets to the library.");

            var sheetNames = GetSelectedSheetNames();
            var kept = new HashSet<string>(sheetNames, StringComparer.OrdinalIgnoreCase);

            string tempSourcePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".xlsx");
            CopyWorkbookFast(srcWb.FullName, tempSourcePath);

            try
            {
                OOXmlExportSanitizer.RewriteForExport(tempSourcePath, kept);
                CopyWorkbookFast(tempSourcePath, xlsxFilePath);
                SheetPruner.Prune(xlsxFilePath, kept);
                EnsureThumbnail(pngThumbPath, "Sheet Asset");
            }
            finally
            {
                try { if (IOFile.Exists(tempSourcePath)) IOFile.Delete(tempSourcePath); } catch { /* ignore */ }
            }
        }

        private string[] GetSelectedSheetNames()
        {
            var selSheets = _xl?.ActiveWindow?.SelectedSheets as Excel.Sheets;
            if (selSheets == null || selSheets.Count == 0)
                throw new InvalidOperationException("Select at least one sheet before saving a Sheet asset.");

            try
            {
                var names = new string[selSheets.Count];
                for (int i = 1; i <= selSheets.Count; i++)
                {
                    dynamic sh = selSheets[i];
                    try { names[i - 1] = (string)sh.Name; }
                    finally { try { Marshal.ReleaseComObject(sh); } catch { } }
                }
                return names;
            }
            finally
            {
                try { Marshal.ReleaseComObject(selSheets); } catch { }
            }
        }

        private static void CopyWorkbookFast(string sourcePath, string destPath)
        {
            var dir = Path.GetDirectoryName(destPath);
            if (!string.IsNullOrEmpty(dir) && !IODir.Exists(dir)) IODir.CreateDirectory(dir);

            if (IOFile.Exists(destPath))
            {
                try { IOFile.Delete(destPath); } catch { /* best-effort */ }
            }

            const int Buffer = 1 << 21;
            using (var src = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete, Buffer))
            using (var dst = new FileStream(destPath, FileMode.Create, FileAccess.Write, FileShare.None, Buffer))
            {
                src.CopyTo(dst, Buffer);
            }
        }

        private void SaveActiveWorkbook(string xlsxFilePath, string pngThumbPath)
        {
            var wb = _xl.ActiveWorkbook;
            if (wb == null) throw new InvalidOperationException("No active workbook.");
            if (string.IsNullOrEmpty(wb.Path) || string.IsNullOrEmpty(wb.FullName) || !IOFile.Exists(wb.FullName))
                throw new InvalidOperationException("Active workbook must be saved locally first.");

            IOFile.Copy(wb.FullName, xlsxFilePath, overwrite: false);
            EnsureThumbnail(pngThumbPath, "Workbook Asset");
        }

        private void InsertSheetsFromLibrary(LibraryAsset asset)
        {
            Excel.Workbook targetWb = _xl.ActiveWorkbook;
            if (targetWb == null) throw new InvalidOperationException("No active workbook to insert into.");

            Excel.Worksheet activeSheet = null;
            Excel.Workbook src = null;

            try
            {
                activeSheet = _xl.ActiveSheet as Excel.Worksheet;
                src = _xl.Workbooks.Open(asset.FilePath, ReadOnly: true);

                int insertIndex = targetWb.Sheets.Count;
                if (activeSheet != null) insertIndex = activeSheet.Index;

                for (int i = 1; i <= src.Sheets.Count; i++)
                {
                    var ws = (Excel.Worksheet)src.Sheets[i];
                    ws.Copy(After: targetWb.Sheets[insertIndex]);
                    insertIndex++;
                    Marshal.ReleaseComObject(ws);
                }
            }
            finally
            {
                if (src != null)
                {
                    try { src.Close(SaveChanges: false); } catch { }
                    Marshal.ReleaseComObject(src);
                }
                if (activeSheet != null) Marshal.ReleaseComObject(activeSheet);
            }
        }

        private void OpenWorkbookFromLibrary(LibraryAsset asset)
        {
            _xl.Workbooks.Open(asset.FilePath);
        }

        private static void EnsureThumbnail(string pngThumbPath, string text)
        {
            try
            {
                using (var bmp = new SD.Bitmap(120, 90))
                using (var g = SD.Graphics.FromImage(bmp))
                {
                    g.Clear(UIColors.SurfaceAlt);
                    using (var f = new SD.Font("Segoe UI", 8.5f))
                    using (var br = new SD.SolidBrush(UIColors.TextSecondary))
                    {
                        g.DrawString(text, f, br, 6, 38);
                    }
                    bmp.Save(pngThumbPath, System.Drawing.Imaging.ImageFormat.Png);
                }
            }
            catch { /* ignore */ }
        }

        private string PromptForAssetName(string promptMessage, string defaultName = "")
        {
            string name = "";
            using (var prompt = new Form
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
                var textLabel = new Label
                {
                    Left = Scale(20),
                    Top = Scale(20),
                    Text = promptMessage,
                    AutoSize = true,
                    Font = new SD.Font("Segoe UI", 9.5f),
                    ForeColor = UIColors.TextPrimary
                };

                var textBox = new TextBox
                {
                    Left = Scale(20),
                    Top = Scale(50),
                    Width = prompt.ClientSize.Width - Scale(40),
                    Text = defaultName,
                    Font = new SD.Font("Segoe UI", 10f),
                    BorderStyle = BorderStyle.FixedSingle
                };

                var ok = CreateModernButton("OK", new SD.Size(Scale(80), Scale(32)), true);
                ok.Left = prompt.ClientSize.Width - Scale(180);
                ok.Top = Scale(90);
                ok.DialogResult = DialogResult.OK;
                ok.Click += (_, __) => prompt.Close();

                var cancel = CreateModernButton("Cancel", new SD.Size(Scale(80), Scale(32)), false);
                cancel.Left = prompt.ClientSize.Width - Scale(90);
                cancel.Top = Scale(90);
                cancel.DialogResult = DialogResult.Cancel;
                cancel.Click += (_, __) => prompt.Close();

                prompt.Controls.Add(textBox);
                prompt.Controls.Add(ok);
                prompt.Controls.Add(cancel);
                prompt.Controls.Add(textLabel);
                prompt.AcceptButton = ok;
                prompt.CancelButton = cancel;

                if (prompt.ShowDialog(this) == DialogResult.OK)
                    name = textBox.Text;
            }
            return name;
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

    public static class ExcelLibraryUtility
    {
        private static Excel.Application XlApp => Globals.ThisAddIn.Application;

        public static void OpenLibrary(ExcelLibraryAssetType? initialTab = null)
        {
            try
            {
                using (var form = new ExcelLibraryForm(XlApp, initialTab))
                    form.ShowDialog();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening Library: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}