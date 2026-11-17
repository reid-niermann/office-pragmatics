using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using SD = System.Drawing;

// VSTO Tools
using Microsoft.Office.Tools;

using OfficePragmatics_Excel.Settings;

namespace OfficePragmatics_Powerpoint.Classes.ExcelLinking
{
    // ==============================================================================================
    // CLASS 1: THE MANAGER (Handles the VSTO Task Pane)
    // ==============================================================================================
    public class LinkManager
    {
        private static CustomTaskPane _taskPane;
        private static LinkManagerControl _userControl;

        public static void ShowLinkManager(PowerPoint.Application app)
        {
            if (_taskPane == null)
            {
                _userControl = new LinkManagerControl();

                // Wire up the CloseRequested event to hide the pane
                _userControl.CloseRequested += () => {
                    if (_taskPane != null) _taskPane.Visible = false;
                };

                _taskPane = Globals.ThisAddIn.CustomTaskPanes.Add(_userControl, "Link Manager");
                _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
                _taskPane.Width = 550;

                _taskPane.VisibleChanged += (s, e) =>
                {
                    if (_taskPane.Visible)
                    {
                        _userControl.Initialize(app);
                    }
                };
            }

            _taskPane.Visible = true;
            _userControl.Initialize(app);
        }
    }

    // ==============================================================================================
    // CLASS 2: THE CONTROL (The UI Logic)
    // ==============================================================================================
    internal class LinkManagerControl : UserControl
    {
        // ================= FIELDS =================
        private PowerPoint.Application ppApp;
        private PowerPoint.Presentation pres;

        public event Action CloseRequested;

        private List<LinkInfo> _allLinks = new List<LinkInfo>();

        // UI Controls
        private Panel headerPanel;
        private Label searchLabel;
        private TextBox searchBox;

        private BetterListView linkListView;
        private Panel footerPanel;
        private Label statusLabel;
        private TableLayoutPanel buttonTable;

        private Button breakLinksButton;
        private Button changeSourceButton;
        private Button refreshButton;
        private Button replaceButton;

        private Label emptyState;
        private Panel emptyHost;
        private Panel bodyPanel;
        private ContextMenuStrip listMenu;

        // ================= DATA CLASS =================
        private class LinkInfo
        {
            public string LinkType { get; set; }
            public string SourceFile { get; set; }
            public string SourceLocation { get; set; }
            public int SlideNumber { get; set; }
            public string ShapeName { get; set; }
            public PowerPoint.Slide Slide { get; set; }
            public PowerPoint.Shape Shape { get; set; }
            public PowerPoint.TextRange TextRange { get; set; }
            public int RunIndex { get; set; }
            public string LinkId { get; set; }
            public Dictionary<string, string> Tags { get; set; }
            public string HyperlinkAddress { get; set; }
            public bool FileExists { get; set; }
        }

        // ================= STYLING =================
        private static class Ui
        {
            public static readonly Color TitleBg = SD.Color.FromArgb(32, 42, 102);
            public static readonly Color TitleText = SD.Color.White;
            public static readonly Color TextPrimary = SD.ColorTranslator.FromHtml("#111827");
            public static readonly Color TextSecondary = SD.ColorTranslator.FromHtml("#6B7280");
            public static readonly Color Border = SD.ColorTranslator.FromHtml("#E5E7EB");
            public static readonly Color HeaderBack = SD.ColorTranslator.FromHtml("#F9FAFB");
            public static readonly Color RowSelect = SD.ColorTranslator.FromHtml("#F3F4F6");
            public static readonly Color DotSuccess = SD.ColorTranslator.FromHtml("#10B981");
            public static readonly Color DotError = SD.ColorTranslator.FromHtml("#EF4444");
            public static readonly Color IconColor = SD.ColorTranslator.FromHtml("#4B5563");
        }

        private const int GWL_STYLE = -16;
        private const int WS_VSCROLL = 0x00200000;
        [DllImport("user32.dll")] private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        // ================= CONSTRUCTOR =================
        public LinkManagerControl()
        {
            this.DoubleBuffered = true;
            this.BackColor = Color.White;
            InitializeModernUI();
        }

        public void Initialize(PowerPoint.Application app)
        {
            this.ppApp = app;
            if (ppApp != null && ppApp.ActivePresentation != null)
            {
                this.pres = ppApp.ActivePresentation;
                LoadLinks();
            }
            LayoutColumns();
        }

        private void AfterHandle(Action action)
        {
            if (IsHandleCreated) BeginInvoke(action);
            else { EventHandler h = null; h = (s, e) => { HandleCreated -= h; action(); }; HandleCreated += h; }
        }

        // ================= KEYBOARD HANDLING (ESC to Close) =================
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Escape)
            {
                CloseRequested?.Invoke();
                return true;
            }
            if (keyData == Keys.F5)
            {
                RefreshButton_Click(this, EventArgs.Empty);
                return true;
            }
            if (keyData == (Keys.Control | Keys.F))
            {
                searchBox.Focus();
                return true;
            }
            if (keyData == Keys.Delete && breakLinksButton.Enabled)
            {
                BreakLinksButton_Click(this, EventArgs.Empty);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        // ================= UI INITIALIZATION =================
        private void InitializeModernUI()
        {
            var mainContainer = new Panel { Dock = DockStyle.Fill, BackColor = Color.White };
            Controls.Add(mainContainer);

            // 1. HEADER
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = Theme.Px(42),
                BackColor = Ui.TitleBg,
                Padding = new Padding(0)
            };

            var headerTable = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 1,
                BackColor = Color.Transparent,
                Padding = new Padding(Theme.Px(10), 0, Theme.Px(10), 0)
            };

            headerTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            headerTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            headerTable.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

            searchLabel = new Label
            {
                Text = "Search:",
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                ForeColor = Ui.TitleText,
                AutoSize = true,
                Anchor = AnchorStyles.Left,
                TextAlign = ContentAlignment.MiddleLeft,
                Margin = new Padding(0, 0, Theme.Px(10), 0)
            };

            var searchContainer = new Panel
            {
                Height = Theme.Px(24),
                BackColor = Color.White,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Padding = new Padding(Theme.Px(5), Theme.Px(4), Theme.Px(5), Theme.Px(4))
            };

            searchBox = new TextBox
            {
                Dock = DockStyle.Fill,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Ui.TextPrimary,
                BorderStyle = BorderStyle.None
            };
            searchBox.TextChanged += SearchBox_TextChanged;

            searchContainer.Controls.Add(searchBox);

            headerTable.Controls.Add(searchLabel, 0, 0);
            headerTable.Controls.Add(searchContainer, 1, 0);
            headerPanel.Controls.Add(headerTable);

            // 2. FOOTER
            footerPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = Theme.Px(100),
                BackColor = SD.ColorTranslator.FromHtml("#F9FAFB"),
                Padding = new Padding(Theme.Px(15), Theme.Px(12), Theme.Px(15), Theme.Px(5))
            };
            footerPanel.Paint += (s, e) => {
                using (var p = new Pen(Ui.Border)) e.Graphics.DrawLine(p, 0, 0, footerPanel.Width, 0);
            };

            // Button Table
            buttonTable = new TableLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = Theme.Px(50),
                ColumnCount = 4,
                RowCount = 1,
                Margin = new Padding(0)
            };
            buttonTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            buttonTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            buttonTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));
            buttonTable.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 25f));

            replaceButton = CreateButton("Replace...", false);
            replaceButton.Click += ReplaceButton_Click;

            refreshButton = CreateButton("Refresh", false);
            refreshButton.Click += RefreshButton_Click;

            changeSourceButton = CreateButton("Source...", false);
            changeSourceButton.Click += ChangeSourceButton_Click;

            breakLinksButton = CreateButton("Break", false);
            breakLinksButton.ForeColor = SD.ColorTranslator.FromHtml("#B91C1C");
            breakLinksButton.Click += BreakLinksButton_Click;

            buttonTable.Controls.Add(breakLinksButton, 0, 0);
            buttonTable.Controls.Add(changeSourceButton, 1, 0);
            buttonTable.Controls.Add(refreshButton, 2, 0);
            buttonTable.Controls.Add(replaceButton, 3, 0);

            // Status Label
            statusLabel = new Label
            {
                Text = "Ready",
                AutoSize = false,
                TextAlign = ContentAlignment.BottomLeft,
                ForeColor = Ui.TextSecondary,
                Font = new Font("Segoe UI", 9F),
                Dock = DockStyle.Bottom,
                Height = Theme.Px(24)
            };

            footerPanel.Controls.Add(buttonTable);
            footerPanel.Controls.Add(statusLabel);

            // 3. Main Body
            bodyPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(0) };

            linkListView = new BetterListView
            {
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                MultiSelect = true,
                Dock = DockStyle.Fill,
                BackColor = Color.White,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 9.5F),
                HeaderStyle = ColumnHeaderStyle.Nonclickable,
                ShowItemToolTips = true,
                OwnerDraw = true
            };
            Theme.DoubleBuffer(linkListView);

            // Columns
            linkListView.Columns.Add("T", 30);
            linkListView.Columns.Add("Source", 100);
            linkListView.Columns.Add("Location", 80);
            linkListView.Columns.Add("#", 30);
            linkListView.Columns.Add("", 30);
            linkListView.Columns.Add("", 30);

            linkListView.DrawColumnHeader += LinkListView_DrawColumnHeader;
            linkListView.DrawItem += (s, e) => { };
            linkListView.DrawSubItem += LinkListView_DrawSubItem;
            linkListView.SelectedIndexChanged += LinkListView_SelectedIndexChanged;
            linkListView.Resize += (s, e) => LayoutColumns();
            linkListView.MouseUp += LinkListView_MouseUp;

            emptyHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, Visible = false };
            emptyState = new Label
            {
                Text = "No links found",
                AutoSize = false,
                Dock = DockStyle.Fill,
                ForeColor = Ui.TextSecondary,
                Font = new Font("Segoe UI", 12F),
                TextAlign = ContentAlignment.MiddleCenter
            };
            emptyHost.Controls.Add(emptyState);

            listMenu = new ContextMenuStrip();
            listMenu.Items.Add("Break Links", null, delegate { BreakLinksButton_Click(null, EventArgs.Empty); });
            listMenu.Items.Add("Change Source…", null, delegate { ChangeSourceButton_Click(null, EventArgs.Empty); });
            listMenu.Items.Add("-");
            listMenu.Items.Add("Refresh", null, delegate { RefreshButton_Click(null, EventArgs.Empty); });
            listMenu.Items.Add("Open Source Folder", null, delegate { OpenSourceFolder_Context(); });
            linkListView.ContextMenuStrip = listMenu;

            bodyPanel.Controls.Add(linkListView);
            bodyPanel.Controls.Add(emptyHost);

            mainContainer.Controls.Add(bodyPanel);
            mainContainer.Controls.Add(footerPanel);
            mainContainer.Controls.Add(headerPanel);

            UpdateActionButtons();
        }

        private Button CreateButton(string text, bool isPrimary)
        {
            var b = new Button
            {
                Text = text,
                Dock = DockStyle.Fill,
                Margin = new Padding(3),
                Font = new Font("Segoe UI", 10F),
                UseVisualStyleBackColor = false,
                Cursor = Cursors.Hand
            };
            if (isPrimary) Theme.PrimaryButton(b); else Theme.GhostButton(b);
            return b;
        }

        // ================= DRAWING =================
        private void LinkListView_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            using (var b = new SolidBrush(Ui.HeaderBack)) e.Graphics.FillRectangle(b, e.Bounds);
            using (var p = new Pen(Ui.Border)) e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);

            if (e.ColumnIndex < 4)
            {
                var textRect = new Rectangle(e.Bounds.X + 2, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height);
                TextRenderer.DrawText(e.Graphics, e.Header.Text.ToUpper(), new Font("Segoe UI", 7F, FontStyle.Bold), textRect, Ui.TextSecondary,
                    TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            }
        }

        private void LinkListView_DrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            bool selected = (e.ItemState & ListViewItemStates.Selected) == ListViewItemStates.Selected;
            using (var b = new SolidBrush(selected ? Ui.RowSelect : Color.White))
                e.Graphics.FillRectangle(b, e.Bounds);
            using (var p = new Pen(Ui.Border))
                e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);

            var link = (LinkInfo)e.Item.Tag;

            if (e.ColumnIndex == 0) // Type (Centered)
            {
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, linkListView.Font, e.Bounds, Ui.TextPrimary,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }
            else if (e.ColumnIndex == 4) // Status (Dot)
            {
                DrawStatusDot(e.Graphics, e.Bounds, link.FileExists);
            }
            else if (e.ColumnIndex == 5) // Icon (Folder)
            {
                int iconSize = Theme.Px(14);
                int x = e.Bounds.X + (e.Bounds.Width - iconSize) / 2;
                int y = e.Bounds.Y + (e.Bounds.Height - iconSize) / 2;
                using (var pen = new Pen(Ui.IconColor, 1.5f))
                {
                    e.Graphics.DrawRectangle(pen, x, y + 3, iconSize, iconSize - 3);
                    e.Graphics.DrawLine(pen, x, y + 3, x, y);
                    e.Graphics.DrawLine(pen, x, y, x + (iconSize / 2), y);
                    e.Graphics.DrawLine(pen, x + (iconSize / 2), y, x + (iconSize / 2), y + 3);
                }
            }
            else
            {
                var textRect = new Rectangle(e.Bounds.X + Theme.Px(5), e.Bounds.Y, e.Bounds.Width - Theme.Px(5), e.Bounds.Height);
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, linkListView.Font, textRect, Ui.TextPrimary,
                    TextFormatFlags.EndEllipsis | TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            }
        }

        private void DrawStatusDot(Graphics g, Rectangle bounds, bool isOk)
        {
            g.SmoothingMode = SmoothingMode.AntiAlias;
            Color dotColor = isOk ? Ui.DotSuccess : Ui.DotError;
            int dotSize = Theme.Px(8);
            int x = bounds.X + (bounds.Width - dotSize) / 2;
            int y = bounds.Y + (bounds.Height - dotSize) / 2;
            using (var brush = new SolidBrush(dotColor)) g.FillEllipse(brush, x, y, dotSize, dotSize);
            g.SmoothingMode = SmoothingMode.Default;
        }

        // ================= LAYOUT =================
        private void LayoutColumns()
        {
            if (linkListView.Columns.Count < 6) return;
            int w = linkListView.ClientSize.Width - (IsVScrollVisible() ? SystemInformation.VerticalScrollBarWidth : 0);
            if (w < 50) return;

            int typeColWidth = Theme.Px(30);
            int iconColWidth = Theme.Px(30);
            int slideColWidth = Theme.Px(30);
            int statusColWidth = Theme.Px(30);
            int fixedWidth = typeColWidth + iconColWidth + slideColWidth + statusColWidth;

            int availableWidth = w - fixedWidth;

            linkListView.Columns[0].Width = typeColWidth;
            linkListView.Columns[1].Width = (int)(availableWidth * 0.50); // Source
            linkListView.Columns[2].Width = (int)(availableWidth * 0.50); // Location
            linkListView.Columns[3].Width = slideColWidth;
            linkListView.Columns[4].Width = statusColWidth;
            linkListView.Columns[5].Width = iconColWidth;
        }

        private bool IsVScrollVisible()
        {
            int style = GetWindowLong(linkListView.Handle, GWL_STYLE);
            return (style & WS_VSCROLL) == WS_VSCROLL;
        }

        // ================= EVENTS & LOGIC =================
        private void LinkListView_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Left) return;
            var hitInfo = linkListView.HitTest(e.Location);
            if (hitInfo.Item != null && hitInfo.SubItem != null && hitInfo.Item.SubItems.IndexOf(hitInfo.SubItem) == 5)
            {
                OpenFolderForLink((LinkInfo)hitInfo.Item.Tag);
            }
        }

        private void OpenFolderForLink(LinkInfo link)
        {
            if (link == null || string.IsNullOrWhiteSpace(link.SourceFile)) return;
            try
            {
                string path = link.SourceFile;
                if (!File.Exists(path) && !Directory.Exists(path))
                {
                    path = Path.GetDirectoryName(path);
                    if (!Directory.Exists(path)) return;
                    Process.Start("explorer.exe", path);
                }
                else
                {
                    Process.Start("explorer.exe", "/select,\"" + path + "\"");
                }
            }
            catch { }
        }

        private void OpenSourceFolder_Context()
        {
            if (linkListView.SelectedItems.Count > 0) OpenFolderForLink((LinkInfo)linkListView.SelectedItems[0].Tag);
        }

        private void SearchBox_TextChanged(object sender, EventArgs e) => PopulateList(searchBox.Text);

        private void UpdateActionButtons()
        {
            bool hasSel = linkListView.SelectedItems.Count > 0;
            breakLinksButton.Enabled = hasSel;
            changeSourceButton.Enabled = hasSel;
            breakLinksButton.ForeColor = hasSel ? SD.ColorTranslator.FromHtml("#B91C1C") : Color.FromArgb(160, 160, 160);
        }

        private void LinkListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateActionButtons();
            if (linkListView.SelectedItems.Count == 0 || ppApp == null) return;
            var link = (LinkInfo)linkListView.SelectedItems[0].Tag;
            try
            {
                ppApp.ActiveWindow.View.GotoSlide(link.SlideNumber);
                link.Shape.Select(Office.MsoTriState.msoTrue);
                if (link.LinkType == "Text" && link.TextRange != null) link.TextRange.Select();
            }
            catch { }
        }

        // ================= REPLACE LOGIC =================
        private void ReplaceButton_Click(object sender, EventArgs e)
        {
            using (var dlg = new ReplaceDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    string find = dlg.FindText;
                    string replace = dlg.ReplaceText;
                    if (string.IsNullOrEmpty(find)) return;

                    int count = 0;
                    foreach (var link in _allLinks)
                    {
                        if (!string.IsNullOrEmpty(link.SourceFile) && link.SourceFile.Contains(find))
                        {
                            try
                            {
                                string newPath = link.SourceFile.Replace(find, replace);

                                if (link.Shape != null)
                                {
                                    link.Shape.Tags.Add("Pragmatics_SourceWorkbookPath", newPath);
                                }

                                if (link.LinkType == "Text" && !string.IsNullOrEmpty(link.HyperlinkAddress))
                                {
                                    if (link.HyperlinkAddress.Contains(find))
                                    {
                                        string newAddr = link.HyperlinkAddress.Replace(find, replace);
                                        if (link.TextRange != null && link.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink != null)
                                        {
                                            link.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Hyperlink.Address = newAddr;
                                        }
                                        link.HyperlinkAddress = newAddr;
                                    }
                                }

                                link.SourceFile = newPath;
                                link.FileExists = File.Exists(newPath);
                                count++;
                            }
                            catch { }
                        }
                    }

                    MessageBox.Show($"Updated {count} link(s).", "Replace Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    LoadLinks();
                }
            }
        }

        private void LoadLinks()
        {
            statusLabel.Text = "Scanning...";
            Application.DoEvents();
            _allLinks.Clear();

            if (pres != null)
            {
                int currentSlide = 0;
                foreach (PowerPoint.Slide slide in pres.Slides)
                {
                    currentSlide++;
                    foreach (PowerPoint.Shape shape in slide.Shapes)
                    {
                        if (shape.Tags["Pragmatics_LinkedChart_ID"] != "")
                        {
                            var path = shape.Tags["Pragmatics_SourceWorkbookPath"];
                            _allLinks.Add(new LinkInfo
                            {
                                LinkType = "Chart",
                                SourceFile = path,
                                SourceLocation = shape.Tags["Pragmatics_SourceWorksheet"] + "!" + shape.Tags["Pragmatics_SourceChart"],
                                SlideNumber = slide.SlideIndex,
                                ShapeName = shape.Name,
                                Slide = slide,
                                Shape = shape,
                                LinkId = shape.Tags["Pragmatics_LinkedChart_ID"],
                                Tags = GetShapeTags(shape),
                                FileExists = File.Exists(path)
                            });
                        }
                        if (shape.Tags["Pragmatics_LinkedRange_ID"] != "")
                        {
                            var path = shape.Tags["Pragmatics_SourceWorkbookPath"];
                            _allLinks.Add(new LinkInfo
                            {
                                LinkType = shape.Tags["Pragmatics_ImportType"] == "Table" ? "Table" : "Range",
                                SourceFile = path,
                                SourceLocation = shape.Tags["Pragmatics_SourceWorksheet"] + "!" + shape.Tags["Pragmatics_RangeName"],
                                SlideNumber = slide.SlideIndex,
                                ShapeName = shape.Name,
                                Slide = slide,
                                Shape = shape,
                                LinkId = shape.Tags["Pragmatics_LinkedRange_ID"],
                                Tags = GetShapeTags(shape),
                                FileExists = File.Exists(path)
                            });
                        }
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue && shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;
                            for (int i = 1; i <= textRange.Runs().Count; i++)
                            {
                                var run = textRange.Runs(i);
                                var actionSetting = run.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                                var hyperlink = actionSetting.Hyperlink;
                                if (hyperlink != null && hyperlink.Address != null && hyperlink.Address.StartsWith("pragmatics-textlink://"))
                                {
                                    try
                                    {
                                        var uri = new Uri(hyperlink.Address);
                                        string query = uri.Query.StartsWith("?") ? uri.Query.Substring(1) : uri.Query;
                                        var qParams = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                                        foreach (var part in query.Split('&'))
                                        {
                                            var kv = part.Split('=');
                                            if (kv.Length == 2) qParams[kv[0]] = Uri.UnescapeDataString(kv[1]);
                                        }
                                        string wbPath = qParams.ContainsKey("wb") ? qParams["wb"] : "";
                                        string wsName = qParams.ContainsKey("ws") ? qParams["ws"] : "";
                                        string nrName = qParams.ContainsKey("nr") ? qParams["nr"] : "";
                                        if (wbPath != "" && wsName != "" && nrName != "")
                                        {
                                            _allLinks.Add(new LinkInfo
                                            {
                                                LinkType = "Text",
                                                SourceFile = wbPath,
                                                SourceLocation = wsName + "!" + nrName,
                                                SlideNumber = slide.SlideIndex,
                                                ShapeName = shape.Name,
                                                Slide = slide,
                                                Shape = shape,
                                                TextRange = run,
                                                RunIndex = i,
                                                HyperlinkAddress = hyperlink.Address,
                                                FileExists = File.Exists(wbPath)
                                            });
                                        }
                                    }
                                    catch { }
                                }
                                if (hyperlink != null) Marshal.ReleaseComObject(hyperlink);
                                if (actionSetting != null) Marshal.ReleaseComObject(actionSetting);
                                if (run != null) Marshal.ReleaseComObject(run);
                            }
                        }
                    }
                }
            }

            statusLabel.Text = $"Found {_allLinks.Count} link(s)";
            PopulateList(searchBox.Text);
        }

        private string GetTypeCode(string linkType)
        {
            if (string.IsNullOrEmpty(linkType)) return "";
            switch (linkType.ToLower())
            {
                case "chart": return "C";
                case "range": return "R";
                case "table": return "R";
                case "text": return "T";
                default: return linkType.Substring(0, 1).ToUpper();
            }
        }

        private void PopulateList(string filter)
        {
            linkListView.BeginUpdate();
            linkListView.Items.Clear();
            var displayList = _allLinks;

            if (!string.IsNullOrWhiteSpace(filter))
            {
                filter = filter.ToLower();
                displayList = _allLinks.Where(x => x.SourceFile.ToLower().Contains(filter) || x.LinkType.ToLower().Contains(filter) || Path.GetFileName(x.SourceFile).ToLower().Contains(filter)).ToList();
            }

            foreach (var link in displayList)
            {
                var item = new ListViewItem(GetTypeCode(link.LinkType));
                item.SubItems.Add(Path.GetFileName(link.SourceFile));
                item.SubItems.Add(link.SourceLocation);
                item.SubItems.Add(link.SlideNumber.ToString());
                item.SubItems.Add("");
                item.SubItems.Add("");
                item.Tag = link;
                item.ToolTipText = link.SourceFile;
                linkListView.Items.Add(item);
            }

            linkListView.EndUpdate();
            bool none = linkListView.Items.Count == 0;
            linkListView.Visible = !none;
            emptyHost.Visible = none;
            if (none) emptyState.Text = string.IsNullOrEmpty(filter) ? "No links found" : "No matches";

            AfterHandle(LayoutColumns);
            UpdateActionButtons();
        }

        private Dictionary<string, string> GetShapeTags(PowerPoint.Shape shape)
        {
            var tags = new Dictionary<string, string>();
            for (int i = 1; i <= shape.Tags.Count; i++) tags[shape.Tags.Name(i)] = shape.Tags.Value(i);
            return tags;
        }

        private void BreakLinksButton_Click(object sender, EventArgs e)
        {
            if (linkListView.SelectedItems.Count == 0) return;
            if (MessageBox.Show($"Are you sure you want to break {linkListView.SelectedItems.Count} link(s)?\nThis action cannot be undone.", "Confirm Break Links", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

            statusLabel.Text = "Breaking links...";
            Application.DoEvents();

            foreach (ListViewItem item in linkListView.SelectedItems)
            {
                var link = (LinkInfo)item.Tag;
                try
                {
                    if (link.LinkType == "Text")
                    {
                        if (link.TextRange != null)
                            link.TextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick].Action = PowerPoint.PpActionType.ppActionNone;
                    }
                    else
                    {
                        var pragmaticsTags = new[] { "Pragmatics_LinkedChart_ID", "Pragmatics_LinkedRange_ID", "Pragmatics_SourceWorkbookPath", "Pragmatics_SourceWorksheet", "Pragmatics_SourceChart", "Pragmatics_RangeName", "Pragmatics_ImportType" };
                        foreach (var t in pragmaticsTags) if (!string.IsNullOrEmpty(link.Shape.Tags[t])) link.Shape.Tags.Add(t, "");
                    }
                }
                catch { }
            }
            LoadLinks();
        }

        private void ChangeSourceButton_Click(object sender, EventArgs e)
        {
            if (linkListView.SelectedItems.Count == 0) return;

            var uniqueSourceFiles = linkListView.SelectedItems.Cast<ListViewItem>().Select(i => ((LinkInfo)i.Tag).SourceFile).Distinct().ToList();
            string initialDir = (uniqueSourceFiles.Count > 0 && !string.IsNullOrEmpty(uniqueSourceFiles[0])) ? Path.GetDirectoryName(uniqueSourceFiles[0]) : "";

            using (var openFileDialog = new OpenFileDialog { Filter = "Excel Files|*.xlsx;*.xlsm;*.xls|All Files|*.*", Title = "Select New Source File", InitialDirectory = initialDir })
            {
                if (openFileDialog.ShowDialog() != DialogResult.OK) return;
                string newSourceFile = openFileDialog.FileName;

                foreach (ListViewItem item in linkListView.SelectedItems)
                {
                    var link = (LinkInfo)item.Tag;
                    try
                    {
                        if (link.LinkType == "Text" && !string.IsNullOrEmpty(link.HyperlinkAddress))
                        {
                            // Standard update
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(link.Shape.Tags["Pragmatics_SourceWorkbookPath"]))
                                link.Shape.Tags.Add("Pragmatics_SourceWorkbookPath", newSourceFile);
                        }
                    }
                    catch { }
                }
                LoadLinks();
            }
        }

        private void RefreshButton_Click(object sender, EventArgs e) => LoadLinks();

        private class BetterListView : ListView
        {
            public BetterListView() { SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true); UpdateStyles(); }
            protected override void OnNotifyMessage(Message m) { if (m.Msg != 0x14) base.OnNotifyMessage(m); }
        }

        // ================= REVISED REPLACE DIALOG (FIXED BUTTON SIZE & ALIGNMENT) =================
        private class ReplaceDialog : Form
        {
            public string FindText { get { return txtFind.Text; } }
            public string ReplaceText { get { return txtReplace.Text; } }

            private TextBox txtFind;
            private TextBox txtReplace;

            public ReplaceDialog()
            {
                Text = "Replace in Paths";
                Size = new Size(Theme.Px(450), Theme.Px(230));
                FormBorderStyle = FormBorderStyle.FixedDialog;
                StartPosition = FormStartPosition.CenterParent;
                MaximizeBox = false; MinimizeBox = false;
                BackColor = Color.White;
                Font = new Font("Segoe UI", 9F);

                var mainLayout = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    RowCount = 5,
                    ColumnCount = 1,
                    Padding = new Padding(Theme.Px(15))
                };
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

                var lblFind = new Label { Text = "Find in path:", AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
                txtFind = new TextBox { Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 15) };

                var lblReplace = new Label { Text = "Replace with:", AutoSize = true, Margin = new Padding(0, 0, 0, 5) };
                txtReplace = new TextBox { Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 15) };

                // [FIXED] Explicit button panel with defined heights to prevent collapsing
                var btnPanel = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    FlowDirection = FlowDirection.RightToLeft,
                    AutoSize = false,
                    Margin = new Padding(0)
                };

                // [FIXED] Explicit sizes on buttons
                var btnCancel = new Button
                {
                    Text = "Cancel",
                    Size = new Size(Theme.Px(80), Theme.Px(30)),
                    DialogResult = DialogResult.Cancel,
                    Margin = new Padding(Theme.Px(10), 0, 0, 0),
                    UseVisualStyleBackColor = true
                };

                var btnOk = new Button
                {
                    Text = "Replace All",
                    Size = new Size(Theme.Px(90), Theme.Px(30)),
                    DialogResult = DialogResult.OK,
                    UseVisualStyleBackColor = true
                };

                // Add right-most first due to FlowDirection.RightToLeft
                btnPanel.Controls.Add(btnCancel);
                btnPanel.Controls.Add(btnOk);

                mainLayout.Controls.Add(lblFind, 0, 0);
                mainLayout.Controls.Add(txtFind, 0, 1);
                mainLayout.Controls.Add(lblReplace, 0, 2);
                mainLayout.Controls.Add(txtReplace, 0, 3);
                mainLayout.Controls.Add(btnPanel, 0, 4);

                Controls.Add(mainLayout);
                AcceptButton = btnOk;
                CancelButton = btnCancel;
            }
        }
    }
}