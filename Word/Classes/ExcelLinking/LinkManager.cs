// LinkManager.cs (.NET Framework 4.8 / C# 7.3)
// High-level coordination class for managing Excel links in Word documents.

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// High-level manager for Excel link operations in Word documents.
    /// Provides a unified interface for creating, updating, and managing links.
    /// </summary>
    public class LinkManager : IDisposable
    {
        private readonly Word.Application _wordApp;
        private Excel.Application _excelApp;
        private bool _excelOwnedByUs;
        private bool _disposed;

        /// <summary>
        /// Creates a new LinkManager for the specified Word application.
        /// </summary>
        public LinkManager(Word.Application wordApp)
        {
            _wordApp = wordApp ?? throw new ArgumentNullException(nameof(wordApp));
        }

        #region Static Entry Points (Ribbon Commands)

        /// <summary>
        /// Shows the Link Manager dialog for viewing and managing all links.
        /// </summary>
        public static void ShowLinkManager(Word.Application wordApp)
        {
            if (wordApp == null) return;

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                MessageBox.Show("No active document.", "Link Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (doc == null)
            {
                MessageBox.Show("No active document.", "Link Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (var dialog = new LinkManagerDialog(wordApp, doc))
            {
                dialog.ShowDialog();
            }
        }

        #endregion

        /// <summary>
        /// Gets or creates the Excel application instance.
        /// </summary>
        public Excel.Application ExcelApp
        {
            get
            {
                if (_excelApp == null)
                {
                    try
                    {
                        _excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                        _excelOwnedByUs = false;
                    }
                    catch
                    {
                        _excelApp = new Excel.Application();
                        _excelApp.Visible = false;
                        _excelOwnedByUs = true;
                    }
                }
                return _excelApp;
            }
        }

        /// <summary>
        /// Creates a chart link from Excel to Word.
        /// </summary>
        public Word.Shape CreateChartLink(
            string workbookPath,
            string sheetName,
            string chartName,
            Word.Range insertRange = null)
        {
            var doc = GetActiveDocument();
            if (doc == null) return null;

            if (insertRange == null)
            {
                insertRange = _wordApp.Selection.Range;
            }

            return ChartLinkUtility.CreateChartLink(
                doc, ExcelApp, workbookPath, sheetName, chartName, insertRange);
        }

        /// <summary>
        /// Creates a range link from Excel to Word.
        /// </summary>
        public Word.Shape CreateRangeLink(
            string workbookPath,
            string sheetName,
            string rangeAddress,
            Word.Range insertRange = null)
        {
            var doc = GetActiveDocument();
            if (doc == null) return null;

            if (insertRange == null)
            {
                insertRange = _wordApp.Selection.Range;
            }

            return RangeLinkUtility.CreateRangeLink(
                doc, ExcelApp, workbookPath, sheetName, rangeAddress, insertRange);
        }

        /// <summary>
        /// Creates a text link from an Excel cell to Word.
        /// </summary>
        public Word.ContentControl CreateTextLink(
            string workbookPath,
            string sheetName,
            string cellAddress,
            Word.Range insertRange = null,
            string formatString = null)
        {
            var doc = GetActiveDocument();
            if (doc == null) return null;

            if (insertRange == null)
            {
                insertRange = _wordApp.Selection.Range;
            }

            return TextLinkUtility.CreateTextLink(
                doc, ExcelApp, workbookPath, sheetName, cellAddress, insertRange, formatString);
        }

        /// <summary>
        /// Updates all Excel links in the active document.
        /// </summary>
        public bool UpdateAllLinks(bool showProgress = true)
        {
            var doc = GetActiveDocument();
            if (doc == null) return false;

            return UpdateLinks.UpdateAllLinks(doc, showProgress);
        }

        private Word.Document GetActiveDocument()
        {
            try
            {
                return _wordApp.ActiveDocument;
            }
            catch
            {
                return null;
            }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            if (_excelOwnedByUs && _excelApp != null)
            {
                try
                {
                    _excelApp.Quit();
                    Marshal.ReleaseComObject(_excelApp);
                }
                catch { }
            }
            _excelApp = null;
        }
    }

    /// <summary>
    /// Dialog for viewing and managing Excel links in the document.
    /// </summary>
    internal class LinkManagerDialog : Form
    {
        private ListView lvLinks;
        private Button btnRefreshAll;
        private Button btnRefreshSelected;
        private Button btnRemove;
        private Button btnClose;
        private Label lblSummary;
        private readonly Word.Application _wordApp;
        private readonly Word.Document _doc;

        public LinkManagerDialog(Word.Application wordApp, Word.Document doc)
        {
            _wordApp = wordApp;
            _doc = doc;
            InitializeComponent();
            LoadLinks();
        }

        private void InitializeComponent()
        {
            this.Text = "Excel Link Manager";
            this.Size = new System.Drawing.Size(700, 450);
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MinimumSize = new System.Drawing.Size(500, 300);

            // Summary label
            lblSummary = new Label
            {
                Location = new System.Drawing.Point(12, 12),
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 9f, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblSummary);

            // List view
            lvLinks = new ListView
            {
                Location = new System.Drawing.Point(12, 40),
                Size = new System.Drawing.Size(560, 320),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = true
            };
            lvLinks.Columns.Add("Type", 80);
            lvLinks.Columns.Add("Workbook", 150);
            lvLinks.Columns.Add("Sheet", 100);
            lvLinks.Columns.Add("Target", 120);
            lvLinks.Columns.Add("Status", 80);
            this.Controls.Add(lvLinks);

            // Buttons
            int btnX = 585;
            btnRefreshAll = new Button
            {
                Text = "Refresh All",
                Location = new System.Drawing.Point(btnX, 40),
                Width = 90,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnRefreshAll.Click += BtnRefreshAll_Click;
            this.Controls.Add(btnRefreshAll);

            btnRefreshSelected = new Button
            {
                Text = "Refresh",
                Location = new System.Drawing.Point(btnX, 75),
                Width = 90,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnRefreshSelected.Click += BtnRefreshSelected_Click;
            this.Controls.Add(btnRefreshSelected);

            btnRemove = new Button
            {
                Text = "Remove",
                Location = new System.Drawing.Point(btnX, 110),
                Width = 90,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            btnRemove.Click += BtnRemove_Click;
            this.Controls.Add(btnRemove);

            btnClose = new Button
            {
                Text = "Close",
                DialogResult = DialogResult.Cancel,
                Location = new System.Drawing.Point(btnX, 330),
                Width = 90,
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            this.Controls.Add(btnClose);

            this.CancelButton = btnClose;
        }

        private void LoadLinks()
        {
            lvLinks.Items.Clear();

            int chartCount = 0, rangeCount = 0, textCount = 0;

            try
            {
                // Load shape links
                foreach (Word.Shape shape in _doc.Shapes)
                {
                    if (ChartLinkUtility.IsChartLink(shape))
                    {
                        var info = ChartLinkUtility.GetChartLinkInfo(shape);
                        var item = new ListViewItem("Chart");
                        item.SubItems.Add(CloudPathHelper.GetFileName(info.WorkbookPath ?? ""));
                        item.SubItems.Add(info.SheetName ?? "");
                        item.SubItems.Add(info.ChartName ?? "");
                        item.SubItems.Add("OK");
                        item.Tag = shape;
                        lvLinks.Items.Add(item);
                        chartCount++;
                    }
                    else if (RangeLinkUtility.IsRangeLink(shape))
                    {
                        var info = RangeLinkUtility.GetRangeLinkInfo(shape);
                        var item = new ListViewItem("Range");
                        item.SubItems.Add(CloudPathHelper.GetFileName(info.WorkbookPath ?? ""));
                        item.SubItems.Add(info.SheetName ?? "");
                        item.SubItems.Add(info.RangeAddress ?? "");
                        item.SubItems.Add("OK");
                        item.Tag = shape;
                        lvLinks.Items.Add(item);
                        rangeCount++;
                    }
                }

                // Load text links
                foreach (Word.ContentControl cc in _doc.ContentControls)
                {
                    if (TextLinkUtility.IsTextLink(cc))
                    {
                        var info = TextLinkUtility.GetTextLinkInfo(cc);
                        var item = new ListViewItem("Text");
                        item.SubItems.Add(CloudPathHelper.GetFileName(info.WorkbookPath ?? ""));
                        item.SubItems.Add(info.SheetName ?? "");
                        item.SubItems.Add(info.CellAddress ?? "");
                        item.SubItems.Add("OK");
                        item.Tag = cc;
                        lvLinks.Items.Add(item);
                        textCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading links: {ex.Message}");
            }

            lblSummary.Text = $"Links: {chartCount} charts, {rangeCount} ranges, {textCount} text";
        }

        private void BtnRefreshAll_Click(object sender, EventArgs e)
        {
            UpdateLinks.UpdateAllLinks(_doc, true);
            LoadLinks();
        }

        private void BtnRefreshSelected_Click(object sender, EventArgs e)
        {
            if (lvLinks.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select one or more links to refresh.", "Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            foreach (ListViewItem item in lvLinks.SelectedItems)
            {
                if (item.Tag is Word.Shape shape)
                {
                    UpdateLinks.UpdateSingleLink(shape);
                }
                else if (item.Tag is Word.ContentControl cc)
                {
                    // For text links, do a quick refresh
                    Excel.Application excelApp = null;
                    try
                    {
                        excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                        using (var wbCache = new WorkbookCache(excelApp))
                        {
                            TextLinkUtility.RefreshTextLink(cc, excelApp, wbCache, null);
                        }
                    }
                    catch { }
                }
            }

            LoadLinks();
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (lvLinks.SelectedItems.Count == 0)
            {
                MessageBox.Show("Select one or more links to remove.", "Remove", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var result = MessageBox.Show($"Remove {lvLinks.SelectedItems.Count} selected link(s)?",
                "Confirm Remove", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result != DialogResult.Yes)
                return;

            foreach (ListViewItem item in lvLinks.SelectedItems)
            {
                try
                {
                    if (item.Tag is Word.Shape shape)
                    {
                        shape.Delete();
                    }
                    else if (item.Tag is Word.ContentControl cc)
                    {
                        cc.Delete(false); // Keep contents
                    }
                }
                catch { }
            }

            LoadLinks();
        }
    }
}