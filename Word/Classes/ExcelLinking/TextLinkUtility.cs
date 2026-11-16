// TextLinkUtility.cs (.NET Framework 4.8 / C# 7.3)
// Text linking utilities for Word - updates Word content controls with Excel cell values.
// Word uses ContentControls instead of hyperlinks (PowerPoint limitation).

using System;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// Provides text linking functionality between Excel cells and Word content controls.
    /// Unlike PowerPoint (which uses hyperlinks), Word uses ContentControls for text links.
    /// </summary>
    internal static class TextLinkUtility
    {
        private const string TagPrefix = "Pragmatics_TextLink_";
        private const string TagWorkbookPath = "WbPath";
        private const string TagSheetName = "SheetName";
        private const string ContentControlTitle = "Pragmatics_TextLink";

        #region Public Import Methods (Ribbon Entry Points)

        /// <summary>
        /// Imports an Excel cell value as linked text.
        /// Shows a dialog to select the cell from an open Excel workbook.
        /// </summary>
        public static Word.ContentControl ImportExcelText(Word.Application wordApp)
        {
            if (wordApp == null) return null;

            Word.Document doc;
            try
            {
                doc = wordApp.ActiveDocument;
            }
            catch
            {
                MessageBox.Show("No active document.", "Import Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (doc == null)
            {
                MessageBox.Show("No active document.", "Import Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            Excel.Application excelApp;
            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch
            {
                MessageBox.Show("Excel is not running. Please open an Excel workbook first.",
                    "Import Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            if (excelApp.Workbooks.Count == 0)
            {
                MessageBox.Show("No Excel workbooks are open.",
                    "Import Text", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }

            using (var dialog = new TextLinkSelectionDialog(excelApp))
            {
                if (dialog.ShowDialog() != DialogResult.OK)
                    return null;

                string wbPath = dialog.SelectedWorkbookPath;
                string sheetName = dialog.SelectedSheetName;
                string cellAddress = dialog.SelectedCellAddress;
                string formatString = dialog.SelectedFormatString;

                if (string.IsNullOrEmpty(wbPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(cellAddress))
                    return null;

                Word.Range insertRange = wordApp.Selection.Range;
                return CreateTextLink(doc, excelApp, wbPath, sheetName, cellAddress, insertRange, formatString);
            }
        }

        #endregion

        /// <summary>
        /// Creates a text link by inserting a content control linked to an Excel cell.
        /// </summary>
        public static Word.ContentControl CreateTextLink(
            Word.Document doc,
            Excel.Application excelApp,
            string workbookPath,
            string sheetName,
            string cellAddress,
            Word.Range insertRange,
            string formatString = null)
        {
            if (doc == null || excelApp == null || insertRange == null)
                return null;

            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(sheetName) || string.IsNullOrEmpty(cellAddress))
                return null;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range cell = null;

            using (var perfScope = new ExcelPerformanceScope(excelApp))
            using (var wbCache = new WorkbookCache(excelApp))
            {
                try
                {
                    // Normalize path for cross-user compatibility
                    string normalizedPath = CloudPathHelper.NormalizePath(workbookPath);

                    // Get workbook
                    wb = wbCache.GetOrOpen(workbookPath, out bool newlyOpened, out string actualPath);
                    if (wb == null)
                    {
                        MessageBox.Show($"Could not open workbook:\n{workbookPath}",
                            "Text Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    // Find worksheet
                    try
                    {
                        ws = wb.Worksheets[sheetName] as Excel.Worksheet;
                    }
                    catch
                    {
                        MessageBox.Show($"Could not find worksheet '{sheetName}' in workbook.",
                            "Text Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    if (ws == null)
                        return null;

                    // Get cell
                    try
                    {
                        cell = ws.Range[cellAddress];
                    }
                    catch
                    {
                        MessageBox.Show($"Invalid cell address: '{cellAddress}'",
                            "Text Link Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return null;
                    }

                    if (cell == null)
                        return null;

                    // Get cell value
                    string cellValue = GetFormattedCellValue(cell, formatString);

                    // Create content control
                    Word.ContentControl cc = doc.ContentControls.Add(
                        Word.WdContentControlType.wdContentControlText,
                        insertRange);

                    cc.Title = ContentControlTitle;
                    cc.Tag = BuildTag(normalizedPath, sheetName, cellAddress, formatString);
                    cc.Range.Text = cellValue;

                    // Optional: style the content control
                    cc.LockContentControl = false; // Allow editing but we'll update it
                    cc.LockContents = false;

                    return cc;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating text link: {ex.Message}");
                    return null;
                }
                finally
                {
                    if (cell != null)
                    {
                        try { Marshal.ReleaseComObject(cell); } catch { }
                    }
                    if (ws != null)
                    {
                        try { Marshal.ReleaseComObject(ws); } catch { }
                    }
                }
            }
        }

        /// <summary>
        /// Refreshes a text link content control with the current Excel cell value.
        /// </summary>
        public static bool RefreshTextLink(
            Word.ContentControl cc,
            Excel.Application excelApp,
            WorkbookCache wbCache,
            WarningCollector warnings)
        {
            if (cc == null || excelApp == null || wbCache == null)
                return false;

            if (!IsTextLink(cc))
                return false;

            // Parse tag
            var linkInfo = ParseTag(cc.Tag);
            if (linkInfo.WorkbookPath == null)
            {
                warnings?.AddWarning("Text link has invalid tag format");
                return false;
            }

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Range cell = null;

            try
            {
                // Get workbook
                wb = wbCache.GetOrOpen(linkInfo.WorkbookPath, out bool newlyOpened, out string actualPath);
                if (wb == null)
                {
                    warnings?.AddWarning($"Could not open workbook: {CloudPathHelper.GetFileName(linkInfo.WorkbookPath)}");
                    return false;
                }

                // Find worksheet
                try
                {
                    ws = wb.Worksheets[linkInfo.SheetName] as Excel.Worksheet;
                }
                catch
                {
                    warnings?.AddWarning($"Worksheet '{linkInfo.SheetName}' not found");
                    return false;
                }

                if (ws == null)
                    return false;

                // Get cell
                try
                {
                    cell = ws.Range[linkInfo.CellAddress];
                }
                catch
                {
                    warnings?.AddWarning($"Invalid cell: '{linkInfo.CellAddress}'");
                    return false;
                }

                if (cell == null)
                    return false;

                // Get formatted value
                string newValue = GetFormattedCellValue(cell, linkInfo.FormatString);

                // Update content control
                cc.Range.Text = newValue;

                // Update tag with normalized path if it changed
                string normalizedPath = CloudPathHelper.NormalizePath(actualPath);
                if (!string.Equals(normalizedPath, linkInfo.WorkbookPath, StringComparison.OrdinalIgnoreCase))
                {
                    cc.Tag = BuildTag(normalizedPath, linkInfo.SheetName, linkInfo.CellAddress, linkInfo.FormatString);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error refreshing text link: {ex.Message}");
                warnings?.AddWarning($"Error refreshing text link: {ex.Message}");
                return false;
            }
            finally
            {
                if (cell != null)
                {
                    try { Marshal.ReleaseComObject(cell); } catch { }
                }
                if (ws != null)
                {
                    try { Marshal.ReleaseComObject(ws); } catch { }
                }
            }
        }

        /// <summary>
        /// Checks if a content control is a text link.
        /// </summary>
        public static bool IsTextLink(Word.ContentControl cc)
        {
            if (cc == null) return false;

            return string.Equals(cc.Title, ContentControlTitle, StringComparison.OrdinalIgnoreCase) ||
                   (cc.Tag?.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase) ?? false);
        }

        /// <summary>
        /// Gets text link information from a content control.
        /// </summary>
        public static (string WorkbookPath, string SheetName, string CellAddress, string FormatString) GetTextLinkInfo(
            Word.ContentControl cc)
        {
            if (cc == null || !IsTextLink(cc))
                return (null, null, null, null);

            return ParseTag(cc.Tag);
        }

        #region Private Methods

        private static string BuildTag(string wbPath, string sheetName, string cellAddr, string format)
        {
            // Format: Pragmatics_TextLink_|WbPath|SheetName|CellAddr|Format
            string formatPart = format ?? "";
            return $"{TagPrefix}|{wbPath}|{sheetName}|{cellAddr}|{formatPart}";
        }

        private static (string WorkbookPath, string SheetName, string CellAddress, string FormatString) ParseTag(string tag)
        {
            if (string.IsNullOrEmpty(tag))
                return (null, null, null, null);

            if (!tag.StartsWith(TagPrefix, StringComparison.OrdinalIgnoreCase))
                return (null, null, null, null);

            string data = tag.Substring(TagPrefix.Length);
            string[] parts = data.Split('|');

            if (parts.Length < 4)
                return (null, null, null, null);

            return (
                parts[0],                                    // WorkbookPath
                parts[1],                                    // SheetName
                parts[2],                                    // CellAddress
                parts.Length > 3 ? parts[3] : null          // FormatString
            );
        }

        private static string GetFormattedCellValue(Excel.Range cell, string formatString)
        {
            if (cell == null)
                return string.Empty;

            try
            {
                object value = cell.Value2;

                if (value == null)
                    return string.Empty;

                // If no custom format, use Excel's displayed text
                if (string.IsNullOrEmpty(formatString))
                {
                    return cell.Text?.ToString() ?? value.ToString();
                }

                // Apply custom format
                if (value is double numValue)
                {
                    return numValue.ToString(formatString, CultureInfo.CurrentCulture);
                }
                else if (value is DateTime dateValue)
                {
                    return dateValue.ToString(formatString, CultureInfo.CurrentCulture);
                }
                else
                {
                    return value.ToString();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error formatting cell value: {ex.Message}");

                try
                {
                    return cell.Text?.ToString() ?? cell.Value2?.ToString() ?? string.Empty;
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        #endregion
    }

    /// <summary>
    /// Dialog for selecting an Excel cell for text linking.
    /// </summary>
    internal class TextLinkSelectionDialog : Form
    {
        private ComboBox cboWorkbook;
        private ComboBox cboSheet;
        private TextBox txtCell;
        private TextBox txtFormat;
        private Button btnOK;
        private Button btnCancel;
        private Button btnSelectFromExcel;
        private readonly Excel.Application _excelApp;

        public string SelectedWorkbookPath { get; private set; }
        public string SelectedSheetName { get; private set; }
        public string SelectedCellAddress { get; private set; }
        public string SelectedFormatString { get; private set; }

        public TextLinkSelectionDialog(Excel.Application excelApp)
        {
            _excelApp = excelApp;
            InitializeComponent();
            LoadWorkbooks();
        }

        private void InitializeComponent()
        {
            this.Text = "Select Excel Cell for Text Link";
            this.Size = new System.Drawing.Size(420, 260);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterParent;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            var lblWorkbook = new Label { Text = "Workbook:", Location = new System.Drawing.Point(20, 20), AutoSize = true };
            cboWorkbook = new ComboBox { Location = new System.Drawing.Point(100, 17), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };
            cboWorkbook.SelectedIndexChanged += CboWorkbook_SelectedIndexChanged;

            var lblSheet = new Label { Text = "Worksheet:", Location = new System.Drawing.Point(20, 55), AutoSize = true };
            cboSheet = new ComboBox { Location = new System.Drawing.Point(100, 52), Width = 280, DropDownStyle = ComboBoxStyle.DropDownList };

            var lblCell = new Label { Text = "Cell:", Location = new System.Drawing.Point(20, 90), AutoSize = true };
            txtCell = new TextBox { Location = new System.Drawing.Point(100, 87), Width = 100 };

            btnSelectFromExcel = new Button { Text = "Select", Location = new System.Drawing.Point(210, 85), Width = 70 };
            btnSelectFromExcel.Click += BtnSelectFromExcel_Click;

            var lblFormat = new Label { Text = "Format:", Location = new System.Drawing.Point(20, 125), AutoSize = true };
            txtFormat = new TextBox { Location = new System.Drawing.Point(100, 122), Width = 180 };
            var lblFormatHint = new Label { Text = "(e.g., #,##0.00 or yyyy-MM-dd)", Location = new System.Drawing.Point(100, 145), AutoSize = true, ForeColor = System.Drawing.SystemColors.GrayText };

            btnOK = new Button { Text = "OK", DialogResult = DialogResult.OK, Location = new System.Drawing.Point(220, 180), Width = 75 };
            btnOK.Click += BtnOK_Click;

            btnCancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new System.Drawing.Point(305, 180), Width = 75 };

            this.Controls.AddRange(new Control[] { lblWorkbook, cboWorkbook, lblSheet, cboSheet, lblCell, txtCell, btnSelectFromExcel, lblFormat, txtFormat, lblFormatHint, btnOK, btnCancel });
            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;
        }

        private void LoadWorkbooks()
        {
            cboWorkbook.Items.Clear();
            foreach (Excel.Workbook wb in _excelApp.Workbooks)
            {
                cboWorkbook.Items.Add(new WorkbookItem { Name = wb.Name, Path = wb.FullName });
            }
            if (cboWorkbook.Items.Count > 0)
                cboWorkbook.SelectedIndex = 0;
        }

        private void CboWorkbook_SelectedIndexChanged(object sender, EventArgs e)
        {
            cboSheet.Items.Clear();

            if (cboWorkbook.SelectedItem is WorkbookItem wbItem)
            {
                try
                {
                    Excel.Workbook wb = _excelApp.Workbooks[wbItem.Name];
                    foreach (Excel.Worksheet ws in wb.Worksheets)
                    {
                        cboSheet.Items.Add(ws.Name);
                    }
                    if (cboSheet.Items.Count > 0)
                        cboSheet.SelectedIndex = 0;
                }
                catch { }
            }
        }

        private void BtnSelectFromExcel_Click(object sender, EventArgs e)
        {
            try
            {
                this.WindowState = FormWindowState.Minimized;

                MessageBox.Show("Select a single cell in Excel and click OK.",
                    "Select Cell", MessageBoxButtons.OK, MessageBoxIcon.Information);

                Excel.Range selection = _excelApp.Selection as Excel.Range;
                if (selection != null)
                {
                    // Use just the first cell if multiple selected
                    Excel.Range cell = selection.Cells[1, 1] as Excel.Range;
                    Excel.Worksheet ws = cell.Worksheet;
                    Excel.Workbook wb = ws.Parent as Excel.Workbook;

                    for (int i = 0; i < cboWorkbook.Items.Count; i++)
                    {
                        if (cboWorkbook.Items[i] is WorkbookItem wbItem && wbItem.Name == wb.Name)
                        {
                            cboWorkbook.SelectedIndex = i;
                            break;
                        }
                    }

                    for (int i = 0; i < cboSheet.Items.Count; i++)
                    {
                        if (cboSheet.Items[i].ToString() == ws.Name)
                        {
                            cboSheet.SelectedIndex = i;
                            break;
                        }
                    }

                    txtCell.Text = cell.Address[false, false];
                }

                this.WindowState = FormWindowState.Normal;
                this.BringToFront();
            }
            catch (Exception ex)
            {
                this.WindowState = FormWindowState.Normal;
                MessageBox.Show($"Error getting selection: {ex.Message}",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (cboWorkbook.SelectedItem is WorkbookItem wbItem)
            {
                SelectedWorkbookPath = wbItem.Path;
                SelectedSheetName = cboSheet.SelectedItem as string;
                SelectedCellAddress = txtCell.Text.Trim();
                SelectedFormatString = string.IsNullOrWhiteSpace(txtFormat.Text) ? null : txtFormat.Text.Trim();
            }
        }

        private class WorkbookItem
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public override string ToString() => Name;
        }
    }
}