// TextLinkUtility.cs (.NET Framework 4.8 / C# 7.3)

using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OfficePragmatics_Powerpoint;

namespace OfficePragmatics_PowerPoint.ExcelLinking
{
    public static class TextLinkUtility
    {
        /* ═════════════════════  IMPORT / CREATE TEXT LINK  ═════════════════════ */

        /// <summary>
        /// Alias for CreateTextLink - used by ribbon controller.
        /// </summary>
        public static void ImportExcelText(PowerPoint.Application ppApp)
        {
            CreateTextLink(ppApp);
        }

        public static void CreateTextLink(PowerPoint.Application ppApp)
        {
            Excel.Application excelApp = null;
            Excel.Range selectedCell = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            PowerPoint.Selection ppSel = null;
            PowerPoint.TextRange selectedTextRange = null;

            try
            {
                // Get Excel application
                try
                {
                    excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    MessageBox.Show("Excel is not running. Please open Excel and select a cell.",
                                    "Excel Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Get selected cell in Excel
                selectedCell = excelApp.Selection as Excel.Range;
                if (selectedCell == null)
                {
                    MessageBox.Show("Please select a single cell in Excel.",
                                    "No Cell Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Ensure only one cell is selected
                if (selectedCell.Cells.Count > 1)
                {
                    MessageBox.Show("Please select only a single cell for text linking.",
                                    "Multiple Cells Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Get workbook and worksheet
                ws = selectedCell.Worksheet;
                wb = ws.Parent as Excel.Workbook;

                if (wb == null || string.IsNullOrEmpty(wb.FullName))
                {
                    MessageBox.Show("Please save the Excel workbook before creating a text link.",
                                    "Workbook Not Saved", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Get PowerPoint selection
                ppSel = ppApp.ActiveWindow.Selection;
                if (ppSel.Type != PowerPoint.PpSelectionType.ppSelectionText)
                {
                    MessageBox.Show("Please select some text in PowerPoint to link.",
                                    "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selectedTextRange = ppSel.TextRange;
                if (selectedTextRange == null || selectedTextRange.Length == 0)
                {
                    MessageBox.Show("Please select some text in PowerPoint to link.",
                                    "No Text Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Use performance scope for Excel operations
                using (new ExcelPerformanceScope(excelApp))
                {
                    // Create a unique named range for this cell
                    string uniqueId = Guid.NewGuid().ToString().Replace("-", "_");
                    string namedRangeName = $"Pragmatics_TextLink_{uniqueId}";

                    // Add the named range
                    wb.Names.Add(Name: namedRangeName, RefersTo: selectedCell, Visible: false);

                    // Get cell value
                    string cellValue = selectedCell.Text?.ToString() ?? "";
                    if (string.IsNullOrEmpty(cellValue))
                    {
                        cellValue = " "; // Use space for empty cells to maintain the link
                    }

                    // Normalize path for cross-user compatibility
                    string normalizedWbPath = CloudPathHelper.NormalizePath(wb.FullName);

                    // Build the hyperlink URI
                    string hyperlinkAddress =
                        "pragmatics-textlink://link" +
                        "?wb=" + Uri.EscapeDataString(normalizedWbPath) +
                        "&ws=" + Uri.EscapeDataString(ws.Name) +
                        "&nr=" + Uri.EscapeDataString(namedRangeName);

                    // Cache font properties before modifying text
                    PowerPoint.Font font = selectedTextRange.Font;
                    string fontName = font.Name;
                    float fontSize = font.Size;
                    Office.MsoTriState fontBold = font.Bold;
                    Office.MsoTriState fontItalic = font.Italic;
                    Office.MsoTriState fontUnderline = font.Underline;
                    int fontColor = font.Color.RGB;
                    Marshal.ReleaseComObject(font);

                    // Get the parent shape for later use
                    PowerPoint.Shape parentShape = null;
                    try
                    {
                        parentShape = selectedTextRange.Parent.Parent as PowerPoint.Shape;
                    }
                    catch { }

                    // Replace the selected text with the cell value
                    int startPos = selectedTextRange.Start;
                    selectedTextRange.Text = cellValue;

                    // Get the new text range after replacement
                    PowerPoint.TextRange newTextRange = null;
                    if (parentShape != null)
                    {
                        newTextRange = parentShape.TextFrame.TextRange.Characters(startPos, cellValue.Length);
                    }
                    else
                    {
                        newTextRange = selectedTextRange;
                    }

                    // Restore font properties
                    PowerPoint.Font newFont = newTextRange.Font;
                    newFont.Name = fontName;
                    newFont.Size = fontSize;
                    newFont.Bold = fontBold;
                    newFont.Italic = fontItalic;
                    newFont.Underline = fontUnderline;
                    newFont.Color.RGB = fontColor;
                    Marshal.ReleaseComObject(newFont);

                    // Add the hyperlink (this creates the link)
                    PowerPoint.ActionSetting actionSetting =
                        newTextRange.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                    actionSetting.Action = PowerPoint.PpActionType.ppActionHyperlink;

                    PowerPoint.Hyperlink hyperlink = actionSetting.Hyperlink;
                    hyperlink.Address = hyperlinkAddress;
                    hyperlink.SubAddress = "";
                    hyperlink.ScreenTip = $"Linked to: {ws.Name}!{selectedCell.Address}";

                    // Style the hyperlink to match original text (hide underline if needed)
                    try
                    {
                        if (parentShape != null && parentShape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            var tf2 = parentShape.TextFrame2;
                            var tr2 = tf2.TextRange;
                            var linkedRange2 = tr2.Characters[startPos, cellValue.Length];
                            var font2 = linkedRange2.Font;

                            // Set text color
                            font2.Fill.ForeColor.RGB = fontColor;

                            // Handle underline - if original had no underline, make it invisible
                            if (fontUnderline == Office.MsoTriState.msoFalse)
                            {
                                // Get background color and set underline to match
                                int bgColor = GetBackgroundColor(parentShape);
                                font2.UnderlineStyle = Office.MsoTextUnderlineType.msoUnderlineSingleLine;
                                font2.UnderlineColor.RGB = bgColor;
                            }

                            Marshal.ReleaseComObject(font2);
                            Marshal.ReleaseComObject(linkedRange2);
                            Marshal.ReleaseComObject(tr2);
                            Marshal.ReleaseComObject(tf2);
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"TextFrame2 styling failed: {ex.Message}");
                    }

                    Marshal.ReleaseComObject(hyperlink);
                    Marshal.ReleaseComObject(actionSetting);

                    if (newTextRange != selectedTextRange)
                    {
                        Marshal.ReleaseComObject(newTextRange);
                    }

                    if (parentShape != null)
                    {
                        Marshal.ReleaseComObject(parentShape);
                    }

                    System.Diagnostics.Debug.WriteLine($"Created text link: {namedRangeName} -> {cellValue}");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error creating text link: {ex.Message}",
                                "Text Link Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (selectedTextRange != null) Marshal.ReleaseComObject(selectedTextRange);
                if (ppSel != null) Marshal.ReleaseComObject(ppSel);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (selectedCell != null) Marshal.ReleaseComObject(selectedCell);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /* ═════════════════════  HELPER: GET BACKGROUND COLOR  ═════════════════════ */
        private static int GetBackgroundColor(PowerPoint.Shape shape)
        {
            int defaultBg = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            if (shape == null) return defaultBg;

            try
            {
                // Check shape's own fill
                if (shape.Fill.Visible == Office.MsoTriState.msoTrue &&
                    shape.Fill.Type == Office.MsoFillType.msoFillSolid)
                {
                    return shape.Fill.ForeColor.RGB;
                }

                // Navigate up to find slide/layout/master background
                object parent = shape.Parent;

                while (parent != null)
                {
                    if (parent is PowerPoint.Slide slide)
                    {
                        if (slide.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                            slide.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                            slide.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                        {
                            return slide.Background.Fill.ForeColor.RGB;
                        }

                        // Check layout
                        PowerPoint.CustomLayout layout = slide.CustomLayout;
                        if (layout != null)
                        {
                            if (layout.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                                layout.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                                layout.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                            {
                                int color = layout.Background.Fill.ForeColor.RGB;
                                Marshal.ReleaseComObject(layout);
                                return color;
                            }

                            // Check master
                            if (layout.Parent is PowerPoint.Master master)
                            {
                                if (master.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                                    master.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                                {
                                    int color = master.Background.Fill.ForeColor.RGB;
                                    Marshal.ReleaseComObject(layout);
                                    Marshal.ReleaseComObject(master);
                                    return color;
                                }
                                Marshal.ReleaseComObject(master);
                            }
                            Marshal.ReleaseComObject(layout);
                        }
                        break;
                    }

                    if (parent is PowerPoint.CustomLayout cl)
                    {
                        if (cl.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                            cl.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                            cl.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                        {
                            return cl.Background.Fill.ForeColor.RGB;
                        }
                        break;
                    }

                    if (parent is PowerPoint.Master m)
                    {
                        if (m.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                            m.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                        {
                            return m.Background.Fill.ForeColor.RGB;
                        }
                        break;
                    }

                    try
                    {
                        dynamic dyn = parent;
                        parent = dyn.Parent;
                    }
                    catch
                    {
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting background color: {ex.Message}");
            }

            return defaultBg;
        }
    }
}