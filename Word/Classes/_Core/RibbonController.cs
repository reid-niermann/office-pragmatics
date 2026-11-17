// RibbonController.cs
using Microsoft.Office.Core;
using OfficePragmatics_Word;
using OfficePragmatics_Word.Alignment;
using OfficePragmatics_Word.Page;
using OfficePragmatics_Word.Settings;
using OfficePragmatics_Word.Shape;
using OfficePragmatics_Word.Size;
using OfficePragmatics_Word.Text;
using OfficePragmatics_Word.Utilities;
using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using OfficePragmatics_Word.ExcelLinking;

namespace Office_Pragmatics_Word
{
    public class RibbonController
    {
        // ══════════════════════════════════════════════════════════════════
        // Callback for custom images
        // ══════════════════════════════════════════════════════════════════
        public Bitmap GetCustomImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Alignment
                case "btnAlignLeft": return Properties.Resources.btnAlignLeft;
                case "btnAlignTop": return Properties.Resources.btnAlignTop;
                case "btnAlignHorizontal": return Properties.Resources.btnAlignHorizontal;
                case "btnAlignRight": return Properties.Resources.btnAlignRight;
                case "btnAlignBottom": return Properties.Resources.btnAlignBottom;
                case "btnAlignVertical": return Properties.Resources.btnAlignVertical;
                case "btnAlignCenter": return Properties.Resources.btnAlignCenter;
                case "btnAlignMiddle": return Properties.Resources.btnAlignMiddle;
                case "btnAlignTogether": return Properties.Resources.btnAlignTogether;

                // Size and Position
                case "btnSameHW": return Properties.Resources.btnSameHW;
                case "btnSameHeight": return Properties.Resources.btnSameHeight;
                case "btnSameWidth": return Properties.Resources.btnSameWidth;
                case "btnExtendTop": return Properties.Resources.btnExtendTop;
                case "btnExtendBottom": return Properties.Resources.btnExtendBottom;
                case "btnExtendLeft": return Properties.Resources.btnExtendLeft;
                case "btnExtendRight": return Properties.Resources.btnExtendRight;
                case "btnFitHorz": return Properties.Resources.btnFitHorz;
                case "btnFitVert": return Properties.Resources.btnFitVert;
                case "b_RatioResize": return Properties.Resources.b_RatioResize;
                case "b_Rotation": return Properties.Resources.b_Rotation;
                case "btnCycleSwap": return Properties.Resources.btnCycleSwap;
                case "btnAutoGrid": return Properties.Resources.btnAutoGrid;
                case "btnCellCenter": return Properties.Resources.btnCellCenter;

                // Library
                case "btnLibrary": return Properties.Resources.btnLibrary;

                // Doc Tools
                case "btnFormEditor": return Properties.Resources.btnFormEditor;
                case "btnDocCompare": return Properties.Resources.btnDocCompare;
                case "btnSmartEyeDrop": return Properties.Resources.btnSmartEyeDrop;
                case "btnNumberToWords": return Properties.Resources.btnNumberToWords;
                case "btnFlattenColors": return Properties.Resources.btnFlattenColors;
                case "btnLineUp": return Properties.Resources.btnLineUp;
                case "btnLineDown": return Properties.Resources.btnLineDown;

                // XLS Linking
                case "galImportChart": return Properties.Resources.galImportChart;
                case "b_vectorChartImport": return Properties.Resources.b_vectorChartImport;
                case "b_pictureChartImport": return Properties.Resources.b_pictureChartImport;
                case "b_chartChartImport": return Properties.Resources.b_chartChartImport;
                case "galImportTable": return Properties.Resources.galImportTable;
                case "b_pictureRangeImport": return Properties.Resources.b_pictureChartImport;
                case "b_tableRangeImport": return Properties.Resources.b_chartChartImport;
                case "galImportText": return Properties.Resources.galImportText;
                case "btnImportText": return Properties.Resources.galImportText;
                case "btnRefreshLinks": return Properties.Resources.btnRefreshLinks;
                case "btnLinkManager": return Properties.Resources.btnLinkManager;

                // Page
                case "btnPageHop": return Properties.Resources.btnSlideHop;

                // File
                case "btnVersionUp": return Properties.Resources.btnVersionUp;
                case "btnSaveAsPDF": return Properties.Resources.btnSaveAsPDF;
                case "btnSelectExport": return Properties.Resources.btnSelectExport;
                case "btnSettings": return Properties.Resources.btnSettings;
                case "galFileTools": return Properties.Resources.galFileTools;
                case "b_Show_Explorer": return Properties.Resources.b_Show_Explorer;
                case "b_Copy_Path": return Properties.Resources.b_Copy_Path;

                default: return Properties.Resources.TBU;

            }
        }

        // ══════════════════════════════════════════════════════════════════
        // Callback for labels
        // ══════════════════════════════════════════════════════════════════
        public string GetLabel(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Tab
                case "tabPragmatics": return "Pragmatics";
                // Alignment
                case "btnAlignLeft": return "Align Left";
                case "btnAlignTop": return "Align Top";
                case "btnAlignHorizontal": return "Align Horizontal";
                case "btnAlignRight": return "Align Right";
                case "btnAlignBottom": return "Align Bottom";
                case "btnAlignVertical": return "Align Vertical";
                case "btnAlignCenter": return "Align Center";
                case "btnAlignMiddle": return "Align Middle";
                case "btnAlignTogether": return "Align Together";
                // Size and Position
                case "btnSameHW": return "Same H-W";
                case "btnSameHeight": return "Same Height";
                case "btnSameWidth": return "Same Width";
                case "btnExtendTop": return "Extend Top";
                case "btnExtendBottom": return "Extend Bottom";
                case "btnExtendLeft": return "Extend Left";
                case "btnExtendRight": return "Extend Right";
                case "btnFitHorz": return "Fit Horizontal";
                case "btnFitVert": return "Fit Vertical";
                case "b_RatioResize": return "Ratio Resize";
                case "b_Rotation": return "Rotation";
                case "btnCycleSwap": return "Cycle Swap";
                case "btnAutoGrid": return "Auto Grid";
                case "btnCellCenter": return "Cell Center";
                // Library
                case "btnLibrary": return "Library Assets";
                // Doc Tools
                case "btnFormEditor": return "Form Fill";
                case "btnDocCompare": return "Redline Compare";
                case "btnSmartEyeDrop": return "Smart Eye Drop";
                case "btnNumberToWords": return "Number to Words";
                case "btnFlattenColors": return "Flatten Colors";
                case "btnLineUp": return "+ Line Spacing";
                case "btnLineDown": return "- Line Spacing";
                // XLS Linking
                case "galImportChart": return "Import Chart";
                case "b_vectorChartImport": return "Vector Image";
                case "b_pictureChartImport": return "PNG Image";
                case "b_chartChartImport": return "Chart Object";
                case "galImportTable": return "Import Table";
                case "b_pictureRangeImport": return "PNG Image";
                case "b_tableRangeImport": return "Native Table";
                case "galImportText": return "Import Text";
                case "btnImportText": return "Native Text";
                case "btnLinkManager": return "Link Manager";
                case "btnRefreshLinks": return "Refresh Links";
                // Page
                case "btnPageHop": return "Go to Page";
                // File
                case "btnVersionUp": return "Version Up";
                case "btnSaveAsPDF": return "Save PDF";
                case "btnSelectExport": return "Export Selection";
                case "galFileTools": return "File Tools";
                case "b_Show_Explorer": return "Show in Explorer";
                case "b_Copy_Path": return "Copy File Path";
                case "btnSettings": return "Settings";
                default: return string.Empty;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // Callback for keytips (keyboard accelerators)
        // ══════════════════════════════════════════════════════════════════
        public string GetKeytip(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Tab
                case "tabPragmatics": return "X";

                // Alignment
                case "btnAlignLeft": return "L";
                case "btnAlignTop": return "T";
                case "btnAlignHorizontal": return "H";
                case "btnAlignRight": return "R";
                case "btnAlignBottom": return "B";
                case "btnAlignVertical": return "V";
                case "btnAlignCenter": return "C";
                case "btnAlignMiddle": return "M";
                case "btnAlignTogether": return "N";

                // Size and Position
                case "btnSameHW": return "A";
                case "btnSameHeight": return "S";
                case "btnSameWidth": return "D";
                case "btnExtendTop": return "EU";
                case "btnExtendBottom": return "ED";
                case "btnExtendLeft": return "EL";
                case "btnExtendRight": return "ER";
                case "btnFitHorz": return "PV";
                case "btnFitVert": return "PH";
                case "b_RatioResize": return "PS";
                case "b_Rotation": return "PR";
                case "btnCycleSwap": return "1";
                case "btnAutoGrid": return "2";
                case "btnCellCenter": return "3";

                // Library
                case "btnLibrary": return "0";

                // Doc Tools
                case "btnFormEditor": return "F";
                case "btnDocCompare": return "E";
                case "btnSmartEyeDrop": return "I";
                case "btnNumberToWords": return "W";
                case "btnFlattenColors": return "KF";
                case "btnLineUp": return ",";
                case "btnLineDown": return ".";

                // XLS Linking
                case "galImportChart": return "QC";
                case "galImportTable": return "QR";
                case "galImportText": return "QT";
                case "btnImportText": return "YVT";
                case "btnLinkManager": return "QM";
                case "btnRefreshLinks": return "QU";

                // Page
                case "btnPageHop": return "X";

                // File
                case "btnVersionUp": return "[";
                case "btnSaveAsPDF": return "]";
                case "btnSelectExport": return "\\";
                case "galFileTools": return "=";
                case "btnSettings": return "'";
                default: return string.Empty;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // Callbacks for screentips and supertips
        // ══════════════════════════════════════════════════════════════════
        public string GetScreentip(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Only define screentips that differ from the label
                // case "": return "";

                default: return GetLabel(control);
            }
        }

        public string GetSupertip(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Alignment
                case "btnAlignLeft": return "Aligns shapes to the left. Single shape selections will align to page edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignTop": return "Aligns shapes to the top. Single shape selections will align to page edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignHorizontal": return "Evenly distributes selected shapes horizontally with equal spacing between them. Requires at least three shapes.";
                case "btnAlignRight": return "Aligns shapes to the right. Single shape selections will align to page edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignBottom": return "Aligns shapes to the bottom. Single shape selections will align to page edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignVertical": return "Evenly distributes selected shapes vertically with equal spacing between them. Requires at least three shapes.";
                case "btnAlignCenter": return "Centers shapes horizontally. Single shape selections will center on page; multiple selected shapes will align to the last selected shape's horizontal center.";
                case "btnAlignMiddle": return "Centers shapes vertically. Single shape selections will center on page; multiple selected shapes will align to the last selected shape's vertical middle.";
                case "btnAlignTogether": return "Snaps two shapes together edge-to-edge. Detects the relative position of the shapes and moves the first shape to touch the second. Requires exactly two shapes.";

                // Size and Position
                case "btnSameHW": return "Resizes all selected shapes to match the last selected shape's width and height. Accounts for shape borders when calculating dimensions.";
                case "btnSameHeight": return "Resizes all selected shapes to match the last selected shape's height. Accounts for shape borders when calculating dimensions.";
                case "btnSameWidth": return "Resizes all selected shapes to match the last selected shape's width. Accounts for shape borders when calculating dimensions.";
                case "btnExtendTop": return "Stretches shapes upward. Single shape extends to page top; multiple shapes extend to the last selected shape's top edge.";
                case "btnExtendBottom": return "Stretches shapes downward. Single shape extends to page bottom; multiple shapes extend to the last selected shape's bottom edge.";
                case "btnExtendLeft": return "Stretches shapes leftward. Single shape extends to page left; multiple shapes extend to the last selected shape's left edge.";
                case "btnExtendRight": return "Stretches shapes rightward. Single shape extends to page right; multiple shapes extend to the last selected shape's right edge.";
                case "btnFitHorz": return "Resizes shapes to fill horizontal space. Single shape fits to page width with margins; multiple shapes fit within the last selected shape with specified gaps.";
                case "btnFitVert": return "Resizes shapes to fill vertical space. Single shape fits to page height with margins; multiple shapes fit within the last selected shape with specified gaps.";
                case "b_RatioResize": return "Scales shapes by a specified ratio. Prompts for resize method (width, height, or both) and ratio value. Scales proportionally from center.";
                case "b_Rotation": return "Matches rotation angle of all selected shapes to the last selected shape. Preserves shape positions while adjusting rotation.";
                case "btnCycleSwap": return "Rotates the positions of selected shapes. Each shape moves to the next shape's location in the selection order.";
                case "btnAutoGrid": return "Arranges selected shapes into a tidy grid layout. Automatically calculates optimal rows and columns for the selection.";
                case "btnCellCenter": return "Centers selected shapes within their respective grid cells. Useful for aligning shapes to invisible gridlines or guides.";

                // Library
                case "btnLibrary": return "Opens the asset library manager. Browse, organize, and insert reusable content including shapes, text blocks, and custom libraries. Ability to create a shared company or team library (path set in Office Pragmatics settings).";

                // Doc Tools
                case "btnFormEditor": return "Opens the form builder. Design templates with fillable field types (Text, Date, Yes/No, Multiple Choice), and prompts. Fill form mode presents a questionnaire interface to populate all fields at once. Useful for template documents such as M&A NDAs.";
                case "btnDocCompare": return "Compare two versions of DOCX, DOC or PDF files and generate outputs with tracked changes / redlines.";
                case "btnSmartEyeDrop": return "Captures colors from anywhere on screen to dynamically modify fill, outline, and font colors. Click to pick a color and automatically apply it to the selected shape's fill or outline.";
                case "btnNumberToWords": return "Converts numeric values to written words with original number in parentheses. Works on selected text (converts all numbers found) or at cursor position (converts nearest number). Example: '123' becomes 'one hundred twenty-three (123)'.";
                case "btnFlattenColors": return "Converts theme colors to RGB values. Breaks the link to theme colors, making colors fixed and preventing unintended changes when themes are modified.";
                case "btnLineUp": return "Increases line spacing by 2 points. Adds vertical spacing between lines in selected text or paragraphs.";
                case "btnLineDown": return "Decreases line spacing by 2 points. Reduces vertical spacing between lines in selected text or paragraphs.";

                // XLS Linking
                case "galImportChart": return "Imports charts from Excel. Choose vector image, PNG image, or native chart object with optional live data linking.";
                case "b_vectorChartImport": return "Imports Excel chart as vector shapes. Converts chart to editable Word shapes, breaking the link to Excel.";
                case "b_pictureChartImport": return "Imports Excel chart as PNG image. Pastes chart as a static picture for maximum compatibility and file size efficiency.";
                case "b_chartChartImport": return "Imports Excel chart as native object. Creates a linked chart that updates when the source Excel file changes.";
                case "galImportTable": return "Imports tables from Excel. Choose PNG image or native Word table with optional live data linking.";
                case "b_pictureRangeImport": return "Imports Excel range as PNG image. Pastes selected cells as a static picture for consistent formatting across systems.";
                case "b_tableRangeImport": return "Imports Excel range as native table. Converts selected cells to a Word table with optional live data linking.";
                case "galImportText": return "Imports text from Excel. Paste cell values as formatted text with automatic positioning and styling.";
                case "btnImportText": return "Imports Excel text as text blocks. Pastes selected cell values as individual or merged text elements in Word.";
                case "btnLinkManager": return "Opens the link manager dialog. View, edit, break, or refresh all Excel links in the document from a central interface.";
                case "btnRefreshLinks": return "Refreshes all Excel links in the document. Updates linked charts, tables, and text with the latest data from source files.";

                // Page
                case "btnPageHop": return "Jumps to a specific page number. Quick navigation tool for moving directly to any page in the document.";

                // File
                case "btnVersionUp": return "Increments the version number in the filename. Saves a new copy with the version counter increased (e.g., _v3 becomes _v4).";
                case "btnSaveAsPDF": return "Exports document to PDF format. Saves the entire document or selected pages as a PDF file with quality options.";
                case "btnSelectExport": return "Exports selected content. Saves current selection as a new document or as individual files.";
                case "galFileTools": return "File and folder utilities. Quick access to common file operations like showing in Explorer, copying paths, and QAT setup.";
                case "b_Show_Explorer": return "Opens File Explorer with the current file selected. Quickly navigate to the document's folder location.";
                case "b_Copy_Path": return "Copies the full file path to clipboard. Pastes the complete file path including drive, folder, and filename.";
                case "btnSettings": return "Opens add-in settings dialog. Configure preferences, keyboard shortcuts, default behaviors, and feature options.";

                default: return string.Empty;
            }
        }

        // ══════════════════════════════════════════════════════════════════
        // Universal action handler for all button clicks
        // ══════════════════════════════════════════════════════════════════
        public void OnUniversalAction(IRibbonControl control)
        {
            try
            {
                switch (control.Id)
                {
                    // Alignment
                    case "btnAlignLeft":
                        AlignmentHelper.AlignLeft(GetSelection());
                        break;
                    case "btnAlignTop":
                        AlignmentHelper.AlignTop(GetSelection());
                        break;
                    case "btnAlignHorizontal":
                        AlignmentHelper.DistributeHorizontal(GetShapeRange());
                        break;
                    case "btnAlignRight":
                        AlignmentHelper.AlignRight(GetSelection());
                        break;
                    case "btnAlignBottom":
                        AlignmentHelper.AlignBottom(GetSelection());
                        break;
                    case "btnAlignVertical":
                        AlignmentHelper.DistributeVertical(GetShapeRange());
                        break;
                    case "btnAlignCenter":
                        AlignmentHelper.AlignCenter(GetSelection());
                        break;
                    case "btnAlignMiddle":
                        AlignmentHelper.AlignMiddle(GetSelection());
                        break;
                    case "btnAlignTogether":
                        ShapeUtility.SnapTogether(GetSelection());
                        break;

                    // Size and Position
                    case "btnSameHW":
                        SizeHelper.SameSize(GetShapeRange());
                        break;
                    case "btnSameHeight":
                        SizeHelper.SameHeight(GetShapeRange());
                        break;
                    case "btnSameWidth":
                        SizeHelper.SameWidth(GetShapeRange());
                        break;
                    case "btnExtendTop":
                        SizeHelper.ExtendTop(GetShapeRange());
                        break;
                    case "btnExtendBottom":
                        SizeHelper.ExtendBottom(GetShapeRange());
                        break;
                    case "btnExtendLeft":
                        SizeHelper.ExtendLeft(GetShapeRange());
                        break;
                    case "btnExtendRight":
                        SizeHelper.ExtendRight(GetShapeRange());
                        break;
                    case "btnFitHorz":
                        SizeHelper.FitToHorizontal(GetShapeRange());
                        break;
                    case "btnFitVert":
                        SizeHelper.FitToVertical(GetShapeRange());
                        break;
                    case "b_RatioResize":
                        SizeHelper.RatioResize(GetShapeRange());
                        break;
                    case "b_Rotation":
                        AlignmentHelper.SameRotation(GetSelection());
                        break;
                    case "btnCycleSwap":
                        ShapeUtility.PositionSwap(GetSelection());
                        break;
                    case "btnAutoGrid":
                        ShapeUtility.AutoGrid(GetShapeRange());
                        break;
                    case "btnCellCenter":
                        ShapeUtility.CellCenter(GetApplication());
                        break;

                    case "btnLibrary": LibraryUtility.OpenLibrary(); break;

                    // Doc Tools
                    case "btnFormEditor":
                        FormEditor.ShowFormEditor();
                        break;
                    case "btnDocCompare":
                        DocComparison.RunDocComparison();
                        break;
                    case "btnSmartEyeDrop":
                        ColorDropperUtility.SmartColorPicker();
                        break;
                    case "btnNumberToWords":
                        TextUtility.ConvertNumberToWords(GetApplication());
                        break;
                    case "btnFlattenColors":
                        ColorUtility.ColorFreeze(GetApplication());
                        break;
                    case "btnLineUp":
                        TextUtility.ChangeParagraphSpacing(GetApplication(), +2f);
                        break;
                    case "btnLineDown":
                        TextUtility.ChangeParagraphSpacing(GetApplication(), -2f);
                        break;

                    // XLS Linking
                    case "b_vectorChartImport":
                         ChartLinkUtility.ImportExcelChartAsVector(GetApplication()); 
                        break;
                    case "b_pictureChartImport":
                         ChartLinkUtility.ImportExcelChartAsPicture(GetApplication()); 
                        break;
                    case "b_chartChartImport":
                         ChartLinkUtility.ImportExcelChartAsNative(GetApplication()); 
                        break;
                    case "b_pictureRangeImport":
                         RangeLinkUtility.ImportExcelRangeAsPicture(GetApplication()); 
                        break;
                    case "b_tableRangeImport":
                         RangeLinkUtility.ImportExcelRangeAsTable(GetApplication()); 
                        break;
                    case "btnImportText":
                         TextLinkUtility.ImportExcelText(GetApplication()); 
                        break;
                    case "btnRefreshLinks":
                         UpdateLinks.UpdateAllLinks(GetApplication()); 
                        break;
                    case "btnLinkManager":
                         LinkManager.ShowLinkManager(GetApplication()); 
                        break;

                    // Page
                    case "btnPageHop":
                        PageUtility.GoToPage(GetApplication());
                        break;

                    // File
                    case "btnVersionUp":
                        FileUtility.VersionUp(GetApplication());
                        break;
                    case "btnSaveAsPDF":
                        FileUtility.SavePDF(GetApplication());
                        break;
                    case "btnSelectExport":
                        FileUtility.ExportPages(GetApplication());
                        break;
                    case "btnSettings":
                        SettingsForm.ShowSettingsDialog();
                        break;
                    case "b_Show_Explorer":
                        FileUtility.OpenContainingFolder(GetApplication());
                        break;
                    case "b_Copy_Path":
                        FileUtility.CopyFilePathToClipboard(GetApplication());
                        break;

                    default:
                        MessageBox.Show($"Handler not implemented for: {control.Id}");
                        break;
                }
            }
            catch { }
        }

        // ══════════════════════════════════════════════════════════════════
        // Helper methods for accessing the Word Application object
        // ══════════════════════════════════════════════════════════════════

        private Word.Application GetApplication() => Globals.ThisAddIn.Application;
        private Word.Selection GetSelection() => GetApplication().Selection;

        // Note: Word doesn't have a direct equivalent to PowerPoint's ShapeRange
        private Word.ShapeRange GetShapeRange()
        {
            if (GetSelection().ShapeRange != null && GetSelection().ShapeRange.Count > 0)
            {
                return GetSelection().ShapeRange;
            }

            // If no shapes are selected, you might want to handle this differently
            throw new InvalidOperationException("No shapes are currently selected.");
        }
    }
}