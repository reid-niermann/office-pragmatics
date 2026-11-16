// RibbonController.cs
using Microsoft.Office.Core;
using OfficePragmatics_Powerpoint.Classes.ExcelLinking;
using OfficePragmatics_PowerPoint;
using OfficePragmatics_PowerPoint.Alignment;
using OfficePragmatics_PowerPoint.Chart;
using OfficePragmatics_PowerPoint.Color;
using OfficePragmatics_PowerPoint.ExcelLinking;
using OfficePragmatics_PowerPoint.File;
using OfficePragmatics_PowerPoint.Image;
using OfficePragmatics_PowerPoint.Map;
using OfficePragmatics_PowerPoint.Settings;
using OfficePragmatics_PowerPoint.Shape;
using OfficePragmatics_PowerPoint.Size;
using OfficePragmatics_PowerPoint.Slide;
using OfficePragmatics_PowerPoint.Text;
using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_Powerpoint
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
                case "btnAlignLeft": return Properties.Resources.btnAlignLeft;
                case "btnAlignTop": return Properties.Resources.btnAlignTop;
                case "btnAlignHorizontal": return Properties.Resources.btnAlignHorizontal;
                case "btnAlignRight": return Properties.Resources.btnAlignRight;
                case "btnAlignBottom": return Properties.Resources.btnAlignBottom;
                case "btnAlignVertical": return Properties.Resources.btnAlignVertical;
                case "btnAlignCenter": return Properties.Resources.btnAlignCenter;
                case "btnAlignMiddle": return Properties.Resources.btnAlignMiddle;
                case "btnAlignTogether": return Properties.Resources.btnAlignTogether;
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
                case "b_Move": return Properties.Resources.b_Move;
                case "b_Rotation": return Properties.Resources.b_Rotation;
                case "btnCycleSwap": return Properties.Resources.btnCycleSwap;
                case "btnAutoGrid": return Properties.Resources.btnAutoGrid;
                case "btnCellCenter": return Properties.Resources.btnCellCenter;
                case "btnInsertIcon": return Properties.Resources.btnInsertIcon;
                case "btnPlotMap": return Properties.Resources.btnPlotMap;
                case "btnInsertLogo": return Properties.Resources.btnInsertLogo;
                case "btnAlphabetize": return Properties.Resources.btnAlphabetize;
                case "btnRemoveWhite": return Properties.Resources.btnRemoveWhite;
                case "btnResizeLogos": return Properties.Resources.btnResizeLogos;
                case "btnAutoSelect": return Properties.Resources.btnAutoSelect;
                case "btnSmartEyeDrop": return Properties.Resources.btnSmartEyeDrop;
                case "btnRedaction": return Properties.Resources.btnRedaction;
                case "btnTBUBox": return Properties.Resources.btnTBUBox;
                case "btnFootnote": return Properties.Resources.btnFootnote;
                case "btnFlattenColors": return Properties.Resources.btnFlattenColors;
                case "btnZeroMargin": return Properties.Resources.btnZeroMargin;
                case "btnLineUp": return Properties.Resources.btnLineUp;
                case "btnLineDown": return Properties.Resources.btnLineDown;
                case "btnToggleAutofit": return Properties.Resources.btnToggleAutofit;
                case "btnConformEdges": return Properties.Resources.btnConformEdges;
                case "galPitchOther": return Properties.Resources.galPitchOther;
                case "b_AltFind": return Properties.Resources.b_AltFind;
                case "b_Duplicate": return Properties.Resources.b_Duplicate;
                case "btnLibrary": return Properties.Resources.btnLibrary;
                case "btnPasteFormat": return Properties.Resources.btnPasteFormat;
                case "btnCopyData": return Properties.Resources.btnCopyData;
                case "btnShapesToAxis": return Properties.Resources.btnShapesToAxis;
                case "btnAltTextSnap": return Properties.Resources.btnAltTextSnap;
                case "btnMaximize": return Properties.Resources.btnMaximize;
                case "btnCenterPie": return Properties.Resources.btnCenterPie;
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
                case "btnSlideHop": return Properties.Resources.btnSlideHop;
                case "btnVersionUp": return Properties.Resources.btnVersionUp;
                case "btnSaveAsPDF": return Properties.Resources.btnSaveAsPDF;
                case "btnSelectExport": return Properties.Resources.btnSelectExport;
                case "btnCustomColors": return Properties.Resources.btnCustomColors;
                case "btnSettings": return Properties.Resources.btnSettings;
                case "galFileTools": return Properties.Resources.galFileTools;
                case "b_Show_Explorer": return Properties.Resources.b_Show_Explorer;
                case "b_Copy_Path": return Properties.Resources.b_Copy_Path;
                case "b_QAT": return Properties.Resources.b_QAT;
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
                case "btnSameHW": return "Same Height & Width";
                case "btnSameHeight": return "Same Height";
                case "btnSameWidth": return "Same Width";
                case "btnExtendTop": return "Extend Top";
                case "btnExtendBottom": return "Extend Bottom";
                case "btnExtendLeft": return "Extend Left";
                case "btnExtendRight": return "Extend Right";
                case "btnFitHorz": return "Fit Horizontal";
                case "btnFitVert": return "Fit Vertical";
                case "b_RatioResize": return "Ratio Resize";
                case "b_Move": return "Measure/Move";
                case "b_Rotation": return "Align Rotations";
                case "btnCycleSwap": return "Cycle Swap";
                case "btnAutoGrid": return "Auto Grid";
                case "btnCellCenter": return "Cell Center";
                // Graphics
                case "btnInsertIcon": return "Insert Icon";
                case "btnInsertLogo": return "Insert Logo";
                case "btnPlotMap": return "Plot Map";
                case "btnAlphabetize": return "Sort Logos";
                case "btnRemoveWhite": return "Remove White";
                case "btnResizeLogos": return "Resize Logos";
                // Library
                case "btnLibrary": return "Library Assets";
                // Pitch Tools
                case "btnSmartEyeDrop": return "Smart Eye Dropper";
                case "btnAutoSelect": return "Auto Select Shapes";
                case "btnRedaction": return "Apply Redaction Overlay";
                case "btnTBUBox": return "Insert TBU Box";
                case "btnFootnote": return "Insert Footnote";
                case "btnFlattenColors": return "Flatten Colors";
                case "btnZeroMargin": return "Remove Text Margins";
                case "btnToggleAutofit": return "Toggle Autofit";
                case "btnConformEdges": return "Conform AutoShape Edges";
                case "btnLineUp": return "Increase Line Spacing";
                case "btnLineDown": return "Decrease Line Spacing";
                case "galPitchOther": return "Other Tools";
                case "b_AltFind": return "Alt-Text Find";
                case "b_Duplicate": return "Duplicate Across Slides";
                // Chart Tools
                case "btnPasteFormat": return "Paste Chart Format";
                case "btnCopyData": return "Copy Underlying Data";
                case "btnAltTextSnap": return "Snap Shapes to X-Axis";
                case "btnShapesToAxis": return "Match Logos to X-Axis";
                case "btnMaximize": return "Maximize Plot Area";
                case "btnCenterPie": return "Center Pie Chart";
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
                case "galLinkMatch": return "Links";
                case "btnLinkManager": return "Link Manager";
                case "btnRefreshLinks": return "Refresh Links";
                // Presentation
                case "btnSlideHop": return "Go to Slide";
                // File
                case "btnVersionUp": return "Version Up";
                case "btnSaveAsPDF": return "Save PDF";
                case "btnSelectExport": return "Export Selection";
                case "btnCustomColors": return "Custom Colors";
                case "galFileTools": return "File Tools";
                case "b_Show_Explorer": return "Show in Explorer";
                case "b_Copy_Path": return "Copy File Path";
                case "b_QAT": return "Default QAT";
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
                case "b_Move": return "PM";
                case "b_Rotation": return "PR";
                case "btnCycleSwap": return "1";
                case "btnAutoGrid": return "2";
                case "btnCellCenter": return "3";
                // Graphics
                case "btnInsertIcon": return "4";
                case "btnInsertLogo": return "5";
                case "btnPlotMap": return "6";
                case "btnAlphabetize": return "7";
                case "btnRemoveWhite": return "8";
                case "btnResizeLogos": return "9";
                // Library
                case "btnLibrary": return "0";
                // Pitch Tools
                case "btnSmartEyeDrop": return "I";
                case "btnAutoSelect": return "O";
                case "btnRedaction": return "KR";
                case "btnTBUBox": return "U";
                case "btnFootnote": return "F";
                case "btnFlattenColors": return "KF";
                case "btnZeroMargin": return "G";
                case "btnToggleAutofit": return "J";
                case "btnConformEdges": return "KA";
                case "btnLineUp": return ",";
                case "btnLineDown": return ".";
                case "galPitchOther": return "KO";
                // Chart Tools
                case "btnPasteFormat": return "WF";
                case "btnCopyData": return "WD";
                case "btnAltTextSnap": return "WS";
                case "btnShapesToAxis": return "WL";
                case "btnMaximize": return "WP";
                case "btnCenterPie": return "WC";
                // XLS Linking
                case "galImportChart": return "QC";
                case "galImportTable": return "QR";
                case "galImportText": return "QT";
                case "galLinkMatch": return "QS";
                case "btnLinkManager": return "QM";
                case "btnRefreshLinks": return "QU";
                // Presentation
                case "btnSlideHop": return "X";
                // File
                case "btnVersionUp": return "[";
                case "btnSaveAsPDF": return "]";
                case "btnSelectExport": return "\\";
                case "btnCustomColors": return "-";
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
                case "btnAlignLeft": return "Aligns shapes to the left. Single shape selections will align to slide edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignTop": return "Aligns shapes to the top. Single shape selections will align to slide edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignHorizontal": return "Evenly distributes selected shapes horizontally with equal spacing between them. Requires at least three shapes.";
                case "btnAlignRight": return "Aligns shapes to the right. Single shape selections will align to slide edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignBottom": return "Aligns shapes to the bottom. Single shape selections will align to slide edge; multiple selected shapes will align to the last selected shape. Aligns to shape border/outline edge by default.";
                case "btnAlignVertical": return "Evenly distributes selected shapes vertically with equal spacing between them. Requires at least three shapes.";
                case "btnAlignCenter": return "Centers shapes horizontally. Single shape selections will center on slide; multiple selected shapes will align to the last selected shape's horizontal center.";
                case "btnAlignMiddle": return "Centers shapes vertically. Single shape selections will center on slide; multiple selected shapes will align to the last selected shape's vertical middle.";
                case "btnAlignTogether": return "Snaps two shapes together edge-to-edge. Detects the relative position of the shapes and moves the first shape to touch the second. Requires exactly two shapes.";
                // Size and Position
                case "btnSameHW": return "Resizes all selected shapes to match the last selected shape's width and height. Accounts for shape borders when calculating dimensions.";
                case "btnSameHeight": return "Resizes all selected shapes to match the last selected shape's height. Accounts for shape borders when calculating dimensions.";
                case "btnSameWidth": return "Resizes all selected shapes to match the last selected shape's width. Accounts for shape borders when calculating dimensions.";
                case "btnExtendTop": return "Stretches shapes upward. Single shape extends to slide top; multiple shapes extend to the last selected shape's top edge.";
                case "btnExtendBottom": return "Stretches shapes downward. Single shape extends to slide bottom; multiple shapes extend to the last selected shape's bottom edge.";
                case "btnExtendLeft": return "Stretches shapes leftward. Single shape extends to slide left; multiple shapes extend to the last selected shape's left edge.";
                case "btnExtendRight": return "Stretches shapes rightward. Single shape extends to slide right; multiple shapes extend to the last selected shape's right edge.";
                case "btnFitHorz": return "Resizes shapes to fill horizontal space. Single shape fits to slide width with margins; multiple shapes fit within the last selected shape with specified gaps.";
                case "btnFitVert": return "Resizes shapes to fill vertical space. Single shape fits to slide height with margins; multiple shapes fit within the last selected shape with specified gaps.";
                case "b_RatioResize": return "Scales shapes by a specified ratio. Prompts for resize method (width, height, or both) and ratio value. Scales proportionally from center.";
                case "b_Move": return "Interactive measurement and positioning tool. Displays distance measurements between selected shapes and nearby objects or slide edges. Includes live preview and precise numeric input for moving shapes.";
                case "b_Rotation": return "Matches rotation angle of all selected shapes to the last selected shape. Preserves shape positions while adjusting rotation.";
                case "btnCycleSwap": return "Rotates the positions of selected shapes. Each shape moves to the next shape's location in the selection order.";
                case "btnAutoGrid": return "Arranges selected shapes into a tidy grid layout. Automatically calculates optimal rows and columns for the selection.";
                case "btnCellCenter": return "Centers selected shapes within their respective grid cells. Useful for aligning shapes to invisible gridlines or guides.";
                // Graphics
                case "btnInsertIcon": return "Search Freepik for SVG icons and insert them; they can be recolored like any PowerPoint shape. Icons appear immediately to the right of the active slide. Requires a Freepik API key in Settings > API Keys.";
                case "btnInsertLogo": return "Search and insert logos from the web with automatic harmonization and sizing. Logos will appear immediately to the right of the active slide.";
                case "btnPlotMap": return "Inserts svg map of the continental U.S. based on provided list of locations. Useful for plotting locations within PowerPoint. Custom plot markers can be applied (based on the active shape prior to function activation).";
                case "btnAlphabetize": return "Sorts selected logo images alphabetically, while maintaining original spacing and alignment. Sorting function based on alternative text while is added through out Insert Logo tool. Logos added to the presentation manually will need to be reviewed for alternative text.";
                case "btnRemoveWhite": return "Removes white backgrounds from selected images. Converts white pixels to transparent, useful for cleaning up imported graphics.";
                case "btnResizeLogos": return "Harmonizes logo sizes across selection. Resizes logos proportionally based on their visual area to create balanced compositions.";
                // Library
                case "btnLibrary": return "Opens the asset library manager. Browse, organize, and insert reusable content including shapes, slides, and custom libraries. Ability to create a shared company or team library (path set in Office Pragmatics settings).";
                // Pitch Tools
                case "btnSmartEyeDrop": return "Captures colors from anywhere on screen to dynamically modify fill, outline, and font colors. Click to pick a color and automatically apply it to the selected shape's fill or outline.";
                case "btnAutoSelect": return "Selects all shapes of the same type or style. Options include selecting by shape type, color, size, or formatting attributes.";
                case "btnRedaction": return "Creates redaction overlay for text and shapes. Adds black boxes to hide sensitive content in presentations. Note, these overlays can be removed in professional PDF editors. For confidential, sensitive documents, apply redaction in a professional PDF editor.";
                case "btnTBUBox": return "Inserts a 'To Be Updated' placeholder box. Marks areas requiring future updates.";
                case "btnFootnote": return "Adds formatted footnotes to slides. Inserts footnote text boxes with consistent styling and automatic numbering.";
                case "btnFlattenColors": return "Converts theme colors to RGB values. Breaks the link to theme colors, making colors fixed and preventing unintended changes.";
                case "btnZeroMargin": return "Removes internal margins from text boxes. Sets all text box margins to zero for maximum text area utilization.";
                case "btnToggleAutofit": return "Toggles AutoFit settings for selected text boxes. Cycles between AutoFit options: none, shrink text on overflow, or resize shape.";
                case "btnConformEdges": return "Aligns shape edges to match. Adjusts the bounding box of connector shapes to match their actual drawn path.";
                case "btnLineUp": return "Increases line spacing by 2 points. Adds vertical spacing between lines in selected text boxes or shapes.";
                case "btnLineDown": return "Decreases line spacing by 2 points. Reduces vertical spacing between lines in selected text boxes or shapes.";
                case "galPitchOther": return "Additional pitch deck tools. Access supplementary features for presentation formatting and cleanup.";
                case "b_AltFind": return "Advanced find and replace tool. Search and replace text across slides with support for formatting, alt text, and shape properties.";
                case "b_Duplicate": return "Duplicates selected shapes across multiple slides. Copies shapes to specified slides while maintaining position and formatting.";
                // Chart Tools
                case "btnPasteFormat": return "Applies copied chart formatting to selected charts. Transfers colors, fonts, axes, and other formatting from one chart to another.";
                case "btnCopyData": return "Copies chart data to clipboard. Extracts the underlying data table from the selected chart for pasting into Excel or other charts.";
                case "btnAltTextSnap": return "Snaps data labels to chart axes. Aligns selected text shapes or labels to precise axis values for accurate positioning.";
                case "btnShapesToAxis": return "Positions logos along chart axes. Places selected shapes at specific data points on chart axes with proper alignment.";
                case "btnMaximize": return "Maximizes the plot area size. Expands the chart's plot area to fill available space by minimizing margins and legend areas.";
                case "btnCenterPie": return "Centers pie charts. Ensures that the plot area is centered inside the chart area.";
                // XLS Linking
                case "galImportChart": return "Imports charts from Excel with live link tracking. Select a chart in Excel first, then choose import format. Links can be refreshed to update with source data changes.";
                case "b_vectorChartImport": return "Imports Excel chart as SVG vector image with live link. Chart becomes an editable vector graphic in PowerPoint while maintaining link metadata for future updates.";
                case "b_pictureChartImport": return "Imports Excel chart as Enhanced Metafile (EMF) picture with live link. Creates a high-quality static image while maintaining link metadata for future updates.";
                case "b_chartChartImport": return "Imports Excel chart as native PowerPoint chart object with live link. Creates fully editable chart in PowerPoint while maintaining connection to Excel source for data updates.";
                case "galImportTable": return "Imports Excel ranges as pictures or native tables with live link tracking. Select cells in Excel first, then choose import format. Links can be refreshed to update with source data.";
                case "b_pictureRangeImport": return "Imports Excel range as Enhanced Metafile (EMF) picture with live link. Creates a high-quality static image while maintaining link metadata for future updates.";
                case "b_tableRangeImport": return "Imports Excel range as native PowerPoint table with live link. Converts cells to editable PowerPoint table while maintaining connection to Excel source for data updates.";
                case "galImportText": return "Imports text from Excel with live link tracking. Select a single cell in Excel, then import as formatted text box. Link maintains reference for future updates.";
                case "btnImportText": return "Imports single Excel cell as text box with live link. Creates a formatted text shape in PowerPoint while maintaining connection to Excel cell for content updates.";
                case "galLinkMatch": return "Link management tools for Excel-linked content. Refresh individual links or manage all links in presentation from central interface.";
                case "btnLinkManager": return "Opens link manager dialog showing all Excel-linked objects in presentation. View source paths, refresh individual items, break links, or batch update all links at once.";
                case "btnRefreshLinks": return "Refreshes all linked content in presentation. Updates charts, tables, and text boxes tagged with Pragmatics link metadata from their source workbooks. Handles version upgrades automatically.";
                // Presentation
                case "btnSlideHop": return "Jumps to a specific slide number. Quick navigation tool for moving directly to any slide in the presentation.";
                // File
                case "btnVersionUp": return "Increments the version number in the filename. Saves a new copy with the version counter increased (e.g., _v3 becomes _v4).";
                case "btnSaveAsPDF": return "Exports presentation to PDF format. Saves the entire presentation or selected slides as a PDF file with quality options.";
                case "btnSelectExport": return "Exports selected slides or shapes. Saves current selection as a new presentation or as individual image files.";
                case "btnCustomColors": return "Manages custom color palettes. Add, edit, or remove custom colors for quick access in shape formatting dialogs.";
                case "galFileTools": return "File and folder utilities. Quick access to common file operations like showing in Explorer, copying paths, and QAT setup.";
                case "b_Show_Explorer": return "Opens File Explorer with the current file selected. Quickly navigate to the presentation's folder location.";
                case "b_Copy_Path": return "Copies the full file path to clipboard. Pastes the complete file path including drive, folder, and filename.";
                case "b_QAT": return "Restores default Quick Access Toolbar. Resets QAT to the recommended set of frequently used commands.";
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
                    case "btnAlignLeft": AlignmentHelper.AlignLeft(GetSelection()); break;
                    case "btnAlignTop": AlignmentHelper.AlignTop(GetSelection()); break;
                    case "btnAlignHorizontal": AlignmentHelper.DistributeHorizontal(GetShapeRange()); break;
                    case "btnAlignRight": AlignmentHelper.AlignRight(GetSelection()); break;
                    case "btnAlignBottom": AlignmentHelper.AlignBottom(GetSelection()); break;
                    case "btnAlignVertical": AlignmentHelper.DistributeVertical(GetShapeRange()); break;
                    case "btnAlignCenter": AlignmentHelper.AlignCenter(GetSelection()); break;
                    case "btnAlignMiddle": AlignmentHelper.AlignMiddle(GetSelection()); break;
                    case "btnAlignTogether": ShapeUtility.SnapTogether(GetSelection()); break;
                    case "btnSameHW": SizeHelper.SameSize(GetShapeRange()); break;
                    case "btnSameHeight": SizeHelper.SameHeight(GetShapeRange()); break;
                    case "btnSameWidth": SizeHelper.SameWidth(GetShapeRange()); break;
                    case "btnExtendTop": SizeHelper.ExtendTop(GetShapeRange()); break;
                    case "btnExtendBottom": SizeHelper.ExtendBottom(GetShapeRange()); break;
                    case "btnExtendLeft": SizeHelper.ExtendLeft(GetShapeRange()); break;
                    case "btnExtendRight": SizeHelper.ExtendRight(GetShapeRange()); break;
                    case "btnFitHorz": SizeHelper.FitToHorizontal(GetShapeRange()); break;
                    case "btnFitVert": SizeHelper.FitToVertical(GetShapeRange()); break;
                    case "b_RatioResize": SizeHelper.RatioResize(GetShapeRange()); break;
                    case "b_Move": ShapeUtility.MeasureMove(GetSelection()); break;
                    case "b_Rotation": AlignmentHelper.SameRotation(GetSelection()); break;
                    case "btnCycleSwap": ShapeUtility.PositionSwap(GetSelection()); break;
                    case "btnAutoGrid": ShapeUtility.AutoGrid(GetShapeRange()); break;
                    case "btnCellCenter": ShapeUtility.CellCenter(GetApplication()); break;
                    case "btnLibrary": LibraryUtility.OpenLibrary(); break;
                    case "btnInsertIcon": IconUtility.IconLookup(GetApplication()); break;
                    case "btnPlotMap": MapUtility.InsertMap(GetApplication()); break;
                    case "btnInsertLogo": LogoUtility.LogoLookup(GetApplication()); break;
                    case "btnAlphabetize": ImageUtility.SortLogos(GetSelection()); break;
                    case "btnRemoveWhite": ImageUtility.RemoveWhiteBackground(GetShapeRange()); break;
                    case "btnResizeLogos": ImageUtility.HarmonizeLogos(GetSelection()); break;
                    case "btnAutoSelect": ShapeUtility.AutoSelect(GetApplication()); break;
                    case "btnSmartEyeDrop": ColorDropperUtility.SmartColorPicker(); break;
                    case "btnRedaction": SlideRedaction.Redact(GetApplication()); break;
                    case "btnTBUBox": SlideUtility.InsertTBU(GetApplication()); break;
                    case "btnFootnote": SlideUtility.InsertFootnote(GetApplication()); break;
                    case "btnFlattenColors": ColorUtility.ColorFreeze(GetApplication()); break;
                    case "btnZeroMargin": TextUtility.RemoveMargin(GetShapeRange()); break;
                    case "btnLineUp": TextUtility.ChangeParagraphSpacing(GetApplication(), +2f); break;
                    case "btnLineDown": TextUtility.ChangeParagraphSpacing(GetApplication(), -2f); break;
                    case "btnToggleAutofit": ShapeUtility.ToggleAutoFit(GetSelection()); break;
                    case "btnConformEdges": ShapeUtility.ConformAutoShapes(GetSelection()); break;
                    case "b_AltFind": TextUtility.AltFind(GetApplication()); break;
                    case "b_Duplicate": ShapeUtility.duplicateAcrossSlides(GetApplication()); break;
                    case "btnPasteFormat": ChartUtility.PasteChartFormat(GetShapeRange()); break;
                    case "btnCopyData": ChartUtility.CopyChartData(GetSelection()); break;
                    case "btnShapesToAxis": ChartUtility.ShapesToChartAxis(GetApplication()); break;
                    case "btnAltTextSnap": ChartUtility.ShapeToChartAltText(GetApplication()); break;
                    case "btnMaximize": ChartUtility.MaxPlotArea(GetShapeRange()); break;
                    case "btnCenterPie": ChartUtility.CenterPlotArea(Globals.ThisAddIn.Application.ActiveWindow.Selection); break;
                    case "b_vectorChartImport": ChartLinkUtility.ImportExcelChartAsVector(GetApplication()); break;
                    case "b_pictureChartImport": ChartLinkUtility.ImportExcelChartAsPicture(GetApplication()); break;
                    case "b_chartChartImport": ChartLinkUtility.ImportExcelChartAsNative(GetApplication()); break;
                    case "b_pictureRangeImport": RangeLinkUtility.ImportExcelRangeAsPicture(GetApplication()); break;
                    case "b_tableRangeImport": RangeLinkUtility.ImportExcelRangeAsTable(GetApplication()); break;
                    case "btnImportText": TextLinkUtility.ImportExcelText(GetApplication()); break;
                    case "btnRefreshLinks": UpdateLinks.UpdateAllLinks(GetApplication()); break;
                    case "btnLinkManager": LinkManager.ShowLinkManager(GetApplication()); break;
                    case "btnSlideHop": GoToSlideUtility.GoToSlide(GetApplication()); break;
                    case "btnVersionUp": FileUtility.VersionUp(GetApplication()); break;
                    case "btnSaveAsPDF": FileUtility.SavePDF(GetApplication()); break;
                    case "btnSelectExport": ExportUtility.ExportSelection(GetApplication()); break;
                    case "btnCustomColors": CustomColorManager.ShowDialog(Globals.ThisAddIn.Application); break;
                    case "btnSettings": SettingsForm.ShowSettingsDialog(); break;
                    case "b_Show_Explorer": FileUtility.OpenContainingFolder(); break;
                    case "b_Copy_Path": FileUtility.CopyFilePathToClipboard(); break;
                    case "b_QAT": FileUtility.SetDefaultQuickAccessToolbar(GetApplication()); break;

                    default:
                        MessageBox.Show($"Handler not implemented for: {control.Id}");
                        break;
                }
            }
            catch { }
        }

        // ══════════════════════════════════════════════════════════════════
        // Helper methods for accessing the PowerPoint Application object
        // ══════════════════════════════════════════════════════════════════
        private PowerPoint.Application GetApplication() => Globals.ThisAddIn.Application;
        private PowerPoint.Selection GetSelection() => GetApplication().ActiveWindow.Selection;
        private PowerPoint.ShapeRange GetShapeRange() => GetSelection().ShapeRange;
    }
}