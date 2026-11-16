using Microsoft.Office.Core;
using OfficePragmatics_Excel.Alignment;
using OfficePragmatics_Excel.Classes;
using OfficePragmatics_Excel.Classes.Charting;
using OfficePragmatics_Excel.Features;
using OfficePragmatics_Excel.Library;
using OfficePragmatics_Excel.Settings;
using OfficePragmatics_Excel.Size;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    public class RibbonController
    {

        public Bitmap GetCustomImage(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Main buttons
                case "btnTracePrecedents": return Properties.Resources.btnTracePrecedents;
                case "btnTraceDependents": return Properties.Resources.btnTraceDependents;
                case "btnDrillPrecedents": return Properties.Resources.btnDrillPrecedents;
                case "btnDrillDependents": return Properties.Resources.btnDrillDependents;
                case "btnInsertSymbols": return Properties.Resources.btnInsertSymbols;
                case "btnSheetSelect": return Properties.Resources.btnSheetSelect;
                case "btnDeepSearch": return Properties.Resources.btnDeepSearch;
                case "btnCleanFile": return Properties.Resources.btnCleanFile;
                case "btnVisualCheck": return Properties.Resources.btnVisualCheck;
                case "btnLibrary": return Properties.Resources.btnLibrary;
                case "galImportChart": return Properties.Resources.btnChart;
                case "galImportTable": return Properties.Resources.btnRange;
                case "galImportText": return Properties.Resources.btnText;
                case "btnVersionUp": return Properties.Resources.btnVersionUp;
                case "btnSaveAsPDF": return Properties.Resources.btnSaveAsPDF;
                case "btnExportSheet": return Properties.Resources.btnExportSheet;
                case "galFileTools": return Properties.Resources.btnFileTools;
                case "btnSettings": return Properties.Resources.btnSettings;
                case "btnRefresh": return Properties.Resources.btnRefresh;
                case "btnRemoveAllTracePrecedentArrows": return Properties.Resources.b_RemoveFootnotes;
                case "btnRemoveAllTraceDependentArrows": return Properties.Resources.b_RemoveFootnotes;
                case "galFormulaTools": return Properties.Resources.galFormulaTools;
                case "galPasteTools": return Properties.Resources.galPasteTools;
                case "galNumberFormats": return Properties.Resources.galNumberFormats;
                case "galFontColors": return Properties.Resources.galFontColors;
                case "galFillColors": return Properties.Resources.galFillColors;
                case "galOtherCycles": return Properties.Resources.galOtherCycles;
                case "galFormatTools": return Properties.Resources.galFormatTools;
                case "galRangeTools": return Properties.Resources.galRangeTools;
                case "galSizingTools": return Properties.Resources.galSizingTools;
                case "galShapeAlignment": return Properties.Resources.galShapeAlignment;
                case "galFormatToolsChart": return Properties.Resources.galFormatToolsChart;
                case "galLabelTools": return Properties.Resources.galLabelTools;
                case "galInsertMetrics": return Properties.Resources.galInsertMetrics;
                case "galDuplicateChart": return Properties.Resources.galDuplicateChart;
                case "galCommonCharts": return Properties.Resources.b_CommonCharts;
                case "btnRevertTrace": return Properties.Resources.btnRevertDependents;
                case "menuPrecedentOptions": return Properties.Resources.btnRemoveAllTraceArrows;
                case "menuDepedentOptions": return Properties.Resources.btnRemoveAllTraceArrows;
                case "b_Magnitude_Plus": return Properties.Resources.b_Magnitude_Plus;
                case "b_Magnitude_Minus": return Properties.Resources.b_Magnitude_Minus;
                case "b_FlipSign": return Properties.Resources.b_FlipSign;
                case "b_FillRight": return Properties.Resources.b_FillRight;
                case "b_FillDown": return Properties.Resources.b_FillDown;
                case "b_ErrorWrapNA": return Properties.Resources.b_ErrorWrapNA;
                case "b_ErrorWrapZero": return Properties.Resources.b_ErrorWrapZero;
                case "b_ConvertToLet": return Properties.Resources.b_ConvertToLet;
                case "b_FormulaWrap": return Properties.Resources.b_FormulaWrap;
                case "b_SolveFormula": return Properties.Resources.b_SolveFormula;
                case "b_AnchorCycle": return Properties.Resources.b_AnchorCycle;
                case "b_pasteExact": return Properties.Resources.b_pasteExact;
                case "b_pasteDuplicate": return Properties.Resources.b_pasteDuplicate;
                case "b_PasteNumberFormat": return Properties.Resources.b_PasteNumberFormat;
                case "b_pasteFillColor": return Properties.Resources.b_pasteFillColor;
                case "b_pasteBorders": return Properties.Resources.b_pasteBorders;
                case "b_pasteCellWidths": return Properties.Resources.b_pasteCellWidths;
                case "b_pasteCellHeights": return Properties.Resources.b_pasteCellHeights;
                case "btnBrandFormat": return Properties.Resources.btnBrandFormat;
                case "btnBrandFillCycle": return Properties.Resources.btnBrandFillCycle;
                case "btnBrandBorderCycle": return Properties.Resources.btnBrandBorderCycle;
                case "btnBrandFontColorCycle": return Properties.Resources.btnBrandFontColorCycle;
                case "btnChartBrandFormat": return Properties.Resources.btnChart;
                case "btnTableBrandFormat": return Properties.Resources.btnRange;
                case "b_NumberCycle": return Properties.Resources.galNumberFormats;
                case "b_CurrencyCycle": return Properties.Resources.b_CurrencyCycle;
                case "b_PercentCycle": return Properties.Resources.b_PercentCycle;
                case "b_DateCycle": return Properties.Resources.b_DateCycle;
                case "b_MultipleCycle": return Properties.Resources.b_MultipleCycle;
                case "b_BinaryCycle": return Properties.Resources.b_BinaryCycle;
                case "b_PlusDecimals": return Properties.Resources.b_PlusDecimals;
                case "b_MinusDecimals": return Properties.Resources.b_MinusDecimals;
                case "b_BBColorCycle": return Properties.Resources.b_BBColorCycle;
                case "b_ModelColorCycle": return Properties.Resources.b_ModelColorCycle;
                case "b_AutoColorSelection": return Properties.Resources.b_AutoColorSelection;
                case "b_AutoTableColor": return Properties.Resources.b_AutoTableColor;
                case "b_ModelFillColorCycle": return Properties.Resources.b_ModelColorCycle;
                case "b_FillThemeColorCycle": return Properties.Resources.b_ThemeColorCycle;
                case "b_UnderlineCycle": return Properties.Resources.galOtherCycles;
                case "b_BorderCycle": return Properties.Resources.b_BorderCycle;
                case "b_FontSizeCycle": return Properties.Resources.b_FontSizeCycle;
                case "b_CellHeightCycle": return Properties.Resources.b_CellHeightCycle;
                case "b_CellWidthCycle": return Properties.Resources.b_CellWidthCycle;
                case "b_FootnoteCycle": return Properties.Resources.btnFootnote;
                case "b_RemoveFootnotes": return Properties.Resources.b_RemoveFootnotes;
                case "b_ListBullets": return Properties.Resources.b_ListBullets;
                case "b_AccountingDots": return Properties.Resources.b_AccountingDots;
                case "b_SumLine": return Properties.Resources.b_SumLine;
                case "b_LeftIndent": return Properties.Resources.b_LeftIndent;
                case "b_RightIndent": return Properties.Resources.b_RightIndent;
                case "b_TransparentToWhite": return Properties.Resources.b_TransparentToWhite;
                case "b_MergeToCenterAcrossSelection": return Properties.Resources.b_MergeToCenterAcrossSelection;
                case "b_hideToGroup": return Properties.Resources.b_hideToGroup;
                case "b_reverseColumns": return Properties.Resources.b_reverseColumns;
                case "b_reverseRows": return Properties.Resources.b_reverseRows;
                case "b_deleteBlankColumns": return Properties.Resources.b_deleteBlankColumns;
                case "b_deleteBlankRows": return Properties.Resources.b_deleteBlankRows;
                case "b_harcodeSelection": return Properties.Resources.b_harcodeSelection;
                case "b_calculateCAGR": return Properties.Resources.b_calculateCAGR;
                case "b_Guides": return Properties.Resources.b_Guides;
                case "b_ResizetoPPT": return Properties.Resources.b_ResizetoPPT;
                case "b_AutoGrid": return Properties.Resources.btnAutoGrid;
                case "b_SameWidth": return Properties.Resources.btnSameWidth;
                case "b_SameHeight": return Properties.Resources.btnSameHeight;
                case "b_SameWidthandHeight": return Properties.Resources.btnSameHW;
                case "b_ExtendLeft": return Properties.Resources.btnExtendLeft;
                case "b_ExtendRight": return Properties.Resources.btnExtendRight;
                case "b_ExtendTop": return Properties.Resources.btnExtendTop;
                case "b_ExtendBottom": return Properties.Resources.btnExtendBottom;
                case "b_FitHorizontal": return Properties.Resources.btnFitHorz;
                case "b_FitVertical": return Properties.Resources.btnFitVert;
                case "b_AlignLeft": return Properties.Resources.btnAlignLeft;
                case "b_AlignRight": return Properties.Resources.btnAlignRight;
                case "b_AlignMiddle": return Properties.Resources.btnAlignMiddle;
                case "b_AlignTop": return Properties.Resources.btnAlignTop;
                case "b_AlignButton": return Properties.Resources.btnAlignBottom;
                case "b_AlignCenter": return Properties.Resources.btnAlignCenter;
                case "b_DistributeHorizonatal": return Properties.Resources.btnAlignHorizontal;
                case "b_DistributeVertical": return Properties.Resources.btnAlignVertical;
                case "b_SnapTogether": return Properties.Resources.btnAlignTogether;
                case "b_CycleSwap": return Properties.Resources.btnCycleSwap;
                case "b_StackSort": return Properties.Resources.b_StackSort;
                case "b_MaximizePlot": return Properties.Resources.btnMaximize;
                case "b_CenterPie": return Properties.Resources.b_CenterPie;
                case "b_YMinZero": return Properties.Resources.b_YMinZero;
                case "b_HideYAxis": return Properties.Resources.b_HideYAxis;
                case "b_KeyboardControl": return Properties.Resources.b_KeyboardControl;
                case "b_ResetLabels": return Properties.Resources.b_ResetLabels;
                case "b_KeyPoints": return Properties.Resources.b_KeyPoints;
                case "b_SuperScript": return Properties.Resources.b_SuperScript;
                case "b_insertArrowCAGR": return Properties.Resources.b_insertArrowCAGR;
                case "b_insertElbowCAGR": return Properties.Resources.b_insertElbowCAGR;
                case "b_AddStackedTotals": return Properties.Resources.b_AddStackedTotals;
                case "b_AddStatLine": return Properties.Resources.b_AddStatLine;
                case "b_AddMetricBubbles": return Properties.Resources.b_AddMetricBubbles;
                case "b_DuplicateToSheets": return Properties.Resources.b_DuplicateToSheets;
                case "b_OffsetChart": return Properties.Resources.b_DuplicateToRows;
                case "b_ChartStackedWaterfall": return Properties.Resources.b_ChartWaterfall;
                case "b_ChartMarimekko": return Properties.Resources.b_ChartMarimekko;
                case "b_ChartGantt": return Properties.Resources.b_ChartGantt;
                case "b_ChartBubble": return Properties.Resources.b_ChartBubble;
                case "b_vectorChartImport": return Properties.Resources.b_vectorChartImport;
                case "b_pictureChartImport": return Properties.Resources.b_pictureChartImport;
                case "b_chartChartImport": return Properties.Resources.b_chartChartImport;
                case "b_pictureRangeImport": return Properties.Resources.b_pictureChartImport;
                case "b_tableRangeImport": return Properties.Resources.b_chartChartImport;
                case "btnImportText": return Properties.Resources.btnImportText;
                case "b_Show_Explorer": return Properties.Resources.b_Show_Explorer;
                case "b_Copy_Path": return Properties.Resources.b_Copy_Path;
                case "b_hidePageBreaks": return Properties.Resources.b_hidePageBreaks;
                default: return null;
            }
        }

        public string GetLabel(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Tab
                case "tabPragmatics": return "Pragmatics";
                // Tracing
                case "btnTracePrecedents": return "Trace Precedents";
                case "btnTraceDependents": return "Trace Dependents";
                case "btnRevertTrace": return "Revert Backwards";
                case "btnDrillPrecedents": return "Precedent Arrows";
                case "btnRemoveAllTracePrecedentArrows": return "Remove Arrows";
                case "btnDrillDependents": return "Dependent Arrows";
                case "btnRemoveAllTraceDependentArrows": return "Remove Arrows";
                // Formula Tools
                case "galFormulaTools": return "Formula Tools";
                case "b_Magnitude_Plus": return "Increase Magnitude";
                case "b_Magnitude_Minus": return "Decrease Magnitude";
                case "b_FlipSign": return "Flip Formula Sign ";
                case "b_FillRight": return "Smart Fill Right";
                case "b_FillDown": return "Smart Fill Down";
                case "b_ErrorWrapNA": return "IFERROR(NA)";
                case "b_ErrorWrapZero": return "IFERROR(0)";
                case "b_ConvertToLet": return "LET Optimization";
                case "b_FormulaWrap": return "Formula Wrap";
                case "b_SolveFormula": return "Simplify Formula";
                case "b_AnchorCycle": return "Anchor Cycle";
                // Paste Tools
                case "galPasteTools": return "Paste Tools";
                case "b_pasteExact": return "Exact Formulas";
                case "b_pasteDuplicate": return "Smart Exact";
                case "b_PasteNumberFormat": return "Number Format";
                case "b_pasteFillColor": return "Fill Color";
                case "b_pasteBorders": return "Range Border";
                case "b_pasteCellWidths": return "Column Width";
                case "b_pasteCellHeights": return "Row Height";
                // Formatting
                case "btnBrandFormat": return "Company Format";
                case "btnChartBrandFormat": return "Apply Chart Format";
                case "btnTableBrandFormat": return "Apply Table Format";
                case "btnBrandFillCycle": return "Fill Color Cycle";
                case "btnBrandBorderCycle": return "Border Color Cycle";
                case "btnBrandFontColorCycle": return "Font Color Cycle";

                case "galNumberFormats": return "Number Formats";
                case "b_NumberCycle": return "Number Cycle";
                case "b_CurrencyCycle": return "Currency Cycle";
                case "b_PercentCycle": return "Percent Cycle";
                case "b_DateCycle": return "Date Cycle";
                case "b_MultipleCycle": return "Multiple Cycle";
                case "b_BinaryCycle": return "Binary Cycle";
                case "b_PlusDecimals": return "Decimals (+)";
                case "b_MinusDecimals": return "Decimals (–)";

                case "galFontColors": return "Font Colors";
                case "b_BBColorCycle": return "Blue Black Cycle";
                case "b_ModelColorCycle": return "Model Color Cycle";
                case "b_AutoColorSelection": return "Auto Color Selection";

                case "galFillColors": return "Fill Colors";
                case "b_AutoTableColor": return "Auto-Table Color";
                case "b_ModelFillColorCycle": return "Model Color Cycle";
                case "b_FillThemeColorCycle": return "Theme Color Cycle";

                case "galOtherCycles": return "Other Cycles";
                case "b_UnderlineCycle": return "Underline Cycle";
                case "b_BorderCycle": return "Border Cycle";
                case "b_FontSizeCycle": return "Font Size Cycle";
                case "b_CellHeightCycle": return "Row Height Cycle";
                case "b_CellWidthCycle": return "Column Width Cycle";
                case "b_FootnoteCycle": return "Footnote Cycle";
                case "b_RemoveFootnotes": return "Remove Footnotes";
                case "galFormatTools": return "Format Tools";
                case "b_ListBullets": return "List Bullets";
                case "b_AccountingDots": return "Accounting Dots";
                case "b_SumLine": return "Sum Line";
                case "b_LeftIndent": return "Left Indent";
                case "b_RightIndent": return "Right Indent";
                case "b_TransparentToWhite": return "Transparent → White";
                case "galRangeTools": return "Range Tools";
                case "b_MergeToCenterAcrossSelection": return "Merge‑Center Across";
                case "b_hideToGroup": return "Hide → Group";
                case "b_reverseColumns": return "Reverse Columns";
                case "b_reverseRows": return "Reverse Rows";
                case "b_deleteBlankColumns": return "Delete Blank Cols";
                case "b_deleteBlankRows": return "Delete Blank Rows";
                case "b_harcodeSelection": return "Hardcode Selection";
                case "b_calculateCAGR": return "Calculate CAGR";
                case "galSizingTools": return "Sizing Tools";
                case "b_Guides": return "PPT Guides";
                case "b_ResizetoPPT": return "Resize to PPT";
                case "b_AutoGrid": return "Auto Grid";
                case "b_SameWidth": return "Same Width";
                case "b_SameHeight": return "Same Height";
                case "b_SameWidthandHeight": return "Same Width and Height";
                case "b_ExtendLeft": return "Extend Left";
                case "b_ExtendRight": return "Extend Right";
                case "b_ExtendTop": return "Extend Top";
                case "b_ExtendBottom": return "Extend Bottom";
                case "b_FitHorizontal": return "Fit Horizontal";
                case "b_FitVertical": return "Fit Vertical";
                case "galShapeAlignment": return "Shape Alignment";
                case "b_AlignLeft": return "Align Left";
                case "b_AlignRight": return "Align Right";
                case "b_AlignMiddle": return "Align Middle";
                case "b_AlignTop": return "Align Top";
                case "b_AlignButton": return "Align Bottom";
                case "b_AlignCenter": return "Align Center";
                case "b_DistributeHorizonatal": return "Distribute Horizontal";
                case "b_DistributeVertical": return "Distribute Vertical";
                case "b_SnapTogether": return "Snap Together";
                case "b_CycleSwap": return "Cycle Swap";
                case "btnInsertSymbols": return "Insert Symbols";
                // Sheet Tools
                case "btnSheetSelect": return "Sheet Select";
                case "btnDeepSearch": return "Deep Search";
                case "btnVisualCheck": return "Visual Check";
                case "btnCleanFile": return "Clean File";
                // Library
                case "btnLibrary": return "Library Assets";
                // Charting
                case "galFormatToolsChart": return "Format Tools";
                case "b_StackSort": return "Sort Stacked Columns";
                case "b_MaximizePlot": return "Maximize Plot Area";
                case "b_CenterPie": return "Center Pie Chart";
                case "b_YMinZero": return "Set Y‑Min to 0";
                case "b_HideYAxis": return "Hide Chart Y‑Axis";
                case "galLabelTools": return "Label Tools";
                case "b_KeyboardControl": return "Key Label Control";
                case "b_ResetLabels": return "Reset Labels";
                case "b_KeyPoints": return "Insert Points";
                case "b_SuperScript": return "Superscript";
                case "galInsertMetrics": return "Insert Metrics";
                case "b_insertArrowCAGR": return "Arrow CAGR";
                case "b_insertElbowCAGR": return "Elbow CAGR";
                case "b_AddStackedTotals": return "Stacked Total";
                case "b_AddStatLine": return "Stat Line";
                case "b_AddMetricBubbles": return "Metric Bubbles";
                case "galDuplicateChart": return "Duplicate";
                case "b_DuplicateToSheets": return "Copy Across Sheets";
                case "b_OffsetChart": return "Copy Across Offset";
                case "galCommonCharts": return "Common Charts";
                case "b_ChartStackedWaterfall": return "Stacked Waterfall";
                case "b_ChartMarimekko": return "Marimekko";
                case "b_ChartGantt": return "Gantt";
                case "b_ChartBubble": return "Bubble";
                // File
                case "btnVersionUp": return "Version Up";
                case "btnSaveAsPDF": return "Save as PDF";
                case "btnExportSheet": return "Export Sheet";
                case "galFileTools": return "File Tools";
                case "b_Show_Explorer": return "Show in Explorer";
                case "b_Copy_Path": return "Copy File Path";
                case "b_hidePageBreaks": return "Hide Page Breaks";
                case "btnSettings": return "Settings";
                default: return string.Empty;
            }
        }

        public string GetKeytip(IRibbonControl control)
        {
            switch (control.Id)
            {
                // Tab
                case "tabPragmatics": return "X";
                
                // Tracing
                case "btnTracePrecedents": return "TP";
                case "sbTraceDependents": return "TD";
                case "sbPrecedentArrows": return "AP";
                case "sbDependentArrows": return "AD";
                
                // Formula Tools
                case "galFormulaTools": return "F";
                
                // Paste Tools
                case "galPasteTools": return "P";
                
                // Formatting
                case "splitBtnBrandFormat": return "C";
                case "galNumberFormats": return "N";
                case "galFontColors": return "B";
                case "galFillColors": return "H";
                case "galOtherCycles": return "O";
                case "galFormatTools": return "M";
                case "galRangeTools": return "R";
                
                // Sizing & Alignment
                case "galSizingTools": return "Z";
                case "galShapeAlignment": return "L";
                case "btnInsertSymbols": return "I";
                
                // Library
                case "btnLibrary": return "0";
                
                // Sheet Tools
                case "btnSheetSelect": return "X";
                case "btnDeepSearch": return "Q";
                case "btnVisualCheck": return "U";
                case "btnCleanFile": return "K";
                
                // Charting
                case "galFormatToolsChart": return "W";
                case "galLabelTools": return "V";
                case "galInsertMetrics": return "G";
                case "galDuplicateChart": return "J";
                case "galCommonCharts": return "E";
                
                // File
                case "btnVersionUp": return "[";
                case "btnSaveAsPDF": return "]";
                case "btnExportSheet": return "\\";
                case "galFileTools": return "=";
                case "btnSettings": return "'";
                
                default: return string.Empty;
            }
        }

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
            string baseSupertip;

            switch (control.Id)
            {
                case "btnTracePrecedents": baseSupertip = "Displays all precedent cells and calculations that flow into the selected formula, with step through and drill down functionality across ranges, sheets, and workbooks. Enables enhanced model auditability beyond Excel's native single cell tracing functionality."; break;
                case "btnTraceDependents": baseSupertip = "Displays all dependent formulas on the selected cell. Enables enhanced model auditability beyond Excel's native single cell tracing functionality."; break;
                case "btnRevertTrace": baseSupertip = "Goes back to the selected cell subsequent to any tracing through Office Pragmatics."; break;
                case "btnDrillPrecedents": baseSupertip = "Draws precedent arrows for all precedent cells for the current selection, providing enhanced visibility beyond Excel's standard single arrow tracing."; break;
                case "btnRemoveAllTracePrecedentArrows": baseSupertip = "Clears all precedent and dependent arrows from the active worksheet."; break;
                case "btnDrillDependents": baseSupertip = "Draws dependent arrows for all dependent cells for the current selection, providing enhanced visibility beyond Excel's standard single arrow tracing."; break;
                case "btnRemoveAllTraceDependentArrows": baseSupertip = "Clears all precedent and dependent arrows from the active worksheet."; break;
                case "b_Magnitude_Plus": baseSupertip = "Increases the selected cell by three orders of magnitude (i.e., multiplies the range values/formulas by 1,000). Useful for changing the scale of data quickly."; break;
                case "b_Magnitude_Minus": baseSupertip = "Decreases the selected cell by three orders of magnitude (i.e., multiplies the range values/formulas by 0.001). Useful for changing the scale of data quickly."; break;
                case "b_FlipSign": baseSupertip = "Flips the outputs sign of the current selection (i.e, multiplies the range values/formulas by -1). Positive outputs will become negative and negative outputs will become positive in the selected range."; break;
                case "b_FillRight": baseSupertip = "Takes the selected cell's formula, and intelligently copies that cell to the right based on relative worksheet context and expected user intent."; break;
                case "b_FillDown": baseSupertip = "Takes the selected cell's formula, and intelligently copies that cell down based on relative worksheet context and expected user intent."; break;
                case "b_ErrorWrapNA": baseSupertip = "Wraps the selected cell(s) inside IFERROR function(s) to return NA on error as text, rather than #N/A"; break;
                case "b_ErrorWrapZero": baseSupertip = "Wraps the selected cell(s) inside IFERROR function(s) to return the value 0 on error, rather than #N/A"; break;
                case "b_ConvertToLet": baseSupertip = "Converts complex formulas with repetitive cell references to use the LET function for better readability and performance with intermediate variables."; break;
                case "b_FormulaWrap": baseSupertip = "Adds specified text to front and back of all formulas in the selected range. Enables insertion of a new wrapping formula without multi-step find and replace workarounds."; break;
                case "b_SolveFormula": baseSupertip = "Replaces select formulas with direct links to the resulting cells for auditing and analysis. Built for XLOOKUP, VLOOKUP, HLOOKUP, INDEX, MATCH, MIN, MAX, CHOOSE, TRANSPOSE, SUMIF, SUMIFS, INDIRECT, OFFSET"; break;
                case "b_AnchorCycle": baseSupertip = "Cycles through reference types (relative, absolute, mixed) for cell references in formulas. Quickly change A1 to $A$1 to $A1 to A$1."; break;
                case "b_pasteExact": baseSupertip = "Pastes formulas exactly as copied without adjusting cell references. Maintains absolute positioning of all references."; break;
                case "b_pasteDuplicate": baseSupertip = "Intelligently paste formulas with exact references where appropriate and relative references where needed."; break;
                case "b_PasteNumberFormat": baseSupertip = "Paste only the number formatting from copied cells, preserving existing values, formulas and other formatting in the destination."; break;
                case "b_pasteFillColor": baseSupertip = "Paste only the background fill from copied cells, preserving formulas and all other formatting."; break;
                case "b_pasteBorders": baseSupertip = "Paste only the cell broders from copied cells, preserving formulas and all other formatting."; break;
                case "b_pasteCellWidths": baseSupertip = "Paste column widths from copied cells."; break;
                case "b_pasteCellHeights": baseSupertip = "Paste row heights from copied cells."; break;
                case "btnBrandFormat": baseSupertip = "Automatically applies brand formatting to selected cell ranges or charts."; break;
                case "btnChartBrandFormat": baseSupertip = "Applies the default company brand format to a selected chart. Use the import function within Office Pragmatics settings on the Charts tab to set the active chart as the default company format."; break;
                case "btnTableBrandFormat": baseSupertip = "Applies the default company brand format to a selected cell range. Use the import function within Office Pragmatics settings on the Tables tab to set an active range as the default company format."; break;
                case "btnBrandFillCycle": baseSupertip = "Cycles through the company's primary colors and applies to the selected range's fill color. Company colors can be set in the Office Pragmatics settings menu under the Colors tab."; break;
                case "btnBrandBorderCycle": baseSupertip = "Cycles through the company's primary colors and applies to the selected range's border color. Company colors can be set in the Office Pragmatics settings menu under the Colors tab."; break;
                case "btnBrandFontColorCycle": baseSupertip = "Cycles through the company's primary colors and applies to the selected range's font color. Company colors can be set in the Office Pragmatics settings menu under the Colors tab."; break;
                case "b_NumberCycle": baseSupertip = "Cycles through common basic number formats used in financial modeling and presentation for the selected cell range."; break;
                case "b_CurrencyCycle": baseSupertip = "Cycles through common currency formats (e.g., $, €, £, ¥) for the selected cell range. Useful for quickly applying localized currency formatting."; break;
                case "b_PercentCycle": baseSupertip = "Cycles through common percentage formats (e.g., 0%, 0.0%, 0.00%) for the selected cell range. Ideal for presenting data as proportions or growth rates."; break;
                case "b_DateCycle": baseSupertip = "Cycles through common date formats (e.g., m/d/yyyy, mm-yy, mmmm d, yyyy) for the selected cell range. Simplifies the process of standardizing date presentations."; break;
                case "b_MultipleCycle": baseSupertip = "Cycles through formats that append 'x' to numbers (e.g., 1x, 1.0x, 1.00x) for the selected cell range. Useful for displaying multiples in financial models."; break;
                case "b_BinaryCycle": baseSupertip = "Cycles through binary formats (e.g., Yes/No, True/False, Y/N, On/Off) for the selected cell range. Ideal for representing binary data in a clear and concise manner."; break;
                case "b_PlusDecimals": baseSupertip = "Increases the number of decimal places displayed in the selected cell range. Useful for improving precision in numerical data."; break;
                case "b_MinusDecimals": baseSupertip = "Decreases the number of decimal places displayed in the selected cell range. Useful for simplifying numerical data presentation."; break;
                case "b_BBColorCycle": baseSupertip = "Cycles through the two primary font colors used in financial modeling: blue (inputs), black (calculations)."; break;
                case "b_ModelColorCycle": baseSupertip = "Cycles through standard font colors used in financial modeling: blue (inputs), black (calculations), green (references to other worksheets), red (references to other workbooks), purple (links to other data providers)."; break;
                case "b_AutoColorSelection": baseSupertip = "Automatically applies appropriate standard financial modeling font colors based on cell content type: blue (inputs), black (calculations), green (references to other worksheets), red (references to other workbooks), purple (links to other data providers)."; break;
                case "b_AutoTableColor": baseSupertip = "Automatically applies alternating row colors for a selected range to create a professional table appearance."; break;

                // Fill color cycles
                case "b_ModelFillColorCycle": baseSupertip = "Cycles through standard model fill colors (pale yellow for inputs, light blue, light red, light gray, white, then transparent) and applies the selection's interior accordingly. Useful for quickly highlighting inputs and structure consistently across a model."; break;
                case "b_FillThemeColorCycle": baseSupertip = "Cycles through the workbook's theme accent colors (Accent1–Accent6 and Light1/white), then transparent. Applies theme-consistent fills so selection colors remain consistent when the workbook theme changes."; break;

                case "b_UnderlineCycle": baseSupertip = "Cycles through common underline styles (single, single accounting, double, double accounting, none) for the selected cell range. Useful for applying professional accounting and financial presentation underlines."; break;
                case "b_BorderCycle": baseSupertip = "Cycles through common border patterns (bottom, top, left, right, all borders, no borders) for the selected cell range. Displays a visual indicator showing which border style was applied."; break;
                case "b_FontSizeCycle": baseSupertip = "Cycles through common font sizes (10pt, 18pt, 7pt) for the selected cell range. Useful for quickly adjusting text prominence in financial models and presentations."; break;
                case "b_CellHeightCycle": baseSupertip = "Cycles through common row heights (5pt for thin dividers, 20pt for content rows, reset to default) for the selected rows. Simplifies standardizing row spacing across worksheets."; break;
                case "b_CellWidthCycle": baseSupertip = "Cycles through common column widths (12pt standard, 2pt thin spacers, 25pt wide content, reset to default) for the selected columns. Enables quick column sizing for consistent model layouts."; break;
                case "b_FootnoteCycle": baseSupertip = "Adds or increments superscript footnote numbers (1), (2), (3)... at the end of cell text in the selected range. Useful for adding sequential footnote references to financial statements and tables."; break;
                case "b_RemoveFootnotes": baseSupertip = "Removes all superscript footnote numbers from the end of cell text in the selected range. Cleans up footnote references while preserving the main cell content."; break;

                case "b_ListBullets": baseSupertip = "Cycles through list formats (dash, bullet, number, letter, roman numeral, none) for the selected single-column range. Automatically skips blank cells and maintains sequential numbering."; break;
                case "b_AccountingDots": baseSupertip = "Applies or removes accounting dots leader format (@*.) to the selected cell range. Useful for creating professional financial statement layouts with dot leaders."; break;
                case "b_SumLine": baseSupertip = "Toggles single accounting underline formatting with thin row height (5pt) for the selected rows. Ideal for creating professional sum lines in financial statements and tables."; break;
                case "b_LeftIndent": baseSupertip = "Decreases the indent level of text in the selected cell range by one level. Helps with hierarchical formatting of financial statements and nested lists."; break;
                case "b_RightIndent": baseSupertip = "Increases the indent level of text in the selected cell range by one level (max 15 levels). Useful for creating visual hierarchy in financial models and reports."; break;
                case "b_TransparentToWhite": baseSupertip = "Converts all transparent (no fill) cell backgrounds to solid white fill in the selected range or used range. Ensures consistent white backgrounds across the worksheet."; break;
                case "b_MergeToCenterAcrossSelection": baseSupertip = "Converts merged cells to Center Across Selection alignment, which preserves individual cell structures while maintaining centered appearance. Prevents common Excel model issues caused by traditional cell merging."; break;
                case "b_hideToGroup": baseSupertip = "Converts all hidden rows and columns on the active worksheet into collapsed outline groups. Provides better visibility and control over hidden content in financial models."; break;
                case "b_reverseColumns": baseSupertip = "Reverses the order of columns in the selected range (left-to-right becomes right-to-left). Includes progress indicator and Esc key support for large selections."; break;
                case "b_reverseRows": baseSupertip = "Reverses the order of rows in the selected range (top-to-bottom becomes bottom-to-top). Includes progress indicator and Esc key support for large selections."; break;
                case "b_deleteBlankColumns": baseSupertip = "Deletes all completely blank columns within the selected range. Helps clean up models and reduce file size by removing unnecessary empty columns."; break;

                case "b_deleteBlankRows": baseSupertip = "Deletes all completely blank rows within the selected range. Helps clean up models and reduce file size by removing unnecessary empty rows."; break;
                case "b_harcodeSelection": baseSupertip = "Converts all formulas in the selected range (or used range if no selection) to their calculated values. Permanently replaces formulas with static values for finalizing models for external distribution."; break;
                case "b_calculateCAGR": baseSupertip = "Inserts a Compound Annual Growth Rate (CAGR) formula between two selected cells. Automatically calculates the annualized growth rate over the time period, ideal for financial analysis and trend modeling."; break;
                case "b_Guides": baseSupertip = "Inserts PowerPoint-style alignment guides (thin reference lines) on the worksheet. Useful for precise positioning and alignment of shapes, charts, and objects for professional presentation layouts."; break;
                case "b_ResizetoPPT": baseSupertip = "Resizes the selected shape or chart to standard PowerPoint slide dimensions. Ensures objects are properly sized for seamless transition between Excel and PowerPoint presentations."; break;
                case "b_AutoGrid": baseSupertip = "Snaps all selected shapes to the nearest cell grid boundaries within the current page break area. Aligns objects with worksheet cells for clean, structured layouts that maintain consistency across the model."; break;
                case "b_SameWidth": baseSupertip = "Makes all selected shapes the same width as the last-selected (anchor) shape. Accounts for border thickness to ensure visual consistency across multiple objects."; break;
                case "b_SameHeight": baseSupertip = "Makes all selected shapes the same height as the last-selected (anchor) shape. Accounts for border thickness to ensure visual consistency across multiple objects."; break;
                case "b_SameWidthandHeight": baseSupertip = "Makes all selected shapes the same width and height as the last-selected (anchor) shape. Creates uniformly sized objects while accounting for border thickness."; break;
                case "b_ExtendLeft": baseSupertip = "Stretches selected shapes leftward to align with either the page edge (single shape) or the left edge of the anchor shape (multiple shapes). Maintains shape positioning while extending width."; break;
                case "b_ExtendRight": baseSupertip = "Stretches selected shapes rightward to align with either the page edge (single shape) or the right edge of the anchor shape (multiple shapes). Maintains shape positioning while extending width."; break;
                case "b_ExtendTop": baseSupertip = "Stretches selected shapes upward to align with either the page edge (single shape) or the top edge of the anchor shape (multiple shapes). Maintains shape positioning while extending height."; break;
                case "b_ExtendBottom": baseSupertip = "Stretches selected shapes downward to align with either the page edge (single shape) or the bottom edge of the anchor shape (multiple shapes). Maintains shape positioning while extending height."; break;
                case "b_FitHorizontal": baseSupertip = "Distributes shapes to fit horizontally within available space. For a single shape, fits to worksheet width with user-defined margins. For multiple shapes, fits within anchor shape boundaries with specified gaps."; break;
                case "b_FitVertical": baseSupertip = "Distributes shapes to fit vertically within available space. For a single shape, fits to worksheet height with user-defined margins. For multiple shapes, fits within anchor shape boundaries with specified gaps."; break;
                case "b_AlignLeft": baseSupertip = "Aligns selected shapes to the left edge of the page (single shape) or to the left edge of the anchor shape (multiple shapes). Accounts for border thickness when alignment to border setting is enabled."; break;
                case "b_AlignRight": baseSupertip = "Aligns selected shapes to the right edge of the page (single shape) or to the right edge of the anchor shape (multiple shapes). Accounts for border thickness when alignment to border setting is enabled."; break;
                case "b_AlignMiddle": baseSupertip = "Centers selected shapes vertically within the page (single shape) or relative to the vertical center of the anchor shape (multiple shapes). Useful for creating balanced vertical layouts."; break;
                case "b_AlignTop": baseSupertip = "Aligns selected shapes to the top edge of the page (single shape) or to the top edge of the anchor shape (multiple shapes). Accounts for border thickness when alignment to border setting is enabled."; break;
                case "b_AlignButton": baseSupertip = "Aligns selected shapes to the bottom edge of the page (single shape) or to the bottom edge of the anchor shape (multiple shapes). Accounts for border thickness when alignment to border setting is enabled."; break;
                case "b_AlignCenter": baseSupertip = "Centers selected shapes horizontally within the page (single shape) or relative to the horizontal center of the anchor shape (multiple shapes). Useful for creating balanced horizontal layouts."; break;
                case "b_DistributeHorizonatal": baseSupertip = "Evenly distributes three or more selected shapes horizontally across the page break area. Creates equal spacing between objects for professional-looking aligned layouts."; break;
                case "b_DistributeVertical": baseSupertip = "Evenly distributes three or more selected shapes vertically across the page break area. Creates equal spacing between objects for professional-looking aligned layouts."; break;
                case "btnDeepSearch": baseSupertip = "Opens an advanced search tool that finds text, formulas, values, and formatting across the entire workbook. Provides enhanced search capabilities beyond Excel's native find function with filtering and navigation features."; break;

                case "btnVisualCheck": baseSupertip = "Applies visual formula consistency highlighting for the selected range using textured patterns to cells with matching formulas: horizontal stripes for row consistency, vertical stripes for column consistency, and crosshatch for both directions."; break;
                case "btnCleanFile": baseSupertip = "Performs comprehensive workbook cleanup including removing unused/broken defined names, deleting unused custom number formats, clearing unused rows and columns, and resetting to default theme. Creates automatic backup before cleaning."; break;
                case "btnLibrary": baseSupertip = "Opens the Library Assets manager to access and insert saved sheets, workbooks, and other reusable components into the active workbook."; break;
                case "btnRefresh": baseSupertip = "Refreshes all data connections and linked objects in the active workbook."; break;
                case "b_StackSort": baseSupertip = "Sorts the series in a stacked column chart by size for better visual clarity. Reorders data series from largest to smallest values for improved chart readability."; break;
                case "b_MaximizePlot": baseSupertip = "Expands the plot area to fill the maximum available space within the selected chart object. Reduces margins and maximizes data visualization area."; break;
                case "b_CenterPie": baseSupertip = "Centers the plot area within a pie chart frame and optimizes the radial layout. Ensures the pie is properly positioned for balanced presentation."; break;
                case "b_YMinZero": baseSupertip = "Sets the Y-axis minimum value to zero on the selected chart. Ensures charts start at zero for accurate visual representation of relative values."; break;
                case "b_HideYAxis": baseSupertip = "Hides the Y-axis on the selected chart while preserving the chart data and formatting. Useful for creating cleaner, presentation-ready visualizations."; break;
                case "b_KeyboardControl": baseSupertip = "Enables keyboard-based control for positioning and formatting chart data labels. Use arrow keys to move labels and shortcuts to apply formatting changes quickly."; break;
                case "b_ResetLabels": baseSupertip = "Resets all data labels in the selected chart to their default positions and formatting. Removes custom positioning and reverts labels to Excel's automatic placement."; break;
                case "b_KeyPoints": baseSupertip = "Adds data labels to selected key data points in the chart. Useful for highlighting specific values without cluttering the chart with all labels."; break;
                case "b_SuperScript": baseSupertip = "Converts footnote references in chart labels to superscript format. Automatically detects and formats parenthetical numbers as superscript for professional footnote styling."; break;
                case "b_insertArrowCAGR": baseSupertip = "Inserts a straight arrow connector showing the Compound Annual Growth Rate (CAGR) between two selected chart data points. Automatically calculates and displays the growth rate percentage."; break;
                case "b_insertElbowCAGR": baseSupertip = "Inserts an elbow-style connector showing the Compound Annual Growth Rate (CAGR) between two selected chart data points. Automatically calculates and displays the growth rate percentage with a right-angle connector."; break;
                case "b_AddStackedTotals": baseSupertip = "Adds data labels showing the total values at the top of each stacked column. Displays the sum of all series values for each category in a stacked chart."; break;
                case "b_AddStatLine": baseSupertip = "Inserts a statistical reference line (mean, median, or target) on the selected chart. Helps visualize central tendency or benchmark values against chart data."; break;
                case "b_AddMetricBubbles": baseSupertip = "Inserts metric bubbles (callout shapes) displaying calculated statistics or custom text annotations on the selected chart. Useful for highlighting key insights and data points."; break;
                case "b_DuplicateToSheets": baseSupertip = "Duplicates the selected chart across multiple worksheets in the workbook. Prompts for sheet selection and maintains chart formatting while updating data source references."; break;
                case "b_OffsetChart": baseSupertip = "Duplicates the selected chart with a specified offset in rows and columns. Creates copies of charts positioned consistently for dashboard layouts or multi-chart presentations."; break;
                case "galCommonCharts": baseSupertip = "Provides quick access to commonly used custom chart types including stacked waterfall, Marimekko, Gantt, and bubble charts."; break;
                case "b_ChartStackedWaterfall": baseSupertip = "Creates a stacked waterfall chart showing cumulative effects of sequential positive and negative values. Useful for variance analysis and bridge charts in financial reporting."; break;
                case "b_ChartMarimekko": baseSupertip = "Creates a Marimekko (mosaic) chart with variable-width columns for displaying two-dimensional categorical data. Both width and height represent data values."; break;
                case "b_ChartGantt": baseSupertip = "Creates a Gantt chart from selected data showing project timelines and task durations. Useful for project management and schedule visualization."; break;
                case "b_ChartBubble": baseSupertip = "Creates an advanced bubble chart with three or four dimensions of data. Bubble size and position represent different variables for multi-dimensional analysis."; break;

                case "btnVersionUp": baseSupertip = "Saves a new version of the workbook with an incremented version number in the filename (e.g., Model_v1.xlsx → Model_v2.xlsx). Preserves original file and creates versioned copy."; break;
                case "btnSaveAsPDF": baseSupertip = "Exports the active workbook or selected sheets to PDF format. Provides options for page setup, print area, and PDF quality settings."; break;
                case "btnExportSheet": baseSupertip = "Exports the active sheet as a separate workbook file. Creates a standalone Excel file containing only the current worksheet."; break;
                case "galFileTools": baseSupertip = "Provides quick access to file management utilities including Explorer integration and file path operations."; break;
                case "b_Show_Explorer": baseSupertip = "Opens Windows Explorer and selects the current workbook file. Quickly navigates to the file location in the file system."; break;
                case "b_Copy_Path": baseSupertip = "Copies the full file path of the active workbook to the clipboard. Useful for sharing file locations or creating hyperlinks."; break;
                case "b_hidePageBreaks": baseSupertip = "Toggles the visibility of page break lines in all worksheets. Hides or shows the dashed lines that indicate print page boundaries for cleaner on-screen viewing."; break;
                case "btnSettings": baseSupertip = "Opens the Office Pragmatics settings dialog to configure add-in preferences, custom colors, brand formats, and keyboard shortcuts."; break;
                case "btnSheetSelect": baseSupertip = "Opens the Sheet Selector dialog for quick navigation between worksheets. Provides search and filter capabilities for workbooks with many sheets."; break;

                default: baseSupertip = string.Empty; break;
            }

            // Append keyboard shortcut if available
            return AppendShortcutToSupertip(baseSupertip, control.Id);
        }

        private string AppendShortcutToSupertip(string baseSupertip, string buttonId)
        {
            if (string.IsNullOrEmpty(baseSupertip))
                return baseSupertip;

            try
            {
                // Get the method name for this button
                string methodName = AddInUtilities.GetMethodNameForButtonId(buttonId);

                if (!string.IsNullOrEmpty(methodName))
                {
                    // Get the current shortcut text (already formatted for display)
                    string shortcutText = AddInUtilities.GetShortcutText(methodName);

                    if (!string.IsNullOrEmpty(shortcutText))
                    {
                        // Append with two line breaks
                        return baseSupertip + "\n\nShortcut: " + shortcutText;
                    }
                }
            }
            catch
            {
                // If anything fails, just return the base supertip
            }

            return baseSupertip;
        }

        public void OnUniversalAction(IRibbonControl control)
        {

            try
            {
                Excel.Application xlApp = Globals.ThisAddIn.Application;
                Excel.ShapeRange sel = GetSelectedShapeRange(xlApp);
                Excel.Worksheet ws = xlApp.ActiveSheet as Excel.Worksheet;

                switch (control.Id)
                {
                    case "btnTracePrecedents": TracePrecedent.Run(Globals.ThisAddIn.Application); break;
                    case "btnTraceDependents": TraceDependent.Run(Globals.ThisAddIn.Application); break;
                    case "btnRevertTrace": RevertBackwards.Run(Globals.ThisAddIn.Application); break;
                    case "btnDrillPrecedents": PrecedentArrows.DrillDownPrecedents(); break;
                    case "btnRemoveAllTracePrecedentArrows": PrecedentArrows.RemoveAllTracerArrows(); break;
                    case "btnDrillDependents": DependentArrows.DrillDownDependents(); break;
                    case "btnRemoveAllTraceDependentArrows": DependentArrows.RemoveAllTracerArrows(); break;
                    case "b_Magnitude_Plus": FormulaTools.Magnitude_Plus(); break;
                    case "b_Magnitude_Minus": FormulaTools.Magnitude_Minus(); break;
                    case "b_FlipSign": FormulaTools.FlipSign(); break;
                    case "b_FillRight": AutoFill.FillRight(); break;
                    case "b_FillDown": AutoFill.FillDown(); break;
                    case "b_ErrorWrapNA": FormulaTools.ErrorWrapNA(); break;
                    case "b_ErrorWrapZero": FormulaTools.ErrorWrapZero(); break;
                    case "b_ConvertToLet": FormulaTools.ConvertToLet(); break;
                    case "b_FormulaWrap": FormulaTools.FormulaWrap(); break;
                    case "b_SolveFormula": FormulaTools.SolveFormula(); break;
                    case "b_AnchorCycle": FormulaTools.AnchorCycle(); break;
                    case "b_pasteExact": PasteTools.PasteExact(); break;
                    case "b_pasteDuplicate": PasteTools.PasteDuplicate(); break;
                    case "b_PasteNumberFormat": PasteTools.PasteNumberFormat(); break;
                    case "b_pasteFillColor": PasteTools.PasteFillColor(); break;
                    case "b_pasteBorders": PasteTools.PasteBorders(); break;
                    case "b_pasteCellWidths": PasteTools.pasteCellWidths(); break;
                    case "b_pasteCellHeights": PasteTools.pasteCellHeights(); break;
                    case "btnBrandFormat": AutoBrand.Run(); break;
                    case "btnTableBrandFormat": OfficePragmatics_Excel.Classes.Tables.TableFormatter.ApplyToSelection(OfficePragmatics_Excel.Settings.SettingsForm.CurrentSettings.TableFormat); break;
                    case "btnChartBrandFormat": Classes.Brand.ChartFormat.ApplyBrandFormattingToSelection(SettingsForm.CurrentSettings.BrandCharts); break;
                    case "b_NumberCycle": NumberFormats.RunNumberCycle(); break;
                    case "b_CurrencyCycle": NumberFormats.RunCurrencyCycle(); break;
                    case "b_PercentCycle": NumberFormats.RunPercentCycle(); break;
                    case "b_DateCycle": NumberFormats.RunDateCycle(); break;
                    case "b_MultipleCycle": NumberFormats.RunMultipleCycle(); break;
                    case "b_BinaryCycle": NumberFormats.RunBinaryCycle(); break;
                    case "b_PlusDecimals": NumberFormats.RunPlusDecimals(); break;
                    case "b_MinusDecimals": NumberFormats.RunMinusDecimals(); break;
                    case "b_BBColorCycle": FontColor.RunBBColorCycle(); break;
                    case "b_ModelColorCycle": FontColor.RunModelColorCycle(); break;
                    case "b_AutoColorSelection": FontColor.RunAutoColor(); break;
                    case "b_AutoTableColor": FillColors.RunTableFormattingCycle(); break;
                    case "b_ModelFillColorCycle": FillColors.RunFillCycle(); break;
                    case "b_FillThemeColorCycle": FillColors.RunThemeColorCycle(); break;
                    case "b_UnderlineCycle": OtherFormatting.RunUnderlineCycle(); break;
                    case "b_BorderCycle": OtherFormatting.RunBorderCycle(); break;
                    case "b_FontSizeCycle": OtherFormatting.RunFontSizeCycle(); break;
                    case "b_CellHeightCycle": OtherFormatting.RunRowHeightCycle(); break;
                    case "b_CellWidthCycle": OtherFormatting.RunColumnWidthCycle(); break;
                    case "b_FootnoteCycle": OtherFormatting.RunFootnoteCycle(); break;
                    case "b_RemoveFootnotes": OtherFormatting.RunFootnoteRemove(); break;
                    case "b_ListBullets": OtherFormatting.RunListBullets(); break;
                    case "b_AccountingDots": OtherFormatting.RunAccountingDots(); break;
                    case "b_SumLine": OtherFormatting.RunSumLine(); break;
                    case "b_LeftIndent": OtherFormatting.RunIndentDecrease(); break;
                    case "b_RightIndent": OtherFormatting.RunIndentIncrease(); break;
                    case "b_TransparentToWhite": FillColors.RunConvertTransparentToWhite(); break;
                    case "b_MergeToCenterAcrossSelection": CellRangeFormatting.mergeToCenterAcrossSelection(); break;
                    case "b_hideToGroup": CellRangeFormatting.hideToGroup(); break;
                    case "b_reverseColumns": CellRangeFormatting.ReverseColumns(); break;
                    case "b_reverseRows": CellRangeFormatting.ReverseRows(); break;
                    case "b_deleteBlankColumns": CellRangeFormatting.deleteBlankColumns(); break;
                    case "b_deleteBlankRows": CellRangeFormatting.deleteBlankRows(); break;
                    case "b_harcodeSelection": CellRangeFormatting.HardcodeFormulasInSelectionOrUsedRange(); break;
                    case "b_calculateCAGR": CellRangeFormatting.InsertCagrFormula(); break;
                    case "b_Guides": CellRangeFormatting.InsertGuide(); break;
                    case "b_ResizetoPPT": CellRangeFormatting.ResizeToPowerPoint(); break;
                    case "b_AutoGrid": AlignmentHelper.SnapSelectionToCellGrid(ws, sel); break;
                    case "b_SameWidth": SizeHelper.SameWidth(sel); break;
                    case "b_SameHeight": SizeHelper.SameHeight(sel); break;
                    case "b_SameWidthandHeight": SizeHelper.SameSize(sel); break;
                    case "b_ExtendLeft": SizeHelper.ExtendLeft(ws, sel); break;
                    case "b_ExtendRight": SizeHelper.ExtendRight(ws, sel); break;
                    case "b_ExtendTop": SizeHelper.ExtendTop(ws, sel); break;
                    case "b_ExtendBottom": SizeHelper.ExtendBottom(ws, sel); break;
                    case "b_FitHorizontal": SizeHelper.FitToHorizontal(sel); break;
                    case "b_FitVertical": SizeHelper.FitToVertical(sel); break;
                    case "b_AlignLeft": AlignmentHelper.AlignLeft(ws, sel); break;
                    case "b_AlignRight": AlignmentHelper.AlignRight(ws, sel); break;
                    case "b_AlignTop": AlignmentHelper.AlignTop(ws, sel); break;
                    case "b_AlignButton": AlignmentHelper.AlignBottom(ws, sel); break;
                    case "b_AlignCenter": AlignmentHelper.AlignCenter(ws, sel); break;
                    case "b_AlignMiddle": AlignmentHelper.AlignMiddle(ws, sel); break;
                    case "b_DistributeHorizonatal": AlignmentHelper.DistributeInPage(ws, sel, true); break;
                    case "b_DistributeVertical": AlignmentHelper.DistributeInPage(ws, sel, false); break;
                    case "b_SnapTogether": AlignmentHelper.SnapTogether(sel); break;
                    case "b_CycleSwap": AlignmentHelper.CycleSwap(sel); break;
                    case "btnSheetSelect": SheetSelector.Show(Globals.ThisAddIn.Application); break;
                    case "btnDeepSearch": SuperFind.Show(); break;
                    case "btnCleanFile": WorkbookCleaner.ComprehensiveClean(); break;
                    case "btnVisualCheck": VisualCheckHelper.ToggleVisualCheck(); break;
                    case "b_StackSort": FormatTools.StackSort(); break;
                    case "b_MaximizePlot": FormatTools.MaxPlotArea(); break;
                    case "b_CenterPie": FormatTools.CenterPlot(); break;
                    case "b_YMinZero": FormatTools.ZeroYAxis(); break;
                    case "b_HideYAxis": FormatTools.HideYAxis(); break;
                    case "b_KeyboardControl": LabelTools.KeyboardControlChartLabels(); break;
                    case "b_ResetLabels": LabelTools.ResetLabels(); break;
                    case "b_DuplicateToSheets": ChartDuplicator.DuplicateSelectedChart(); break;
                    case "b_OffsetChart": ChartOffsetDuplicator.DuplicateSelectedChartByOffset(); break;
                    case "b_AddMetricBubbles": StatBubbles.InsertMetricsBubbles(); break;
                    case "b_KeyPoints": LabelTools.LabelPoints(); break;
                    case "b_SuperScript": LabelTools.SuperscriptFootnotes(); break;
                    case "b_insertArrowCAGR": new CAGRArrow().insertArrowCAGR(); break;
                    case "b_insertElbowCAGR": new CAGRArrow().insertElbowCAGR(); break;
                    case "b_AddStackedTotals": StackedColumn.StackedTotal(); break;
                    case "b_AddStatLine": StatLines.AddMedianLine(); break;
                    case "b_ChartStackedWaterfall": WaterfallChartHelper.InsertStackedWaterfall(); break;
                    case "btnVersionUp": FileTools.VersionUp(Globals.ThisAddIn.Application); break;
                    case "btnSaveAsPDF": FileTools.SavePDF(Globals.ThisAddIn.Application); break;
                    case "btnExportSheet": ExportSelectionTool.ExportSelection(Globals.ThisAddIn.Application); break;
                    case "b_Show_Explorer": FileTools.OpenContainingFolder(); break;
                    case "b_Copy_Path": FileTools.CopyFilePathToClipboard(); break;
                    case "b_hidePageBreaks": FileTools.HidePageBreaks(Globals.ThisAddIn.Application); break;
                    case "btnSettings": SettingsForm.ShowSettingsDialog(); break;
                    case "btnLibrary": ExcelLibraryUtility.OpenLibrary(); break;

                    default:
                        MessageBox.Show($"Handler not implemented for: {control.Id}");
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error executing {control.Id}: {ex.Message}");
            }
        }

        private static Excel.ShapeRange GetSelectedShapeRange(Excel.Application xlApp)
        {
            try
            {
                if (xlApp.Selection is Excel.ShapeRange sr) return sr;
                if (xlApp.Selection is Excel.Shape sh) return sh.Parent.Shapes.Range(new object[] { sh.Name });
                dynamic dynSel = xlApp.Selection;
                try { return (Excel.ShapeRange)dynSel.ShapeRange; } catch { }
                return null;
            }
            catch { return null; }
        }
    }
}