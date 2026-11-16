using OfficePragmatics_Excel.Classes;
using OfficePragmatics_Excel.Classes.Charting;
using OfficePragmatics_Excel.Library;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class AddInUtilities
    {
        public static (string Key, string Method)[] _shortcuts =
        {
            ("^c",      nameof(RunCopyAndStore)),
            ("^+1",     nameof(RunNumberCycle)),
            ("^+3",     nameof(RunDateCycle)),
            ("^+4",     nameof(RunCurrencyCycle)),
            ("^+5",     nameof(RunPercentCycle)),
            ("^+8",     nameof(RunMultipleCycle)),
            ("^+9",     nameof(RunBinaryCycle)),
            ("^,",      nameof(RunPlusDecimals)),
            ("^.",      nameof(RunMinusDecimals)),
            ("^+,",     nameof(Magnitude_Plus)),
            ("^+.",     nameof(Magnitude_Minus)),
            ("^+{[}",   nameof(TracePrecedents)),
            ("^+{]}",   nameof(TraceDependents)),
            ("^+{F5}",  nameof(RevertBackward)),
            ("^+x",     nameof(SheetSelectorShow)),
            ("^;",      nameof(RunBBColorCycle)),
            ("^'",      nameof(RunModelColorCycle)),
            ("^+r",     nameof(FillRight)),
            ("^+d",     nameof(FillDown)),
            ("^+e",     nameof(ErrorWrapNA)),
            ("^+/",     nameof(FlipSign)),
            ("^%{[}",   nameof(DrillPrecedents)),
            ("^%+{[}",  nameof(RemoveAllTracePrecedentArrows)),
            ("^%{]}",   nameof(DrillDependents)),
            ("^%+{]}",  nameof(RemoveAllTraceDependentArrows)),
            ("^+w",     nameof(ErrorWrapZero)),
            ("^+l",     nameof(ConvertToLet)),
            ("^+f",     nameof(FormulaWrap)),
            ("^+s",     nameof(SolveFormula)),
            ("^+k",     nameof(AnchorCycle)),
            ("^%v",     nameof(PasteExact)),
            ("^%d",     nameof(PasteDuplicate)),
            ("^%n",     nameof(PasteNumberFormat)),
            ("^%c",     nameof(PasteFillColor)),
            ("^%b",     nameof(PasteBorders)),
            ("^%w",     nameof(PasteCellWidths)),
            ("^%h",     nameof(PasteCellHeights)),
            ("^+;",     nameof(RunAutoColor)),
            ("^%t",     nameof(RunTableFormattingCycle)),
            ("^%f",     nameof(RunFillCycle)),
            ("^%+t",    nameof(RunThemeColorCycle)),
            ("^%+c",    nameof(RunConvertTransparentToWhite)),
            ("^+u",     nameof(RunUnderlineCycle)),
            ("^+b",     nameof(RunBorderCycle)),
            ("^%+z",    nameof(RunFontSizeCycle)),
            ("^%+r",    nameof(RunRowHeightCycle)),
            ("^%+w",    nameof(RunColumnWidthCycle)),
            ("^%+n",    nameof(RunFootnoteCycle)),
            ("^%+x",    nameof(RunFootnoteRemove)),
            ("^%l",     nameof(RunListBullets)),
            ("^%a",     nameof(RunAccountingDots)),
            ("^%s",     nameof(RunSumLine)),
            ("^{[}",    nameof(RunIndentDecrease)),
            ("^{]}",    nameof(RunIndentIncrease)),
            ("^+m",     nameof(MergeToCenterAcrossSelection)),
            ("^+g",     nameof(HideToGroup)),
            ("^%+{LEFT}",  nameof(ReverseColumns)),
            ("^%+{UP}",    nameof(ReverseRows)),
            ("^%+y", nameof(DeleteBlankColumns)),
            ("^%+q",  nameof(DeleteBlankRows)),
            ("^+h",     nameof(HardcodeSelection)),
            ("^%+g",    nameof(CalculateCAGR)),
            ("^%g",     nameof(InsertGuides)),
            ("^%p",     nameof(ResizeToPPT)),
            ("^%+1",    nameof(StackSort)),
            ("^%+2",    nameof(MaximizePlot)),
            ("^%+3",    nameof(CenterPie)),
            ("^%+0",    nameof(YMinZero)),
            ("^%+9",    nameof(HideYAxis)),
            ("^%k",     nameof(KeyboardControl)),
            ("^%r",     nameof(ResetLabels)),
            ("^%+d",    nameof(DuplicateToSheets)),
            ("^%+o",    nameof(OffsetChart)),
            ("^%+m",    nameof(AddMetricBubbles)),
            ("^%+k",    nameof(KeyPoints)),
            ("^%+u",    nameof(SuperScript)),
            ("^%+7",    nameof(InsertArrowCAGR)),
            ("^%+8",    nameof(InsertElbowCAGR)),
            ("^%+5",    nameof(AddStackedTotals)),
            ("^%+l",    nameof(AddStatLine)),
            ("^%+4",    nameof(ChartStackedWaterfall)),
            ("^+{F12}", nameof(VersionUp)),
            ("^+{F11}", nameof(SaveAsPDF)),
            ("^+{F10}", nameof(ExportSheet)),
            ("^%o",     nameof(ShowExplorer)),
            ("^%+{BACKSPACE}", nameof(CopyPath)),
            ("^%+b",    nameof(HidePageBreaks)),
            ("^%+f",    nameof(DeepSearch)),
            ("^%+e", nameof(CleanFile)),
            ("^%+v",   nameof(VisualCheck)),
            ("^+j", nameof(ApplyTableBrandFormat)),
            ("^+n", nameof(ApplyChartBrandFormat)),
            ("^%+s",    nameof(ShowSettings)),
            ("^+{F9}", nameof(OpenLibrary)),
        };

        private static Dictionary<string, string> _shortcutLookup;
        private static bool _shortcutsCurrentlyRegistered = false;

        public static string GetShortcutText(string methodName)
        {
            if (_shortcutLookup == null)
                _shortcutLookup = _shortcuts.ToDictionary(s => s.Method, s => s.Key);

            return _shortcutLookup.TryGetValue(methodName, out var onKey)
                ? FormatOnKeyForDisplay(onKey)
                : string.Empty;
        }

        public static string FormatOnKeyForDisplay(string onKey)
        {
            if (string.IsNullOrEmpty(onKey)) return "";

            var parts = new List<string>();
            string key = onKey;

            if (key.Contains("^")) { parts.Add("Ctrl"); key = key.Replace("^", ""); }
            if (key.Contains("+")) { parts.Add("Shift"); key = key.Replace("+", ""); }
            if (key.Contains("%")) { parts.Add("Alt"); key = key.Replace("%", ""); }

            key = key.Replace("{[}", "[").Replace("{]}", "]")
                    .Replace("{LEFT}", "Left").Replace("{RIGHT}", "Right")
                    .Replace("{UP}", "Up").Replace("{DOWN}", "Down")
                    .Replace("{DELETE}", "Del").Replace("{BACKSPACE}", "Backspace")
                    .Replace("{ENTER}", "Enter").Replace("{TAB}", "Tab")
                    .Replace("{F1}", "F1").Replace("{F10}", "F10")
                    .Replace("{F11}", "F11").Replace("{F12}", "F12");

            if (key.StartsWith("{") && key.EndsWith("}"))
                parts.Add(key.Substring(1, key.Length - 2));
            else
                parts.Add(key.ToUpperInvariant());

            return string.Join(" + ", parts);
        }

        public void RegisterShortcuts()
        {
            _shortcutLookup = null;

            if (_shortcutsCurrentlyRegistered) return;

            var app = Globals.ThisAddIn.Application;

            if (app.Workbooks.Count == 0 || !app.Interactive) return;

            bool wasScreenUpdating = app.ScreenUpdating;
            Excel.XlCalculation wasCalculation = app.Calculation;
            bool wasEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.EnableEvents = false;

                foreach (var s in _shortcuts)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(s.Key))
                        {
                            app.OnKey(s.Key, $"'ExecuteAction \"{s.Method}\"'");
                        }
                    }
                    catch (Exception)
                    {
                    }
                }

                _shortcutsCurrentlyRegistered = true;
            }
            catch (Exception)
            {
            }
            finally
            {
                app.EnableEvents = wasEnableEvents;
                app.ScreenUpdating = wasScreenUpdating;
            }
        }

        public void UnregisterShortcuts()
        {
            var app = Globals.ThisAddIn.Application;

            bool wasScreenUpdating = app.ScreenUpdating;
            Excel.XlCalculation wasCalculation = app.Calculation;
            bool wasEnableEvents = app.EnableEvents;

            try
            {
                app.ScreenUpdating = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                foreach (var s in _shortcuts)
                {
                    try
                    {
                        if (!string.IsNullOrEmpty(s.Key))
                            app.OnKey(s.Key, Type.Missing);
                    }
                    catch
                    {
                    }
                }

                _shortcutsCurrentlyRegistered = false;
            }
            finally
            {
                app.ScreenUpdating = wasScreenUpdating;
                app.Calculation = wasCalculation;
                app.EnableEvents = wasEnableEvents;
            }
        }

        public void RunNumberCycle() => NumberFormats.RunNumberCycle();
        public void RunCurrencyCycle() => NumberFormats.RunCurrencyCycle();
        public void RunPercentCycle() => NumberFormats.RunPercentCycle();
        public void RunDateCycle() => NumberFormats.RunDateCycle();
        public void RunMultipleCycle() => NumberFormats.RunMultipleCycle();
        public void RunBinaryCycle() => NumberFormats.RunBinaryCycle();
        public void RunPlusDecimals() => NumberFormats.RunPlusDecimals();
        public void RunMinusDecimals() => NumberFormats.RunMinusDecimals();

        public void ErrorWrapNA() => FormulaTools.ErrorWrapNA();
        public void Magnitude_Plus() => FormulaTools.Magnitude_Plus();
        public void Magnitude_Minus() => FormulaTools.Magnitude_Minus();
        public void FlipSign() => FormulaTools.FlipSign();
        public void TracePrecedents() => TracePrecedent.Run(Globals.ThisAddIn.Application);
        public void TraceDependents() => TraceDependent.Run(Globals.ThisAddIn.Application);
        public void RevertBackward() => RevertBackwards.Run(Globals.ThisAddIn.Application);

        public void FillRight() => AutoFill.FillRight();
        public void FillDown() => AutoFill.FillDown();
        public void RunCopyAndStore() => PasteTools.RunCopyAndStore();

        public void RunBBColorCycle() => FontColor.RunBBColorCycle();
        public void RunModelColorCycle() => FontColor.RunModelColorCycle();

        public void SheetSelectorShow() => SheetSelector.Show(Globals.ThisAddIn.Application);

        public void DrillPrecedents() => PrecedentArrows.DrillDownPrecedents();
        public void RemoveAllTracePrecedentArrows() => PrecedentArrows.RemoveAllTracerArrows();
        public void DrillDependents() => DependentArrows.DrillDownDependents();
        public void RemoveAllTraceDependentArrows() => DependentArrows.RemoveAllTracerArrows();

        public void ErrorWrapZero() => FormulaTools.ErrorWrapZero();
        public void ConvertToLet() => FormulaTools.ConvertToLet();
        public void FormulaWrap() => FormulaTools.FormulaWrap();
        public void SolveFormula() => FormulaTools.SolveFormula();
        public void AnchorCycle() => FormulaTools.AnchorCycle();

        public void PasteExact() => PasteTools.PasteExact();
        public void PasteDuplicate() => PasteTools.PasteDuplicate();
        public void PasteNumberFormat() => PasteTools.PasteNumberFormat();
        public void PasteFillColor() => PasteTools.PasteFillColor();
        public void PasteBorders() => PasteTools.PasteBorders();
        public void PasteCellWidths() => PasteTools.pasteCellWidths();
        public void PasteCellHeights() => PasteTools.pasteCellHeights();

        public void RunAutoColor() => FontColor.RunAutoColor();
        public void RunTableFormattingCycle() => FillColors.RunTableFormattingCycle();
        public void RunFillCycle() => FillColors.RunFillCycle();
        public void RunThemeColorCycle() => FillColors.RunThemeColorCycle();
        public void RunConvertTransparentToWhite() => FillColors.RunConvertTransparentToWhite();

        public void RunUnderlineCycle() => OtherFormatting.RunUnderlineCycle();
        public void RunBorderCycle() => OtherFormatting.RunBorderCycle();
        public void RunFontSizeCycle() => OtherFormatting.RunFontSizeCycle();
        public void RunRowHeightCycle() => OtherFormatting.RunRowHeightCycle();
        public void RunColumnWidthCycle() => OtherFormatting.RunColumnWidthCycle();
        public void RunFootnoteCycle() => OtherFormatting.RunFootnoteCycle();
        public void RunFootnoteRemove() => OtherFormatting.RunFootnoteRemove();
        public void RunListBullets() => OtherFormatting.RunListBullets();
        public void RunAccountingDots() => OtherFormatting.RunAccountingDots();
        public void RunSumLine() => OtherFormatting.RunSumLine();
        public void RunIndentDecrease() => OtherFormatting.RunIndentDecrease();
        public void RunIndentIncrease() => OtherFormatting.RunIndentIncrease();

        public void MergeToCenterAcrossSelection() => CellRangeFormatting.mergeToCenterAcrossSelection();
        public void HideToGroup() => CellRangeFormatting.hideToGroup();
        public void ReverseColumns() => CellRangeFormatting.ReverseColumns();
        public void ReverseRows() => CellRangeFormatting.ReverseRows();
        public void DeleteBlankColumns() => CellRangeFormatting.deleteBlankColumns();
        public void DeleteBlankRows() => CellRangeFormatting.deleteBlankRows();
        public void HardcodeSelection() => CellRangeFormatting.HardcodeFormulasInSelectionOrUsedRange();
        public void CalculateCAGR() => CellRangeFormatting.InsertCagrFormula();
        public void InsertGuides() => CellRangeFormatting.InsertGuide();
        public void ResizeToPPT() => CellRangeFormatting.ResizeToPowerPoint();

        public void StackSort() => FormatTools.StackSort();
        public void MaximizePlot() => FormatTools.MaxPlotArea();
        public void CenterPie() => FormatTools.CenterPlot();
        public void YMinZero() => FormatTools.ZeroYAxis();
        public void HideYAxis() => FormatTools.HideYAxis();
        public void KeyboardControl() => LabelTools.KeyboardControlChartLabels();
        public void ResetLabels() => LabelTools.ResetLabels();
        public void DuplicateToSheets() => ChartDuplicator.DuplicateSelectedChart();
        public void OffsetChart() => ChartOffsetDuplicator.DuplicateSelectedChartByOffset();
        public void AddMetricBubbles() => StatBubbles.InsertMetricsBubbles();
        public void KeyPoints() => LabelTools.LabelPoints();
        public void SuperScript() => LabelTools.SuperscriptFootnotes();
        public void InsertArrowCAGR() => new CAGRArrow().insertArrowCAGR();
        public void InsertElbowCAGR() => new CAGRArrow().insertElbowCAGR();
        public void AddStackedTotals() => StackedColumn.StackedTotal();
        public void AddStatLine() => StatLines.AddMedianLine();
        public void ChartStackedWaterfall() => WaterfallChartHelper.InsertStackedWaterfall();

        public void VersionUp() => FileTools.VersionUp(Globals.ThisAddIn.Application);
        public void SaveAsPDF() => FileTools.SavePDF(Globals.ThisAddIn.Application);
        public void ExportSheet() => ExportSelectionTool.ExportSelection(Globals.ThisAddIn.Application);
        public void ShowExplorer() => FileTools.OpenContainingFolder();
        public void CopyPath() => FileTools.CopyFilePathToClipboard();
        public void HidePageBreaks() => FileTools.HidePageBreaks(Globals.ThisAddIn.Application);
        public void DeepSearch() => SuperFind.Show();
        public void CleanFile() => WorkbookCleaner.ComprehensiveClean();
        public void VisualCheck() => VisualCheckHelper.ToggleVisualCheck();
        public void ApplyTableBrandFormat() =>
            OfficePragmatics_Excel.Classes.Tables.TableFormatter.ApplyToSelection(
                OfficePragmatics_Excel.Settings.SettingsForm.CurrentSettings.TableFormat);
        public void ApplyChartBrandFormat() =>
            Classes.Brand.ChartFormat.ApplyBrandFormattingToSelection(
                Settings.SettingsForm.CurrentSettings.BrandCharts);

        public void ShowSettings() => Settings.SettingsForm.ShowSettingsDialog();
        public void OpenLibrary() => ExcelLibraryUtility.OpenLibrary();

        public static string GetMethodNameForButtonId(string buttonId)
        {
            switch (buttonId)
            {
                case "btnTracePrecedents": return nameof(TracePrecedents);
                case "btnTraceDependents": return nameof(TraceDependents);
                case "btnRevertTrace": return nameof(RevertBackward);
                case "btnDrillPrecedents": return nameof(DrillPrecedents);
                case "btnRemoveAllTracePrecedentArrows": return nameof(RemoveAllTracePrecedentArrows);
                case "btnDrillDependents": return nameof(DrillDependents);
                case "btnRemoveAllTraceDependentArrows": return nameof(RemoveAllTraceDependentArrows);

                case "b_Magnitude_Plus": return nameof(Magnitude_Plus);
                case "b_Magnitude_Minus": return nameof(Magnitude_Minus);
                case "b_FlipSign": return nameof(FlipSign);
                case "b_FillRight": return nameof(FillRight);
                case "b_FillDown": return nameof(FillDown);
                case "b_ErrorWrapNA": return nameof(ErrorWrapNA);
                case "b_ErrorWrapZero": return nameof(ErrorWrapZero);
                case "b_ConvertToLet": return nameof(ConvertToLet);
                case "b_FormulaWrap": return nameof(FormulaWrap);
                case "b_SolveFormula": return nameof(SolveFormula);
                case "b_AnchorCycle": return nameof(AnchorCycle);

                case "b_pasteExact": return nameof(PasteExact);
                case "b_pasteDuplicate": return nameof(PasteDuplicate);
                case "b_PasteNumberFormat": return nameof(PasteNumberFormat);
                case "b_pasteFillColor": return nameof(PasteFillColor);
                case "b_pasteBorders": return nameof(PasteBorders);
                case "b_pasteCellWidths": return nameof(PasteCellWidths);
                case "b_pasteCellHeights": return nameof(PasteCellHeights);

                case "b_NumberCycle": return nameof(RunNumberCycle);
                case "b_CurrencyCycle": return nameof(RunCurrencyCycle);
                case "b_PercentCycle": return nameof(RunPercentCycle);
                case "b_DateCycle": return nameof(RunDateCycle);
                case "b_MultipleCycle": return nameof(RunMultipleCycle);
                case "b_BinaryCycle": return nameof(RunBinaryCycle);
                case "b_PlusDecimals": return nameof(RunPlusDecimals);
                case "b_MinusDecimals": return nameof(RunMinusDecimals);

                case "b_BBColorCycle": return nameof(RunBBColorCycle);
                case "b_ModelColorCycle": return nameof(RunModelColorCycle);
                case "b_AutoColorSelection": return nameof(RunAutoColor);
                case "b_AutoTableColor": return nameof(RunTableFormattingCycle);
                case "b_ModelFillColorCycle": return nameof(RunFillCycle);
                case "b_FillThemeColorCycle": return nameof(RunThemeColorCycle);
                case "b_TransparentToWhite": return nameof(RunConvertTransparentToWhite);

                case "b_UnderlineCycle": return nameof(RunUnderlineCycle);
                case "b_BorderCycle": return nameof(RunBorderCycle);
                case "b_FontSizeCycle": return nameof(RunFontSizeCycle);
                case "b_CellHeightCycle": return nameof(RunRowHeightCycle);
                case "b_CellWidthCycle": return nameof(RunColumnWidthCycle);
                case "b_FootnoteCycle": return nameof(RunFootnoteCycle);
                case "b_RemoveFootnotes": return nameof(RunFootnoteRemove);
                case "b_ListBullets": return nameof(RunListBullets);
                case "b_AccountingDots": return nameof(RunAccountingDots);
                case "b_SumLine": return nameof(RunSumLine);
                case "b_LeftIndent": return nameof(RunIndentDecrease);
                case "b_RightIndent": return nameof(RunIndentIncrease);

                case "b_MergeToCenterAcrossSelection": return nameof(MergeToCenterAcrossSelection);
                case "b_hideToGroup": return nameof(HideToGroup);
                case "b_reverseColumns": return nameof(ReverseColumns);
                case "b_reverseRows": return nameof(ReverseRows);
                case "b_deleteBlankColumns": return nameof(DeleteBlankColumns);
                case "b_deleteBlankRows": return nameof(DeleteBlankRows);
                case "b_harcodeSelection": return nameof(HardcodeSelection);
                case "b_calculateCAGR": return nameof(CalculateCAGR);
                case "b_Guides": return nameof(InsertGuides);
                case "b_ResizetoPPT": return nameof(ResizeToPPT);

                case "b_StackSort": return nameof(StackSort);
                case "b_MaximizePlot": return nameof(MaximizePlot);
                case "b_CenterPie": return nameof(CenterPie);
                case "b_YMinZero": return nameof(YMinZero);
                case "b_HideYAxis": return nameof(HideYAxis);
                case "b_KeyboardControl": return nameof(KeyboardControl);
                case "b_ResetLabels": return nameof(ResetLabels);
                case "b_DuplicateToSheets": return nameof(DuplicateToSheets);
                case "b_OffsetChart": return nameof(OffsetChart);
                case "b_AddMetricBubbles": return nameof(AddMetricBubbles);
                case "b_KeyPoints": return nameof(KeyPoints);
                case "b_SuperScript": return nameof(SuperScript);
                case "b_insertArrowCAGR": return nameof(InsertArrowCAGR);
                case "b_insertElbowCAGR": return nameof(InsertElbowCAGR);
                case "b_AddStackedTotals": return nameof(AddStackedTotals);
                case "b_AddStatLine": return nameof(AddStatLine);
                case "b_ChartStackedWaterfall": return nameof(ChartStackedWaterfall);

                case "btnVersionUp": return nameof(VersionUp);
                case "btnSaveAsPDF": return nameof(SaveAsPDF);
                case "btnExportSheet": return nameof(ExportSheet);
                case "b_Show_Explorer": return nameof(ShowExplorer);
                case "b_Copy_Path": return nameof(CopyPath);
                case "b_hidePageBreaks": return nameof(HidePageBreaks);
                case "btnDeepSearch": return nameof(DeepSearch);
                case "btnCleanFile": return nameof(CleanFile);
                case "btnVisualCheck": return nameof(VisualCheck);
                case "btnTableBrandFormat": return nameof(ApplyTableBrandFormat);
                case "btnChartBrandFormat": return nameof(ApplyChartBrandFormat);

                case "btnSettings": return nameof(ShowSettings);
                case "btnLibrary": return nameof(OpenLibrary);
                case "btnSheetSelect": return nameof(SheetSelectorShow);

                default: return null;
            }
        }
    }
}