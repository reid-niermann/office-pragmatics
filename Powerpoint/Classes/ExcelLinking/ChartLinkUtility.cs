// ChartLinkUtility.cs (.NET Framework 4.8 / C# 7.3)

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OfficePragmatics_Powerpoint;

namespace OfficePragmatics_PowerPoint.ExcelLinking
{
    public static class ChartLinkUtility
    {
        /* ═════════════════════════  LINK BREAKING  ══════════════════════════ */

        /// <summary>
        /// Breaks the Excel data link on a native PowerPoint chart, embedding the data instead.
        /// After calling this, the chart will no longer auto-update when Excel changes.
        /// </summary>
        /// <param name="chartShape">A PowerPoint shape that contains a chart (msoChart type)</param>
        /// <returns>True if link was broken successfully, false otherwise</returns>
        public static bool BreakChartLink(PowerPoint.Shape chartShape)
        {
            if (chartShape == null)
                return false;

            if (chartShape.Type != Office.MsoShapeType.msoChart)
                return false;

            try
            {
                PowerPoint.Chart chart = chartShape.Chart;
                if (chart == null)
                    return false;

                // Check if chart has linked data
                PowerPoint.ChartData chartData = chart.ChartData;
                if (chartData == null)
                    return false;

                // IsLinked property tells us if data is linked to external source
                if (chartData.IsLinked)
                {
                    // BreakLink embeds the data and removes the external link
                    chartData.BreakLink();
                    System.Diagnostics.Debug.WriteLine("Chart link broken - data is now embedded");
                    return true;
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("Chart data is already embedded (no link to break)");
                    return true; // Already embedded, that's fine
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to break chart link: {ex.Message}");
                return false;
            }
        }

        /* ═════════════════════════  SHARED HELPER  ══════════════════════════ */
        /// <summary>
        /// Temporarily resizes an embedded Excel chart (ChartObject) to the given
        /// width/height, executes a copy/export delegate, then restores the original size.
        /// NOTE: Caller must ensure ExcelPerformanceScope is active (handles ScreenUpdating).
        /// </summary>
        private static void CopyOrExportWithTempResize(
            Excel.ChartObject chartObj,
            float newWidth,
            float newHeight,
            Action copyOrExportAction)
        {
            if (chartObj == null)
            {
                copyOrExportAction();
                return;
            }

            double originalW = chartObj.Width;
            double originalH = chartObj.Height;

            try
            {
                chartObj.Width = newWidth;
                chartObj.Height = newHeight;
                copyOrExportAction();
            }
            finally
            {
                chartObj.Width = originalW;
                chartObj.Height = originalH;
            }
        }

        /* ════════════════════════  UPDATE (REFRESH)  ════════════════════════ */
        public static void UpdateLinks(PowerPoint.Application ppApp)
        {
            PowerPoint.Presentation pres = null;
            Excel.Application excelApp = null;
            bool excelWasOpenedByUs = false;
            WorkbookCache wbCache = null;
            WarningCollector warnings = new WarningCollector();
            SlideActivationTracker slideTracker = null;

            try
            {
                if (ppApp?.ActivePresentation == null)
                {
                    MessageBox.Show("No active PowerPoint presentation found.",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                pres = ppApp.ActivePresentation;

                /* ── Obtain or launch Excel ─────────────────────────────── */
                try
                {
                    excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                }
                catch (COMException)
                {
                    excelApp = new Excel.Application();
                    excelWasOpenedByUs = true;
                    excelApp.Visible = true;
                    excelApp.WindowState = Excel.XlWindowState.xlMinimized;
                }

                // Initialize slide tracker
                slideTracker = new SlideActivationTracker(ppApp);

                // Use performance scope for all Excel operations
                using (new ExcelPerformanceScope(excelApp))
                {
                    // Initialize workbook cache for this batch operation
                    wbCache = new WorkbookCache(excelApp);

                    /* ── Loop through all slides & shapes ───────────────────── */
                    foreach (PowerPoint.Slide slide in pres.Slides)
                    {
                        int currentSlideIndex = slide.SlideIndex;

                        for (int sIdx = slide.Shapes.Count; sIdx >= 1; sIdx--)
                        {
                            PowerPoint.Shape shape = null;
                            try
                            {
                                shape = slide.Shapes[sIdx];
                                if (shape.Tags["Pragmatics_LinkedChart_ID"] == "") continue;

                                // Cache tag values once (avoid repeated COM calls)
                                string wbPath = shape.Tags["Pragmatics_SourceWorkbookPath"];
                                string wsName = shape.Tags["Pragmatics_SourceWorksheet"];
                                string chName = shape.Tags["Pragmatics_SourceChart"];
                                string importType = shape.Tags["Pragmatics_ImportType"];
                                string sizeMode = shape.Tags["Pragmatics_SizeMode"];

                                if (string.IsNullOrEmpty(wbPath) ||
                                    string.IsNullOrEmpty(wsName) ||
                                    string.IsNullOrEmpty(chName)) continue;

                                if (string.IsNullOrEmpty(importType)) importType = "Picture";

                                // Cache shape dimensions once
                                float shapeW = shape.Width;
                                float shapeH = shape.Height;
                                float shapeL = shape.Left;
                                float shapeT = shape.Top;

                                /* ── Get workbook (cached) ───────────────────────── */
                                Excel.Workbook wb = wbCache.GetOrOpen(wbPath, out bool wbOpened, out string actualWbPath);
                                if (wb == null)
                                {
                                    warnings.AddWarning($"Workbook not found: {System.IO.Path.GetFileName(wbPath)}");
                                    warnings.IncrementFailed();
                                    continue;
                                }

                                /* ── Get chart ─────────────────────────────────── */
                                Excel.ChartObject embChartObj = null;
                                Excel.Chart chartToUse = null;
                                try
                                {
                                    object sheetItem = wb.Sheets[wsName];
                                    if (sheetItem is Excel.Worksheet ws)
                                    {
                                        foreach (Excel.ChartObject co in ws.ChartObjects())
                                        {
                                            if (co.Name == chName)
                                            {
                                                embChartObj = co;
                                                chartToUse = co.Chart;
                                                break;
                                            }
                                            Marshal.ReleaseComObject(co);
                                        }
                                        Marshal.ReleaseComObject(ws);
                                        if (chartToUse == null)
                                        {
                                            warnings.AddWarning($"Chart '{chName}' not found on '{wsName}'");
                                            warnings.IncrementFailed();
                                            continue;
                                        }
                                    }
                                    else if (sheetItem is Excel.Chart chartSheet)
                                    {
                                        if (chartSheet.Name != chName)
                                        {
                                            warnings.AddWarning($"Chart sheet mismatch: '{chartSheet.Name}' != '{chName}'");
                                            warnings.IncrementFailed();
                                            Marshal.ReleaseComObject(chartSheet);
                                            continue;
                                        }
                                        chartToUse = chartSheet;
                                    }
                                    else
                                    {
                                        warnings.AddWarning($"'{wsName}' is neither sheet nor chart");
                                        warnings.IncrementFailed();
                                        if (sheetItem != null) Marshal.ReleaseComObject(sheetItem);
                                        continue;
                                    }

                                    PowerPoint.ShapeRange newSR = null;
                                    PowerPoint.Shape newSh = null;
                                    try
                                    {
                                        /* ══════  PICTURE  ══════ */
                                        if (importType == "Picture")
                                        {
                                            Action copyPic = () =>
                                                chartToUse.CopyPicture(
                                                    Excel.XlPictureAppearance.xlScreen,
                                                    Excel.XlCopyPictureFormat.xlPicture);

                                            bool needsResize = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase)
                                                && embChartObj != null
                                                && SizeModeUtil.NeedsResize(embChartObj, shapeW, shapeH);

                                            if (needsResize)
                                            {
                                                CopyOrExportWithTempResize(embChartObj, shapeW, shapeH, copyPic);
                                            }
                                            else
                                            {
                                                copyPic();
                                            }

                                            ClipboardHelper.WaitForClipboardReady();

                                            // Only activate slide right before paste
                                            slideTracker.EnsureSlideActive(currentSlideIndex);

                                            newSR = ClipboardHelper.RetryPaste(() =>
                                                slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile));
                                            newSh = newSR[1];
                                        }
                                        /* ══════  SVG  ══════ */
                                        else if (importType == "SVG")
                                        {
                                            string tmpSvg = System.IO.Path.GetTempFileName() + ".svg";
                                            try
                                            {
                                                Action exportSvg = () => chartToUse.Export(tmpSvg, "SVG");

                                                bool needsResize = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase)
                                                    && embChartObj != null
                                                    && SizeModeUtil.NeedsResize(embChartObj, shapeW, shapeH);

                                                if (needsResize)
                                                {
                                                    CopyOrExportWithTempResize(embChartObj, shapeW, shapeH, exportSvg);
                                                }
                                                else
                                                {
                                                    exportSvg();
                                                }

                                                newSh = slide.Shapes.AddPicture(
                                                    tmpSvg,
                                                    Office.MsoTriState.msoFalse,
                                                    Office.MsoTriState.msoTrue,
                                                    shapeL, shapeT, -1, -1);
                                            }
                                            finally
                                            {
                                                if (System.IO.File.Exists(tmpSvg))
                                                {
                                                    try { System.IO.File.Delete(tmpSvg); } catch { }
                                                }
                                            }
                                        }
                                        /* ══════  NATIVE CHART  ══════ */
                                        else if (importType == "Chart")
                                        {
                                            Action copyNative = () =>
                                            {
                                                if (embChartObj != null) embChartObj.Copy();
                                                else chartToUse.Copy();
                                            };

                                            bool needsResize = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase)
                                                && embChartObj != null
                                                && SizeModeUtil.NeedsResize(embChartObj, shapeW, shapeH);

                                            if (needsResize)
                                            {
                                                CopyOrExportWithTempResize(embChartObj, shapeW, shapeH, copyNative);
                                            }
                                            else
                                            {
                                                copyNative();
                                            }

                                            ClipboardHelper.WaitForClipboardReady();

                                            // Only activate slide right before paste
                                            slideTracker.EnsureSlideActive(currentSlideIndex);

                                            try
                                            {
                                                // Try generic Paste first, fall back to PasteOLEObject if needed
                                                newSR = ClipboardHelper.RetryPaste(() => slide.Shapes.Paste());
                                                newSh = newSR[1];
                                                if (newSh.Type != Office.MsoShapeType.msoChart)
                                                {
                                                    newSh.Delete();
                                                    Marshal.ReleaseComObject(newSh);
                                                    newSh = null;
                                                    Marshal.ReleaseComObject(newSR);
                                                    newSR = null;

                                                    // Re-copy and try OLE paste
                                                    if (embChartObj != null) embChartObj.Copy();
                                                    else chartToUse.Copy();
                                                    ClipboardHelper.WaitForClipboardReady();

                                                    newSR = ClipboardHelper.RetryPaste(() =>
                                                        slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteOLEObject));
                                                    newSh = newSR[1];
                                                }
                                            }
                                            catch
                                            {
                                                if (newSR != null)
                                                {
                                                    try { newSR.Delete(); } catch { }
                                                }
                                                throw;
                                            }
                                        }
                                        else
                                        {
                                            warnings.AddWarning($"Unsupported import type '{importType}'");
                                            warnings.IncrementFailed();
                                            continue;
                                        }

                                        if (newSh == null)
                                        {
                                            warnings.AddWarning("Paste produced no shape");
                                            warnings.IncrementFailed();
                                            continue;
                                        }

                                        // Break Excel link on native charts to prevent auto-updating
                                        // Our add-in manages updates through tags, not Excel's live linking
                                        if (importType == "Native" && newSh.Type == Office.MsoShapeType.msoChart)
                                        {
                                            BreakChartLink(newSh);
                                        }

                                        /* ── Copy tags ───────────────────────────── */
                                        bool pathWasUpgraded = actualWbPath != wbPath;

                                        for (int i = 1; i <= shape.Tags.Count; i++)
                                        {
                                            string tagName = shape.Tags.Name(i);
                                            string tagValue = shape.Tags.Value(i);
                                            if (pathWasUpgraded && string.Equals(tagName, "Pragmatics_SourceWorkbookPath", StringComparison.OrdinalIgnoreCase))
                                            {
                                                newSh.Tags.Add(tagName, actualWbPath);
                                                System.Diagnostics.Debug.WriteLine(
                                                    $"Updated chart workbook path in tags: {wbPath} → {actualWbPath}");
                                            }
                                            else
                                            {
                                                newSh.Tags.Add(tagName, tagValue);
                                            }
                                        }

                                        /* ── Position & size per saved policy ───── */
                                        SizeModeUtil.ApplyFromSource(shape, newSh);

                                        string oldName = shape.Name;
                                        shape.Delete();
                                        shape = null;
                                        newSh.Select();
                                        warnings.IncrementUpdated();
                                        System.Diagnostics.Debug.WriteLine($"Replaced '{oldName}' successfully.");
                                    }
                                    finally
                                    {
                                        if (newSh != null) Marshal.ReleaseComObject(newSh);
                                        if (newSR != null) Marshal.ReleaseComObject(newSR);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    warnings.AddWarning($"Chart refresh error: {ex.Message}");
                                    warnings.IncrementFailed();
                                }
                                finally
                                {
                                    if (chartToUse != null && Marshal.IsComObject(chartToUse))
                                        Marshal.ReleaseComObject(chartToUse);
                                    if (embChartObj != null) Marshal.ReleaseComObject(embChartObj);
                                }
                            }
                            finally
                            {
                                if (shape != null) Marshal.ReleaseComObject(shape);
                            }
                        }
                        Marshal.ReleaseComObject(slide);
                    }

                    // Dispose cache inside the performance scope
                    wbCache?.Dispose();
                    wbCache = null;
                }

                // Show single summary instead of multiple MessageBoxes
                warnings.ShowSummary("Chart Links Update Complete");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Update error: {ex.Message}", "Update Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                wbCache?.Dispose();

                if (excelWasOpenedByUs && excelApp?.Workbooks.Count == 0) excelApp?.Quit();
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (pres != null) Marshal.ReleaseComObject(pres);

                // Single GC at end
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /* ═════════════════════  IMPORT – NATIVE CHART  ═════════════════════ */
        public static void ImportExcelChartAsNative(PowerPoint.Application ppApp)
        {
            Excel.Application excelApp = null;
            object selection = null;
            Excel.ChartObject selChartObj = null;
            Excel.Chart activeChart = null;
            PowerPoint.Slide tgtSlide = null;
            PowerPoint.ShapeRange selShapes = null;
            PowerPoint.Shape anchorSh = null;
            PowerPoint.Shape pastedSh = null;
            PowerPoint.ShapeRange pastedSR = null;
            float targW = 0, targH = 0, targL = 0, targT = 0;
            bool hasAnchor = false;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.Selection == null)
                {
                    MessageBox.Show("Select a chart in Excel first.", "Excel",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selection = excelApp.Selection;
                if (selection is Excel.ChartObject co)
                {
                    selChartObj = co;
                    activeChart = co.Chart;
                }
                else if (excelApp.ActiveChart != null)
                {
                    activeChart = excelApp.ActiveChart;
                    if (activeChart.Parent is Excel.ChartObject coPar) selChartObj = coPar;
                }
                else
                {
                    MessageBox.Show("Selection is not a chart.", "Excel",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                /* ── find target slide & (optional) anchor shape ───────── */
                tgtSlide = ppApp.ActiveWindow.View.Slide ??
                           ppApp.ActiveWindow.Selection.SlideRange[1];

                var ppSel = ppApp.ActiveWindow.Selection;
                if (ppSel.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    ppSel.ShapeRange.Count > 0)
                {
                    selShapes = ppSel.ShapeRange;
                    anchorSh = selShapes[1];
                    targW = anchorSh.Width;
                    targH = anchorSh.Height;
                    targL = anchorSh.Left;
                    targT = anchorSh.Top;
                    hasAnchor = true;
                }

                Ribbon1 ribbon = Ribbon1.Instance;
                string match = ribbon?.GetCurrentImportMatchType() ?? "Match Width";

                /* ── copy chart (with optional temp resize) ─────────────── */
                Action copyNative = () =>
                {
                    if (selChartObj != null) selChartObj.Copy();
                    else activeChart.Copy();
                };

                bool needsResize = match == "Match All" && hasAnchor && selChartObj != null
                    && SizeModeUtil.NeedsResize(selChartObj, targW, targH);

                // Use performance scope for copy operation
                using (new ExcelPerformanceScope(excelApp))
                {
                    if (needsResize)
                    {
                        CopyOrExportWithTempResize(selChartObj, targW, targH, copyNative);
                    }
                    else
                    {
                        copyNative();
                    }
                }

                ClipboardHelper.WaitForClipboardReady();

                /* ── paste & adjust ────────────────────────────────────── */
                try
                {
                    // Try generic Paste first, fall back to PasteOLEObject if needed
                    pastedSR = ClipboardHelper.RetryPaste(() => tgtSlide.Shapes.Paste());
                    pastedSh = pastedSR[1];
                    if (pastedSh.Type != Office.MsoShapeType.msoChart)
                    {
                        pastedSh.Delete();
                        Marshal.ReleaseComObject(pastedSh);
                        pastedSh = null;
                        Marshal.ReleaseComObject(pastedSR);
                        pastedSR = null;

                        // Re-copy and try OLE paste
                        if (selChartObj != null) selChartObj.Copy();
                        else activeChart.Copy();
                        ClipboardHelper.WaitForClipboardReady();

                        pastedSR = ClipboardHelper.RetryPaste(() =>
                            tgtSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteOLEObject));
                        pastedSh = pastedSR[1];
                    }
                }
                catch (COMException ex)
                {
                    MessageBox.Show($"Paste failed: {ex.Message}", "Import Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (pastedSh.Type != Office.MsoShapeType.msoChart)
                {
                    pastedSh.Delete();
                    MessageBox.Show($"Unexpected pasted type: {pastedSh.Type}",
                                    "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Break Excel link to prevent auto-updating when source file is open
                // Our add-in manages updates through tags, not Excel's live linking
                BreakChartLink(pastedSh);

                /* ── tag new shape ─────────────────────────────────────── */
                string id = Guid.NewGuid().ToString();
                // Normalize path for cross-user compatibility
                string normalizedWbPath = CloudPathHelper.NormalizePath(activeChart.Application.ActiveWorkbook.FullName);
                pastedSh.Tags.Add("Pragmatics_LinkedChart_ID", id);
                pastedSh.Tags.Add("Pragmatics_SourceWorkbookPath", normalizedWbPath);

                string wsName, chName;
                if (selChartObj != null)
                {
                    wsName = ((Excel.Worksheet)selChartObj.Parent).Name;
                    chName = selChartObj.Name;
                }
                else
                {
                    wsName = activeChart.Name;
                    chName = activeChart.Name;
                }

                pastedSh.Tags.Add("Pragmatics_SourceWorksheet", wsName);
                pastedSh.Tags.Add("Pragmatics_SourceChart", chName);
                pastedSh.Tags.Add("Pragmatics_ImportType", "Chart");
                pastedSh.Tags.Add("Pragmatics_SizeMode", match);

                /* ── size/position ─────────────────────────────────────── */
                if (hasAnchor)
                {
                    if (match == "Match All")
                    {
                        pastedSh.LockAspectRatio = Office.MsoTriState.msoFalse;
                        pastedSh.Width = targW;
                        pastedSh.Height = targH;
                    }
                    else if (match == "Match Height")
                    {
                        pastedSh.LockAspectRatio = Office.MsoTriState.msoTrue;
                        pastedSh.Height = targH;
                    }
                    else
                    {
                        pastedSh.LockAspectRatio = Office.MsoTriState.msoTrue;
                        pastedSh.Width = targW;
                    }
                    pastedSh.Left = targL;
                    pastedSh.Top = targT;
                    anchorSh.Delete();
                }

                pastedSh.Select(Office.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Import error: {ex.Message}", "Import Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (pastedSh != null) Marshal.ReleaseComObject(pastedSh);
                if (pastedSR != null) Marshal.ReleaseComObject(pastedSR);
                if (anchorSh != null) Marshal.ReleaseComObject(anchorSh);
                if (selShapes != null) Marshal.ReleaseComObject(selShapes);
                if (tgtSlide != null) Marshal.ReleaseComObject(tgtSlide);
                if (activeChart != null) Marshal.ReleaseComObject(activeChart);
                if (selChartObj != null) Marshal.ReleaseComObject(selChartObj);
                if (selection != null && Marshal.IsComObject(selection))
                    Marshal.ReleaseComObject(selection);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /* ════════════════  IMPORT – PICTURE (EMF)  ════════════════════════ */
        public static void ImportExcelChartAsPicture(PowerPoint.Application ppApp)
        {
            Excel.Application excelApp = null;
            object selection = null;
            Excel.ChartObject selChartObj = null;
            Excel.Chart activeChart = null;
            PowerPoint.Slide tgtSlide = null;
            PowerPoint.ShapeRange selShapes = null;
            PowerPoint.Shape anchorSh = null;
            PowerPoint.Shape pastedSh = null;
            PowerPoint.ShapeRange pastedSR = null;
            float targW = 0, targH = 0, targL = 0, targT = 0;
            bool hasAnchor = false;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.Selection == null)
                {
                    MessageBox.Show("Select a chart in Excel first.", "Excel",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selection = excelApp.Selection;
                if (selection is Excel.ChartObject co)
                {
                    selChartObj = co;
                    activeChart = co.Chart;
                }
                else if (excelApp.ActiveChart != null)
                {
                    activeChart = excelApp.ActiveChart;
                    if (activeChart.Parent is Excel.ChartObject coPar) selChartObj = coPar;
                }
                else
                {
                    MessageBox.Show("Selection is not a chart.", "Excel",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                /* ── PPT target ────────────────────────────────────────── */
                tgtSlide = ppApp.ActiveWindow.View.Slide ??
                           ppApp.ActiveWindow.Selection.SlideRange[1];

                var ppSel = ppApp.ActiveWindow.Selection;
                if (ppSel.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    ppSel.ShapeRange.Count > 0)
                {
                    selShapes = ppSel.ShapeRange;
                    anchorSh = selShapes[1];
                    targW = anchorSh.Width;
                    targH = anchorSh.Height;
                    targL = anchorSh.Left;
                    targT = anchorSh.Top;
                    hasAnchor = true;
                }

                Ribbon1 ribbon = Ribbon1.Instance;
                string match = ribbon?.GetCurrentImportMatchType() ?? "Match Width";

                /* ── copy picture (with optional temp resize) ──────────── */
                Action copyPic = () =>
                    activeChart.CopyPicture(Excel.XlPictureAppearance.xlScreen,
                                            Excel.XlCopyPictureFormat.xlPicture);

                bool needsResize = match == "Match All" && hasAnchor && selChartObj != null
                    && SizeModeUtil.NeedsResize(selChartObj, targW, targH);

                using (new ExcelPerformanceScope(excelApp))
                {
                    if (needsResize)
                    {
                        CopyOrExportWithTempResize(selChartObj, targW, targH, copyPic);
                    }
                    else
                    {
                        copyPic();
                    }
                }

                ClipboardHelper.WaitForClipboardReady();

                /* ── paste & adjust ────────────────────────────────────── */
                pastedSR = ClipboardHelper.RetryPaste(() =>
                    tgtSlide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile));
                pastedSh = pastedSR[1];

                string id = Guid.NewGuid().ToString();
                // Normalize path for cross-user compatibility
                string normalizedWbPath = CloudPathHelper.NormalizePath(activeChart.Application.ActiveWorkbook.FullName);
                pastedSh.Tags.Add("Pragmatics_LinkedChart_ID", id);
                pastedSh.Tags.Add("Pragmatics_SourceWorkbookPath", normalizedWbPath);

                string wsName, chName;
                if (selChartObj != null)
                {
                    wsName = ((Excel.Worksheet)selChartObj.Parent).Name;
                    chName = selChartObj.Name;
                }
                else
                {
                    wsName = activeChart.Name;
                    chName = activeChart.Name;
                }

                pastedSh.Tags.Add("Pragmatics_SourceWorksheet", wsName);
                pastedSh.Tags.Add("Pragmatics_SourceChart", chName);
                pastedSh.Tags.Add("Pragmatics_ImportType", "Picture");
                pastedSh.Tags.Add("Pragmatics_SizeMode", match);

                if (hasAnchor)
                {
                    if (match == "Match All")
                    {
                        pastedSh.LockAspectRatio = Office.MsoTriState.msoFalse;
                        pastedSh.Width = targW;
                        pastedSh.Height = targH;
                    }
                    else if (match == "Match Height")
                    {
                        pastedSh.LockAspectRatio = Office.MsoTriState.msoTrue;
                        pastedSh.Height = targH;
                    }
                    else
                    {
                        pastedSh.LockAspectRatio = Office.MsoTriState.msoTrue;
                        pastedSh.Width = targW;
                    }
                    pastedSh.Left = targL;
                    pastedSh.Top = targT;
                    anchorSh.Delete();
                }

                pastedSh.Select(Office.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Import error: {ex.Message}", "Import Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (pastedSh != null) Marshal.ReleaseComObject(pastedSh);
                if (pastedSR != null) Marshal.ReleaseComObject(pastedSR);
                if (anchorSh != null) Marshal.ReleaseComObject(anchorSh);
                if (selShapes != null) Marshal.ReleaseComObject(selShapes);
                if (tgtSlide != null) Marshal.ReleaseComObject(tgtSlide);
                if (activeChart != null) Marshal.ReleaseComObject(activeChart);
                if (selChartObj != null) Marshal.ReleaseComObject(selChartObj);
                if (selection != null && Marshal.IsComObject(selection))
                    Marshal.ReleaseComObject(selection);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /* ════════════════  IMPORT – VECTOR (SVG)  ═════════════════════════ */
        public static void ImportExcelChartAsVector(PowerPoint.Application ppApp)
        {
            Excel.Application excelApp = null;
            object selection = null;
            Excel.ChartObject selChartObj = null;
            Excel.Chart activeChart = null;
            PowerPoint.Slide tgtSlide = null;
            PowerPoint.ShapeRange selShapes = null;
            PowerPoint.Shape anchorSh = null;
            PowerPoint.Shape svgSh = null;
            float targW = 0, targH = 0, targL = 0, targT = 0;
            bool hasAnchor = false;
            string tmpSvg = null;

            try
            {
                excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                if (excelApp.Selection == null)
                {
                    MessageBox.Show("Select a chart in Excel first.", "Excel",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                selection = excelApp.Selection;
                if (selection is Excel.ChartObject co)
                {
                    selChartObj = co;
                    activeChart = co.Chart;
                }
                else if (excelApp.ActiveChart != null)
                {
                    activeChart = excelApp.ActiveChart;
                    if (activeChart.Parent is Excel.ChartObject coPar) selChartObj = coPar;
                }
                else
                {
                    MessageBox.Show("Selection is not a chart.", "Excel",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                /* ── PPT target & anchor ───────────────────────────────── */
                tgtSlide = ppApp.ActiveWindow.View.Slide ??
                           ppApp.ActiveWindow.Selection.SlideRange[1];

                var ppSel = ppApp.ActiveWindow.Selection;
                if (ppSel.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    ppSel.ShapeRange.Count > 0)
                {
                    selShapes = ppSel.ShapeRange;
                    anchorSh = selShapes[1];
                    targW = anchorSh.Width;
                    targH = anchorSh.Height;
                    targL = anchorSh.Left;
                    targT = anchorSh.Top;
                    hasAnchor = true;
                }

                Ribbon1 ribbon = Ribbon1.Instance;
                string match = ribbon?.GetCurrentImportMatchType() ?? "Match Width";

                tmpSvg = System.IO.Path.GetTempFileName() + ".svg";
                Action exportSvg = () => activeChart.Export(tmpSvg, "SVG");

                bool needsResize = match == "Match All" && hasAnchor && selChartObj != null
                    && SizeModeUtil.NeedsResize(selChartObj, targW, targH);

                using (new ExcelPerformanceScope(excelApp))
                {
                    if (needsResize)
                    {
                        CopyOrExportWithTempResize(selChartObj, targW, targH, exportSvg);
                    }
                    else
                    {
                        exportSvg();
                    }
                }

                /* ── insert SVG ────────────────────────────────────────── */
                float insLeft = hasAnchor ? targL : 10;
                float insTop = hasAnchor ? targT : 10;

                svgSh = tgtSlide.Shapes.AddPicture(
                            tmpSvg,
                            Office.MsoTriState.msoFalse,
                            Office.MsoTriState.msoTrue,
                            insLeft, insTop, -1, -1);

                string id = Guid.NewGuid().ToString();
                // Normalize path for cross-user compatibility
                string normalizedWbPath = CloudPathHelper.NormalizePath(activeChart.Application.ActiveWorkbook.FullName);
                svgSh.Tags.Add("Pragmatics_LinkedChart_ID", id);
                svgSh.Tags.Add("Pragmatics_SourceWorkbookPath", normalizedWbPath);

                string wsName, chName;
                if (selChartObj != null)
                {
                    wsName = ((Excel.Worksheet)selChartObj.Parent).Name;
                    chName = selChartObj.Name;
                }
                else
                {
                    wsName = activeChart.Name;
                    chName = activeChart.Name;
                }

                svgSh.Tags.Add("Pragmatics_SourceWorksheet", wsName);
                svgSh.Tags.Add("Pragmatics_SourceChart", chName);
                svgSh.Tags.Add("Pragmatics_ImportType", "SVG");
                svgSh.Tags.Add("Pragmatics_SizeMode", match);

                if (hasAnchor)
                {
                    if (match == "Match All")
                    {
                        svgSh.LockAspectRatio = Office.MsoTriState.msoFalse;
                        svgSh.Width = targW;
                        svgSh.Height = targH;
                    }
                    else if (match == "Match Height")
                    {
                        svgSh.LockAspectRatio = Office.MsoTriState.msoTrue;
                        svgSh.Height = targH;
                    }
                    else
                    {
                        svgSh.LockAspectRatio = Office.MsoTriState.msoTrue;
                        svgSh.Width = targW;
                    }
                    svgSh.Left = targL;
                    svgSh.Top = targT;
                    anchorSh.Delete();
                }

                svgSh.Select(Office.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                MessageBox.Show("SVG import error: " + ex.Message, "Import Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (svgSh != null) Marshal.ReleaseComObject(svgSh);
                if (anchorSh != null) Marshal.ReleaseComObject(anchorSh);
                if (selShapes != null) Marshal.ReleaseComObject(selShapes);
                if (tgtSlide != null) Marshal.ReleaseComObject(tgtSlide);
                if (activeChart != null) Marshal.ReleaseComObject(activeChart);
                if (selChartObj != null) Marshal.ReleaseComObject(selChartObj);
                if (selection != null && Marshal.IsComObject(selection))
                    Marshal.ReleaseComObject(selection);
                if (excelApp != null) Marshal.ReleaseComObject(excelApp);
                if (tmpSvg != null && System.IO.File.Exists(tmpSvg))
                {
                    try { System.IO.File.Delete(tmpSvg); } catch { }
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
    }
}