// UpdateLinks.cs (.NET Framework 4.8 / C# 7.3)

using OfficePragmatics_PowerPoint.ExcelLinking;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_Powerpoint.Classes.ExcelLinking
{
    /// <summary>
    /// Types of links that can be updated.
    /// </summary>
    internal enum LinkType
    {
        Text,
        Range,
        Chart
    }

    /// <summary>
    /// Information about a linked shape for batch processing.
    /// Used to group shapes by workbook for efficient updates.
    /// </summary>
    internal class LinkedShapeInfo
    {
        public int SlideIndex { get; set; }
        public int ShapeIndex { get; set; }
        public LinkType LinkType { get; set; }
        public string WorkbookPath { get; set; }
    }

    // Win32: keep a window topmost temporarily
    internal static class Win32Focus
    {
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        private const int SW_SHOW = 5;

        public static void ActivateWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                return;

            ShowWindow(hwnd, SW_SHOW);
            SetForegroundWindow(hwnd);
        }
    }

    internal static class Win32TopMost
    {
        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool SetWindowPos(
            IntPtr hWnd, IntPtr hWndInsertAfter,
            int X, int Y, int cx, int cy, uint uFlags);

        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOACTIVATE = 0x0010;

        public static void MakeTopMost(IntPtr hwnd, bool on)
        {
            SetWindowPos(
                hwnd,
                on ? HWND_TOPMOST : HWND_NOTOPMOST,
                0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);
        }

        public sealed class Scope : IDisposable
        {
            private readonly IntPtr _hwnd;
            public Scope(IntPtr hwnd)
            {
                _hwnd = hwnd;
                MakeTopMost(_hwnd, true);
            }
            public void Dispose() => MakeTopMost(_hwnd, false);
        }
    }

    internal sealed class WindowHandleWrapper : IWin32Window
    {
        public IntPtr Handle { get; }
        public WindowHandleWrapper(IntPtr handle) => Handle = handle;
    }

    internal class UpdateLinks
    {
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

        public static void UpdateAllLinks(PowerPoint.Application ppApp)
        {
            PowerPoint.Presentation pres = null;
            Excel.Application excelApp = null;
            bool excelStartedByUs = false;
            WorkbookCache wbCache = null;
            WarningCollector warnings = new WarningCollector();
            SlideActivationTracker slideTracker = null;

            PowerPoint.Slide originalActiveSlide = null;
            int originalSlideIndex = 1;

            WaitForm waitForm = null;

            IntPtr pptHwnd = IntPtr.Zero;
            try { pptHwnd = (IntPtr)ppApp.HWND; } catch { }

            using (pptHwnd != IntPtr.Zero ? new Win32TopMost.Scope(pptHwnd) : null)
            {
                try
                {
                    waitForm = new WaitForm(pptHwnd);

                    if (pptHwnd != IntPtr.Zero)
                        waitForm.Show(new WindowHandleWrapper(pptHwnd));
                    else
                        waitForm.Show();

                    waitForm.Update();

                    try
                    {
                        originalActiveSlide = ppApp.ActiveWindow.View.Slide;
                        originalSlideIndex = originalActiveSlide.SlideIndex;
                    }
                    catch { originalSlideIndex = 1; }

                    if (ppApp?.ActivePresentation == null || ppApp.Presentations.Count == 0)
                    {
                        waitForm.ShowCompletion("Update Error", "No presentation is open.");
                        return;
                    }

                    pres = ppApp.ActivePresentation;

                    // Get or create Excel
                    try
                    {
                        excelApp = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                    }
                    catch (COMException)
                    {
                        try
                        {
                            excelApp = new Excel.Application();
                            excelStartedByUs = true;

                            // Set properties with error handling - Excel may not be fully ready
                            try { excelApp.Visible = true; } catch { }
                            try { excelApp.WindowState = Excel.XlWindowState.xlMinimized; } catch { }
                        }
                        catch (Exception ex)
                        {
                            waitForm.ShowCompletion("Excel Error", $"Could not start Excel: {ex.Message}");
                            return;
                        }
                    }

                    if (excelApp == null)
                    {
                        waitForm.ShowCompletion("Excel Error", "Could not connect to or start Excel.");
                        return;
                    }

                    // Initialize slide tracker
                    slideTracker = new SlideActivationTracker(ppApp);

                    // Use performance scope for ALL Excel operations
                    using (new ExcelPerformanceScope(excelApp))
                    {
                        wbCache = new WorkbookCache(excelApp, checkForNewerVersions: true);

                        int totalSlides = pres.Slides.Count;

                        // Phase 1: Collect all linked shapes grouped by workbook
                        // This allows us to process all shapes from one workbook together
                        waitForm.SetStatus("Scanning for linked objects...");

                        var linkedShapes = new List<LinkedShapeInfo>();

                        foreach (PowerPoint.Slide slide in pres.Slides)
                        {
                            PowerPoint.Shapes slideShapes = null;
                            int currentSlideIndex = slide.SlideIndex;

                            try
                            {
                                slideShapes = slide.Shapes;

                                for (int shapeIdx = slideShapes.Count; shapeIdx >= 1; shapeIdx--)
                                {
                                    PowerPoint.Shape shape = null;
                                    try
                                    {
                                        shape = slideShapes[shapeIdx];

                                        // Check for range/chart links first (these take precedence)
                                        string wbPath = shape.Tags["Pragmatics_SourceWorkbookPath"];
                                        bool isRangeLink = !string.IsNullOrEmpty(shape.Tags["Pragmatics_LinkedRange_ID"]) &&
                                                          !string.IsNullOrEmpty(wbPath);
                                        bool isChartLink = !string.IsNullOrEmpty(shape.Tags["Pragmatics_LinkedChart_ID"]) &&
                                                          !string.IsNullOrEmpty(wbPath);

                                        if (isRangeLink)
                                        {
                                            linkedShapes.Add(new LinkedShapeInfo
                                            {
                                                SlideIndex = currentSlideIndex,
                                                ShapeIndex = shapeIdx,
                                                LinkType = LinkType.Range,
                                                WorkbookPath = wbPath
                                            });
                                        }
                                        else if (isChartLink)
                                        {
                                            linkedShapes.Add(new LinkedShapeInfo
                                            {
                                                SlideIndex = currentSlideIndex,
                                                ShapeIndex = shapeIdx,
                                                LinkType = LinkType.Chart,
                                                WorkbookPath = wbPath
                                            });
                                        }
                                        else if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                                                 shape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                                        {
                                            // Only add as text link if NOT already a range/chart link
                                            // This avoids double-processing and incorrect failure counts
                                            linkedShapes.Add(new LinkedShapeInfo
                                            {
                                                SlideIndex = currentSlideIndex,
                                                ShapeIndex = shapeIdx,
                                                LinkType = LinkType.Text,
                                                WorkbookPath = "" // Multiple workbooks possible per shape
                                            });
                                        }
                                    }
                                    finally
                                    {
                                        if (shape != null) Marshal.ReleaseComObject(shape);
                                    }
                                }
                            }
                            finally
                            {
                                if (slideShapes != null) Marshal.ReleaseComObject(slideShapes);
                                Marshal.ReleaseComObject(slide);
                            }
                        }

                        // Phase 2: Group by workbook and process
                        // Process text links first (they may span multiple workbooks)
                        // Then process range/chart links grouped by workbook

                        var textLinks = linkedShapes.Where(s => s.LinkType == LinkType.Text).ToList();
                        var rangeChartLinks = linkedShapes.Where(s => s.LinkType != LinkType.Text)
                            .GroupBy(s => s.WorkbookPath, StringComparer.OrdinalIgnoreCase)
                            .OrderBy(g => g.Key) // Consistent order
                            .ToList();

                        // Count total range/chart links for progress tracking
                        int totalRangeChartLinks = rangeChartLinks.Sum(g => g.Count());
                        int processedLinks = 0;

                        // Initialize progress bar (text links + range/chart links)
                        waitForm.InitializeProgress(totalRangeChartLinks);

                        // Process text links first (they don't delete shapes, don't count in progress)
                        if (textLinks.Count > 0)
                        {
                            waitForm.SetStatus($"Updating {textLinks.Count} text links...");
                            foreach (var info in textLinks)
                            {
                                PowerPoint.Slide slide = null;
                                PowerPoint.Shape shape = null;
                                try
                                {
                                    slide = pres.Slides[info.SlideIndex];
                                    shape = slide.Shapes[info.ShapeIndex];
                                    ProcessTextLinks(shape, excelApp, wbCache, warnings);
                                }
                                catch (Exception ex)
                                {
                                    // Don't increment failed here - ProcessTextLinks handles its own counting
                                    // This catch is for shapes that can't be accessed (groups, etc.)
                                    System.Diagnostics.Debug.WriteLine($"Error accessing shape for text link: {ex.Message}");
                                }
                                finally
                                {
                                    if (shape != null) Marshal.ReleaseComObject(shape);
                                    if (slide != null) Marshal.ReleaseComObject(slide);
                                }
                            }
                        }

                        // Process range/chart links grouped by workbook
                        // Process in descending shape order so deletions don't affect indices
                        int wbIndex = 0;
                        int totalWorkbooks = rangeChartLinks.Count;

                        foreach (var wbGroup in rangeChartLinks)
                        {
                            wbIndex++;
                            string wbFileName = CloudPathHelper.GetFileName(wbGroup.Key);

                            // Open workbook once for all shapes from this workbook
                            Excel.Workbook wb = wbCache.GetOrOpen(wbGroup.Key, out bool _, out string actualWbPath);
                            if (wb != null)
                            {
                                wbCache.PrepareWorkbookForCopy(wb);
                            }

                            // Process all shapes from this workbook
                            // Order: by slide ascending, then by shape index DESCENDING
                            // (descending shape index ensures deletions don't affect remaining shapes)
                            foreach (var info in wbGroup
                                .OrderBy(s => s.SlideIndex)
                                .ThenByDescending(s => s.ShapeIndex))
                            {
                                processedLinks++;
                                waitForm.SetProgress(processedLinks, $"Workbook {wbIndex}/{totalWorkbooks}: {wbFileName}");

                                PowerPoint.Slide slide = null;
                                PowerPoint.Shape shape = null;
                                try
                                {
                                    slide = pres.Slides[info.SlideIndex];
                                    shape = slide.Shapes[info.ShapeIndex];

                                    if (info.LinkType == LinkType.Range)
                                    {
                                        ProcessRangeLink(shape, slide, info.SlideIndex, excelApp, wbCache, warnings, slideTracker, ppApp);
                                    }
                                    else if (info.LinkType == LinkType.Chart)
                                    {
                                        ProcessChartLink(shape, slide, info.SlideIndex, excelApp, wbCache, warnings, slideTracker);
                                    }
                                }
                                catch (Exception ex)
                                {
                                    warnings.AddWarning($"Error processing shape: {ex.Message}");
                                    warnings.IncrementFailed();
                                }
                                finally
                                {
                                    if (shape != null) Marshal.ReleaseComObject(shape);
                                    if (slide != null) Marshal.ReleaseComObject(slide);
                                }
                            }
                        }

                        // Dispose cache inside performance scope
                        wbCache?.Dispose();
                        wbCache = null;
                    }

                    // Restore original slide
                    if (originalSlideIndex > 0 && originalSlideIndex <= pres.Slides.Count)
                    {
                        try { ppApp.ActiveWindow.View.GotoSlide(originalSlideIndex); } catch { }
                    }

                    string body = warnings.BuildSummaryMessage(topIssues: 5);
                    waitForm.ShowCompletion("Update Complete", body);
                }
                catch (Exception ex)
                {
                    string errorDetails = $"{ex.GetType().Name}: {ex.Message}";
                    if (ex.InnerException != null)
                        errorDetails += $"\r\n\r\nInner: {ex.InnerException.Message}";
                    errorDetails += $"\r\n\r\nStack:\r\n{ex.StackTrace}";

                    if (waitForm != null && !waitForm.IsDisposed)
                        waitForm.ShowCompletion("Update Error", errorDetails);
                }
                finally
                {
                    wbCache?.Dispose();

                    if (pres != null) Marshal.ReleaseComObject(pres);
                    if (excelApp != null)
                    {
                        try
                        {
                            if (excelStartedByUs && excelApp.Workbooks.Count == 0)
                                excelApp.Quit();
                        }
                        catch { }

                        try { Marshal.ReleaseComObject(excelApp); } catch { }
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        /* ═══════════════════════  TEXT LINKS  ═══════════════════════ */
        private static void ProcessTextLinks(PowerPoint.Shape shape,
                                             Excel.Application excelApp,
                                             WorkbookCache wbCache,
                                             WarningCollector warnings)
        {
            PowerPoint.TextFrame textFrame = null;
            PowerPoint.TextRange shapeTextRange = null;

            try
            {
                textFrame = shape.TextFrame;
                if (textFrame.HasText != Office.MsoTriState.msoTrue) return;

                shapeTextRange = textFrame.TextRange;

                int numberOfRuns = 0;
                PowerPoint.TextRange tempRunsRange = null;
                try
                {
                    tempRunsRange = shapeTextRange.Runs();
                    if (tempRunsRange != null) numberOfRuns = tempRunsRange.Count;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not get run count: {ex.Message}");
                }
                finally
                {
                    if (tempRunsRange != null) Marshal.ReleaseComObject(tempRunsRange);
                }

                if (numberOfRuns > 0)
                {
                    for (int runIndex = 1; runIndex <= numberOfRuns; runIndex++)
                    {
                        PowerPoint.TextRange currentRun = null;
                        PowerPoint.ActionSetting runActionSetting = null;
                        PowerPoint.Hyperlink currentHyperlink = null;
                        Excel.Workbook linkedWb = null;
                        Excel.Worksheet linkedWs = null;
                        Excel.Name namedRangeObject = null;
                        Excel.Range linkedExcelCell = null;

                        try
                        {
                            currentRun = shapeTextRange.Runs(runIndex);
                            if (currentRun == null) continue;

                            runActionSetting = currentRun.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                            if (runActionSetting.Action == PowerPoint.PpActionType.ppActionHyperlink)
                            {
                                currentHyperlink = runActionSetting.Hyperlink;
                                if (currentHyperlink != null &&
                                    currentHyperlink.Address != null &&
                                    currentHyperlink.Address.StartsWith("pragmatics-textlink://"))
                                {
                                    Uri uri;
                                    try { uri = new Uri(currentHyperlink.Address); }
                                    catch (UriFormatException)
                                    {
                                        warnings.AddWarning($"Invalid text link URI");
                                        warnings.IncrementFailed();
                                        continue;
                                    }

                                    if (uri.Scheme != "pragmatics-textlink") continue;

                                    string query = uri.Query.StartsWith("?")
                                        ? uri.Query.Substring(1)
                                        : uri.Query;

                                    var qParams = query
                                        .Split('&')
                                        .Select(p => p.Split('='))
                                        .Where(p => p.Length == 2)
                                        .ToDictionary(
                                            s => s[0],
                                            s => Uri.UnescapeDataString(s[1]),
                                            StringComparer.OrdinalIgnoreCase);

                                    if (!qParams.TryGetValue("wb", out string wbPath) ||
                                        !qParams.TryGetValue("ws", out string wsName) ||
                                        !qParams.TryGetValue("nr", out string nrName))
                                    {
                                        warnings.AddWarning($"Malformed text link");
                                        warnings.IncrementFailed();
                                        continue;
                                    }

                                    linkedWb = wbCache.GetOrOpen(wbPath, out bool _, out string actualWbPath);
                                    if (linkedWb == null)
                                    {
                                        string fileName = CloudPathHelper.GetFileName(wbPath);
                                        if (CloudPathHelper.IsCloudPath(wbPath))
                                        {
                                            warnings.AddWarning($"Text link: Cloud workbook must be open in Excel: {fileName}");
                                        }
                                        else
                                        {
                                            warnings.AddWarning($"Text link: Workbook not found: {fileName}");
                                        }
                                        warnings.IncrementFailed();
                                        continue;
                                    }

                                    bool pathWasUpgraded =
                                        !string.Equals(actualWbPath, wbPath, StringComparison.OrdinalIgnoreCase);

                                    try { linkedWs = linkedWb.Worksheets[wsName]; }
                                    catch
                                    {
                                        warnings.AddWarning($"Sheet '{wsName}' not found");
                                        warnings.IncrementFailed();
                                        continue;
                                    }

                                    if (linkedWs == null)
                                    {
                                        warnings.IncrementFailed();
                                        continue;
                                    }

                                    try
                                    {
                                        namedRangeObject = linkedWb.Names.Item(nrName);
                                        linkedExcelCell = namedRangeObject.RefersToRange;
                                    }
                                    catch (Exception ex)
                                    {
                                        warnings.AddWarning($"Named range '{nrName}' error: {ex.Message}");
                                        warnings.IncrementFailed();
                                        CleanExcelObjects(ref linkedWs, ref namedRangeObject);
                                        continue;
                                    }

                                    if (linkedExcelCell == null)
                                    {
                                        warnings.AddWarning($"No cell from '{nrName}'");
                                        warnings.IncrementFailed();
                                        CleanExcelObjects(ref linkedWs, ref namedRangeObject);
                                        continue;
                                    }

                                    string newVal = (linkedExcelCell.Text?.ToString() ?? "").Trim();
                                    string txtToSet = string.IsNullOrEmpty(newVal) ? " " : newVal;

                                    if (currentRun.Text != txtToSet)
                                    {
                                        // Cache font props AND hyperlink info
                                        PowerPoint.Font rFont = currentRun.Font;
                                        string rFontName = rFont.Name;
                                        float rFontSize = rFont.Size;
                                        Office.MsoTriState rBold = rFont.Bold;
                                        Office.MsoTriState rItalic = rFont.Italic;
                                        Office.MsoTriState rUnderline = rFont.Underline;
                                        int rFontColor = rFont.Color.RGB;
                                        Marshal.ReleaseComObject(rFont);
                                        rFont = null;

                                        string savedHyperlinkAddress = currentHyperlink.Address;
                                        string savedScreenTip = currentHyperlink.ScreenTip;

                                        int runStart = currentRun.Start;
                                        int runLength = currentRun.Length;
                                        int runEnd = runStart + runLength - 1;

                                        PowerPoint.Shape parentShape = currentRun.Parent.Parent as PowerPoint.Shape;
                                        if (parentShape == null) continue;

                                        PowerPoint.TextRange fullTextRange = parentShape.TextFrame.TextRange;
                                        int fullTextLength = fullTextRange.Length;
                                        string textBefore = "";
                                        string textAfter = "";

                                        if (runStart > 1)
                                        {
                                            int beforeLength = runStart - 1;
                                            if (beforeLength > 0 && beforeLength <= fullTextLength)
                                            {
                                                PowerPoint.TextRange beforeRange = fullTextRange.Characters(1, beforeLength);
                                                textBefore = beforeRange.Text;
                                                Marshal.ReleaseComObject(beforeRange);
                                            }
                                        }

                                        if (runEnd < fullTextLength)
                                        {
                                            int afterStart = runEnd + 1;
                                            int afterLength = fullTextLength - runEnd;
                                            if (afterStart > 0 && afterStart <= fullTextLength && afterLength > 0)
                                            {
                                                PowerPoint.TextRange afterRange = fullTextRange.Characters(afterStart, afterLength);
                                                textAfter = afterRange.Text;
                                                Marshal.ReleaseComObject(afterRange);
                                            }
                                        }

                                        Marshal.ReleaseComObject(fullTextRange);

                                        currentRun.Delete();

                                        PowerPoint.TextRange shapeText = parentShape.TextFrame.TextRange;
                                        shapeText.Text = textBefore + txtToSet + textAfter;
                                        Marshal.ReleaseComObject(shapeText);

                                        int insertPosition = textBefore.Length + 1;
                                        PowerPoint.TextRange finalTextRange = parentShape.TextFrame.TextRange;
                                        PowerPoint.TextRange newRun = finalTextRange.Characters(insertPosition, txtToSet.Length);

                                        rFont = newRun.Font;
                                        rFont.Name = rFontName;
                                        rFont.Size = rFontSize;
                                        rFont.Bold = rBold;
                                        rFont.Italic = rItalic;
                                        rFont.Underline = rUnderline;
                                        rFont.Color.RGB = rFontColor;
                                        Marshal.ReleaseComObject(rFont);

                                        PowerPoint.ActionSetting newActionSetting = newRun.ActionSettings[PowerPoint.PpMouseActivation.ppMouseClick];
                                        newActionSetting.Action = PowerPoint.PpActionType.ppActionHyperlink;

                                        PowerPoint.Hyperlink newHyperlink = newActionSetting.Hyperlink;

                                        if (pathWasUpgraded)
                                        {
                                            string newAddress =
                                                "pragmatics-textlink://" +
                                                uri.Host +
                                                "?wb=" + Uri.EscapeDataString(actualWbPath) +
                                                "&ws=" + Uri.EscapeDataString(wsName) +
                                                "&nr=" + Uri.EscapeDataString(nrName);
                                            newHyperlink.Address = newAddress;
                                        }
                                        else
                                        {
                                            newHyperlink.Address = savedHyperlinkAddress;
                                        }

                                        newHyperlink.SubAddress = "";
                                        newHyperlink.ScreenTip = savedScreenTip;

                                        Marshal.ReleaseComObject(newHyperlink);
                                        Marshal.ReleaseComObject(newActionSetting);

                                        try
                                        {
                                            if (parentShape.HasTextFrame == Office.MsoTriState.msoTrue)
                                            {
                                                var tf2 = parentShape.TextFrame2;
                                                var tr2 = tf2.TextRange;
                                                var insertedRange2 = tr2.Characters[insertPosition, txtToSet.Length];
                                                var font2 = insertedRange2.Font;

                                                font2.Fill.ForeColor.RGB = rFontColor;

                                                if (rUnderline == Office.MsoTriState.msoFalse)
                                                {
                                                    int backgroundColorToMatch = GetBackgroundColor(parentShape);
                                                    font2.UnderlineStyle = Office.MsoTextUnderlineType.msoUnderlineSingleLine;
                                                    font2.UnderlineColor.RGB = backgroundColorToMatch;
                                                }
                                                else
                                                {
                                                    font2.UnderlineStyle = Office.MsoTextUnderlineType.msoUnderlineSingleLine;
                                                }

                                                Marshal.ReleaseComObject(font2);
                                                Marshal.ReleaseComObject(insertedRange2);
                                                Marshal.ReleaseComObject(tr2);
                                                Marshal.ReleaseComObject(tf2);
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"TextFrame2 formatting failed: {ex.Message}");
                                        }

                                        Marshal.ReleaseComObject(newRun);
                                        Marshal.ReleaseComObject(finalTextRange);

                                        warnings.IncrementUpdated();
                                    }
                                    else
                                    {
                                        if (pathWasUpgraded)
                                        {
                                            string newAddress =
                                                "pragmatics-textlink://" +
                                                uri.Host +
                                                "?wb=" + Uri.EscapeDataString(actualWbPath) +
                                                "&ws=" + Uri.EscapeDataString(wsName) +
                                                "&nr=" + Uri.EscapeDataString(nrName);

                                            currentHyperlink.Address = newAddress;
                                            currentHyperlink.SubAddress = "";
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            warnings.AddWarning($"Error updating text link: {ex.Message}");
                            warnings.IncrementFailed();
                        }
                        finally
                        {
                            CleanRunObjects(ref currentHyperlink, ref runActionSetting, ref currentRun);
                            CleanExcelObjects(ref linkedWs, ref namedRangeObject, ref linkedExcelCell);
                        }
                    }
                }
            }
            finally
            {
                CleanShapeTextObjects(shapeTextRange, textFrame, shape);
            }
        }

        private static int GetBackgroundColor(PowerPoint.Shape shapeForText)
        {
            int backgroundColorToMatch = ColorTranslator.ToOle(System.Drawing.Color.White);
            if (shapeForText == null) return backgroundColorToMatch;

            PowerPoint.Slide hostingSlide = null;
            PowerPoint.CustomLayout hostingLayout = null;
            PowerPoint.Master hostingMaster = null;

            try
            {
                if (shapeForText.Fill.Visible == Office.MsoTriState.msoTrue &&
                    shapeForText.Fill.Type == Office.MsoFillType.msoFillSolid)
                    return shapeForText.Fill.ForeColor.RGB;

                object currentParentObj = shapeForText.Parent;
                object tempParentToRelease = null;

                while (currentParentObj != null)
                {
                    if (tempParentToRelease != null)
                    {
                        Marshal.ReleaseComObject(tempParentToRelease);
                        tempParentToRelease = null;
                    }

                    if (currentParentObj is PowerPoint.Slide s) { hostingSlide = s; break; }
                    if (currentParentObj is PowerPoint.CustomLayout cl) { hostingLayout = cl; break; }
                    if (currentParentObj is PowerPoint.Master m) { hostingMaster = m; break; }

                    tempParentToRelease = currentParentObj;
                    try
                    {
                        dynamic dyn = currentParentObj;
                        currentParentObj = dyn.Parent;
                    }
                    catch { currentParentObj = null; }
                }

                if (tempParentToRelease != null &&
                    tempParentToRelease != hostingSlide &&
                    tempParentToRelease != hostingLayout &&
                    tempParentToRelease != hostingMaster)
                    Marshal.ReleaseComObject(tempParentToRelease);

                if (hostingSlide != null)
                {
                    if (hostingSlide.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                        hostingSlide.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                        hostingSlide.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                        return hostingSlide.Background.Fill.ForeColor.RGB;
                    else
                    {
                        PowerPoint.CustomLayout sl = null;
                        PowerPoint.Master sm = null;
                        try
                        {
                            sl = hostingSlide.CustomLayout;
                            if (sl.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                                sl.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                                sl.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                                return sl.Background.Fill.ForeColor.RGB;
                            else
                            {
                                if (sl.Parent is PowerPoint.Master masterFromLayout) sm = masterFromLayout;
                                if (sm != null &&
                                    sm.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                                    sm.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                                    return sm.Background.Fill.ForeColor.RGB;
                            }
                        }
                        finally
                        {
                            if (sl != null) Marshal.ReleaseComObject(sl);
                            if (sm != null) Marshal.ReleaseComObject(sm);
                        }
                    }
                }
                else if (hostingLayout != null)
                {
                    if (hostingLayout.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                        hostingLayout.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                        hostingLayout.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                        return hostingLayout.Background.Fill.ForeColor.RGB;
                    else
                    {
                        PowerPoint.Master layoutMaster = null;
                        try
                        {
                            if (hostingLayout.Parent is PowerPoint.Master masterFromLayout)
                                layoutMaster = masterFromLayout;

                            if (layoutMaster != null &&
                                layoutMaster.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                                layoutMaster.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                                return layoutMaster.Background.Fill.ForeColor.RGB;
                        }
                        finally
                        {
                            if (layoutMaster != null) Marshal.ReleaseComObject(layoutMaster);
                        }
                    }
                }
                else if (hostingMaster != null)
                {
                    if (hostingMaster.Background.Fill.Visible == Office.MsoTriState.msoTrue &&
                        hostingMaster.Background.Fill.Type == Office.MsoFillType.msoFillSolid)
                        return hostingMaster.Background.Fill.ForeColor.RGB;
                }

                return backgroundColorToMatch;
            }
            finally
            {
                if (hostingSlide != null) Marshal.ReleaseComObject(hostingSlide);
                if (hostingLayout != null) Marshal.ReleaseComObject(hostingLayout);
                if (hostingMaster != null) Marshal.ReleaseComObject(hostingMaster);
            }
        }

        /* ═══════════════════════  RANGE LINKS  ═══════════════════════ */
        private static void ProcessRangeLink(
                    PowerPoint.Shape shape,
                    PowerPoint.Slide slide,
                    int slideIndex,
                    Excel.Application excelApp,
                    WorkbookCache wbCache,
                    WarningCollector warnings,
                    SlideActivationTracker slideTracker,
                    PowerPoint.Application ppApp)
        {
            string workbookPath = shape.Tags["Pragmatics_SourceWorkbookPath"];
            string worksheetName = shape.Tags["Pragmatics_SourceWorksheet"];
            string rangeName = shape.Tags["Pragmatics_RangeName"];
            string importType = shape.Tags["Pragmatics_ImportType"];
            string sizeMode = shape.Tags["Pragmatics_SizeMode"];

            if (string.IsNullOrEmpty(workbookPath) || string.IsNullOrEmpty(worksheetName) || string.IsNullOrEmpty(rangeName))
                return;

            // 1. Capture Target Dimensions
            float shapeW = shape.Width;
            float shapeH = shape.Height;
            float shapeL = shape.Left;
            float shapeT = shape.Top;

            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;
            Excel.Name namedRange = null;
            Excel.Range currentRange = null;

            try
            {
                wb = wbCache.GetOrOpen(workbookPath, out bool _, out string actualWbPath);
                if (wb == null)
                {
                    warnings.AddWarning($"Range link: Workbook not found: {CloudPathHelper.GetFileName(workbookPath)}");
                    warnings.IncrementFailed();
                    return;
                }

                try { ws = wb.Worksheets[worksheetName]; } catch { }
                if (ws == null)
                {
                    warnings.AddWarning($"Range link: Worksheet '{worksheetName}' not found");
                    warnings.IncrementFailed();
                    return;
                }

                try
                {
                    namedRange = wb.Names.Item(rangeName);
                    currentRange = namedRange.RefersToRange;
                }
                catch { }

                if (currentRange == null)
                {
                    warnings.AddWarning($"Range link: Range '{rangeName}' is invalid");
                    warnings.IncrementFailed();
                    return;
                }

                wbCache.PrepareWorkbookForCopy(wb);

                PowerPoint.Shape newShape = null;
                PowerPoint.ShapeRange newShapeRange = null;

                // --- PICTURE UPDATE LOGIC ---
                if (importType == "Picture")
                {
                    bool originalDisplayPageBreaks = false;
                    try { originalDisplayPageBreaks = ws.DisplayPageBreaks; ws.DisplayPageBreaks = false; } catch { }

                    try
                    {
                        Action copyPicture = () => currentRange.CopyPicture(Excel.XlPictureAppearance.xlScreen, Excel.XlCopyPictureFormat.xlPicture);

                        bool needsResize = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase);

                        if (needsResize)
                        {
                            // UPDATE 1: Pass 'true' for scaleHeight to ensure vertical fit
                            RangeLinkUtility.CopyRangeWithTempResize(currentRange, shapeW, shapeH, true, copyPicture);
                        }
                        else
                        {
                            copyPicture();
                        }

                        ClipboardHelper.WaitForClipboardReady();
                        slideTracker.EnsureSlideActive(slideIndex);

                        newShapeRange = ClipboardHelper.RetryPaste(() => slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile));
                        newShape = newShapeRange[1];
                    }
                    finally
                    {
                        try { ws.DisplayPageBreaks = originalDisplayPageBreaks; } catch { }
                    }
                }
                // --- TABLE UPDATE LOGIC ---
                else if (importType == "Table")
                {
                    // 1. READ STRUCTURE (Shared Utility)
                    RangeLinkUtility.GetVisibleRowsAndColumns(currentRange, out bool[] visibleRows, out bool[] visibleCols, out int rows, out int cols);

                    if (rows == 0 || cols == 0)
                    {
                        warnings.AddWarning("No visible cells in range");
                        warnings.IncrementFailed();
                        return;
                    }

                    var cellData = RangeLinkUtility.BulkReadCellDataVisible(currentRange, visibleRows, visibleCols, rows, cols);
                    RangeLinkUtility.GetRangeDimensionsWithVisibility(currentRange, visibleRows, visibleCols, out float sumColW, out float sumRowH, out float[] colWidths, out float[] rowHeights);

                    // UPDATE 2: Wrap Copy in Resize Logic for Tables
                    // If "Match All", we squeeze Excel row heights before copying so text wrap calculates correctly
                    bool scaleHeight = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase);

                    Action copyOp = () => {
                        currentRange.Copy();
                        System.Threading.Thread.Sleep(50);
                    };

                    // We always apply width resizing (using shapeW) to ensure columns fit constraints.
                    // We only apply height resizing (using shapeH) if scaleHeight is true.
                    RangeLinkUtility.CopyRangeWithTempResize(currentRange, shapeW, shapeH, scaleHeight, copyOp);

                    ClipboardHelper.WaitForClipboardReady();

                    // Ensure slide is active so Paste works
                    slideTracker.EnsureSlideActive(slideIndex);

                    // 3. PASTE (Source Formatting)
                    try
                    {
                        ppApp.CommandBars.ExecuteMso("PasteSourceFormatting");
                        System.Windows.Forms.Application.DoEvents();
                        if (ppApp.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                            newShape = ppApp.ActiveWindow.Selection.ShapeRange[1];
                    }
                    catch
                    {
                        newShapeRange = ClipboardHelper.RetryPaste(() => slide.Shapes.Paste());
                        newShape = newShapeRange[1];
                    }

                    if (newShape != null && newShape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var table = newShape.Table;

                        // Cleanup Ghost Cols
                        while (table.Columns.Count > cols)
                        {
                            try { table.Columns[table.Columns.Count].Delete(); } catch { break; }
                        }

                        // 4. APPLY STRUCTURAL FIXES
                        for (int r = 0; r < rows; r++)
                        {
                            for (int c = 0; c < cols; c++)
                            {
                                var data = cellData[r, c];
                                if (data == null) continue;

                                if (data.IsMerged && data.MergeRowSpan > 0)
                                {
                                    try { table.Cell(r + 1, c + 1).Merge(table.Cell(r + data.MergeRowSpan, c + data.MergeColSpan)); } catch { }
                                }

                                if (!data.IsMerged || (data.MergeRowSpan > 0))
                                {
                                    try { RangeLinkUtility.ApplyStructuralFixes(table.Cell(r + 1, c + 1), data.NumberFormat); } catch { }
                                }
                            }
                        }

                        // 5. RESTORE DIMENSIONS
                        newShape.Width = shapeW;

                        // Restore internal column ratios
                        for (int c = 1; c <= cols; c++)
                        {
                            if (c <= table.Columns.Count)
                                table.Columns[c].Width = (colWidths[c - 1] / sumColW) * shapeW;
                        }

                        // UPDATE 3: Force Height if "Match All"
                        if (scaleHeight)
                        {
                            newShape.Height = shapeH;
                        }
                    }
                }

                // --- FINALIZE UPDATE ---
                if (newShape != null)
                {
                    bool pathWasUpgraded = !string.Equals(actualWbPath, workbookPath, StringComparison.OrdinalIgnoreCase);

                    // Transfer Tags
                    for (int i = 1; i <= shape.Tags.Count; i++)
                    {
                        string tagName = shape.Tags.Name(i);
                        string tagValue = shape.Tags.Value(i);

                        if (pathWasUpgraded && string.Equals(tagName, "Pragmatics_SourceWorkbookPath", StringComparison.OrdinalIgnoreCase))
                            newShape.Tags.Add(tagName, actualWbPath);
                        else
                            newShape.Tags.Add(tagName, tagValue);
                    }

                    // Position and Delete Old
                    newShape.Left = shapeL;
                    newShape.Top = shapeT;

                    // If Picture, apply size policy again (redundant safety)
                    if (importType == "Picture")
                    {
                        SizeModeUtil.ApplyFromSource(shape, newShape);
                    }
                    // For Tables, we already set Width (and optionally Height) above in the table logic block

                    shape.Delete();
                    newShape.Select();
                    warnings.IncrementUpdated();

                    if (newShape != null) Marshal.ReleaseComObject(newShape);
                }

                if (newShapeRange != null) Marshal.ReleaseComObject(newShapeRange);
            }
            catch (Exception ex)
            {
                warnings.AddWarning($"Failed to update range '{rangeName}': {ex.Message}");
                warnings.IncrementFailed();
            }
            finally
            {
                if (currentRange != null) Marshal.ReleaseComObject(currentRange);
                if (namedRange != null) Marshal.ReleaseComObject(namedRange);
                if (ws != null) Marshal.ReleaseComObject(ws);
            }
        }

        /* ═══════════════════════  CHART LINKS  ═══════════════════════ */
        private static void ProcessChartLink(PowerPoint.Shape shape, PowerPoint.Slide slide,
            int slideIndex, Excel.Application excelApp, WorkbookCache wbCache,
            WarningCollector warnings, SlideActivationTracker slideTracker)
        {
            string workbookPath = shape.Tags["Pragmatics_SourceWorkbookPath"];
            string worksheetName = shape.Tags["Pragmatics_SourceWorksheet"];
            string chartName = shape.Tags["Pragmatics_SourceChart"];
            string importType = shape.Tags["Pragmatics_ImportType"];
            string sizeMode = shape.Tags["Pragmatics_SizeMode"];

            if (string.IsNullOrEmpty(workbookPath) ||
                string.IsNullOrEmpty(worksheetName) ||
                string.IsNullOrEmpty(chartName))
                return;

            if (string.IsNullOrEmpty(importType)) importType = "Picture";

            float shapeW = shape.Width;
            float shapeH = shape.Height;
            float shapeL = shape.Left;
            float shapeT = shape.Top;

            Excel.Workbook wb = null;
            Excel.Chart chartToProcess = null;
            Excel.ChartObject embeddedChartObject = null;
            Excel.Worksheet sheetWithChart = null;
            Excel.Chart chartSheetObject = null;

            try
            {
                wb = wbCache.GetOrOpen(workbookPath, out bool _, out string actualWbPath);
                if (wb == null)
                {
                    string fileName = CloudPathHelper.GetFileName(workbookPath);
                    if (CloudPathHelper.IsCloudPath(workbookPath))
                    {
                        warnings.AddWarning($"Chart link: Cloud workbook must be open in Excel: {fileName}");
                    }
                    else
                    {
                        warnings.AddWarning($"Chart link: Workbook not found: {fileName}");
                    }
                    warnings.IncrementFailed();
                    return;
                }

                object excelSheetItem = null;
                try
                {
                    excelSheetItem = wb.Sheets[worksheetName];
                }
                catch (COMException)
                {
                    warnings.AddWarning($"Chart link: Sheet '{worksheetName}' not found");
                    warnings.IncrementFailed();
                    return;
                }

                if (excelSheetItem is Excel.Worksheet excelWs)
                {
                    sheetWithChart = excelWs;
                    bool found = false;
                    foreach (Excel.ChartObject co_iter in sheetWithChart.ChartObjects())
                    {
                        if (co_iter.Name == chartName)
                        {
                            embeddedChartObject = co_iter;
                            chartToProcess = embeddedChartObject.Chart;
                            found = true;
                            break;
                        }
                        Marshal.ReleaseComObject(co_iter);
                    }
                    if (!found)
                    {
                        warnings.AddWarning($"Chart '{chartName}' not found on '{worksheetName}'");
                        warnings.IncrementFailed();
                        return;
                    }
                }
                else if (excelSheetItem is Excel.Chart excelChartSheet)
                {
                    chartSheetObject = excelChartSheet;
                    if (chartSheetObject.Name == chartName)
                    {
                        chartToProcess = chartSheetObject;
                    }
                    else
                    {
                        warnings.AddWarning($"Chart sheet mismatch: '{chartSheetObject.Name}' != '{chartName}'");
                        warnings.IncrementFailed();
                        return;
                    }
                }
                else
                {
                    if (excelSheetItem != null) Marshal.ReleaseComObject(excelSheetItem);
                    warnings.AddWarning($"'{worksheetName}' is neither worksheet nor chart sheet");
                    warnings.IncrementFailed();
                    return;
                }

                if (excelSheetItem != null &&
                    !(excelSheetItem is Excel.Worksheet) &&
                    !(excelSheetItem is Excel.Chart))
                {
                    Marshal.ReleaseComObject(excelSheetItem);
                }

                PowerPoint.Shape newShape = null;
                PowerPoint.ShapeRange newShapeRange = null;

                try
                {
                    if (importType == "Picture")
                    {
                        Action copyPic = () => chartToProcess.CopyPicture(
                            Excel.XlPictureAppearance.xlScreen,
                            Excel.XlCopyPictureFormat.xlPicture);

                        bool needsResize = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase)
                            && embeddedChartObject != null
                            && SizeModeUtil.NeedsResize(embeddedChartObject, shapeW, shapeH);

                        if (needsResize)
                        {
                            CopyOrExportWithTempResize(embeddedChartObject, shapeW, shapeH, copyPic);
                        }
                        else
                        {
                            copyPic();
                        }

                        ClipboardHelper.WaitForClipboardReady();

                        // Only activate slide right before paste
                        slideTracker.EnsureSlideActive(slideIndex);

                        newShapeRange = ClipboardHelper.RetryPaste(() =>
                            slide.Shapes.PasteSpecial(PowerPoint.PpPasteDataType.ppPasteEnhancedMetafile));
                        if (newShapeRange != null && newShapeRange.Count > 0)
                            newShape = newShapeRange[1];
                    }
                    else if (importType == "SVG")
                    {
                        string tempSvgPath = System.IO.Path.GetTempFileName() + ".svg";
                        try
                        {
                            Action exportSvg = () => chartToProcess.Export(tempSvgPath, "SVG");

                            bool needsResize = string.Equals(sizeMode, "Match All", StringComparison.OrdinalIgnoreCase)
                                && embeddedChartObject != null
                                && SizeModeUtil.NeedsResize(embeddedChartObject, shapeW, shapeH);

                            if (needsResize)
                            {
                                CopyOrExportWithTempResize(embeddedChartObject, shapeW, shapeH, exportSvg);
                            }
                            else
                            {
                                exportSvg();
                            }

                            newShape = slide.Shapes.AddPicture(tempSvgPath,
                                Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue,
                                shapeL, shapeT, -1, -1);
                        }
                        finally
                        {
                            if (System.IO.File.Exists(tempSvgPath))
                            {
                                try { System.IO.File.Delete(tempSvgPath); } catch { }
                            }
                        }
                    }
                    else if (importType == "Chart")
                    {
                        if (embeddedChartObject != null) embeddedChartObject.Copy();
                        else chartToProcess.Copy();

                        ClipboardHelper.WaitForClipboardReady();

                        // Only activate slide right before paste
                        slideTracker.EnsureSlideActive(slideIndex);

                        try
                        {
                            // Try generic Paste first, fall back to PasteOLEObject if needed
                            newShapeRange = ClipboardHelper.RetryPaste(() => slide.Shapes.Paste());
                            if (newShapeRange != null && newShapeRange.Count > 0)
                                newShape = newShapeRange[1];

                            if (newShape == null || newShape.Type != Office.MsoShapeType.msoChart)
                            {
                                if (newShape != null)
                                {
                                    newShape.Delete();
                                    Marshal.ReleaseComObject(newShape);
                                    newShape = null;
                                }
                                if (newShapeRange != null)
                                {
                                    Marshal.ReleaseComObject(newShapeRange);
                                    newShapeRange = null;
                                }

                                // Re-copy and try OLE paste
                                if (embeddedChartObject != null) embeddedChartObject.Copy();
                                else chartToProcess.Copy();
                                ClipboardHelper.WaitForClipboardReady();

                                newShapeRange = ClipboardHelper.RetryPaste(() =>
                                    slide.Shapes.PasteSpecial(DataType: PowerPoint.PpPasteDataType.ppPasteOLEObject));
                                if (newShapeRange != null && newShapeRange.Count > 0)
                                    newShape = newShapeRange[1];
                            }
                        }
                        catch (COMException pEx)
                        {
                            warnings.AddWarning($"Paste failed for native chart: {pEx.Message}");
                            throw;
                        }

                        if (newShape == null || newShape.Type != Office.MsoShapeType.msoChart)
                        {
                            if (newShape != null)
                            {
                                newShape.Delete();
                                Marshal.ReleaseComObject(newShape);
                            }
                            warnings.AddWarning($"Pasted object is not a native chart");
                            warnings.IncrementFailed();
                            return;
                        }

                        // Break Excel link to prevent auto-updating when source file is open
                        // Our add-in manages updates through tags, not Excel's live linking
                        ChartLinkUtility.BreakChartLink(newShape);
                    }
                    else
                    {
                        warnings.AddWarning($"Unsupported import type '{importType}'");
                        warnings.IncrementFailed();
                        return;
                    }

                    if (newShape == null)
                    {
                        warnings.AddWarning("Failed to paste/create new shape");
                        warnings.IncrementFailed();
                        return;
                    }

                    bool pathWasUpgraded = !string.Equals(actualWbPath, workbookPath, StringComparison.OrdinalIgnoreCase);

                    for (int i = 1; i <= shape.Tags.Count; i++)
                    {
                        string tagName = shape.Tags.Name(i);
                        string tagValue = shape.Tags.Value(i);

                        if (pathWasUpgraded && string.Equals(tagName, "Pragmatics_SourceWorkbookPath", StringComparison.OrdinalIgnoreCase))
                        {
                            newShape.Tags.Add(tagName, actualWbPath);
                            System.Diagnostics.Debug.WriteLine(
                                $"Updated chart workbook path in tags: {workbookPath} → {actualWbPath}");
                        }
                        else
                        {
                            newShape.Tags.Add(tagName, tagValue);
                        }
                    }

                    newShape.Tags.Add("Pragmatics_ImportType", importType);

                    SizeModeUtil.ApplyFromSource(shape, newShape);

                    string oldShapeName = shape.Name;
                    shape.Delete();
                    newShape.Select();
                    warnings.IncrementUpdated();
                    System.Diagnostics.Debug.WriteLine($"Old shape '{oldShapeName}' deleted and replaced by new chart.");
                }
                finally
                {
                    if (newShape != null) Marshal.ReleaseComObject(newShape);
                    if (newShapeRange != null) Marshal.ReleaseComObject(newShapeRange);
                }
            }
            catch (Exception ex)
            {
                warnings.AddWarning($"Failed to update chart: {ex.Message}");
                warnings.IncrementFailed();
            }
            finally
            {
                if (chartToProcess != null && Marshal.IsComObject(chartToProcess))
                    Marshal.ReleaseComObject(chartToProcess);
                if (embeddedChartObject != null) Marshal.ReleaseComObject(embeddedChartObject);
                if (sheetWithChart != null) Marshal.ReleaseComObject(sheetWithChart);
                if (chartSheetObject != null) Marshal.ReleaseComObject(chartSheetObject);
            }
        }

        /* ═══════════════════════  HELPER CLEANUP METHODS  ═══════════════════════ */
        private static void CleanExcelObjects(ref Excel.Worksheet ws, ref Excel.Name nr)
        {
            if (nr != null) { Marshal.ReleaseComObject(nr); nr = null; }
            if (ws != null) { Marshal.ReleaseComObject(ws); ws = null; }
        }

        private static void CleanExcelObjects(ref Excel.Worksheet ws, ref Excel.Name nr, ref Excel.Range cell)
        {
            if (cell != null) { Marshal.ReleaseComObject(cell); cell = null; }
            CleanExcelObjects(ref ws, ref nr);
        }

        private static void CleanRunObjects(ref PowerPoint.Hyperlink hl,
                                           ref PowerPoint.ActionSetting act,
                                           ref PowerPoint.TextRange run)
        {
            if (hl != null) { Marshal.ReleaseComObject(hl); hl = null; }
            if (act != null) { Marshal.ReleaseComObject(act); act = null; }
            if (run != null) { Marshal.ReleaseComObject(run); run = null; }
        }

        private static void CleanShapeTextObjects(PowerPoint.TextRange tr,
                                                  PowerPoint.TextFrame tf,
                                                  PowerPoint.Shape shp)
        {
            if (tr != null && Marshal.IsComObject(tr))
                Marshal.ReleaseComObject(tr);
            if (tf != null && Marshal.IsComObject(tf))
                Marshal.ReleaseComObject(tf);
            if (shp != null && Marshal.IsComObject(shp))
                Marshal.ReleaseComObject(shp);
        }
    }

    internal class WaitForm : Form
    {
        private readonly Label titleLabel;
        private readonly Label messageLabel;
        private readonly Label progressLabel;
        private readonly ProgressBar progressBar;
        private readonly Button closeButton;
        private readonly Panel contentPanel;
        private readonly Panel headerPanel;
        private readonly Panel buttonPanel;
        private readonly Panel progressPanel;

        private readonly IntPtr _pptHwnd;
        private int _totalItems;
        private int _completedItems;

        private static readonly Color BackgroundColor = Color.FromArgb(250, 250, 250);
        private static readonly Color HeaderColor = Color.FromArgb(245, 247, 250);
        private static readonly Color AccentColor = Color.FromArgb(0, 0, 0);
        private static readonly Color TextPrimaryColor = Color.FromArgb(32, 33, 36);
        private static readonly Color TextSecondaryColor = Color.FromArgb(95, 99, 104);

        public WaitForm(IntPtr pptHwnd)
        {
            _pptHwnd = pptHwnd;

            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.Width = 460;
            this.Height = 280;  // Increased height for progress bar
            this.BackColor = BackgroundColor;
            this.Padding = new Padding(1);
            this.KeyPreview = true;  // Allow form to receive key events first

            this.Paint += (s, e) =>
            {
                using (Pen borderPen = new Pen(Color.FromArgb(200, 200, 200), 1))
                {
                    e.Graphics.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
                }
            };

            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 60,
                BackColor = HeaderColor,
                Padding = new Padding(20, 16, 20, 16)
            };

            titleLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                Text = "Updating Links",
                Font = new Font("Segoe UI Semibold", 13f, FontStyle.Bold),
                ForeColor = TextPrimaryColor,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };

            headerPanel.Controls.Add(titleLabel);

            // Progress panel with bar and counter
            progressPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 50,
                BackColor = BackgroundColor,
                Padding = new Padding(24, 8, 24, 8)
            };

            progressLabel = new Label
            {
                Dock = DockStyle.Top,
                Height = 18,
                AutoSize = false,
                Text = "",
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                ForeColor = TextSecondaryColor,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };

            progressBar = new ProgressBar
            {
                Dock = DockStyle.Bottom,
                Height = 8,
                Style = ProgressBarStyle.Continuous,
                Minimum = 0,
                Maximum = 100,
                Value = 0
            };

            progressPanel.Controls.Add(progressLabel);
            progressPanel.Controls.Add(progressBar);

            contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(24, 12, 24, 12),
                BackColor = BackgroundColor
            };

            messageLabel = new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = false,
                TextAlign = ContentAlignment.TopLeft,
                Font = new Font("Segoe UI", 10f, FontStyle.Regular),
                ForeColor = TextSecondaryColor,
                Text = "Scanning for linked objects...\nPlease wait."
            };

            contentPanel.Controls.Add(messageLabel);

            buttonPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 70,
                BackColor = BackgroundColor,
                Padding = new Padding(24, 12, 24, 20)
            };

            closeButton = new ModernButton
            {
                Text = "Close",
                Visible = false,
                Enabled = true,
                Width = 100,
                Height = 36,
                FlatStyle = FlatStyle.Flat,
                BackColor = AccentColor,
                ForeColor = Color.White,
                Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                DialogResult = DialogResult.OK
            };

            closeButton.FlatAppearance.BorderSize = 0;
            closeButton.FlatAppearance.MouseOverBackColor = Color.FromArgb(0, 103, 184);
            closeButton.FlatAppearance.MouseDownBackColor = Color.FromArgb(0, 90, 158);

            closeButton.Click += (_, __) => this.Close();

            closeButton.Location = new Point(
                (buttonPanel.Width - closeButton.Width) / 2,
                12
            );

            buttonPanel.Resize += (_, __) =>
            {
                if (closeButton != null)
                {
                    closeButton.Location = new Point(
                        (buttonPanel.Width - closeButton.Width) / 2,
                        12
                    );
                }
            };

            buttonPanel.Controls.Add(closeButton);

            // Order matters for Dock: last added is on top
            // We want: header (top), progress (below header), content (fill), button (bottom)
            this.Controls.Add(buttonPanel);    // Dock.Bottom
            this.Controls.Add(contentPanel);   // Dock.Fill
            this.Controls.Add(progressPanel);  // Dock.Top (below header)
            this.Controls.Add(headerPanel);    // Dock.Top

            this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
                          ControlStyles.AllPaintingInWmPaint |
                          ControlStyles.UserPaint, true);
        }

        public void SetStatus(string text)
        {
            if (this.IsHandleCreated && this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => messageLabel.Text = text));
            }
            else
            {
                messageLabel.Text = text;
            }
        }

        /// <summary>
        /// Initializes the progress tracking with the total number of items.
        /// </summary>
        public void InitializeProgress(int total)
        {
            Action ui = () =>
            {
                _totalItems = total;
                _completedItems = 0;
                progressBar.Value = 0;
                progressLabel.Text = $"0 of {total}";
            };

            if (this.IsHandleCreated && this.InvokeRequired)
                this.BeginInvoke(ui);
            else
                ui();
        }

        /// <summary>
        /// Updates the progress counter and bar.
        /// </summary>
        public void SetProgress(int completed, string statusText = null)
        {
            Action ui = () =>
            {
                _completedItems = completed;
                int percent = _totalItems > 0 ? (int)Math.Round(100.0 * completed / _totalItems) : 0;
                percent = Math.Min(100, Math.Max(0, percent));

                progressBar.Value = percent;
                progressLabel.Text = $"{completed} of {_totalItems}";

                if (!string.IsNullOrEmpty(statusText))
                    messageLabel.Text = statusText;
            };

            if (this.IsHandleCreated && this.InvokeRequired)
                this.BeginInvoke(ui);
            else
                ui();
        }

        public void ShowCompletion(string title, string body)
        {
            Action ui = () =>
            {
                titleLabel.Text = string.IsNullOrWhiteSpace(title) ? "Complete" : title;
                messageLabel.Text = body;
                messageLabel.Font = new Font("Segoe UI", 10f, FontStyle.Regular);
                messageLabel.ForeColor = TextPrimaryColor;

                // Set progress to 100% and show completion
                progressBar.Value = 100;
                progressLabel.Text = "Complete";

                closeButton.Visible = true;
                this.AcceptButton = closeButton;  // Enter closes
                this.CancelButton = closeButton;  // Escape closes

                this.TopMost = true;
                Win32Focus.ActivateWindow(this.Handle);
                closeButton.Focus();
            };

            if (this.IsHandleCreated && this.InvokeRequired)
            {
                this.BeginInvoke(ui);
            }
            else
            {
                ui();
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            base.OnFormClosed(e);
            Win32Focus.ActivateWindow(_pptHwnd);
        }

        private class ModernButton : Button
        {
            public ModernButton()
            {
                this.FlatStyle = FlatStyle.Flat;
                this.FlatAppearance.BorderSize = 0;
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);

                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                using (GraphicsPath path = GetRoundedRect(this.ClientRectangle, 6))
                {
                    this.Region = new Region(path);

                    Color bgColor = this.BackColor;
                    if (!this.Enabled)
                    {
                        bgColor = Color.FromArgb(200, 200, 200);
                    }

                    using (SolidBrush brush = new SolidBrush(bgColor))
                    {
                        e.Graphics.FillPath(brush, path);
                    }
                }

                TextRenderer.DrawText(e.Graphics, this.Text, this.Font, this.ClientRectangle,
                    this.ForeColor, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
            }

            private GraphicsPath GetRoundedRect(Rectangle rect, int radius)
            {
                GraphicsPath path = new GraphicsPath();
                int diameter = radius * 2;

                path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
                path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
                path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
                path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
                path.CloseFigure();

                return path;
            }
        }
    }
}