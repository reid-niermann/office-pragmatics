// PerformanceHelpers.cs (.NET Framework 4.8 / C# 7.3)

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_PowerPoint.ExcelLinking
{
    /// <summary>
    /// RAII wrapper that sets Excel to maximum performance mode and restores on dispose.
    /// This is critical for large workbooks with circular references.
    /// </summary>
    internal sealed class ExcelPerformanceScope : IDisposable
    {
        private readonly Excel.Application _app;
        private Excel.XlCalculation _originalCalc;
        private bool _originalScreenUpdating;
        private bool _originalDisplayAlerts;
        private bool _originalEnableEvents;
        private bool _originalAskToUpdateLinks;
        private bool _originalDisplayStatusBar;
        private bool _calcCached;
        private bool _disposed;

        public ExcelPerformanceScope(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));

            // Cache original settings - some properties can throw when no workbooks are open
            try { _originalCalc = app.Calculation; _calcCached = true; }
            catch { _originalCalc = Excel.XlCalculation.xlCalculationAutomatic; _calcCached = false; }

            try { _originalScreenUpdating = app.ScreenUpdating; } catch { _originalScreenUpdating = true; }
            try { _originalDisplayAlerts = app.DisplayAlerts; } catch { _originalDisplayAlerts = true; }
            try { _originalEnableEvents = app.EnableEvents; } catch { _originalEnableEvents = true; }
            try { _originalAskToUpdateLinks = app.AskToUpdateLinks; } catch { _originalAskToUpdateLinks = true; }
            try { _originalDisplayStatusBar = app.DisplayStatusBar; } catch { _originalDisplayStatusBar = true; }

            // Set performance mode - ORDER MATTERS
            // Disable calculation FIRST to prevent recalcs during other changes
            try { app.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }
            try { app.EnableEvents = false; } catch { }
            try { app.ScreenUpdating = false; } catch { }
            try { app.DisplayAlerts = false; } catch { }
            try { app.AskToUpdateLinks = false; } catch { }      // Suppress "Update Links" dialog
            try { app.DisplayStatusBar = false; } catch { }      // Minor perf gain
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;

            try
            {
                // Restore in reverse order
                try { _app.DisplayStatusBar = _originalDisplayStatusBar; } catch { }
                try { _app.AskToUpdateLinks = _originalAskToUpdateLinks; } catch { }
                try { _app.DisplayAlerts = _originalDisplayAlerts; } catch { }
                try { _app.ScreenUpdating = _originalScreenUpdating; } catch { }
                try { _app.EnableEvents = _originalEnableEvents; } catch { }

                // Only restore Calculation if we successfully cached it
                if (_calcCached)
                {
                    try { _app.Calculation = _originalCalc; } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error restoring Excel settings: {ex.Message}");
            }
        }
    }

    /// <summary>
    /// Clipboard helper with optimized retry logic and adaptive waiting.
    /// </summary>
    internal static class ClipboardHelper
    {
        /// <summary>
        /// Waits for clipboard to contain image data with adaptive backoff.
        /// </summary>
        /// <param name="maxWaitMs">Maximum total wait time in milliseconds</param>
        /// <param name="initialRetryMs">Initial retry interval</param>
        /// <returns>True if clipboard is ready, false if timeout</returns>
        public static bool WaitForClipboardReady(int maxWaitMs = 200, int initialRetryMs = 10)
        {
            int elapsed = 0;
            int currentWait = initialRetryMs;

            while (elapsed < maxWaitMs)
            {
                try
                {
                    // Check multiple clipboard formats
                    if (Clipboard.ContainsImage() ||
                        Clipboard.ContainsData("EnhancedMetafile") ||
                        Clipboard.ContainsData(DataFormats.MetafilePict) ||
                        Clipboard.ContainsData("Embed Source"))
                    {
                        return true;
                    }
                }
                catch
                {
                    // Clipboard access can throw if another process is using it
                }

                Thread.Sleep(currentWait);
                elapsed += currentWait;

                // Exponential backoff with cap: 10, 15, 22, 33, 50, 50, 50...
                currentWait = Math.Min(currentWait + (currentWait / 2), 50);
            }

            return false;
        }

        /// <summary>
        /// Retries a paste operation with exponential backoff.
        /// </summary>
        public static T RetryPaste<T>(Func<T> pasteAction, int maxAttempts = 3) where T : class
        {
            Exception lastEx = null;

            for (int attempt = 0; attempt < maxAttempts; attempt++)
            {
                try
                {
                    return pasteAction();
                }
                catch (COMException ex) when (attempt < maxAttempts - 1)
                {
                    lastEx = ex;
                    // Exponential backoff: 30ms, 60ms, 120ms
                    Thread.Sleep(30 * (1 << attempt));
                    System.Diagnostics.Debug.WriteLine($"Paste attempt {attempt + 1} failed, retrying...");
                }
            }

            throw lastEx ?? new Exception("Paste failed after retries");
        }
    }

    /// <summary>
    /// Tracks which slide is currently active to avoid redundant GotoSlide calls.
    /// </summary>
    internal sealed class SlideActivationTracker
    {
        private int _currentSlideIndex = -1;
        private readonly Microsoft.Office.Interop.PowerPoint.Application _ppApp;

        public SlideActivationTracker(Microsoft.Office.Interop.PowerPoint.Application ppApp)
        {
            _ppApp = ppApp ?? throw new ArgumentNullException(nameof(ppApp));
        }

        /// <summary>
        /// Ensures the specified slide is active, only switching if necessary.
        /// Call this only before paste operations that require it.
        /// </summary>
        /// <param name="slideIndex">1-based slide index</param>
        /// <param name="sleepMs">Sleep time after switching (default 30ms)</param>
        public void EnsureSlideActive(int slideIndex, int sleepMs = 30)
        {
            if (_currentSlideIndex != slideIndex)
            {
                try
                {
                    _ppApp.ActiveWindow.View.GotoSlide(slideIndex);
                    _currentSlideIndex = slideIndex;

                    // Only sleep if we actually changed slides
                    if (sleepMs > 0)
                    {
                        Thread.Sleep(sleepMs);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Could not activate slide {slideIndex}: {ex.Message}");
                }
            }
        }
    }
}