// SizeModeUtil.cs (.NET Framework 4.8 / C# 7.3)

using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_PowerPoint.ExcelLinking
{
    /// <summary>
    /// Centralized sizing/position policy for refreshed shapes.
    /// Optimized to minimize COM calls.
    /// </summary>
    internal static class SizeModeUtil
    {
        private const float SizeTolerance = 0.5f;

        /// <summary>
        /// Applies saved sizing policy (Match Width / Match Height / Match All) and position
        /// from <paramref name="source"/> to <paramref name="target"/>.
        /// Falls back to aspect-based behavior if no tag is present.
        /// </summary>
        public static void ApplyFromSource(PowerPoint.Shape source, PowerPoint.Shape target)
        {
            if (source == null || target == null) return;

            // Cache all source properties once to minimize COM calls
            float srcLeft = source.Left;
            float srcTop = source.Top;
            float srcWidth = source.Width;
            float srcHeight = source.Height;
            Office.MsoTriState srcLockAspect = source.LockAspectRatio;
            string mode = source.Tags["Pragmatics_SizeMode"];

            // Position first
            target.Left = srcLeft;
            target.Top = srcTop;

            // Determine sizing mode
            if (string.IsNullOrEmpty(mode))
                mode = (srcLockAspect == Office.MsoTriState.msoFalse) ? "Match All" : "Match Width";

            // Cache target dimensions to check if resize is needed
            float tgtWidth = target.Width;
            float tgtHeight = target.Height;
            bool widthMatches = Math.Abs(tgtWidth - srcWidth) < SizeTolerance;
            bool heightMatches = Math.Abs(tgtHeight - srcHeight) < SizeTolerance;

            switch (mode)
            {
                case "Match All":
                    target.LockAspectRatio = Office.MsoTriState.msoFalse;
                    if (!widthMatches) target.Width = srcWidth;
                    if (!heightMatches) target.Height = srcHeight;
                    break;

                case "Match Height":
                    target.LockAspectRatio = Office.MsoTriState.msoTrue;
                    if (!heightMatches) target.Height = srcHeight;
                    break;

                default: // "Match Width"
                    target.LockAspectRatio = Office.MsoTriState.msoTrue;
                    if (!widthMatches) target.Width = srcWidth;
                    break;
            }
        }

        /// <summary>
        /// Checks if shape dimensions need resizing for "Match All" mode.
        /// </summary>
        public static bool NeedsResize(PowerPoint.Shape shape, float targetWidth, float targetHeight)
        {
            if (shape == null) return false;
            return Math.Abs(shape.Width - targetWidth) > SizeTolerance ||
                   Math.Abs(shape.Height - targetHeight) > SizeTolerance;
        }

        /// <summary>
        /// Checks if Excel chart object dimensions need resizing for "Match All" mode.
        /// </summary>
        public static bool NeedsResize(Excel.ChartObject chartObj, float targetWidth, float targetHeight)
        {
            if (chartObj == null) return false;
            return Math.Abs(chartObj.Width - targetWidth) > SizeTolerance ||
                   Math.Abs(chartObj.Height - targetHeight) > SizeTolerance;
        }

        /// <summary>
        /// Checks if Excel range dimensions need resizing.
        /// </summary>
        public static bool NeedsResize(float currentWidth, float currentHeight, float targetWidth, float targetHeight)
        {
            return Math.Abs(currentWidth - targetWidth) > SizeTolerance ||
                   Math.Abs(currentHeight - targetHeight) > SizeTolerance;
        }
    }
}