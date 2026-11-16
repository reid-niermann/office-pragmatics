// SizeModeUtil.cs (.NET Framework 4.8 / C# 7.3)
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Word.ExcelLinking
{
    /// <summary>
    /// Centralized sizing/position policy for refreshed shapes.
    /// Optimized to minimize COM calls.
    /// </summary>
    internal static class SizeModeUtil
    {

        /// <summary>
        /// Convenience helper for callers that need to know if "Match All" (no aspect lock) is set.
        /// </summary>
        public static bool IsMatchAll(Word.Shape source)
        {
            if (source == null) return false;
            string mode = ShapeTagsHelper.GetTag(source, "Pragmatics_SizeMode");
            return string.Equals(mode, "Match All", System.StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if shape dimensions need resizing for "Match All" mode
        /// </summary>
        public static bool NeedsResize(Word.Shape shape, float targetWidth, float targetHeight)
        {
            if (shape == null) return false;
            return System.Math.Abs(shape.Width - targetWidth) > 0.5f ||
                   System.Math.Abs(shape.Height - targetHeight) > 0.5f;
        }

        /// <summary>
        /// Checks if Excel chart object dimensions need resizing for "Match All" mode
        /// </summary>
        public static bool NeedsResize(Excel.ChartObject chartObj, float targetWidth, float targetHeight)
        {
            if (chartObj == null) return false;
            return System.Math.Abs(chartObj.Width - targetWidth) > 0.5f ||
                   System.Math.Abs(chartObj.Height - targetHeight) > 0.5f;
        }
    }
}