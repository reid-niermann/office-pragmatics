using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    public static class VisualCheckHelper
    {
        // OLE color for the unique highlight pattern
        private static readonly int HighlightColor = ColorTranslator.ToOle(ColorTranslator.FromHtml("#F3BFD7"));
        private static readonly int BackgroundColor = ColorTranslator.ToOle(Color.White);

        /// <summary>
        /// Applies textured highlights to help visualize formula consistency.
        /// Horizontal stripes indicate consistent formulas across columns.
        /// Vertical stripes indicate consistent formulas down rows.
        /// Crosshatch indicates consistency both ways.
        /// </summary>
        public static void VisualCheck()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null) return;

            // Prevent screen updates for performance
            app.ScreenUpdating = false;

            try
            {
                int rows = sel.Rows.Count;
                int cols = sel.Columns.Count;
                var formulas = new string[rows, cols];

                // Capture formulas (or null for constants)
                for (int i = 1; i <= rows; i++)
                {
                    for (int j = 1; j <= cols; j++)
                    {
                        var cell = sel.Cells[i, j] as Excel.Range;
                        formulas[i - 1, j - 1] = cell.HasFormula ? cell.FormulaR1C1 : null;
                    }
                }

                var horiz = new bool[rows, cols];
                var vert = new bool[rows, cols];

                // Detect horizontal consistency
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols - 1; j++)
                    {
                        var f1 = formulas[i, j];
                        var f2 = formulas[i, j + 1];
                        if (!string.IsNullOrEmpty(f1) && f1 == f2)
                        {
                            horiz[i, j] = true;
                            horiz[i, j + 1] = true;
                        }
                    }
                }

                // Detect vertical consistency
                for (int j = 0; j < cols; j++)
                {
                    for (int i = 0; i < rows - 1; i++)
                    {
                        var f1 = formulas[i, j];
                        var f2 = formulas[i + 1, j];
                        if (!string.IsNullOrEmpty(f1) && f1 == f2)
                        {
                            vert[i, j] = true;
                            vert[i + 1, j] = true;
                        }
                    }
                }

                // Apply patterns
                for (int i = 0; i < rows; i++)
                {
                    for (int j = 0; j < cols; j++)
                    {
                        var cell = sel.Cells[i + 1, j + 1] as Excel.Range;
                        bool appliedPattern = false;
                        if (horiz[i, j] && vert[i, j])
                        {
                            cell.Interior.Pattern = Excel.XlPattern.xlPatternGrid;
                            appliedPattern = true;
                        }
                        else if (horiz[i, j])
                        {
                            cell.Interior.Pattern = Excel.XlPattern.xlPatternLightHorizontal;
                            appliedPattern = true;
                        }
                        else if (vert[i, j])
                        {
                            cell.Interior.Pattern = Excel.XlPattern.xlPatternLightVertical;
                            appliedPattern = true;
                        }
                        else
                        {
                            // VisualCheck only adds highlights. Cells that no longer qualify keep any
                            // earlier highlight; the "Remove" command clears highlights wholesale.
                            continue;
                        }

                        if (appliedPattern)
                        {
                            cell.Interior.PatternColor = HighlightColor;
                            cell.Interior.Color = BackgroundColor;
                        }
                    }
                }
            }
            finally
            {
                // Restore screen updates
                app.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// Toggles the VisualCheck formatting on or off for the current selection.
        /// </summary>
        public static void ToggleVisualCheck()
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null) return;

            app.ScreenUpdating = false;
            try
            {
                bool isVisualCheckApplied = false;

                // Check if any cell in the selection has the VisualCheck formatting
                // We only need to find one cell to decide to remove formatting for the whole selection.
                // Iterating through all cells can be slow for very large selections.
                // For performance on huge selections, one might check a sample or just the first cell.
                // However, to be accurate, we check until one is found or all are checked.
                foreach (Excel.Range cell in sel.Cells) // sel.Cells is crucial for iterating individual cells
                {
                    // Ensure PatternColor is not null/DBNull and cast safely
                    if (cell.Interior.PatternColor is int currentOlePatternColor && currentOlePatternColor == HighlightColor)
                    {
                        Excel.XlPattern currentPattern = (Excel.XlPattern)cell.Interior.Pattern;
                        if (currentPattern == Excel.XlPattern.xlPatternGrid ||
                            currentPattern == Excel.XlPattern.xlPatternLightHorizontal ||
                            currentPattern == Excel.XlPattern.xlPatternLightVertical)
                        {
                            isVisualCheckApplied = true;
                            break; // Found one, no need to check further
                        }
                    }
                }

                if (isVisualCheckApplied)
                {
                    RemoveVisualCheckFormatting(sel);
                }
                else
                {
                    VisualCheck(); // This will use the current selection internally
                }
            }
            finally
            {
                app.ScreenUpdating = true;
            }
        }

        /// <summary>
        /// Removes the specific VisualCheck formatting from the given range.
        /// </summary>
        /// <param name="rangeToRemove">The Excel range from which to remove formatting.</param>
        private static void RemoveVisualCheckFormatting(Excel.Range rangeToRemove)
        {
            if (rangeToRemove == null) return;

            // It's more efficient to apply changes to the whole range at once if possible
            rangeToRemove.Interior.Pattern = Excel.XlPattern.xlPatternNone;
            // Reset PatternColor to automatic (no color)
            rangeToRemove.Interior.PatternColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
            // Reset background color. Since VisualCheck sets it to BackgroundColor (White),
            // we can set it back to that, or to automatic for a true reset.
            // Setting to BackgroundColor (White) is consistent with VisualCheck's application.
            rangeToRemove.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
            // If you want to truly reset to Excel's default (which is usually white unless themed):
            // rangeToRemove.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexAutomatic;
        }
    }
}