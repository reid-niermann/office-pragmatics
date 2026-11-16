using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Xml.Linq;

namespace OfficePragmatics_Excel
{
    internal static class FillColors
    {
        // --- Fill Color List ---
        // Last entry is a null sentinel for Transparent
        private static readonly string[] FillCycleColors = {
            "#FFFFCC", // Yellow
            "#DCE6F1", // Light Blue
            "#F1DCDC", // Light Red
            "#F2F2F2", // Light Gray
            "#FFFFFF", // White
            null       // Transparent
        };

        // --- Theme Color Cycle ---
        // Defines the cycle order for theme colors.
        // Last entry is a null sentinel for Transparent.
        private static readonly object[] ThemeColorCycle = {
            Excel.XlThemeColor.xlThemeColorAccent1,
            Excel.XlThemeColor.xlThemeColorAccent2,
            Excel.XlThemeColor.xlThemeColorAccent3,
            Excel.XlThemeColor.xlThemeColorAccent4,
            Excel.XlThemeColor.xlThemeColorAccent5,
            Excel.XlThemeColor.xlThemeColorAccent6,
            Excel.XlThemeColor.xlThemeColorLight1, // Represents White in most themes
            null // Sentinel for Transparent
        };

        // Helper method to cycle interior fill colors
        private static void CycleFillColor()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (!(app.Selection is Excel.Range selection))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a cell or range.",
                        "Selection Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                var firstInterior = selection.Cells[1, 1].Interior;
                object patternObj = firstInterior.Pattern;
                int patternValue = patternObj is int pi ? pi : Convert.ToInt32(patternObj);
                bool isTransparent = patternValue == (int)Excel.XlPattern.xlPatternNone;
                int currentIndex = isTransparent ? FillCycleColors.Length - 1 : -1;

                if (!isTransparent)
                {
                    object colorObj = firstInterior.Color;
                    int currentOle = colorObj is double d ? Convert.ToInt32(d) : 0;
                    for (int i = 0; i < FillCycleColors.Length - 1; i++)
                    {
                        int ole = ColorTranslator.ToOle(ColorTranslator.FromHtml(FillCycleColors[i]));
                        if (ole == currentOle)
                        {
                            currentIndex = i;
                            break;
                        }
                    }
                }

                int nextIndex = (currentIndex + 1) % FillCycleColors.Length;
                string nextColor = FillCycleColors[nextIndex];

                if (nextColor == null)
                {
                    selection.Interior.Pattern = (int)Excel.XlPattern.xlPatternNone;
                }
                else
                {
                    int oleColor = ColorTranslator.ToOle(ColorTranslator.FromHtml(nextColor));
                    selection.Interior.Pattern = (int)Excel.XlPattern.xlPatternSolid;
                    selection.Interior.Color = oleColor;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error applying fill color: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Cycles the fill color of the selected range through the workbook's theme colors.
        /// </summary>
        private static void CycleThemeColor()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (!(app.Selection is Excel.Range selection))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a cell or range.",
                        "Selection Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                var firstInterior = selection.Cells[1, 1].Interior;
                object patternObj = firstInterior.Pattern;
                int patternValue = patternObj is int pi ? pi : Convert.ToInt32(patternObj);
                bool isTransparent = patternValue == (int)Excel.XlPattern.xlPatternNone;

                int currentIndex = -1;

                if (isTransparent)
                {
                    // Current is transparent, which is the last item in our cycle array.
                    currentIndex = ThemeColorCycle.Length - 1;
                }
                else
                {
                    // Check if the current fill is a theme color from our cycle.
                    object themeColorObj = firstInterior.ThemeColor;
                    if (themeColorObj is int || themeColorObj is Excel.XlThemeColor)
                    {
                        var currentThemeColor = (Excel.XlThemeColor)Convert.ToInt32(themeColorObj);

                        // Find the color in our cycle array.
                        for (int i = 0; i < ThemeColorCycle.Length - 1; i++) // Exclude the null sentinel
                        {
                            if ((Excel.XlThemeColor)ThemeColorCycle[i] == currentThemeColor)
                            {
                                currentIndex = i;
                                break;
                            }
                        }
                    }
                    // If the color isn't a theme color in our cycle, currentIndex remains -1,
                    // so the cycle will start from the beginning.
                }

                // Calculate the next index, wrapping around.
                int nextIndex = (currentIndex + 1) % ThemeColorCycle.Length;
                object nextThemeColor = ThemeColorCycle[nextIndex];

                // Apply the next color or formatting.
                if (nextThemeColor == null)
                {
                    // Set to Transparent (no fill).
                    selection.Interior.Pattern = Excel.XlPattern.xlPatternNone;
                }
                else
                {
                    // Apply the next theme color from the cycle.
                    selection.Interior.Pattern = Excel.XlPattern.xlPatternSolid;
                    selection.Interior.ThemeColor = nextThemeColor;
                    // Reset TintAndShade to use the pure theme color.
                    selection.Interior.TintAndShade = 0;
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error applying theme color: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }


        // --- Table Formatting Cycle ---
        private static void CycleTableFormatting()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (!(app.Selection is Excel.Range selection))
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Please select a cell or range.",
                        "Selection Error",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                var fcs = selection.FormatConditions;
                Excel.FormatCondition existing = null;

                // Identify existing cycle rule (expression-based)
                foreach (Excel.FormatCondition fc in fcs)
                {
                    var fcType = (Excel.XlFormatConditionType)fc.Type;
                    if (fcType == Excel.XlFormatConditionType.xlExpression &&
                        (fc.Formula1.StartsWith("=MOD(ROW(),", StringComparison.OrdinalIgnoreCase) ||
                         fc.Formula1.StartsWith("=MOD(COLUMN(),", StringComparison.OrdinalIgnoreCase)))
                    {
                        existing = fc;
                        break;
                    }
                }

                if (existing == null)
                {
                    // 1st cycle: apply row striping
                    int firstRow = selection.Row;
                    int parity = firstRow % 2;
                    string formula = $"=MOD(ROW(),2)={parity}";

                    fcs.Delete();
                    var newFc = fcs.Add(
                        Type: Excel.XlFormatConditionType.xlExpression,
                        Formula1: formula);
                    newFc.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#F2F2F2"));
                }
                else
                {
                    var existingType = (Excel.XlFormatConditionType)existing.Type;
                    if (existingType == Excel.XlFormatConditionType.xlExpression &&
                        existing.Formula1.StartsWith("=MOD(ROW(),", StringComparison.OrdinalIgnoreCase))
                    {
                        // 2nd cycle: apply column striping
                        int firstCol = selection.Column;
                        int parity = firstCol % 2;
                        string formula = $"=MOD(COLUMN(),2)={parity}";

                        fcs.Delete();
                        var newFc = fcs.Add(
                            Type: Excel.XlFormatConditionType.xlExpression,
                            Formula1: formula);
                        newFc.Interior.Color = ColorTranslator.ToOle(ColorTranslator.FromHtml("#F2F2F2"));
                    }
                    else
                    {
                        // 3rd cycle: remove all CF
                        fcs.Delete();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error applying table formatting: {ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        public static void ConvertTransparentToWhite()
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var ws = app.ActiveSheet as Excel.Worksheet;
                if (ws == null) return;

                var sel = app.Selection as Excel.Range;

                // Use selection only if multiple cells; otherwise use the entire UsedRange
                bool useSelection = sel != null && Convert.ToDouble(sel.CountLarge) > 1.0;
                Excel.Range target = useSelection ? sel : ws.UsedRange;
                if (target == null) return;

                // Performance guardrails
                bool prevScreen = app.ScreenUpdating;
                bool prevEvents = app.EnableEvents;
                var prevCalc = app.Calculation;
                app.ScreenUpdating = false;
                app.EnableEvents = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;

                int whiteOle = ColorTranslator.ToOle(Color.White);

                foreach (Excel.Range area in target.Areas)
                {
                    foreach (Excel.Range cell in area.Cells)
                    {
                        var interior = cell.Interior;

                        // Treat as "transparent" if Pattern is None OR ColorIndex is None
                        int pattern = Convert.ToInt32(interior.Pattern);
                        bool patternNone = pattern == (int)Excel.XlPattern.xlPatternNone;

                        object ciObj = interior.ColorIndex; // may be boxed int/double
                        int colorIndex = (ciObj is int i) ? i : Convert.ToInt32(ciObj);
                        bool colorNone = colorIndex == (int)Excel.XlColorIndex.xlColorIndexNone;

                        if (patternNone || colorNone)
                        {
                            interior.Pattern = Excel.XlPattern.xlPatternSolid;
                            interior.Color = whiteOle;
                            interior.TintAndShade = 0;
                        }
                    }
                }

                // restore
                app.ScreenUpdating = prevScreen;
                app.EnableEvents = prevEvents;
                app.Calculation = prevCalc;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error converting transparent fills: {ex.Message}",
                    "Fill Tools",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // Public entry points
        public static void RunFillCycle()
        {
            CycleFillColor();
        }

        public static void RunThemeColorCycle()
        {
            CycleThemeColor();
        }

        public static void RunTableFormattingCycle()
        {
            CycleTableFormatting();
        }


        public static void RunConvertTransparentToWhite()
        {
            ConvertTransparentToWhite();
        }
    }
}