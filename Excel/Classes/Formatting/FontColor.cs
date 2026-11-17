using System;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    internal static class FontColor
    {
        // ---- Colour lists --------------------------------------------------

        private static readonly int[] BBCycleColors =
        {
               ColorTranslator.ToOle(Color.Black),
               ColorTranslator.ToOle(Color.Blue),
        };

        private static readonly int[] ModelCycleColors =
        {
            ColorTranslator.ToOle(Color.Black),
            ColorTranslator.ToOle(Color.Blue),
            ColorTranslator.ToOle(Color.Green),
            ColorTranslator.ToOle(Color.Red),
            ColorTranslator.ToOle(Color.Purple),
            ColorTranslator.ToOle(Color.White)
        };

        // ---- Public entry points -------------------------------------------

        public static void RunBBColorCycle()
        {
            CycleFontColor(BBCycleColors, defaultIndex: 0);
        }

        public static void RunModelColorCycle()
        {
            CycleFontColor(ModelCycleColors, defaultIndex: 0);
        }

        public static void RunAutoColor()
        {
            var app = Globals.ThisAddIn.Application;
            try
            {
                app.ScreenUpdating = false;

                Excel.Range targetRange;
                if (app.Selection is Excel.Range selection && selection.Count > 1)
                {
                    targetRange = selection;
                }
                else
                {
                    targetRange = app.ActiveSheet.UsedRange;
                }

                AutoColorRange(targetRange);
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"An error occurred during the auto-color process:\n{ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally
            {
                if (app != null)
                {
                    app.ScreenUpdating = true;
                }
            }
        }

        // ---- Core helpers ---------------------------------------------------

        public static void AutoColorRange(Excel.Range range)
        {
            int blackColorOle = ColorTranslator.ToOle(Color.Black);
            int blueColorOle = ColorTranslator.ToOle(Color.Blue);
            int greenColorOle = ColorTranslator.ToOle(Color.Green);
            int redColorOle = ColorTranslator.ToOle(Color.Red);
            int purpleColorOle = ColorTranslator.ToOle(Color.Purple);

            string[] financialFunctionSignatures = {
        "FDS(", "PITCHBOOK(", "CIQ(", "BDP(", "BDH(",
        "BDS(", "TR(", "RDATA(", "YCHARTS(", "MSTAR("
    };

            foreach (Excel.Range cell in range.Cells)
            {
                // STEP 1: Always reset the font to black first.
                // This makes the logic re-runnable on every cell change.
                cell.Font.Color = blackColorOle;

                // STEP 2: Apply coloring rules.
                if (cell.HasFormula)
                {
                    string formula = cell.Formula.ToString().ToUpper();

                    // Rule: Financial data provider -> Purple
                    bool isFinancial = false;
                    foreach (string signature in financialFunctionSignatures)
                    {
                        if (formula.Contains(signature))
                        {
                            cell.Font.Color = purpleColorOle;
                            isFinancial = true;
                            break;
                        }
                    }
                    if (isFinancial) continue;

                    // Link checks
                    bool hasExternalLink = formula.Contains("["); // Simplified check
                    bool hasInternalLink = formula.Contains("!");

                    // Rule: Links to another workbook -> Red
                    if (hasExternalLink)
                    {
                        cell.Font.Color = redColorOle;
                        continue;
                    }

                    // Rule: Links to another sheet -> Green
                    if (hasInternalLink)
                    {
                        cell.Font.Color = greenColorOle;
                        continue;
                    }

                    // Rule: Formula with no precedents -> Blue
                    try
                    {
                        Excel.Range precedents = cell.Precedents;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        cell.Font.Color = blueColorOle;
                    }
                }
                else
                {
                    // Rule: Hardcoded number -> Blue
                    if (cell.Value2 is double || cell.Value2 is int || cell.Value2 is decimal)
                    {
                        cell.Font.Color = blueColorOle;
                    }
                }
            }
        }

        private static void CycleFontColor(int[] colors, int defaultIndex)
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

                int currentColor = Convert.ToInt32(selection.Cells[1, 1].Font.Color ?? -1);
                int idx = Array.IndexOf(colors, currentColor);
                int nextIdx = (idx == -1) ? defaultIndex : (idx + 1) % colors.Length;

                selection.Font.Color = colors[nextIdx];
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error applying font colour:\n{ex.Message}",
                    "Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }
    }
}