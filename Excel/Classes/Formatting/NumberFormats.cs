using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel
{
    internal static class NumberFormats
    {
        // --- Format Lists ---

        private static readonly string[] NumberCycleFormats = {
            @"_(#,##0_);(#,##0);_(""–""_);_(@_)",
            @"_(#,##0.0_);(#,##0.0);_(""–""_);_(@_)",
            @"_(#,##0.00_);(#,##0.00);_(""–""_);_(@_)"
        };

        private static readonly string[] CurrencyCycleFormats = {
            @"_($#,##0_);($#,##0);_(""–""_);_(@_)",
            @"_(€#,##0_);(€#,##0);_(""–""_);_(@_)",
            @"_(£#,##0_);(£#,##0);_(""–""_);_(@_)",
            @"_(¥#,##0_);(¥#,##0);_(""–""_);_(@_)"
        };

        private static readonly string[] PercentCycleFormats = {
            @"_(0%_);(0%);_(""–""_);_(@_)",
            @"_(0.0%_);(0.0%);_(""–""_);_(@_)",
            @"_(0.00%_);(0.00%);_(""–""_);_(@_)",
            @"_(""S+""0_);(""S-""0);_(""–""_);_(@_)"
        };

        private static readonly string[] DateCycleFormats = {
            @"_(m/d/yyyy_);_(@_)",
            @"_(mm-yy_);_(@_)",
            @"_(yyyy\A_);_(@_)",
            @"_(yyyy\E_);_(@_)",
            @"_(mmmm d, yyyy_);_(@_)"
        };

        private static readonly string[] MultipleCycleFormats = {
            @"_(#,##0""x""_);(#,##0""x"");_(""–""_);_(@_)",
            @"_(#,##0.0""x""_);(#,##0.0""x"");_(""–""_);_(@_)",
            @"_(#,##0.00""x""_);(#,##0.00""x"");_(""–""_);_(@_)"
        };

        private static readonly string[] BinaryCycleFormats = {
            @"_(""Yes""_);_(""–""_);_(""No""_);_(@_)",
            @"_(""True""_);_(""–""_);_(""False""_);_(@_)",
            @"_(""Y""_);_(""–""_);_(""N""_);_(@_)",
            @"_(""On""_);_(""–""_);_(""Off""_);_(@_)"
        };

        // --- Helper for Cycling ---

        private static void CycleFormat(string[] formats)
        {
            try
            {
                Excel.Application app = Globals.ThisAddIn.Application; // Get application object
                if (app.Selection == null || !(app.Selection is Excel.Range))
                {
                    System.Windows.Forms.MessageBox.Show("Please select a cell or range.", "Selection Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    return;
                }

                Excel.Range selection = app.Selection as Excel.Range;

                // Use the format of the top-left cell to determine the current state
                // NumberFormat returns object, might be null or DBNull.Value
                object currentFormatObj = selection.Cells[1, 1].NumberFormat;
                string currentFormat = currentFormatObj as string ?? string.Empty; // Handle null safely

                int currentIndex = -1;
                for (int i = 0; i < formats.Length; i++)
                {
                    // Case-insensitive comparison might be safer, though formats are usually specific
                    if (string.Equals(currentFormat, formats[i], StringComparison.OrdinalIgnoreCase))
                    {
                        currentIndex = i;
                        break;
                    }
                }

                int nextIndex = (currentIndex == -1) ? 0 : (currentIndex + 1) % formats.Length;
                string newFormat = formats[nextIndex];

                // Apply to the entire selection
                selection.NumberFormat = newFormat;

                // Release COM object (optional but good practice in some scenarios, less critical with Selection)
                // System.Runtime.InteropServices.Marshal.ReleaseComObject(selection);
            }
            catch (Exception ex)
            {
                // Log the error or show a message
                System.Windows.Forms.MessageBox.Show($"Error applying format: {ex.Message}", "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        // --- Public Cycle Functions ---

        public static void RunNumberCycle()
        {
            CycleFormat(NumberCycleFormats);
        }

        public static void RunCurrencyCycle()
        {
            CycleFormat(CurrencyCycleFormats);
        }

        public static void RunPercentCycle() // Applies Date formats as per request description
        {
            CycleFormat(PercentCycleFormats);
        }

        public static void RunDateCycle() // Applies Percent formats as per request description
        {
            CycleFormat(DateCycleFormats);
        }

        public static void RunMultipleCycle()
        {
            CycleFormat(MultipleCycleFormats);
        }

        public static void RunBinaryCycle()
        {
            CycleFormat(BinaryCycleFormats);
        }


        // --- Decimal Place Functions ---

        public static void RunPlusDecimals()
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("DecimalsIncrease");
        }

        public static void RunMinusDecimals()
        {
            Globals.ThisAddIn.Application.CommandBars.ExecuteMso("DecimalsDecrease");
        }

    }
}