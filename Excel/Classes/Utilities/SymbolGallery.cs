// Features/SymbolGallery.cs
// (C# 7.3 / .NET 4.8, VSTO Excel)
using System;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Features
{
    internal static class SymbolGallery
    {
        // Keep this order in sync with the <item> order in Ribbon XML
        internal static readonly string[] Items = new[]
        {
            "✓","✗","–","Δ","Σ","∞","Ω","≠","°","★","₿","Ξ","▲","▼","±","≈","≤","≥","≃","©","™"
        };

        internal static void InsertByIndex(int selectedIndex)
        {
            if (selectedIndex < 0 || selectedIndex >= Items.Length) return;

            var app = Globals.ThisAddIn?.Application as Excel.Application;
            if (app == null) return;

            try
            {
                var rng = app.Selection as Excel.Range;
                if (rng == null)
                {
                    MessageBox.Show(
                        "Select one or more cells first.",
                        "Insert Symbol",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                    return;
                }

                bool oldScreenUpdating = app.ScreenUpdating;
                app.ScreenUpdating = false;
                try
                {
                    string symbol = Items[selectedIndex];

                    // Vectorized write per area for speed
                    foreach (Excel.Range area in rng.Areas)
                        area.Value2 = symbol;
                }
                finally
                {
                    app.ScreenUpdating = oldScreenUpdating;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "Couldn't insert the symbol:\n" + ex.Message,
                    "Insert Symbol",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}