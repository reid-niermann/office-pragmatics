using System;
using Excel = Microsoft.Office.Interop.Excel;
using OfficePragmatics_Excel.Classes.Brand;
using OfficePragmatics_Excel.Classes.Tables;
using OfficePragmatics_Excel.Settings; // Assumed namespace for SettingsForm

namespace OfficePragmatics_Excel
{
    public static class AutoBrand
    {
        public static void Run()
        {
            // ──────────────────────────────────────────────────────────────────────────
            // 1. Check for Chart Selection
            // ──────────────────────────────────────────────────────────────────────────
            // We use the existing helper in ChartFormat to determine if the user's 
            // selection constitutes a valid chart (ActiveChart, Shape, or ChartObject).
            Excel.Chart activeChart = ChartFormat.GetActiveChart();

            if (activeChart != null)
            {
                // Retrieve settings and apply
                var chartSettings = SettingsForm.CurrentSettings.BrandCharts;
                ChartFormat.ApplyBrandFormattingToSelection(chartSettings);
                return;
            }

            // ──────────────────────────────────────────────────────────────────────────
            // 2. Check for Range Selection
            // ──────────────────────────────────────────────────────────────────────────
            if (Globals.ThisAddIn.Application.Selection is Excel.Range)
            {
                // Retrieve settings and apply
                var tableSettings = SettingsForm.CurrentSettings.TableFormat;
                TableFormatter.ApplyToSelection(tableSettings);
                return;
            }
        }
    }
}