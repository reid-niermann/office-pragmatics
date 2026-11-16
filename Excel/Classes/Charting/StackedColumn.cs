using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes.Charting
{
    internal class StackedColumn
    {
        /// <summary>
        /// Adds a total line series to a selected stacked column chart
        /// </summary>
        public static void StackedTotal()
        {
            try
            {
                // Get the application and active chart
                var application = Globals.ThisAddIn.Application;
                var activeChart = application.ActiveChart;
                if (activeChart == null)
                {
                    throw new InvalidOperationException("No chart is currently selected. Please select a stacked column chart.");
                }

                // Verify it's a stacked column chart
                var seriesCollection = activeChart.SeriesCollection() as Excel.SeriesCollection;
                if (seriesCollection == null || seriesCollection.Count == 0)
                {
                    throw new InvalidOperationException("The selected chart has no data series.");
                }

                // Get the first series to determine the data range
                var firstSeries = seriesCollection.Item(1) as Excel.Series;
                var LastSeries = seriesCollection.Item(seriesCollection.Count) as Excel.Series;
                var formula = firstSeries.Formula;
                var formula2 = LastSeries.Formula;

                // Parse the formula to get the data range
                var dataRange = GetDataRangeFromFormula(formula);
                var dataRange2 = GetDataRangeFromFormula(formula2);
                if (dataRange == null)
                {
                    throw new InvalidOperationException("Could not determine the data range from the chart.");
                }

                // Get the worksheet containing the data
                var worksheet = dataRange.Worksheet;

                // Determine the bounds of the data
                int firstDataRow = dataRange.Row;
                int lastDataRow = dataRange2.Row;
                int firstCol = dataRange.Column;
                int lastCol = dataRange.Column + dataRange.Columns.Count - 1;

                // Find the row for the totals (one row below the last data row)
                int totalRow = lastDataRow + 1;

                // Add SUM formulas for each data column
                for (int col = firstCol; col <= lastCol; col++)
                {
                    var totalCell = worksheet.Cells[totalRow, col] as Excel.Range;
                    var columnLetter = GetColumnLetter(col);
                    var sumFormula = $"=SUM({columnLetter}{firstDataRow}:{columnLetter}{lastDataRow})";
                    totalCell.Formula = sumFormula;
                }

                // Create ranges for the new series
                var xValuesRange = worksheet.Range[
                    worksheet.Cells[firstDataRow - 1, firstCol],
                    worksheet.Cells[firstDataRow - 1, lastCol]
                ] as Excel.Range;

                var yValuesRange = worksheet.Range[
                    worksheet.Cells[totalRow, firstCol],
                    worksheet.Cells[totalRow, lastCol]
                ] as Excel.Range;

                // Add the new series as a line
                var newSeries = seriesCollection.NewSeries() as Excel.Series;
                newSeries.Name = "Total";
                newSeries.XValues = xValuesRange;
                newSeries.Values = yValuesRange;

                // Change the new series to a line type
                newSeries.ChartType = Excel.XlChartType.xlLine;

                // Remove markers from the line
                newSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
                newSeries.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;

                // Add data labels positioned above
                newSeries.HasDataLabels = true;
                var dataLabels = newSeries.DataLabels() as Excel.DataLabels;
                dataLabels.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove;

                // Format the line (optional - make it more prominent)
                newSeries.Format.Line.Weight = 2.5f;
                newSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error adding stacked total: {ex.Message}",
                    "Stacked Total Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Extracts the data range from a series formula
        /// </summary>
        private static Excel.Range GetDataRangeFromFormula(string formula)
        {
            try
            {
                // Series formula format: =SERIES(name,xvalues,yvalues,order)
                var parts = formula.Split(',');
                if (parts.Length < 3) return null;

                // Get the Y values range (third parameter)
                var yValuesStr = parts[2].Trim();

                // Remove sheet name if present
                if (yValuesStr.Contains("!"))
                {
                    yValuesStr = yValuesStr.Substring(yValuesStr.IndexOf("!") + 1);
                }

                // Get the application and parse the range
                var app = Globals.ThisAddIn.Application;
                var activeSheet = app.ActiveSheet as Excel.Worksheet;

                return activeSheet.Range[yValuesStr] as Excel.Range;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Converts a column number to Excel column letter
        /// </summary>
        private static string GetColumnLetter(int columnNumber)
        {
            string columnLetter = "";

            while (columnNumber > 0)
            {
                int modulo = (columnNumber - 1) % 26;
                columnLetter = Convert.ToChar(65 + modulo) + columnLetter;
                columnNumber = (columnNumber - modulo) / 26;
            }

            return columnLetter;
        }

        /// <summary>
        /// Alternative method that works with a specific chart object
        /// </summary>
        public static void StackedTotal(Excel.Chart chart)
        {
            // Temporarily make this the active chart
            chart.Activate();

            // Call the main method
            StackedTotal();
        }
    }
}