using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace OfficePragmatics_Excel.Classes.Charting
{
    internal class StatLines
    {
        /// <summary>
        /// Adds a median line series to the currently selected column chart based on the first series
        /// </summary>
        public static void AddMedianLine()
        {
            Application app = null;
            Chart chart = null;
            Series firstSeries = null, medianSeries = null;
            Worksheet ws = null;
            Range valuesRng = null, medianRowRng = null;

            try
            {
                app = Globals.ThisAddIn.Application;
                chart = app.ActiveChart ?? throw new InvalidOperationException("Select a chart first.");

                // Accept only column charts
                if (!chart.ChartType.ToString().Contains("Column"))
                    throw new InvalidOperationException("The selected chart is not a column chart.");

                firstSeries = chart.SeriesCollection(1);

                /* --------------------------------------------------------
                 * 1.  Get the range behind the first series
                 *     (robust against array materialisation).
                 * -------------------------------------------------------*/
                string f = firstSeries.Formula;   // =SERIES(name,xVals,yVals,order)
                int p1 = f.IndexOf('(') + 1, p2 = f.LastIndexOf(')');
                string[] parts = f.Substring(p1, p2 - p1)
                                  .Split(',')  // safe because SERIES formula never
                                  .Select(s => s.Trim()).ToArray();   // nests commas

                if (parts.Length < 3) throw new InvalidOperationException("Can't parse series formula.");
                string yRef = parts[2];                 // e.g. Sheet1!$B$2:$B$10
                valuesRng = app.Evaluate(yRef) as Range
                              ?? throw new InvalidOperationException("Could not resolve Y-value range.");

                ws = valuesRng.Worksheet;

                /* --------------------------------------------------------
                 * 2.  Write the MEDIAN row, directly under the data block
                 * -------------------------------------------------------*/
                int medianRowIdx = valuesRng.Row + valuesRng.Rows.Count;     // row below the block
                int colStart = valuesRng.Column;
                int colEnd = colStart + valuesRng.Columns.Count - 1;

                medianRowRng = ws.Range[ws.Cells[medianRowIdx, colStart], ws.Cells[medianRowIdx, colEnd]];

                // absolute reference to ALL Y-values in the first series
                string fullSeries = valuesRng.Address[true, true];   // e.g. $L$6:$R$6

                // put the same formula in every cell of the new row
                medianRowRng.Formula = $"=MEDIAN({fullSeries})";

                /* --------------------------------------------------------
                 * 3.  Add the MEDIAN line to the chart
                 * -------------------------------------------------------*/
                medianSeries = chart.SeriesCollection().NewSeries();
                medianSeries.Values = medianRowRng;
                medianSeries.XValues = firstSeries.XValues;
                medianSeries.Name = "MEDIAN";
                medianSeries.ChartType = XlChartType.xlLine;

                // simple formatting – adjust to taste
                medianSeries.Format.Line.ForeColor.RGB =
                    System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                medianSeries.Format.Line.Weight = 2f;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Error adding median line:\r\n{ex.Message}",
                    "StatLines", System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Error);
            }
            finally   // tidy up COM references
            {
                void Release(object o) { if (o != null) Marshal.ReleaseComObject(o); }
                Release(medianRowRng); Release(valuesRng); Release(medianSeries);
                Release(firstSeries); Release(ws); Release(chart);
            }
        }
    }
}