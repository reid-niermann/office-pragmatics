using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes.Charting
{
    public class CAGRArrow
    {
        public void insertArrowCAGR()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Check if a chart is selected
            Excel.Chart activeChart = null;
            try
            {
                activeChart = (Excel.Chart)excelApp.ActiveChart;
                if (activeChart == null)
                {
                    MessageBox.Show("Please select a chart first.", "No Chart Selected",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            catch
            {
                MessageBox.Show("Please select a chart first.", "No Chart Selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Get the first series from the chart (assuming it's the data series)
            if (activeChart.SeriesCollection().Count == 0)
            {
                MessageBox.Show("The selected chart has no data series.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Series dataSeries = activeChart.SeriesCollection(1);

            // Get series values and categories
            string valuesFormula = dataSeries.Formula;
            string[] formulaParts = valuesFormula.Split(',');

            // Extract ranges from formula
            string categoryRange = formulaParts[1];
            string valuesRange = formulaParts[2];

            // Remove sheet references if present
            if (categoryRange.Contains("!"))
                categoryRange = categoryRange.Substring(categoryRange.IndexOf("!") + 1);
            if (valuesRange.Contains("!"))
                valuesRange = valuesRange.Substring(valuesRange.IndexOf("!") + 1);

            // Prompt user to select location for intermediate table
            Excel.Range tableLocation = null;
            try
            {
                tableLocation = excelApp.InputBox(
                    "Select a cell for the intermediate calculations table (4 rows x 4 columns will be used):",
                    "Select Table Location",
                    Type: 8) as Excel.Range;

                if (tableLocation == null)
                    return;
            }
            catch
            {
                return; // User cancelled
            }

            // Create intermediate table
            Excel.Range tableRange = worksheet.Range[tableLocation,
                tableLocation.Offset[3, 3]];

            // Clear any existing data
            tableRange.Clear();

            // Get range objects
            Excel.Range catRange = worksheet.Range[categoryRange];
            Excel.Range valRange = worksheet.Range[valuesRange];

            // Determine if ranges are horizontal or vertical
            bool isVertical = catRange.Columns.Count == 1;

            // Get first and last cell references
            string firstValCell = isVertical ?
                valRange.Cells[1, 1].Address[false, false] :
                valRange.Cells[1, 1].Address[false, false];
            string lastValCell = isVertical ?
                valRange.Cells[valRange.Rows.Count, 1].Address[false, false] :
                valRange.Cells[1, valRange.Columns.Count].Address[false, false];

            string firstCatCell = isVertical ?
                catRange.Cells[1, 1].Address[false, false] :
                catRange.Cells[1, 1].Address[false, false];
            string lastCatCell = isVertical ?
                catRange.Cells[catRange.Rows.Count, 1].Address[false, false] :
                catRange.Cells[1, catRange.Columns.Count].Address[false, false];

            // Row 1: CAGR calculation
            tableLocation.Value = "Label";
            tableLocation.Offset[0, 2].Formula =
                $"=({lastValCell}/{firstValCell})^(1/(COLUMNS({valuesRange})-1))-1";
            tableLocation.Offset[0, 2].NumberFormat = "\"CAGR: \" 0.0%";

            // Row 2: Y-Offset
            tableLocation.Offset[1, 0].Value = "Y-Offset";
            tableLocation.Offset[1, 1].Formula =
                $"=({lastValCell}-{firstValCell})/(COLUMNS({valuesRange})-1)";

            // Row 3: X-Axis
            tableLocation.Offset[2, 0].Value = "X-Axis";
            tableLocation.Offset[2, 1].Value = 1;
            tableLocation.Offset[2, 3].Formula = $"=COLUMNS({categoryRange})";
            tableLocation.Offset[2, 2].Formula =
                $"=AVERAGE({tableLocation.Offset[2, 1].Address[false, false]}," +
                $"{tableLocation.Offset[2, 3].Address[false, false]})";

            // Row 4: Y-Axis (values + offset)
            tableLocation.Offset[3, 0].Value = "Y-Axis";
            string offsetCell = tableLocation.Offset[1, 1].Address[false, false];
            tableLocation.Offset[3, 1].Formula = $"={firstValCell}+{offsetCell}";
            tableLocation.Offset[3, 3].Formula = $"={lastValCell}+{offsetCell}";
            tableLocation.Offset[3, 2].Formula =
                $"=AVERAGE({tableLocation.Offset[3, 1].Address[false, false]}," +
                $"{tableLocation.Offset[3, 3].Address[false, false]})";

            // Add new series to chart
            Excel.Series cagrSeries = activeChart.SeriesCollection().NewSeries();
            cagrSeries.Name = "CAGR_Series";

            // Set X and Y values for the series
            string xValuesRange = $"{tableLocation.Offset[2, 1].Address}:" +
                $"{tableLocation.Offset[2, 3].Address}";
            string yValuesRange = $"{tableLocation.Offset[3, 1].Address}:" +
                $"{tableLocation.Offset[3, 3].Address}";

            cagrSeries.XValues = worksheet.Range[xValuesRange];
            cagrSeries.Values = worksheet.Range[yValuesRange];

            // Change series to Scatter with Straight Lines
            cagrSeries.ChartType = Excel.XlChartType.xlXYScatterLines;

            // Hide the axis
            activeChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).Delete();
            activeChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary).Delete();

            // Format the series with arrow
            cagrSeries.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            cagrSeries.Format.Line.Weight = 1;
            cagrSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(
                System.Drawing.Color.Black);

            // Add arrow end
            cagrSeries.Format.Line.EndArrowheadStyle =
                Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen;
            cagrSeries.Format.Line.EndArrowheadLength =
                Microsoft.Office.Core.MsoArrowheadLength.msoArrowheadLong;
            cagrSeries.Format.Line.EndArrowheadWidth =
                Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWide;

            // Remove markers
            cagrSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

            // Add data label to show CAGR percentage
            cagrSeries.HasDataLabels = true;

            // Clear all data labels first
            foreach (Excel.DataLabel lbl in cagrSeries.DataLabels())
            {
                lbl.Delete();
            }

            // Add label only to the middle point (point 2)
            cagrSeries.Points(2).HasDataLabel = true;
            Excel.DataLabel label = cagrSeries.Points(2).DataLabel;
            label.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove;
            label.Format.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Cell that holds the CAGR formula
            Excel.Range cagrCell = tableLocation.Offset[0, 2];

            // Tell the data-label to use that cell's value
            label.Formula = "=" + cagrCell.Address[
                RowAbsolute: true,
                ColumnAbsolute: true,
                ReferenceStyle: Excel.XlReferenceStyle.xlA1,
                External: true          // prepends the sheet name—safer if the chart is on another sheet
            ];

            MessageBox.Show("CAGR arrow has been added successfully!", "Success",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public void insertElbowCAGR()
        {
            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Workbook workbook = excelApp.ActiveWorkbook;
            Excel.Worksheet worksheet = workbook.ActiveSheet;

            // Check if a chart is selected
            Excel.Chart activeChart = null;
            try
            {
                activeChart = (Excel.Chart)excelApp.ActiveChart;
                if (activeChart == null)
                {
                    MessageBox.Show("Please select a chart first.", "No Chart Selected",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
            }
            catch
            {
                MessageBox.Show("Please select a chart first.", "No Chart Selected",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Get the first series from the chart (assuming it's the data series)
            if (activeChart.SeriesCollection().Count == 0)
            {
                MessageBox.Show("The selected chart has no data series.", "No Data",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Series dataSeries = activeChart.SeriesCollection(1);

            // Get series values and categories
            string valuesFormula = dataSeries.Formula;
            string[] formulaParts = valuesFormula.Split(',');

            // Extract ranges from formula
            string categoryRange = formulaParts[1];
            string valuesRange = formulaParts[2];

            // Remove sheet references if present
            if (categoryRange.Contains("!"))
                categoryRange = categoryRange.Substring(categoryRange.IndexOf("!") + 1);
            if (valuesRange.Contains("!"))
                valuesRange = valuesRange.Substring(valuesRange.IndexOf("!") + 1);

            // Get range objects to determine column count
            Excel.Range catRange = worksheet.Range[categoryRange];
            int columnCount = catRange.Columns.Count;
            bool useAverageColumn = columnCount > 4;

            // Determine table size based on column count
            int tableColumns = useAverageColumn ? 5 : 4;
            string tableSize = useAverageColumn ? "6 rows x 6 columns" : "5 rows x 5 columns";

            // Prompt user to select location for intermediate table
            Excel.Range tableLocation = null;
            try
            {
                tableLocation = excelApp.InputBox(
                    $"Select a cell for the intermediate calculations table ({tableSize} will be used):",
                    "Select Table Location",
                    Type: 8) as Excel.Range;

                if (tableLocation == null)
                    return;
            }
            catch
            {
                return; // User cancelled
            }

            // Create intermediate table
            Excel.Range tableRange = worksheet.Range[tableLocation,
                tableLocation.Offset[4, tableColumns]];

            // Clear any existing data
            tableRange.Clear();

            Excel.Range valRange = worksheet.Range[valuesRange];

            // Determine if ranges are horizontal or vertical
            bool isVertical = catRange.Columns.Count == 1;

            // Get first and last cell references
            string firstValCell = isVertical ?
                valRange.Cells[1, 1].Address[false, false] :
                valRange.Cells[1, 1].Address[false, false];
            string lastValCell = isVertical ?
                valRange.Cells[valRange.Rows.Count, 1].Address[false, false] :
                valRange.Cells[1, valRange.Columns.Count].Address[false, false];

            string firstCatCell = isVertical ?
                catRange.Cells[1, 1].Address[false, false] :
                catRange.Cells[1, 1].Address[false, false];
            string lastCatCell = isVertical ?
                catRange.Cells[catRange.Rows.Count, 1].Address[false, false] :
                catRange.Cells[1, catRange.Columns.Count].Address[false, false];

            // Row 1: Label with CAGR calculation
            tableLocation.Value = "Label";
            tableLocation.Offset[0, 3].Formula =
                $"=({lastValCell}/{firstValCell})^(1/(COLUMNS({valuesRange})-1))-1";
            tableLocation.Offset[0, 3].NumberFormat = "\"CAGR: \" 0.0%";

            // Row 2: Y-Offset
            tableLocation.Offset[1, 0].Value = "Y-Offset";
            tableLocation.Offset[1, 1].Formula =
                $"=({lastValCell}-{firstValCell})/(COLUMNS({valuesRange})-1)";

            // Row 3: Y-Hook (same formula as Y-Offset)
            tableLocation.Offset[2, 0].Value = "Y-Hook";
            tableLocation.Offset[2, 1].Formula =
                $"=({lastValCell}-{firstValCell})/(COLUMNS({valuesRange})-1)";

            // Row 4: X-Axis
            tableLocation.Offset[3, 0].Value = "X-Axis";
            string offsetCell = tableLocation.Offset[1, 1].Address[false, false];
            string hookCell = tableLocation.Offset[2, 1].Address[false, false];

            if (useAverageColumn)
            {
                // 5 columns: 1, 1, average, COLUMNS, COLUMNS
                tableLocation.Offset[3, 1].Value = 1;
                tableLocation.Offset[3, 2].Value = 1;
                tableLocation.Offset[3, 4].Formula = $"=COLUMNS({categoryRange})";
                tableLocation.Offset[3, 5].Formula = $"=COLUMNS({categoryRange})";
                // Middle column = average of left and right
                tableLocation.Offset[3, 3].Formula =
                    $"=AVERAGE({tableLocation.Offset[3, 2].Address[false, false]}," +
                    $"{tableLocation.Offset[3, 4].Address[false, false]})";
            }
            else
            {
                // 4 columns: 1, 1, COLUMNS, COLUMNS
                tableLocation.Offset[3, 1].Value = 1;
                tableLocation.Offset[3, 2].Value = 1;
                tableLocation.Offset[3, 3].Formula = $"=COLUMNS({categoryRange})";
                tableLocation.Offset[3, 4].Formula = $"=COLUMNS({categoryRange})";
            }

            // Row 5: Y-Axis
            tableLocation.Offset[4, 0].Value = "Y-Axis";

            if (useAverageColumn)
            {
                // 5 columns
                // First column: first series value + Y-Offset
                tableLocation.Offset[4, 1].Formula = $"={firstValCell}+{offsetCell}";

                // Middle three columns: MAX of series range + Y-Offset + Y-Hook
                tableLocation.Offset[4, 2].Formula = $"=MAX({valuesRange})+{offsetCell}+{hookCell}";
                tableLocation.Offset[4, 3].Formula = $"=MAX({valuesRange})+{offsetCell}+{hookCell}";
                tableLocation.Offset[4, 4].Formula = $"=MAX({valuesRange})+{offsetCell}+{hookCell}";

                // Last column: last series value + Y-Offset
                tableLocation.Offset[4, 5].Formula = $"={lastValCell}+{offsetCell}";
            }
            else
            {
                // 4 columns
                // First column: first series value + Y-Offset
                tableLocation.Offset[4, 1].Formula = $"={firstValCell}+{offsetCell}";

                // Middle two columns: MAX of series range + Y-Offset + Y-Hook
                tableLocation.Offset[4, 2].Formula = $"=MAX({valuesRange})+{offsetCell}+{hookCell}";
                tableLocation.Offset[4, 3].Formula = $"=MAX({valuesRange})+{offsetCell}+{hookCell}";

                // Last column: last series value + Y-Offset
                tableLocation.Offset[4, 4].Formula = $"={lastValCell}+{offsetCell}";
            }

            // Add new series to chart
            Excel.Series cagrSeries = activeChart.SeriesCollection().NewSeries();
            cagrSeries.Name = "CAGR_Elbow_Series";

            // Set X and Y values for the series
            string xValuesRange, yValuesRange;
            int labelPointIndex;

            if (useAverageColumn)
            {
                // All 5 points
                xValuesRange = $"{tableLocation.Offset[3, 1].Address}:" +
                    $"{tableLocation.Offset[3, 5].Address}";
                yValuesRange = $"{tableLocation.Offset[4, 1].Address}:" +
                    $"{tableLocation.Offset[4, 5].Address}";
                labelPointIndex = 3; // Middle point
            }
            else
            {
                // All 4 points
                xValuesRange = $"{tableLocation.Offset[3, 1].Address}:" +
                    $"{tableLocation.Offset[3, 4].Address}";
                yValuesRange = $"{tableLocation.Offset[4, 1].Address}:" +
                    $"{tableLocation.Offset[4, 4].Address}";
                labelPointIndex = 2; // Second point (between the two middle points)
            }

            cagrSeries.XValues = worksheet.Range[xValuesRange];
            cagrSeries.Values = worksheet.Range[yValuesRange];

            // Change series to Scatter with Straight Lines
            cagrSeries.ChartType = Excel.XlChartType.xlXYScatterLines;

            // Hide the axis
            activeChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).Delete();
            activeChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary).Delete();

            // Format the series with arrow
            cagrSeries.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            cagrSeries.Format.Line.Weight = 1;
            cagrSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(
                System.Drawing.Color.Black);

            // Add arrow end at the last point
            cagrSeries.Format.Line.EndArrowheadStyle =
                Microsoft.Office.Core.MsoArrowheadStyle.msoArrowheadOpen;
            cagrSeries.Format.Line.EndArrowheadLength =
                Microsoft.Office.Core.MsoArrowheadLength.msoArrowheadLong;
            cagrSeries.Format.Line.EndArrowheadWidth =
                Microsoft.Office.Core.MsoArrowheadWidth.msoArrowheadWide;

            // Remove markers
            cagrSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;

            // Add data label to show CAGR percentage
            cagrSeries.HasDataLabels = true;

            // Clear all data labels first
            foreach (Excel.DataLabel lbl in cagrSeries.DataLabels())
            {
                lbl.Delete();
            }

            // Add label to the appropriate point
            cagrSeries.Points(labelPointIndex).HasDataLabel = true;
            Excel.DataLabel label = cagrSeries.Points(labelPointIndex).DataLabel;
            label.Position = Excel.XlDataLabelPosition.xlLabelPositionAbove;
            label.Format.TextFrame2.WordWrap = Microsoft.Office.Core.MsoTriState.msoFalse;

            // Cell that holds the CAGR formula
            Excel.Range cagrCell = tableLocation.Offset[0, 3];

            // Tell the data-label to use that cell's value
            label.Formula = "=" + cagrCell.Address[
                RowAbsolute: true,
                ColumnAbsolute: true,
                ReferenceStyle: Excel.XlReferenceStyle.xlA1,
                External: true          // prepends the sheet name—safer if the chart is on another sheet
            ];

            MessageBox.Show("CAGR elbow has been added successfully!", "Success",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}