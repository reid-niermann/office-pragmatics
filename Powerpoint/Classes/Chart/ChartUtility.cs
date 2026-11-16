using Microsoft.VisualBasic;
using System;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_PowerPoint.Chart
{
    public static class ChartUtility
    {

        public static void PasteChartFormat(PowerPoint.ShapeRange sel)
        {
            try
            {
                var c1 = sel[1].Chart;
                var c2 = sel[2].Chart;
                string tmpl = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Microsoft", "Templates", "Charts", "temp.crtx");
                c1.SaveChartTemplate(tmpl);
                c2.ApplyChartTemplate(tmpl);
                if (System.IO.File.Exists(tmpl)) System.IO.File.Delete(tmpl);
            }
            catch
            {
                MessageBox.Show("Invalid selection. Please select two chart shapes.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void ShapesToChartAxis(PowerPoint.Application app)
        {
            try
            {
                int axisType = int.Parse(Interaction.InputBox("1 = horizontal axis, 2 = vertical axis", "Axis type", "1"));
                var sel = app.ActiveWindow.Selection.ShapeRange;
                int num = sel.Count;
                var chart = sel[num].Chart;
                var cats = (object[])chart.Axes(PowerPoint.XlAxisType.xlCategory, PowerPoint.XlAxisGroup.xlPrimary).CategoryNames;
                int catCount = cats.Length;
                double plotH = chart.PlotArea.InsideHeight, plotW = chart.PlotArea.InsideWidth;
                double baseL = chart.PlotArea.InsideLeft + chart.ChartArea.Left + sel[num].Left;
                double baseT = chart.PlotArea.InsideTop + chart.ChartArea.Top + sel[num].Top;

                if (axisType == 1) // horizontal
                {
                    int rows = 1, cols = catCount;
                    double totalH = 1, totalW = plotW;
                    double hGap = totalH / rows, wGap = totalW / cols;
                    double thisL = baseL, thisT = baseT + plotH + 25;

                    for (int k = 0; k < cols; k++)
                    {
                        thisL += wGap;
                        string target = cats[k]?.ToString() ?? "";
                        for (int z = 1; z < sel.Count; z++)
                        {
                            string alt = sel[z].AlternativeText?.TrimEnd('\r', '\n') ?? "";
                            if (string.Equals(target, alt, StringComparison.OrdinalIgnoreCase))
                            {
                                sel[z].Top = (float)(thisT - (hGap / 2) - (sel[z].Height / 2));
                                sel[z].Left = (float)(thisL - (wGap / 2) - (sel[z].Width / 2));
                                break;
                            }
                        }
                    }
                }
                else // vertical
                {
                    int rows = catCount, cols = 1;
                    double totalH = plotH, totalW = 1;
                    double hGap = totalH / rows, wGap = totalW / cols;
                    double thisL = baseL - 45, thisT = baseT + plotH;

                    for (int k = 0; k < rows; k++)
                    {
                        thisT -= hGap;
                        string target = cats[k]?.ToString() ?? "";
                        for (int z = 1; z < sel.Count; z++)
                        {
                            string alt = sel[z].AlternativeText?.TrimEnd('\r', '\n') ?? "";
                            if (string.Equals(target, alt, StringComparison.OrdinalIgnoreCase))
                            {
                                sel[z].Top = (float)(thisT + (hGap / 2) - (sel[z].Height / 2));
                                sel[z].Left = (float)(thisL - (wGap / 2) - (sel[z].Width / 2));
                                break;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        public static void ShapeToChartAltText(PowerPoint.Application app)
        {
            try
            {
                string input = Interaction.InputBox("1 = horizontal axis, 2 = vertical axis", "Axis type", "1");
                if (!int.TryParse(input, out int axisType)) return;
                var sel = app.ActiveWindow.Selection.ShapeRange;
                int chartIdx = sel.Count;
                var chart = sel[chartIdx].Chart;
                int numberOfObj = chart.ChartGroups(1).CategoryCollection.Count;
                double plotH = chart.PlotArea.InsideHeight, plotW = chart.PlotArea.InsideWidth;
                double plotL = chart.PlotArea.InsideLeft + chart.ChartArea.Left + sel[chartIdx].Left;
                double plotT = chart.PlotArea.InsideTop + chart.ChartArea.Top + sel[chartIdx].Top;

                if (axisType == 1)
                {
                    int numRow = 1, numCol = numberOfObj;
                    double totalH = 1, totalW = plotW;
                    int count = sel.Count - 1;
                    double startL = plotL, startT = plotT + plotH + 25;
                    double hGap = totalH / numRow, wGap = totalW / numCol;
                    int s = 1;
                    double thisT = startT;

                    for (int j = 1; j <= numRow; j++)
                    {
                        double thisL = startL;
                        thisT += hGap;
                        for (int k = 1; k <= numCol; k++)
                        {
                            if (j == numRow && (count % (numRow * numCol)) != 0)
                            {
                                int slots = numRow * numCol, b = count % slots;
                                int lastC = numCol - (slots - b);
                                double w2 = totalW / lastC;
                                thisL += w2;
                                var shp = sel[s];
                                shp.Top = (float)(thisT - (hGap / 2) - (shp.Height / 2));
                                shp.Left = (float)(thisL - (w2 / 2) - (shp.Width / 2));
                                s++;
                                if (k == lastC) return;
                            }
                            else
                            {
                                thisL += wGap;
                                var shp = sel[s];
                                shp.Top = (float)(thisT - (hGap / 2) - (shp.Height / 2));
                                shp.Left = (float)(thisL - (wGap / 2) - (shp.Width / 2));
                                s++;
                            }
                        }
                    }
                }
                else if (axisType == 2)
                {
                    int numRow = numberOfObj, numCol = 1;
                    double totalH = plotH, totalW = 1;
                    int count = sel.Count - 1;
                    double startL = plotL - 45, startT = plotT;
                    double hGap = totalH / numRow, wGap = totalW / numCol;
                    int s = 1;
                    double thisT = startT;

                    for (int j = 1; j <= numRow; j++)
                    {
                        double thisL = startL;
                        thisT += hGap;
                        for (int k = 1; k <= numCol; k++)
                        {
                            if (j == numRow && (count % (numRow * numCol)) != 0)
                            {
                                int slots = numRow * numCol, b = count % slots;
                                int lastR = numCol - (slots - b);
                                double w2 = totalW / lastR;
                                thisL += w2;
                                var shp = sel[s];
                                shp.Top = (float)(thisT - (hGap / 2) - (shp.Height / 2));
                                shp.Left = (float)(thisL - (w2 / 2) - (shp.Width / 2));
                                s++;
                                if (k == lastR) return;
                            }
                            else
                            {
                                thisL += wGap;
                                var shp = sel[s];
                                shp.Top = (float)(thisT - (hGap / 2) - (shp.Height / 2));
                                shp.Left = (float)(thisL - (wGap / 2) - (shp.Width / 2));
                                s++;
                            }
                        }
                    }
                }
            }
            catch { }
        }

        public static void MaxPlotArea(PowerPoint.ShapeRange sr)
        {
            try
            {
                var chart = sr[1].Chart;
                double pvTitleW = 0, svTitleW = 0, pcTitleH = 0;
                double lgOffW = 0, lgOffH = 0, lgOffL = 0, lgOffT = 0;
                int legendPos = 4;

                if (chart.HasLegend)
                {
                    legendPos = int.Parse(Interaction.InputBox("Legend position (1=left,2=right,3=top,4=bottom)", "Legend", "4"));
                    switch (legendPos)
                    {
                        case 1: lgOffW = -chart.Legend.Width; lgOffL = chart.Legend.Width; break;
                        case 2: lgOffW = -chart.Legend.Width; break;
                        case 3: lgOffH = -chart.Legend.Height; lgOffT = chart.Legend.Height; break;
                        case 4: lgOffH = -chart.Legend.Height; break;
                    }
                }

                if (chart.HasAxis(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary))
                {
                    var axP = chart.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary);
                    if (axP.HasTitle)
                    {
                        pvTitleW = axP.AxisTitle.Width;
                        axP.AxisTitle.Left = (float)(-chart.ChartArea.Left + lgOffL);
                    }
                }

                if (chart.HasAxis(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary))
                {
                    var axS = chart.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlSecondary);
                    if (axS.HasTitle)
                    {
                        svTitleW = axS.AxisTitle.Width;
                        axS.AxisTitle.Left = (float)(-chart.ChartArea.Left + pvTitleW + chart.PlotArea.Width + lgOffL);
                    }
                }

                if (chart.HasAxis(PowerPoint.XlAxisType.xlCategory, PowerPoint.XlAxisGroup.xlPrimary))
                {
                    var axC = chart.Axes(PowerPoint.XlAxisType.xlCategory, PowerPoint.XlAxisGroup.xlPrimary);
                    if (axC.HasTitle) pcTitleH = axC.AxisTitle.Height;
                }

                chart.PlotArea.Height = (float)(chart.ChartArea.Height - pcTitleH + lgOffH);
                chart.PlotArea.Top = (float)(-chart.ChartArea.Top + lgOffT);
                chart.PlotArea.Left = (float)(-chart.ChartArea.Left + pvTitleW + lgOffL + 0.75);
                chart.PlotArea.Width = (float)(chart.ChartArea.Width - pvTitleW - svTitleW + lgOffW - 1.5);

                if (chart.HasLegend)
                {
                    switch (legendPos)
                    {
                        case 1: chart.Legend.Left = -chart.ChartArea.Left; break;
                        case 2: chart.Legend.Left = chart.PlotArea.Left + chart.PlotArea.Width + pvTitleW; break;
                        case 3: chart.Legend.Top = -chart.ChartArea.Top; break;
                        case 4: chart.Legend.Top = chart.PlotArea.Top + chart.PlotArea.Height + pcTitleH; break;
                    }
                }
            }
            catch { }
        }

        public static void CenterPlotArea(PowerPoint.Selection sel)
        {
            try
            {
                if (sel == null ||
                    sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.ShapeRange == null || sel.ShapeRange.Count < 1 ||
                    sel.ShapeRange[1].HasChart != Office.MsoTriState.msoTrue)
                {
                    MessageBox.Show("Select a single chart shape first.", "Center Plot",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                var chart = sel.ShapeRange[1].Chart;

                // Default adjustments
                double adjLeft = 0.0, adjTop = 0.0;

                if (chart.HasLegend)
                {
                    // Use legend size as defaults (negative), mirroring the VBA behavior
                    double legendW = (double)chart.Legend.Width;
                    double legendH = (double)chart.Legend.Height;

                    string defLeft = (-legendW).ToString(CultureInfo.InvariantCulture);
                    string defTop = (-legendH).ToString(CultureInfo.InvariantCulture);

                    string inLeft = Interaction.InputBox(
                        "Enter left adjustment. Legend width is " + legendW.ToString("0.##"),
                        "Axis/Legend Adjustment – Left", defLeft);

                    string inTop = Interaction.InputBox(
                        "Enter top adjustment. Legend height is " + legendH.ToString("0.##"),
                        "Axis/Legend Adjustment – Top", defTop);

                    // Parse with invariant first, then fall back to current culture
                    double tmp;
                    if (double.TryParse(inLeft, NumberStyles.Float, CultureInfo.InvariantCulture, out tmp) ||
                        double.TryParse(inLeft, NumberStyles.Float, CultureInfo.CurrentCulture, out tmp))
                        adjLeft = tmp;

                    if (double.TryParse(inTop, NumberStyles.Float, CultureInfo.InvariantCulture, out tmp) ||
                        double.TryParse(inTop, NumberStyles.Float, CultureInfo.CurrentCulture, out tmp))
                        adjTop = tmp;
                }

                // Center PlotArea within ChartArea, then apply user adjustments
                double chartW = (double)chart.ChartArea.Width;
                double chartH = (double)chart.ChartArea.Height;
                double plotW = (double)chart.PlotArea.Width;
                double plotH = (double)chart.PlotArea.Height;

                double chartLeft = (double)chart.ChartArea.Left;
                double chartTop = (double)chart.ChartArea.Top;

                double newLeft = ((chartW - plotW) / 2.0) - chartLeft + adjLeft;
                double newTop = ((chartH - plotH) / 2.0) - chartTop + adjTop;

                // Interop expects Single in many cases; cast to float to avoid COM type fuss
                chart.PlotArea.Left = (float)newLeft;
                chart.PlotArea.Top = (float)newTop;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not center the plot area:\r\n" + ex.Message,
                    "Center Plot", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void CopyChartData(PowerPoint.Selection sel)
        {
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0 || sel.ShapeRange[1].HasChart != Office.MsoTriState.msoTrue)
            {
                MessageBox.Show("Select a chart first.");
                return;
            }

            try
            {
                var chart = sel.ShapeRange[1].Chart;
                chart.ChartData.Workbook.Application.Visible = true;
                chart.ChartData.Workbook.ActiveSheet.Range["A1"].CurrentRegion.Copy();
                MessageBox.Show("Chart data copied – paste it anywhere.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not copy chart data:\r\n" + ex.Message);
            }
        }
    }
}