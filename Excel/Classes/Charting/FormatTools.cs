//  FormatTools.cs
//  OfficePragmatics – Excel VSTO add-in   (C# 7.2, .NET Framework 4.x)

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;          //  MsoTriState enum, etc.

namespace OfficePragmatics_Excel.Classes.Charting
{
    /// <summary>Formatting helpers for embedded Excel charts.</summary>
    internal static class FormatTools
    {

        /// <summary>Re-orders stacked-column series so the largest last value plots at the bottom.</summary>
        public static void StackSort()
        {
            var app = Globals.ThisAddIn.Application;
            Excel.Chart chart = null;
            Excel.SeriesCollection sc = null;
            var info = new List<(Excel.Series ser, double last)>();

            try
            {
                chart = app.ActiveChart;
                if (chart == null)
                {
                    MessageBox.Show("Please select a stacked column chart first.",
                                    "No Chart Selected",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                if (chart.ChartType != Excel.XlChartType.xlColumnStacked &&
                    chart.ChartType != Excel.XlChartType.xlColumnStacked100)
                {
                    MessageBox.Show("StackSort works only on stacked column charts.",
                                    "Wrong Chart Type",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                sc = chart.SeriesCollection();
                if (sc == null || sc.Count == 0) return;

                for (int i = 1; i <= sc.Count; i++)
                {
                    var s = sc.Item(i);
                    var vals = s.Values as Array;
                    double last = 0;

                    if (vals != null)
                    {
                        for (int j = vals.Length - 1; j >= 0; j--)
                        {
                            if (vals.GetValue(j) is double d &&
                                !double.IsNaN(d) && !double.IsInfinity(d))
                            { last = d; break; }
                        }
                    }
                    info.Add((s, last));
                }

                var sorted = info.OrderByDescending(t => t.last).ToList();
                for (int i = 0; i < sorted.Count; i++)
                {
                    if (sorted[i].ser.PlotOrder != i + 1)
                        sorted[i].ser.PlotOrder = i + 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred in StackSort:\r\n{ex}",
                                "StackSort Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                foreach (var t in info) Marshal.ReleaseComObject(t.ser);
                if (sc != null) Marshal.ReleaseComObject(sc);
                if (chart != null) Marshal.ReleaseComObject(chart);
            }
        }

        /// <summary>Hide both Y-axes (labels, ticks, gridlines, line).</summary>
        public static void HideYAxis()
        {
            var ch = Globals.ThisAddIn.Application.ActiveChart;
            if (ch == null)
            {
                MessageBox.Show("Please select a chart first.",
                                "No Chart Selected",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            HideAxisCompletely(ch, Excel.XlAxisGroup.xlPrimary);
            HideAxisCompletely(ch, Excel.XlAxisGroup.xlSecondary);

            Marshal.ReleaseComObject(ch);
        }

        /// <summary>Force both Y-axes to start at zero (if the axis exists and is linear).</summary>
        public static void ZeroYAxis()
        {
            var ch = Globals.ThisAddIn.Application.ActiveChart;
            if (ch == null)
            {
                MessageBox.Show("Please select a chart first.",
                                "No Chart Selected",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ForceAxisStartsAtZero(ch, Excel.XlAxisGroup.xlPrimary);
            ForceAxisStartsAtZero(ch, Excel.XlAxisGroup.xlSecondary);

            Marshal.ReleaseComObject(ch);
        }

        /// <summary>
        /// Maximise the plot area (skips pie/doughnut). Determines legend side automatically.
        /// Two passes (default) let Excel recalc text sizes between runs.
        /// </summary>
        public static void MaxPlotArea(int passes = 2)
        {
            var app = Globals.ThisAddIn.Application;
            Excel.Chart ch = app.ActiveChart;

            if (ch == null)
            {
                MessageBox.Show("Please select a chart first.",
                                "No Chart Selected",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (IsPieLike(ch.ChartType))
            {
                MessageBox.Show("MaxPlotArea is not applicable to pie or doughnut charts.",
                                "Chart Type Skipped",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            for (int i = 0; i < passes; i++)
                ApplyMaxPlotArea(ch);
        }

        /// <summary>
        /// Centre the *plot area* inside the chart (works for all chart types).
        /// Ideal for pie / doughnut charts but safe for any chart.
        /// </summary>
        public static void CenterPlot()
        {
            var app = Globals.ThisAddIn.Application;

            Excel.Chart ch = null;
            Excel.Legend legend = null;
            Excel.ChartTitle cTitle = null;

            try
            {
                ch = app.ActiveChart;
                if (ch == null)
                {
                    MessageBox.Show("Please select a chart first.",
                                    "No Chart Selected",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                /*── build an inner rectangle that excludes legend & title ─*/
                double innerLeft = 0;
                double innerTop = 0;
                double innerWidth = ch.ChartArea.Width;
                double innerHeight = ch.ChartArea.Height;

                /* legend */
                if (ch.HasLegend)
                {
                    legend = ch.Legend;
                    switch (legend.Position)
                    {
                        case Excel.XlLegendPosition.xlLegendPositionLeft:
                            innerLeft += legend.Width;
                            innerWidth -= legend.Width;
                            break;

                        case Excel.XlLegendPosition.xlLegendPositionRight:
                            innerWidth -= legend.Width;
                            break;

                        case Excel.XlLegendPosition.xlLegendPositionTop:
                            innerTop += legend.Height;
                            innerHeight -= legend.Height;
                            break;

                        default: /* bottom / corner / custom → treat as bottom */
                            innerHeight -= legend.Height;
                            break;
                    }
                }

                /* chart title */
                if (ch.HasTitle)
                {
                    cTitle = ch.ChartTitle;
                    innerTop += cTitle.Height;
                    innerHeight -= cTitle.Height;
                }

                /*── centre the plot area within that rectangle ───────────*/
                double newLeft = innerLeft + (innerWidth - ch.PlotArea.Width) / 2.0;
                double newTop = innerTop + (innerHeight - ch.PlotArea.Height) / 2.0;

                ch.PlotArea.Left = (float)newLeft;
                ch.PlotArea.Top = (float)newTop;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred in CenterPlot:\r\n{ex}",
                                "CenterPlot Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (cTitle != null) Marshal.ReleaseComObject(cTitle);
                if (legend != null) Marshal.ReleaseComObject(legend);
                if (ch != null) Marshal.ReleaseComObject(ch);
            }
        }

        /*────────────────────────────  PRIVATE HELPERS  ──────────────────────────*/

        private static bool IsPieLike(Excel.XlChartType t) =>
               t == Excel.XlChartType.xlPie
            || t == Excel.XlChartType.xlPieExploded
            || t == Excel.XlChartType.xlDoughnut
            || t == Excel.XlChartType.xlDoughnutExploded;

        /// <summary>
        /// Try to get the specified axis; return <c>null</c> if Excel says it does not exist.
        /// </summary>
        private static Excel.Axis TryGetAxis(
            Excel.Chart ch,
            Excel.XlAxisType type,
            Excel.XlAxisGroup grp = Excel.XlAxisGroup.xlPrimary)
        {
            try { return (Excel.Axis)ch.Axes(type, grp); }
            catch { return null; }
        }

        private static void HideAxisCompletely(Excel.Chart ch, Excel.XlAxisGroup grp)
        {
            var ax = TryGetAxis(ch, Excel.XlAxisType.xlValue, grp);
            if (ax == null) return;

            ax.TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
            ax.MajorTickMark = Excel.XlTickMark.xlTickMarkNone;
            ax.MinorTickMark = Excel.XlTickMark.xlTickMarkNone;
            ax.HasMajorGridlines = false;
            ax.HasMinorGridlines = false;

            // Office 2007+ : hide the actual axis line
            ax.Format.Line.Visible = Office.MsoTriState.msoFalse;

            Marshal.ReleaseComObject(ax);
        }

        private static void ForceAxisStartsAtZero(Excel.Chart ch, Excel.XlAxisGroup grp)
        {
            var ax = TryGetAxis(ch, Excel.XlAxisType.xlValue, grp);
            if (ax == null) return;

            // log-scale axes cannot be forced to zero – skip them gracefully
            if (ax.ScaleType == Excel.XlScaleType.xlScaleLogarithmic)
            {
                Marshal.ReleaseComObject(ax);
                return;
            }

            ax.MinimumScale = 0;      // set the value first
            ax.MinimumScaleIsAuto = false;  // then lock it
            ax.CrossesAt = 0;      // ensure X-axis crosses at zero

            Marshal.ReleaseComObject(ax);
        }

        /// <summary>One optimisation pass; auto-detects legend side (no InputBox).</summary>
        private static void ApplyMaxPlotArea(Excel.Chart ch)
        {
            const Excel.XlAxisType xlValue = Excel.XlAxisType.xlValue;
            const Excel.XlAxisType xlCat = Excel.XlAxisType.xlCategory;

            Excel.Legend legend = null;
            Excel.Axis axP = null, axS = null, axC = null;

            double pvAxisW = 0, svAxisW = 0;
            double pvTitleW = 0, svTitleW = 0, pcTitleH = 0;
            double lgOffW = 0, lgOffH = 0, lgOffL = 0, lgOffT = 0;

            try
            {
                /*── legend offsets (auto side detection) ────────────────*/
                if (ch.HasLegend)
                {
                    legend = ch.Legend;
                    switch (legend.Position)
                    {
                        case Excel.XlLegendPosition.xlLegendPositionLeft:
                            lgOffW = -legend.Width;
                            lgOffL = legend.Width;
                            break;

                        case Excel.XlLegendPosition.xlLegendPositionRight:
                            lgOffW = -legend.Width;
                            break;

                        case Excel.XlLegendPosition.xlLegendPositionTop:
                            lgOffH = -legend.Height;
                            lgOffT = legend.Height;
                            break;

                        default: /* bottom / corner / custom → treat as bottom */
                            lgOffH = -legend.Height;
                            break;
                    }
                }

                /*── primary value axis ──────────────────────────────────*/
                if ((axP = TryGetAxis(ch, xlValue, Excel.XlAxisGroup.xlPrimary)) != null)
                {
                    pvAxisW = axP.Width;
                    if (axP.HasTitle)
                    {
                        pvTitleW = axP.AxisTitle.Width;
                        axP.AxisTitle.Left = (float)lgOffL;
                    }
                }

                /*── secondary value axis ────────────────────────────────*/
                if ((axS = TryGetAxis(ch, xlValue, Excel.XlAxisGroup.xlSecondary)) != null)
                {
                    svAxisW = axS.Width;
                    if (axS.HasTitle)
                    {
                        svTitleW = axS.AxisTitle.Width;
                        axS.AxisTitle.Left = (float)(pvTitleW + ch.PlotArea.Width + lgOffL);
                    }
                }

                /*── category axis title height ──────────────────────────*/
                if ((axC = TryGetAxis(ch, xlCat, Excel.XlAxisGroup.xlPrimary)) != null)
                {
                    if (axC.HasTitle) pcTitleH = axC.AxisTitle.Height;
                }

                /*── resize / reposition plot area ───────────────────────*/
                ch.PlotArea.Height = (float)(ch.ChartArea.Height - pcTitleH + lgOffH);
                ch.PlotArea.Top = (float)lgOffT;
                ch.PlotArea.Left = (float)(pvTitleW + lgOffL + 1.5);
                ch.PlotArea.Width = (float)(ch.ChartArea.Width
                                           - pvAxisW - svAxisW
                                           + lgOffW - 1.5);

                /*── move legend back where it belongs ───────────────────*/
                if (legend != null)
                {
                    switch (legend.Position)
                    {
                        case Excel.XlLegendPosition.xlLegendPositionLeft:
                            legend.Left = 0; break;

                        case Excel.XlLegendPosition.xlLegendPositionRight:
                            legend.Left = ch.PlotArea.Left + ch.PlotArea.Width + pvTitleW;
                            break;

                        case Excel.XlLegendPosition.xlLegendPositionTop:
                            legend.Top = 0; break;

                        default: /* bottom / custom */
                            legend.Top = ch.PlotArea.Top + ch.PlotArea.Height + pcTitleH;
                            break;
                    }
                }
            }
            finally
            {
                if (axC != null) Marshal.ReleaseComObject(axC);
                if (axS != null) Marshal.ReleaseComObject(axS);
                if (axP != null) Marshal.ReleaseComObject(axP);
                if (legend != null) Marshal.ReleaseComObject(legend);
            }
        }
    }
}