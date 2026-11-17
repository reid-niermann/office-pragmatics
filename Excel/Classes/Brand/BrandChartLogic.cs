// ──────────────────────────────────────────────────────────────────────────────
// BrandChartLogic.cs  (C# 7.3 / .NET 4.8)
// ──────────────────────────────────────────────────────────────────────────────
using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes.Brand
{
    #region ▌ Data models ▌
    public class BrandChartSettings
    {
        public List<ChartFormatSettings> SecondaryCharts { get; set; } = new List<ChartFormatSettings>();
    }

    public class ChartFormatSettings
    {
        public bool IsSet { get; set; }
        public bool IsBaseChart { get; set; }

        public Excel.XlChartType ChartType { get; set; }

        // If this format is just an alias, point to another chart type
        public Excel.XlChartType? SameAsChartType { get; set; }

        // (Name removed from UI; keep internally if you like, but unused)
        public string Name { get; set; }

        public bool? HasLegend { get; set; }
        public Excel.XlLegendPosition? LegendPosition { get; set; }

        public bool? HasTitle { get; set; }
        public FontFormat TitleFont { get; set; } = new FontFormat();

        public List<SeriesFormat> SeriesFormats { get; set; } = new List<SeriesFormat>();

        public AxisFormat PrimaryValueAxis { get; set; } = new AxisFormat();
        public AxisFormat PrimaryCategoryAxis { get; set; } = new AxisFormat();
        public AxisFormat SecondaryValueAxis { get; set; } = new AxisFormat();
        public AxisFormat SecondaryCategoryAxis { get; set; } = new AxisFormat();

        public int? GapWidth { get; set; }
        public int? Overlap { get; set; }

        public PlotAreaPadding PlotArea { get; set; } = new PlotAreaPadding();
    }

    public class SeriesFormat
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public Excel.XlAxisGroup? AxisGroup { get; set; }

        public FillFormat Fill { get; set; } = new FillFormat();
        public LineFormat Border { get; set; } = new LineFormat();
        public MarkerFormat Marker { get; set; } = new MarkerFormat();
        public DataLabelFormat DataLabels { get; set; } = new DataLabelFormat();

        public List<PointFormat> Points { get; set; } = new List<PointFormat>();
    }

    public class PointFormat
    {
        public int Index { get; set; }
        public FillFormat Fill { get; set; } = new FillFormat();
        public LineFormat Border { get; set; } = new LineFormat();
        public int? Explosion { get; set; }
    }

    public class MarkerFormat
    {
        public bool IsSet { get; set; }
        public bool? IsVisible { get; set; }
        public Excel.XlMarkerStyle? Style { get; set; }
        public int? Size { get; set; }
        public Color? ForeColor { get; set; }
        public Color? BackColor { get; set; }
    }

    public class AxisFormat
    {
        public bool IsSet { get; set; }
        public bool? IsVisible { get; set; }
        public bool? HasMajorGridlines { get; set; }
        public LineFormat MajorGridlinesFormat { get; set; } = new LineFormat();
        public bool? HasMinorGridlines { get; set; }
        public Excel.XlTickLabelPosition? TickLabelPosition { get; set; }
        public AxisTitleFormat Title { get; set; } = new AxisTitleFormat();
    }

    public class AxisTitleFormat
    {
        public bool IsSet { get; set; }
        public bool? HasTitle { get; set; }
        public string Text { get; set; }
        public FontFormat Font { get; set; } = new FontFormat();
    }

    public class DataLabelFormat
    {
        public bool IsSet { get; set; }
        public bool? IsVisible { get; set; }
        public Excel.XlDataLabelPosition? Position { get; set; }
        public bool? ShowValue { get; set; }
        public FontFormat Font { get; set; } = new FontFormat();
    }

    public class FontFormat
    {
        public bool IsSet { get; set; }
        public string Name { get; set; }
        public double? Size { get; set; }
        public bool? Bold { get; set; }
        public Color? Color { get; set; }
    }

    public class FillFormat
    {
        public bool IsSet { get; set; }
        public Color? ForeColor { get; set; }
        public MsoTriState Visible { get; set; } = MsoTriState.msoFalse;
    }

    public class LineFormat
    {
        public bool IsSet { get; set; }
        public Color? ForeColor { get; set; }
        public double? Weight { get; set; }
        public MsoLineDashStyle? DashStyle { get; set; }
        public MsoTriState Visible { get; set; } = MsoTriState.msoFalse;
    }

    public class PlotAreaPadding
    {
        public bool IsSet { get; set; }
        public bool? Auto { get; set; }
        public double? Left { get; set; }
        public double? Top { get; set; }
        public double? Width { get; set; }
        public double? Height { get; set; }
    }
    #endregion

    #region ▌ Core engine ▌
    public static class ChartFormat
    {
        public static void ApplyBrandFormattingToSelection(BrandChartSettings settings)
        {
            if (settings == null) return;

            Excel.Chart chart = GetActiveChart();
            if (chart == null)
            {
                MessageBox.Show("Please select a single chart to format.", "Pragmatics",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Base pass
            var baseFmt = settings.SecondaryCharts.FirstOrDefault(s => s.IsBaseChart && s.IsSet);
            if (baseFmt != null) ApplySettingsToChart(chart, baseFmt);

            // Specific or alias
            var fmt = ResolveChartSettings(settings.SecondaryCharts, chart.ChartType);
            if (fmt != null && fmt != baseFmt) ApplySettingsToChart(chart, fmt);
        }

        private static ChartFormatSettings ResolveChartSettings(List<ChartFormatSettings> list, Excel.XlChartType type)
        {
            // Direct
            var direct = list.FirstOrDefault(x => x.ChartType == type && x.IsSet);
            if (direct != null) return direct;

            // Placeholder with SameAs
            var alias = list.FirstOrDefault(x => x.ChartType == type && x.SameAsChartType.HasValue);
            if (alias != null)
            {
                var pointed = list.FirstOrDefault(x => x.ChartType == alias.SameAsChartType.Value && x.IsSet);
                if (pointed != null) return pointed;
            }
            return null;
        }

        public static ChartFormatSettings ImportSettingsFromChart(Excel.Chart chart)
        {
            if (chart == null) throw new ArgumentNullException(nameof(chart));

            var set = new ChartFormatSettings
            {
                IsSet = true,
                Name = chart.Name,
                ChartType = chart.ChartType,
                HasTitle = chart.HasTitle,
                HasLegend = chart.HasLegend
            };

            if (chart.HasTitle)
            {
                set.TitleFont = new FontFormat
                {
                    IsSet = true,
                    Name = chart.ChartTitle.Font.Name,
                    Size = (double)chart.ChartTitle.Font.Size,
                    Bold = (bool)chart.ChartTitle.Font.Bold,
                    Color = ColorTranslator.FromOle(Convert.ToInt32(chart.ChartTitle.Font.Color))
                };
            }

            if (chart.HasLegend)
                set.LegendPosition = chart.Legend.Position;

            try { dynamic cg = chart.ChartGroups(1); set.GapWidth = cg.GapWidth; set.Overlap = cg.Overlap; } catch { }

            ImportAxis(chart, Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary, set.PrimaryValueAxis);
            ImportAxis(chart, Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary, set.PrimaryCategoryAxis);
            ImportAxis(chart, Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary, set.SecondaryValueAxis);
            ImportAxis(chart, Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary, set.SecondaryCategoryAxis);

            try
            {
                dynamic dChart = (dynamic)chart;
                dynamic ip = dChart.InnerPlotPosition;
                set.PlotArea = new PlotAreaPadding
                {
                    IsSet = true,
                    Auto = ip.Auto,
                    Left = (double)ip.Left,
                    Top = (double)ip.Top,
                    Width = (double)ip.Width,
                    Height = (double)ip.Height
                };
            }
            catch { }

            var sc = (Excel.SeriesCollection)chart.SeriesCollection();
            int idxCounter = 1;

            foreach (Excel.Series s in sc)
            {
                var sf = new SeriesFormat { Name = s.Name, Index = idxCounter++, AxisGroup = SafeGetAxisGroup(s) };

                TryImportFill(s.Format.Fill, sf.Fill);
                TryImportLine(s.Format.Line, sf.Border);
                TryImportMarkers(s, sf.Marker);
                TryImportDataLabels(s, sf.DataLabels);
                TryImportPoints(set.ChartType, s, sf.Points);

                set.SeriesFormats.Add(sf);
            }

            return set;
        }

        #region --- import helpers ---
        private static Excel.XlAxisGroup? SafeGetAxisGroup(Excel.Series s)
        {
            try { return s.AxisGroup; } catch { return null; }
        }

        private static void TryImportFill(dynamic fill, FillFormat ff)
        {
            ff.IsSet = true;
            Color? col = null;
            MsoTriState vis = MsoTriState.msoFalse;
            try { vis = fill.Visible; } catch { }
            try { col = ColorTranslator.FromOle(fill.ForeColor.RGB); } catch { }

            ff.Visible = (vis == MsoTriState.msoTrue || col.HasValue) ? MsoTriState.msoTrue : MsoTriState.msoFalse;
            ff.ForeColor = col;
        }

        private static void TryImportLine(dynamic line, LineFormat lf)
        {
            try
            {
                bool has = line.Visible == MsoTriState.msoTrue && line.Weight > 0.0001;
                lf.IsSet = true;
                lf.Visible = has ? MsoTriState.msoTrue : MsoTriState.msoFalse;
                if (has)
                {
                    lf.ForeColor = ColorTranslator.FromOle(line.ForeColor.RGB);
                    lf.Weight = line.Weight;
                    lf.DashStyle = line.DashStyle;
                }
            }
            catch { }
        }

        private static void TryImportMarkers(Excel.Series s, MarkerFormat mf)
        {
            try
            {
                var style = s.MarkerStyle;
                bool visible = style != Excel.XlMarkerStyle.xlMarkerStyleNone;
                mf.IsSet = true;
                mf.IsVisible = visible;
                if (visible)
                {
                    mf.Style = style;
                    mf.Size = s.MarkerSize;
                    try { mf.ForeColor = ColorTranslator.FromOle(s.MarkerForegroundColor); } catch { }
                    try { mf.BackColor = ColorTranslator.FromOle(s.MarkerBackgroundColor); } catch { }
                }
            }
            catch { mf.IsVisible = false; }
        }

        private static void TryImportDataLabels(Excel.Series s, DataLabelFormat dlf)
        {
            dlf.IsSet = true;
            try
            {
                bool has = false;
                try { has = s.HasDataLabels; } catch { }
                dlf.IsVisible = has;
                if (has)
                {
                    Excel.DataLabels dl = s.DataLabels();
                    dlf.Position = dl.Position;
                    dlf.ShowValue = dl.ShowValue;
                    dlf.Font = new FontFormat
                    {
                        IsSet = true,
                        Name = dl.Font.Name,
                        Size = (double?)dl.Font.Size,
                        Bold = (bool?)dl.Font.Bold,
                        Color = ColorTranslator.FromOle(Convert.ToInt32(dl.Font.Color))
                    };
                }
            }
            catch { dlf.IsVisible = false; }
        }

        private static void TryImportPoints(Excel.XlChartType ct, Excel.Series s, List<PointFormat> list)
        {
            if (!IsPointColoredChartType(ct)) return;
            try
            {
                int count = s.Points().Count;
                for (int i = 1; i <= count; i++)
                {
                    var pf = new PointFormat { Index = i };
                    dynamic pt = s.Points(i);
                    TryImportFill(pt.Format.Fill, pf.Fill);
                    TryImportLine(pt.Format.Line, pf.Border);
                    try { pf.Explosion = pt.Explosion; } catch { }
                    list.Add(pf);
                }
            }
            catch { }
        }

        private static bool IsPointColoredChartType(Excel.XlChartType ct)
        {
            string name = ct.ToString();
            return name.IndexOf("Pie", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("Doughnut", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static void ImportAxis(Excel.Chart chart, Excel.XlAxisType type, Excel.XlAxisGroup group, AxisFormat tgt)
        {
            try
            {
                Excel.Axis ax = chart.Axes(type, group);
                tgt.IsSet = true;

                tgt.IsVisible = ax.Format.Line.Visible == MsoTriState.msoTrue;
                tgt.TickLabelPosition = ax.TickLabelPosition;

                tgt.HasMajorGridlines = ax.HasMajorGridlines;
                if (ax.HasMajorGridlines)
                {
                    tgt.MajorGridlinesFormat = new LineFormat { IsSet = true };
                    TryImportLine(ax.MajorGridlines.Format.Line, tgt.MajorGridlinesFormat);
                }

                tgt.HasMinorGridlines = ax.HasMinorGridlines;

                var t = tgt.Title;
                t.IsSet = true;
                try
                {
                    bool hasTitle = ax.HasTitle;
                    t.HasTitle = hasTitle;
                    if (hasTitle)
                    {
                        dynamic at = ax.AxisTitle;
                        t.Text = (string)at.Text;
                        t.Font = new FontFormat
                        {
                            IsSet = true,
                            Name = at.Font.Name,
                            Size = (double?)at.Font.Size,
                            Bold = (bool?)at.Font.Bold,
                            Color = ColorTranslator.FromOle(Convert.ToInt32(at.Font.Color))
                        };
                    }
                }
                catch { t.HasTitle = false; }
            }
            catch { }
        }
        #endregion

        #region --- apply helpers ---
        private static void ApplySettingsToChart(Excel.Chart chart, ChartFormatSettings s)
        {
            if (s.HasTitle.HasValue)
            {
                chart.HasTitle = s.HasTitle.Value;
                if (chart.HasTitle && s.TitleFont?.IsSet == true)
                    ApplyFont(chart.ChartTitle.Font, s.TitleFont);
            }
            if (s.HasLegend.HasValue)
            {
                chart.HasLegend = s.HasLegend.Value;
                if (chart.HasLegend && s.LegendPosition.HasValue)
                    chart.Legend.Position = s.LegendPosition.Value;
            }

            try
            {
                foreach (dynamic cg in chart.ChartGroups())
                {
                    if (s.GapWidth.HasValue) cg.GapWidth = s.GapWidth.Value;
                    if (s.Overlap.HasValue) cg.Overlap = s.Overlap.Value;
                }
            }
            catch { }

            ApplyAxis(chart, Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary, s.PrimaryValueAxis);
            ApplyAxis(chart, Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary, s.PrimaryCategoryAxis);
            ApplyAxis(chart, Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary, s.SecondaryValueAxis);
            ApplyAxis(chart, Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlSecondary, s.SecondaryCategoryAxis);

            if (s.PlotArea.IsSet)
            {
                try
                {
                    dynamic dChart = (dynamic)chart;
                    dynamic ip = dChart.InnerPlotPosition;
                    if (s.PlotArea.Auto.HasValue && s.PlotArea.Auto.Value == false) ip.Auto = false;
                    if (s.PlotArea.Left.HasValue) ip.Left = s.PlotArea.Left.Value;
                    if (s.PlotArea.Top.HasValue) ip.Top = s.PlotArea.Top.Value;
                    if (s.PlotArea.Width.HasValue) ip.Width = s.PlotArea.Width.Value;
                    if (s.PlotArea.Height.HasValue) ip.Height = s.PlotArea.Height.Value;
                }
                catch { }
            }

            var sc = (Excel.SeriesCollection)chart.SeriesCollection();
            int total = sc.Count;

            foreach (var sf in s.SeriesFormats)
            {
                Excel.Series ts = null;
                try { ts = sc.Item(sf.Name); } catch { }
                if (ts == null && sf.Index > 0 && sf.Index <= total)
                {
                    try { ts = sc.Item(sf.Index); } catch { }
                }
                if (ts == null) continue;

                if (sf.AxisGroup.HasValue) { try { ts.AxisGroup = sf.AxisGroup.Value; } catch { } }

                // Fill
                if (sf.Fill.IsSet)
                {
                    if (sf.Fill.Visible == MsoTriState.msoFalse)
                    {
                        ts.Format.Fill.Visible = MsoTriState.msoFalse;
                    }
                    else
                    {
                        ts.Format.Fill.Visible = MsoTriState.msoTrue;
                        try { ts.Format.Fill.Solid(); } catch { }
                        if (sf.Fill.ForeColor.HasValue)
                            ts.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(sf.Fill.ForeColor.Value);
                    }
                }

                // Border
                if (sf.Border.IsSet)
                {
                    if (sf.Border.Visible == MsoTriState.msoTrue)
                    {
                        ts.Format.Line.Visible = MsoTriState.msoTrue;
                        if (sf.Border.ForeColor.HasValue)
                            ts.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(sf.Border.ForeColor.Value);
                        if (sf.Border.Weight.HasValue)
                            ts.Format.Line.Weight = (float)sf.Border.Weight.Value;
                        if (sf.Border.DashStyle.HasValue)
                            ts.Format.Line.DashStyle = sf.Border.DashStyle.Value;
                    }
                    else
                    {
                        ts.Format.Line.Visible = MsoTriState.msoFalse;
                        try { ts.Format.Line.Weight = 0f; } catch { }
                    }
                }

                // Markers
                if (sf.Marker.IsSet && sf.Marker.IsVisible.HasValue)
                {
                    bool showMarkers = sf.Marker.IsVisible.Value;
                    try
                    {
                        if (!showMarkers)
                        {
                            ts.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
                        }
                        else
                        {
                            ts.MarkerStyle = sf.Marker.Style ?? Excel.XlMarkerStyle.xlMarkerStyleAutomatic;
                            if (sf.Marker.Size.HasValue) ts.MarkerSize = sf.Marker.Size.Value;
                            if (sf.Marker.ForeColor.HasValue)
                                ts.MarkerForegroundColor = ColorTranslator.ToOle(sf.Marker.ForeColor.Value);
                            if (sf.Marker.BackColor.HasValue)
                                ts.MarkerBackgroundColor = ColorTranslator.ToOle(sf.Marker.BackColor.Value);
                        }
                    }
                    catch { }
                }

                // Data labels
                if (sf.DataLabels.IsSet && sf.DataLabels.IsVisible.HasValue)
                {
                    bool show = sf.DataLabels.IsVisible.Value;
                    ts.HasDataLabels = show;

                    if (show)
                    {
                        try { ts.ApplyDataLabels(); } catch { }
                        Excel.DataLabels dl = ts.DataLabels();
                        if (sf.DataLabels.Position.HasValue) dl.Position = sf.DataLabels.Position.Value;
                        if (sf.DataLabels.ShowValue.HasValue) dl.ShowValue = sf.DataLabels.ShowValue.Value;
                        if (sf.DataLabels.Font?.IsSet == true) ApplyFont(dl.Font, sf.DataLabels.Font);
                    }
                }

                // Points
                if (sf.Points != null && sf.Points.Count > 0)
                {
                    try
                    {
                        int count = ts.Points().Count;
                        foreach (var pf in sf.Points)
                        {
                            if (pf.Index < 1 || pf.Index > count) continue;
                            dynamic pt = ts.Points(pf.Index);

                            if (pf.Fill.IsSet)
                            {
                                if (pf.Fill.Visible == MsoTriState.msoFalse)
                                {
                                    pt.Format.Fill.Visible = MsoTriState.msoFalse;
                                }
                                else
                                {
                                    pt.Format.Fill.Visible = MsoTriState.msoTrue;
                                    try { pt.Format.Fill.Solid(); } catch { }
                                    if (pf.Fill.ForeColor.HasValue)
                                        pt.Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(pf.Fill.ForeColor.Value);
                                }
                            }

                            if (pf.Border.IsSet)
                            {
                                if (pf.Border.Visible == MsoTriState.msoTrue)
                                {
                                    pt.Format.Line.Visible = MsoTriState.msoTrue;
                                    if (pf.Border.ForeColor.HasValue)
                                        pt.Format.Line.ForeColor.RGB = ColorTranslator.ToOle(pf.Border.ForeColor.Value);
                                    if (pf.Border.Weight.HasValue)
                                        pt.Format.Line.Weight = (float)pf.Border.Weight.Value;
                                    if (pf.Border.DashStyle.HasValue)
                                        pt.Format.Line.DashStyle = pf.Border.DashStyle.Value;
                                }
                                else
                                {
                                    pt.Format.Line.Visible = MsoTriState.msoFalse;
                                    try { pt.Format.Line.Weight = 0f; } catch { }
                                }
                            }

                            if (pf.Explosion.HasValue) { try { pt.Explosion = pf.Explosion.Value; } catch { } }
                        }
                    }
                    catch { }
                }
            }
        }

        // explicit show/hide & style helper for gridlines
        private static void ApplyGridlines(Excel.Gridlines gl, bool turnOn, LineFormat fmt)
        {
            try
            {
                var ln = gl.Format.Line;
                if (!turnOn)
                {
                    ln.Visible = MsoTriState.msoFalse;
                    return;
                }

                ln.Visible = MsoTriState.msoTrue;
                if (fmt != null && fmt.IsSet)
                {
                    if (fmt.ForeColor.HasValue) ln.ForeColor.RGB = ColorTranslator.ToOle(fmt.ForeColor.Value);
                    if (fmt.Weight.HasValue) ln.Weight = (float)fmt.Weight.Value;
                    if (fmt.DashStyle.HasValue) ln.DashStyle = fmt.DashStyle.Value;
                }
                else
                {
                    try { ln.Weight = 0.75f; } catch { }
                }
            }
            catch { }
        }

        private static void ApplyAxis(Excel.Chart chart,
                                      Excel.XlAxisType type,
                                      Excel.XlAxisGroup group,
                                      AxisFormat src)
        {
            if (!src.IsSet) return;

            try
            {
                Excel.Axis ax = chart.Axes(type, group);

                if (src.IsVisible.HasValue)
                    ax.Format.Line.Visible = src.IsVisible.Value ? MsoTriState.msoTrue : MsoTriState.msoFalse;

                if (src.TickLabelPosition.HasValue)
                    ax.TickLabelPosition = src.TickLabelPosition.Value;

                if (src.HasMajorGridlines.HasValue)
                {
                    ax.HasMajorGridlines = src.HasMajorGridlines.Value;
                    ApplyGridlines(ax.MajorGridlines, src.HasMajorGridlines.Value, src.MajorGridlinesFormat);
                }

                if (src.HasMinorGridlines.HasValue)
                {
                    ax.HasMinorGridlines = src.HasMinorGridlines.Value;
                    if (!src.HasMinorGridlines.Value)
                    {
                        try { ax.MinorGridlines.Format.Line.Visible = MsoTriState.msoFalse; } catch { }
                    }
                }

                if (src.Title.IsSet && src.Title.HasTitle.HasValue)
                {
                    bool show = src.Title.HasTitle.Value;
                    try { ax.HasTitle = show; } catch { }
                    if (show)
                    {
                        dynamic at = ax.AxisTitle;
                        try { at.Text = src.Title.Text ?? string.Empty; } catch { }
                        if (src.Title.Font?.IsSet == true) ApplyFont(at.Font, src.Title.Font);
                    }
                }
            }
            catch { }
        }

        private static void ApplyFont(dynamic fontObj, FontFormat ff)
        {
            try
            {
                if (!string.IsNullOrEmpty(ff.Name)) fontObj.Name = ff.Name;
                if (ff.Size.HasValue) fontObj.Size = (float)ff.Size.Value;
                if (ff.Bold.HasValue) fontObj.Bold = ff.Bold.Value;
                if (ff.Color.HasValue) fontObj.Color = ColorTranslator.ToOle(ff.Color.Value);
            }
            catch { }
        }

        public static Excel.Chart GetActiveChart()
        {
            try
            {
                if (Globals.ThisAddIn.Application.ActiveChart != null)
                    return Globals.ThisAddIn.Application.ActiveChart;

                if (Globals.ThisAddIn.Application.ActiveWindow.Selection is Excel.Shape shp &&
                    shp.HasChart == MsoTriState.msoTrue)
                    return shp.Chart;

                if (Globals.ThisAddIn.Application.Selection is Excel.ChartObject co)
                    return co.Chart;
            }
            catch { }
            return null;
        }
        #endregion
    }
    #endregion
}
