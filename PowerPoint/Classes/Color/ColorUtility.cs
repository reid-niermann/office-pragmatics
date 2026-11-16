using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_PowerPoint.Color
{
    public static class ColorUtility
    {
        public static void ColorFreeze(PowerPoint.Application app)
        {
            if (app == null || app.ActivePresentation == null)
            {
                MessageBox.Show("No active presentation found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var pres = app.ActivePresentation;
            var sel = app.ActiveWindow.Selection;
            var shapesToProcess = new List<PowerPoint.Shape>();
            var slidesToProcess = new List<PowerPoint.Slide>();

            try
            {
                // Determine what to process based on selection
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null && sel.ShapeRange.Count > 0)
                {
                    shapesToProcess.AddRange(sel.ShapeRange.Cast<PowerPoint.Shape>());
                    Debug.WriteLine($"Processing {shapesToProcess.Count} selected shapes.");
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionSlides && sel.SlideRange != null && sel.SlideRange.Count > 0)
                {
                    slidesToProcess.AddRange(sel.SlideRange.Cast<PowerPoint.Slide>());
                    Debug.WriteLine($"Processing {slidesToProcess.Count} selected slides.");
                }
                else // Default: process all slides if no specific shape or slide selection
                {
                    if (pres.Slides != null && pres.Slides.Count > 0)
                    {
                        slidesToProcess.AddRange(pres.Slides.Cast<PowerPoint.Slide>());
                        Debug.WriteLine($"Processing all {slidesToProcess.Count} slides in the presentation.");
                    }
                    else
                    {
                        MessageBox.Show("No slides found in the presentation to process.", "Freeze Colors", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                }

                // If slides are being processed, gather their items and shapes from masters/layouts
                var shapesFromSlidesAndMasters = new List<PowerPoint.Shape>();
                foreach (var slide in slidesToProcess)
                {
                    Debug.WriteLine($"Gathering items from Slide {slide.SlideIndex}");

                    // Slide Background
                    if (slide.Background?.Fill != null && slide.Background.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        Debug.WriteLine("  Freezing slide.Background.Fill");
                        FreezeFillFormat(slide.Background.Fill);
                    }
                    // Custom Layout Background (if it doesn't follow master and is visible)
                    if (slide.FollowMasterBackground == Office.MsoTriState.msoFalse &&
                        slide.CustomLayout?.Background?.Fill != null &&
                        slide.CustomLayout.Background.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        Debug.WriteLine("  Freezing slide.CustomLayout.Background.Fill");
                        FreezeFillFormat(slide.CustomLayout.Background.Fill);
                    }
                    // Slide Master Background (freeze it as items on slide/layout might be transparent to it)
                    if (slide.Design?.SlideMaster?.Background?.Fill != null &&
                        slide.Design.SlideMaster.Background.Fill.Visible == Office.MsoTriState.msoTrue)
                    {
                        Debug.WriteLine("  Freezing slide.Design.SlideMaster.Background.Fill");
                        FreezeFillFormat(slide.Design.SlideMaster.Background.Fill);
                    }

                    // Shapes on the slide itself
                    if (slide.Shapes != null) shapesFromSlidesAndMasters.AddRange(slide.Shapes.Cast<PowerPoint.Shape>());

                    // Shapes on the slide master (process all, placeholders will be handled or skipped if not applicable)
                    if (slide.Design?.SlideMaster?.Shapes != null)
                    {
                        shapesFromSlidesAndMasters.AddRange(slide.Design.SlideMaster.Shapes.Cast<PowerPoint.Shape>());
                    }
                    // Shapes on the custom layout
                    if (slide.CustomLayout?.Shapes != null)
                    {
                        shapesFromSlidesAndMasters.AddRange(slide.CustomLayout.Shapes.Cast<PowerPoint.Shape>());
                    }
                }

                // Combine explicitly selected shapes with those from slides/masters
                shapesToProcess.AddRange(shapesFromSlidesAndMasters);
                var distinctShapes = shapesToProcess.Distinct(new ShapeEqualityComparer()).ToList();
                Debug.WriteLine($"Total distinct shapes to process: {distinctShapes.Count}");

                foreach (var shp in distinctShapes)
                {
                    ProcessShape(shp);
                }

                MessageBox.Show("Attempted to freeze theme-based colors to explicit RGB (excluding tables, charts, and SmartArt).", "Freeze Colors", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"ColorFreeze Error: {ex.ToString()}");
            }
        }

        private class ShapeEqualityComparer : IEqualityComparer<PowerPoint.Shape>
        {
            public bool Equals(PowerPoint.Shape x, PowerPoint.Shape y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (x is null || y is null) return false;
                try { return x.Id == y.Id; } catch { return false; }
            }

            public int GetHashCode(PowerPoint.Shape obj)
            {
                try { return obj?.Id.GetHashCode() ?? 0; } catch { return 0; }
            }
        }

        private static void ProcessShape(PowerPoint.Shape shape)
        {
            if (shape == null) return;

            // — skip non‑content placeholders on masters/layouts —
            try
            {
                if ((shape.Parent is PowerPoint.Master || shape.Parent is PowerPoint.CustomLayout) &&
                    shape.Type == Office.MsoShapeType.msoPlaceholder &&
                    !(new[]{
                        PowerPoint.PpPlaceholderType.ppPlaceholderTitle,
                        PowerPoint.PpPlaceholderType.ppPlaceholderSubtitle,
                        PowerPoint.PpPlaceholderType.ppPlaceholderHeader,
                        PowerPoint.PpPlaceholderType.ppPlaceholderFooter,
                        PowerPoint.PpPlaceholderType.ppPlaceholderDate,
                        PowerPoint.PpPlaceholderType.ppPlaceholderSlideNumber
                            }).Contains(shape.PlaceholderFormat.Type))
                {
                    return;
                }
            }
            catch { /* ignore */ }

            // — recurse into grouped shapes —
            if (shape.Type == Office.MsoShapeType.msoGroup)
            {
                foreach (PowerPoint.Shape child in shape.GroupItems)
                    ProcessShape(child);
                return;
            }

            // — tables: freeze each cell’s fill —
            if (shape.HasTable == Office.MsoTriState.msoTrue)
            {
                var tbl = shape.Table;
                for (int r = 1; r <= tbl.Rows.Count; r++)
                    for (int c = 1; c <= tbl.Columns.Count; c++)
                        if (tbl.Cell(r, c)?.Shape?.Fill != null)
                            FreezeFillFormat(tbl.Cell(r, c).Shape.Fill);
                return;
            }

            // — charts: only freeze series & points (no background touches) —
            if (shape.HasChart == Office.MsoTriState.msoTrue)
            {
                dynamic seriesCollection = shape.Chart.SeriesCollection();
                for (int i = 1; i <= seriesCollection.Count; i++)
                {
                    dynamic series = seriesCollection.Item(i);
                    try { FreezeChartFill(series.Format.Fill); } catch { }
                    try { FreezeChartLine(series.Format.Line); } catch { }

                    dynamic points = series.Points();
                    for (int p = 1; p <= points.Count; p++)
                    {
                        dynamic pt = points.Item(p);
                        try { FreezeChartFill(pt.Format.Fill); } catch { }
                        try { FreezeChartLine(pt.Format.Line); } catch { }
                    }
                }
                return;
            }

            // — everything else: normal shapes & text —
            if (shape.Fill != null)
                FreezeFillFormat(shape.Fill);
            if (shape.Line != null)
                FreezeLineFormat(shape.Line);
            if (shape.HasTextFrame == Office.MsoTriState.msoTrue &&
                shape.TextFrame?.HasText == Office.MsoTriState.msoTrue)
            {
                FreezeTextFrame(shape.TextFrame);
            }
        }

        // for PowerPoint.Chart shapes
        private static void FreezeChartFill(dynamic fill)
        {
            if (fill == null) return;
            try
            {
                int rgb = fill.ForeColor.RGB;
                fill.ForeColor.RGB = rgb;
            }
            catch { }
        }

        private static void FreezeChartLine(dynamic line)
        {
            if (line == null) return;

            // **NEW**: if no outline present, skip
            try
            {
                if (line.Visible != Office.MsoTriState.msoTrue)
                    return;

                int rgb = line.ForeColor.RGB;
                line.ForeColor.RGB = rgb;
            }
            catch
            {
                // ignore anything else
            }
        }

        #region ColorFormat Freezing
        private static void FreezeColor(PowerPoint.ColorFormat cf)
        {
            if (cf == null) return;
            try
            {
                if (cf.Type == Office.MsoColorType.msoColorTypeScheme || cf.Type == Office.MsoColorType.msoColorTypeMixed)
                {
                    int originalRGB = cf.RGB;
                    cf.RGB = originalRGB;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing PowerPoint.ColorFormat (Type: {cf?.Type}): {ex.Message}"); }
        }

        private static void FreezeColor(Office.ColorFormat cf)
        {
            if (cf == null) return;
            try
            {
                if (cf.Type == Office.MsoColorType.msoColorTypeScheme || cf.Type == Office.MsoColorType.msoColorTypeMixed)
                {
                    int originalRGB = cf.RGB;
                    cf.RGB = originalRGB;
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing Office.ColorFormat (Type: {cf?.Type}): {ex.Message}"); }
        }
        #endregion

        #region FillFormat Freezing
        private static void FreezeFillFormat(PowerPoint.FillFormat fill)
        {
            if (fill == null || fill.Visible == Office.MsoTriState.msoFalse) return;
            try
            {
                switch (fill.Type)
                {
                    case Office.MsoFillType.msoFillSolid:
                        float transparency = fill.Transparency;
                        FreezeColor(fill.ForeColor);
                        fill.Transparency = transparency;
                        break;
                    case Office.MsoFillType.msoFillGradient:
                        if (fill.GradientStops != null)
                        {
                            foreach (Office.GradientStop gs in fill.GradientStops)
                            {
                                float stopTransparency = gs.Transparency;
                                FreezeColor(gs.Color);
                                gs.Transparency = stopTransparency;
                            }
                        }
                        break;
                    case Office.MsoFillType.msoFillPatterned:
                        FreezeColor(fill.ForeColor);
                        FreezeColor(fill.BackColor);
                        break;
                    case Office.MsoFillType.msoFillBackground:
                    case Office.MsoFillType.msoFillPicture:
                    case Office.MsoFillType.msoFillTextured:
                    case Office.MsoFillType.msoFillMixed:
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error freezing PowerPoint.FillFormat (Type: {fill?.Type}): {ex.Message}");
            }
        }
        #endregion

        #region LineFormat Freezing
        private static void FreezeLineFormat(PowerPoint.LineFormat line)
        {
            if (line == null || line.Visible == Office.MsoTriState.msoFalse) return;
            try
            {
                float transparency = line.Transparency;
                FreezeColor(line.ForeColor);
                line.Transparency = transparency;

                if (line.Style != Office.MsoLineStyle.msoLineSingle && line.Style != Office.MsoLineStyle.msoLineStyleMixed)
                {
                    try { FreezeColor(line.BackColor); }
                    catch (Exception bcex) { Debug.WriteLine($"Note: Could not freeze PowerPoint.Line.BackColor: {bcex.Message}"); }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing PowerPoint.LineFormat: {ex.Message}"); }
        }

        #endregion

        #region TextFrame/Font Freezing
        private static void FreezeTextFrame(PowerPoint.TextFrame tf)
        {
            if (tf == null || tf.HasText == Office.MsoTriState.msoFalse) return;
            try
            {
                if (tf.TextRange?.Font != null)
                {
                    FreezeColor(tf.TextRange.Font.Color);
                }

                if (tf.TextRange?.Runs() != null)
                {
                    for (int i = 1; i <= tf.TextRange.Runs().Count; i++)
                    {
                        PowerPoint.TextRange run = tf.TextRange.Runs(i);
                        if (run?.Font != null)
                        {
                            FreezeColor(run.Font.Color);
                        }
                    }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing TextFrame: {ex.Message}"); }
        }
        #endregion
    }
}