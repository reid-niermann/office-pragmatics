using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace Office_Pragmatics_Word
{
    public static class ColorUtility
    {
        public static void ColorFreeze(Word.Application app)
        {
            if (app == null || app.ActiveDocument == null)
            {
                MessageBox.Show("No active document found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            var doc = app.ActiveDocument;
            var sel = app.Selection;

            var shapesToProcess = new List<Word.Shape>();
            var inlineChartsToProcess = new List<dynamic>(); // Chart objects from InlineShapes

            try
            {
                // Prefer selection if it contains shapes
                if (sel != null && SafeHasShapeSelection(sel))
                {
                    try
                    {
                        foreach (Word.Shape s in sel.ShapeRange)
                            shapesToProcess.Add(s);

                        Debug.WriteLine($"Processing {shapesToProcess.Count} selected shapes.");
                    }
                    catch { /* ignore */ }
                }
                // Else, selection might have inline charts
                else if (sel != null && SafeHasInlineShapes(sel))
                {
                    try
                    {
                        foreach (Word.InlineShape ils in sel.InlineShapes)
                        {
                            if (SafeHasChart(ils))
                            {
                                try { inlineChartsToProcess.Add(ils.Chart); } catch { }
                            }
                        }
                        Debug.WriteLine($"Processing {inlineChartsToProcess.Count} inline charts from selection.");
                    }
                    catch { }
                }
                // Else, process whole document (shapes + inline charts)
                else
                {
                    shapesToProcess.AddRange(EnumerateAllShapes(doc));
                    inlineChartsToProcess.AddRange(EnumerateAllInlineCharts(doc));
                    Debug.WriteLine($"Processing all shapes ({shapesToProcess.Count}) and inline charts ({inlineChartsToProcess.Count}).");
                }

                // Process shapes
                foreach (var shp in shapesToProcess)
                    ProcessShape(shp);

                // Process inline charts
                foreach (var ch in inlineChartsToProcess)
                    ProcessChart(ch);

                // Tables (selection-scoped if available; else whole document)
                if (sel?.Range?.Tables != null && sel.Range.Tables.Count > 0)
                    FreezeTables(sel.Range.Tables);
                else
                    FreezeTables(doc.Tables);

                // Optional (scoped) text font color freeze in selection (outside shapes)
                if (sel?.Range?.Characters != null && sel.Range.Characters.Count > 0)
                    FreezeWordRangeFontColors(sel.Range);

                MessageBox.Show(
                    "Attempted to freeze theme-based colors to explicit RGB (shapes, charts, shape text, tables).",
                    "Freeze Colors",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"An error occurred: {ex.Message}\n\n{ex.StackTrace}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Debug.WriteLine($"ColorFreeze Error: {ex}");
            }
        }

        private static bool SafeHasShapeSelection(Word.Selection sel)
        {
            try { return sel.ShapeRange != null && sel.ShapeRange.Count > 0; } catch { return false; }
        }

        private static bool SafeHasInlineShapes(Word.Selection sel)
        {
            try { return sel.InlineShapes != null && sel.InlineShapes.Count > 0; } catch { return false; }
        }

        private static bool SafeHasChart(Word.InlineShape ils)
        {
            try { return ils.Chart != null; } catch { return false; }
        }

        private static IEnumerable<Word.Shape> EnumerateAllShapes(Word.Document doc)
        {
            var list = new List<Word.Shape>();

            // Main document shapes
            try
            {
                if (doc.Shapes != null && doc.Shapes.Count > 0)
                    foreach (Word.Shape s in doc.Shapes) list.Add(s);
            }
            catch { }

            // Header & Footer shapes
            try
            {
                foreach (Word.Section sec in doc.Sections)
                {
                    foreach (Word.HeaderFooter hf in sec.Headers)
                    {
                        if (hf?.Shapes != null && hf.Shapes.Count > 0)
                            foreach (Word.Shape s in hf.Shapes) list.Add(s);
                    }
                    foreach (Word.HeaderFooter hf in sec.Footers)
                    {
                        if (hf?.Shapes != null && hf.Shapes.Count > 0)
                            foreach (Word.Shape s in hf.Shapes) list.Add(s);
                    }
                }
            }
            catch { }

            return list;
        }

        private static IEnumerable<dynamic> EnumerateAllInlineCharts(Word.Document doc)
        {
            var charts = new List<dynamic>();

            try
            {
                foreach (Word.InlineShape ils in doc.InlineShapes)
                {
                    if (SafeHasChart(ils))
                    {
                        try { charts.Add(ils.Chart); } catch { }
                    }
                }

                foreach (Word.Section sec in doc.Sections)
                {
                    foreach (Word.HeaderFooter hf in sec.Headers)
                    {
                        try
                        {
                            foreach (Word.InlineShape ils in hf.Range.InlineShapes)
                                if (SafeHasChart(ils)) { try { charts.Add(ils.Chart); } catch { } }
                        }
                        catch { }
                    }
                    foreach (Word.HeaderFooter hf in sec.Footers)
                    {
                        try
                        {
                            foreach (Word.InlineShape ils in hf.Range.InlineShapes)
                                if (SafeHasChart(ils)) { try { charts.Add(ils.Chart); } catch { } }
                        }
                        catch { }
                    }
                }
            }
            catch { }

            return charts;
        }

        private static void ProcessShape(Word.Shape shape)
        {
            if (shape == null) return;

            // Recurse into grouped shapes
            try
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    foreach (Word.Shape child in shape.GroupItems)
                        ProcessShape(child);
                    return;
                }
            }
            catch { }

            // Charts on shapes
            try
            {
                if (shape.Chart != null)
                {
                    ProcessChart(shape.Chart);
                    return;
                }
            }
            catch { /* not a chart */ }

            // Normal shapes: fills, lines, text
            try
            {
                if (shape.Fill != null && shape.Fill.Visible == Office.MsoTriState.msoTrue)
                    FreezeFillFormat(shape.Fill); // Word.FillFormat overload
            }
            catch { }

            try
            {
                if (shape.Line != null && shape.Line.Visible == Office.MsoTriState.msoTrue)
                    FreezeLineFormat(shape.Line); // Word.LineFormat overload
            }
            catch { }

            try
            {
                if (shape.TextFrame != null)
                    FreezeTextFrame(shape.TextFrame);
            }
            catch { }
        }

        private static void ProcessChart(dynamic chartOrSeriesContainer)
        {
            try
            {
                dynamic seriesCollection = chartOrSeriesContainer.SeriesCollection();
                int count = seriesCollection.Count;
                for (int i = 1; i <= count; i++)
                {
                    dynamic series = seriesCollection.Item(i);
                    try { FreezeChartFill(series.Format.Fill); } catch { }
                    try { FreezeChartLine(series.Format.Line); } catch { }

                    try
                    {
                        dynamic points = series.Points();
                        int pc = points.Count;
                        for (int p = 1; p <= pc; p++)
                        {
                            dynamic pt = points.Item(p);
                            try { FreezeChartFill(pt.Format.Fill); } catch { }
                            try { FreezeChartLine(pt.Format.Line); } catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        // ---------- Chart helpers (dynamic) ----------
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
            try
            {
                if (line.Visible != Office.MsoTriState.msoTrue) return;
                int rgb = line.ForeColor.RGB;
                line.ForeColor.RGB = rgb;
            }
            catch { }
        }

        // ---------- ColorFormat (both Office and Word overloads) ----------
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

        private static void FreezeColor(Word.ColorFormat cf)
        {
            if (cf == null) return;
            try
            {
                // Word's ColorFormat also exposes RGB; self-assign is sufficient and safest across types.
                int originalRGB = cf.RGB;
                cf.RGB = originalRGB;
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing Word.ColorFormat: {ex.Message}"); }
        }

        // ---------- FillFormat ----------
        private static void FreezeFillFormat(Word.FillFormat fill)
        {
            if (fill == null || fill.Visible == Office.MsoTriState.msoFalse) return;
            try
            {
                switch (fill.Type)
                {
                    case Office.MsoFillType.msoFillSolid:
                        float transparency = fill.Transparency;
                        FreezeColor(fill.ForeColor);   // Word.ColorFormat
                        fill.Transparency = transparency;
                        break;

                    case Office.MsoFillType.msoFillGradient:
                        try
                        {
                            var stops = fill.GradientStops as Office.GradientStops;
                            if (stops != null)
                            {
                                foreach (Office.GradientStop gs in stops)
                                {
                                    float stopTransparency = gs.Transparency;
                                    FreezeColor(gs.Color); // Office.ColorFormat
                                    gs.Transparency = stopTransparency;
                                }
                            }
                        }
                        catch { }
                        break;

                    case Office.MsoFillType.msoFillPatterned:
                        FreezeColor(fill.ForeColor);   // Word.ColorFormat
                        FreezeColor(fill.BackColor);   // Word.ColorFormat
                        break;

                    case Office.MsoFillType.msoFillBackground:
                    case Office.MsoFillType.msoFillPicture:
                    case Office.MsoFillType.msoFillTextured:
                    case Office.MsoFillType.msoFillMixed:
                        // leave as-is
                        break;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error freezing Word.FillFormat (Type: {fill?.Type}): {ex.Message}");
            }
        }

        // (Optional) Office.FillFormat overload - not typically used by Word.Shape, but kept for parity
        private static void FreezeFillFormat(Office.FillFormat fill)
        {
            if (fill == null || fill.Visible == Office.MsoTriState.msoFalse) return;
            try
            {
                switch (fill.Type)
                {
                    case Office.MsoFillType.msoFillSolid:
                        float transparency = fill.Transparency;
                        FreezeColor(fill.ForeColor);   // Office.ColorFormat
                        fill.Transparency = transparency;
                        break;

                    case Office.MsoFillType.msoFillGradient:
                        if (fill.GradientStops != null)
                        {
                            foreach (Office.GradientStop gs in fill.GradientStops)
                            {
                                float stopTransparency = gs.Transparency;
                                FreezeColor(gs.Color);  // Office.ColorFormat
                                gs.Transparency = stopTransparency;
                            }
                        }
                        break;

                    case Office.MsoFillType.msoFillPatterned:
                        FreezeColor(fill.ForeColor);   // Office.ColorFormat
                        FreezeColor(fill.BackColor);   // Office.ColorFormat
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
                Debug.WriteLine($"Error freezing Office.FillFormat (Type: {fill?.Type}): {ex.Message}");
            }
        }

        // ---------- LineFormat ----------
        private static void FreezeLineFormat(Word.LineFormat line)
        {
            if (line == null || line.Visible == Office.MsoTriState.msoFalse) return;
            try
            {
                float transparency = line.Transparency;
                FreezeColor(line.ForeColor); // Word.ColorFormat
                line.Transparency = transparency;

                if (line.Style != Office.MsoLineStyle.msoLineSingle && line.Style != Office.MsoLineStyle.msoLineStyleMixed)
                {
                    try { FreezeColor(line.BackColor); } // Word.ColorFormat
                    catch (Exception bcex) { Debug.WriteLine($"Note: Could not freeze Word.LineFormat.BackColor: {bcex.Message}"); }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing Word.LineFormat: {ex.Message}"); }
        }

        // (Optional) Office.LineFormat overload - not typically used by Word.Shape, but kept for parity
        private static void FreezeLineFormat(Office.LineFormat line)
        {
            if (line == null || line.Visible == Office.MsoTriState.msoFalse) return;
            try
            {
                float transparency = line.Transparency;
                FreezeColor(line.ForeColor); // Office.ColorFormat
                line.Transparency = transparency;

                if (line.Style != Office.MsoLineStyle.msoLineSingle && line.Style != Office.MsoLineStyle.msoLineStyleMixed)
                {
                    try { FreezeColor(line.BackColor); } // Office.ColorFormat
                    catch (Exception bcex) { Debug.WriteLine($"Note: Could not freeze Office.LineFormat.BackColor: {bcex.Message}"); }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing Office.LineFormat: {ex.Message}"); }
        }

        // ---------- TextFrame ----------
        private static void FreezeTextFrame(Word.TextFrame tf)
        {
            if (tf == null) return;
            try
            {
                // Word.TextFrame.HasText is effectively an int (-1 true, 0 false)
                bool hasText = false;
                try { hasText = Convert.ToInt32(tf.HasText) != 0; } catch { }
                if (!hasText) return;

                Word.Range r = tf.TextRange;
                if (r == null) return;

                try
                {
                    var words = r.Words;
                    int count = words.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Word.Range w = null;
                        try { w = words[i]; } catch { }
                        if (w == null) continue;

                        try
                        {
                            var c = w.Font.Color;
                            w.Font.Color = c; // force explicit RGB if theme-linked
                        }
                        catch { }
                    }
                }
                catch
                {
                    try
                    {
                        var c = r.Font.Color;
                        r.Font.Color = c;
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing TextFrame: {ex.Message}"); }
        }

        // ---------- Tables ----------
        private static void FreezeTables(Word.Tables tables)
        {
            if (tables == null || tables.Count == 0) return;

            try
            {
                foreach (Word.Table t in tables)
                {
                    // Borders
                    try
                    {
                        foreach (Word.Border b in t.Borders)
                        {
                            try
                            {
                                if (b.LineStyle != Word.WdLineStyle.wdLineStyleNone)
                                {
                                    var bc = b.Color;
                                    b.Color = bc;
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }

                    // Shading + inner text colors
                    try
                    {
                        int rows = t.Rows.Count;
                        int cols = t.Columns.Count;
                        for (int r = 1; r <= rows; r++)
                        {
                            for (int c = 1; c <= cols; c++)
                            {
                                Word.Cell cell = null;
                                try { cell = t.Cell(r, c); } catch { }
                                if (cell == null) continue;

                                try
                                {
                                    var sh = cell.Shading;
                                    var bg = sh.BackgroundPatternColor;
                                    var fg = sh.ForegroundPatternColor;
                                    sh.BackgroundPatternColor = bg;
                                    sh.ForegroundPatternColor = fg;
                                }
                                catch { }

                                try
                                {
                                    FreezeWordRangeFontColors(cell.Range);
                                }
                                catch { }
                            }
                        }
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing tables: {ex.Message}"); }
        }

        // ---------- Word text color freeze for a Range ----------
        private static void FreezeWordRangeFontColors(Word.Range rng)
        {
            if (rng == null) return;
            try
            {
                var words = rng.Words;
                int count = words.Count;
                for (int i = 1; i <= count; i++)
                {
                    Word.Range w = null;
                    try { w = words[i]; } catch { }
                    if (w == null) continue;

                    try
                    {
                        var c = w.Font.Color;
                        w.Font.Color = c;
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Debug.WriteLine($"Error freezing range font colors: {ex.Message}"); }
        }
    }
}