using Microsoft.VisualBasic;
using OfficePragmatics_PowerPoint.Settings;
using System;
using System.Collections;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_PowerPoint.Slide
{
    public static class SlideUtility
    {
        public static void InsertTBU(PowerPoint.Application app)
        {
            try
            {
                var pres = app.ActivePresentation;
                float w = pres.PageSetup.SlideWidth;

                // ─── shape ───
                var slide = app.ActiveWindow.View.Slide;
                var box = slide.Shapes.AddShape(
                                Office.MsoAutoShapeType.msoShapeRectangle,
                                w - 110, 10, 100, 50);

                box.Name = "TBU Box";
                box.Fill.ForeColor.RGB = 9428735;
                box.Line.Weight = 1;
                box.Line.ForeColor.RGB = 4994804;
                box.Line.DashStyle = Office.MsoLineDashStyle.msoLineDash;

                /* turn all visual effects off */
                box.Shadow.Visible = Office.MsoTriState.msoFalse;
                box.Glow.Radius = 0;
                box.SoftEdge.Radius = 0;
                box.Reflection.Type = Office.MsoReflectionType.msoReflectionTypeNone;
                box.ThreeD.Visible = Office.MsoTriState.msoFalse;

                // ─── text ───
                var tf = box.TextFrame;
                tf.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeNone;
                tf.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;

                const float marginPts = 10f;
                tf.MarginLeft = marginPts;
                tf.MarginRight = marginPts;
                tf.MarginTop = marginPts;
                tf.MarginBottom = marginPts;

                var tr = tf.TextRange;
                tr.Text = "TBU";
                tr.Font.Name = "Arial";
                tr.Font.Size = 10;
                tr.Font.Color.RGB = 4994804;
                tr.Font.Bold = Office.MsoTriState.msoFalse;
                tr.Font.Shadow = Office.MsoTriState.msoFalse;
                tr.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignCenter;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InsertTBU Error: {ex.Message}");
            }
        }

        public static void InsertFootnote(PowerPoint.Application app)
        {
            try
            {
                var settings = SettingsManager.Current;

                var pres = app.ActivePresentation;
                int idx = app.ActiveWindow.View.Slide.SlideIndex;
                var slide = pres.Slides[idx];

                var fn = slide.Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    settings.FootnoteLeft,
                    settings.FootnoteTop,
                    settings.FootnoteWidth,
                    50);

                fn.Name = "Footnote Box";

                fn.Shadow.Visible = Office.MsoTriState.msoFalse;
                fn.Fill.Visible = Office.MsoTriState.msoFalse;
                fn.Line.Visible = Office.MsoTriState.msoFalse;

                var tr = fn.TextFrame.TextRange;

                // CHANGED: "Note: " (space) instead of "Note:\t" (tab)
                tr.Text = "Note: Placeholder.\rPlaceholder\rSource: Placeholder";

                tr.Font.Name = settings.FootnoteFontName;
                tr.Font.Size = settings.FootnoteFontSize;
                tr.Font.Color.RGB = 0;
                tr.Font.Shadow = Office.MsoTriState.msoFalse;

                fn.TextFrame2.MarginLeft = fn.TextFrame2.MarginRight = fn.TextFrame2.MarginTop = fn.TextFrame2.MarginBottom = 0;
                fn.TextFrame2.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorTop;
                fn.TextFrame2.TextRange.ParagraphFormat.Alignment = Office.MsoParagraphAlignment.msoAlignLeft;

                if (fn.TextFrame.HasText == Office.MsoTriState.msoTrue && fn.TextFrame.Ruler != null && fn.TextFrame.Ruler.Levels.Count >= 1)
                {
                    // Sets the paragraph indent (where the text wraps) to 12
                    fn.TextFrame.Ruler.Levels[1].LeftMargin = 12;

                    // Sets the bullet position to 0 (hanging indent effect)
                    fn.TextFrame.Ruler.Levels[1].FirstMargin = 0;

                    fn.TextFrame.Ruler.TabStops.Add(PowerPoint.PpTabStopType.ppTabStopLeft, 12);
                }

                var p2 = tr.Paragraphs(2, 1);
                p2.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletNumbered;
                p2.ParagraphFormat.Bullet.Style = (PowerPoint.PpNumberedBulletStyle)settings.FootnoteBulletStyle;

                var p3 = tr.Paragraphs(3, 1);
                p3.Font.Italic = Office.MsoTriState.msoTrue;

                fn.Left = settings.FootnoteLeft;
                fn.Top = settings.FootnoteTop;
                fn.Width = settings.FootnoteWidth;

                fn.TextFrame.AutoSize = PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"InsertFootnote Error: {ex.Message}");
            }
        }
    }
}