// Redaction.cs
using System;
using System.Linq;
using System.Collections;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_PowerPoint.Slide
{
    public static class SlideRedaction
    {
        private static void ShowRedactionWarning()
        {
            MessageBox.Show(
                "IMPORTANT PRIVACY NOTICE:\n\n" +
                "This redaction only visually obscures content with a black overlay. " +
                "If this file is exported to PDF, the underlying text and images will still remain in the file structure and can be recovered using software tools.\n\n" +
                "For confidential information requiring true redaction:\n" +
                "• Use Adobe Acrobat or similar professional redaction software\n" +
                "• Export to PDF and apply proper redaction tools\n\n" +
                "More robust redaction functionality is in our product development pipeline.",
                "Redaction Privacy Warning",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public static void Redact(PowerPoint.Application app)
        {
            var pres = app.ActivePresentation;
            var sel = app.ActiveWindow.Selection;

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
            {
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return;
                var tr2 = sel.TextRange2;
                float left = tr2.BoundLeft, top = tr2.BoundTop, width = tr2.BoundWidth, height = tr2.BoundHeight;
                string key = $"{left}_{top}_{width}_{height}".Replace(".", "_").Replace(",", "_"); // Sanitize key
                var existing = ((IEnumerable)slide.Shapes).Cast<PowerPoint.Shape>().FirstOrDefault(s => s.Tags["TextRedact"] == key);

                if (existing != null) existing.Delete();
                else
                {
                    var overlay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left, top, width, height);
                    overlay.Fill.Solid();
                    overlay.Fill.ForeColor.RGB = System.Drawing.Color.Black.ToArgb() & 0xFFFFFF;
                    overlay.Line.Visible = Office.MsoTriState.msoFalse;
                    overlay.Tags.Add("TextRedact", key);
                    overlay.Name = "TextRedaction_" + key;
                    overlay.ZOrder(Office.MsoZOrderCmd.msoBringToFront);

                    // Show warning after applying redaction
                    ShowRedactionWarning();
                }
                return;
            }

            if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
                if (slide == null) return;

                bool redactionApplied = false;
                foreach (PowerPoint.Shape shp in sel.ShapeRange)
                {
                    bool isTextShape = shp.HasTextFrame == Office.MsoTriState.msoTrue && shp.TextFrame.HasText == Office.MsoTriState.msoTrue;
                    bool isPicture = shp.Type == Office.MsoShapeType.msoPicture;
                    bool isTable = shp.HasTable == Office.MsoTriState.msoTrue;
                    bool isChart = shp.HasChart == Office.MsoTriState.msoTrue;
                    if (!isTextShape && !isPicture && !isTable && !isChart && shp.Type != Office.MsoShapeType.msoAutoShape) continue;

                    string key = shp.Id.ToString();
                    var existing = ((IEnumerable)slide.Shapes).Cast<PowerPoint.Shape>().FirstOrDefault(s => s.Tags["RedactedOverlay"] == key);

                    if (existing != null) existing.Delete();
                    else
                    {
                        var overlay = slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, shp.Left, shp.Top, shp.Width, shp.Height);
                        overlay.Fill.Solid();
                        overlay.Fill.ForeColor.RGB = System.Drawing.Color.Black.ToArgb() & 0xFFFFFF;
                        overlay.Line.Visible = Office.MsoTriState.msoFalse;
                        overlay.Tags.Add("RedactedOverlay", key);
                        overlay.Name = "RedactionOverlay_" + key;
                        overlay.ZOrder(Office.MsoZOrderCmd.msoBringToFront);
                        redactionApplied = true;
                    }
                }

                // Show warning if any redaction was applied
                if (redactionApplied)
                {
                    ShowRedactionWarning();
                }
                return;
            }

            var result = MessageBox.Show("No specific text or shapes selected for redaction.\nDo you want to remove ALL existing redactions from this presentation?",
                                         "Clear All Redactions?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                foreach (PowerPoint.Slide sl in pres.Slides)
                {
                    for (int i = sl.Shapes.Count; i >= 1; i--)
                    {
                        PowerPoint.Shape s = null;
                        try
                        {
                            s = sl.Shapes[i];
                            if (s.Tags["RedactedOverlay"] != "" || s.Tags["TextRedact"] != "")
                            {
                                s.Delete();
                            }
                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error processing shape for redaction removal: {s?.Name} - {ex.Message}");
                        }
                        finally
                        {
                            if (s != null && System.Runtime.InteropServices.Marshal.IsComObject(s))
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(s);
                        }
                    }
                    if (System.Runtime.InteropServices.Marshal.IsComObject(sl))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(sl);
                }
                MessageBox.Show("All redactions removed.", "Redaction Cleared", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}