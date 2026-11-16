using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_PowerPoint.Text
{
    public static class TextUtility
    {
        public static void AltFind(PowerPoint.Application app)
        {
            string term = Interaction.InputBox("Enter text to find in Alternative Text:", "Alt‑Find");
            if (string.IsNullOrWhiteSpace(term)) return;
            var pres = app.ActivePresentation;
            if (pres == null || pres.Slides == null) return;
            var hits = new List<Tuple<PowerPoint.Shape, int>>();

            for (int s = 1; s <= pres.Slides.Count; s++)
            {
                PowerPoint.Slide slide = null;
                PowerPoint.Shapes shapesInSlide = null;
                try
                {
                    slide = pres.Slides[s];
                    shapesInSlide = slide.Shapes;
                    foreach (PowerPoint.Shape shp in shapesInSlide)
                    {
                        try
                        {
                            if (!string.IsNullOrEmpty(shp.AlternativeText) && shp.AlternativeText.IndexOf(term, StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                hits.Add(new Tuple<PowerPoint.Shape, int>(shp, s));
                            }
                        }
                        catch (Exception exShapeAltText)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error reading alt text for shape {shp?.Name} on slide {s}: {exShapeAltText.Message}");
                        }
                        // Intentionally not releasing shp here as it's added to 'hits' list
                        // It will be released after use or when 'hits' is cleared.
                    }
                }
                catch (Exception exSlide)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing slide {s} for AltFind: {exSlide.Message}");
                }
                finally
                {
                    if (shapesInSlide != null && System.Runtime.InteropServices.Marshal.IsComObject(shapesInSlide))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(shapesInSlide);
                    if (slide != null && System.Runtime.InteropServices.Marshal.IsComObject(slide))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(slide);
                }
            }

            if (hits.Count == 0)
            {
                MessageBox.Show("No instances found.", "Alt‑Find");
                return;
            }

            int idx = 0;
            while (idx < hits.Count)
            {
                var entry = hits[idx];
                PowerPoint.Shape hitShape = null;
                try
                {
                    hitShape = entry.Item1; // Get the shape from the tuple
                    app.ActiveWindow.View.GotoSlide(entry.Item2);
                    hitShape.Select(); // Select the already obtained shape object
                    var result = MessageBox.Show($"Found on slide {entry.Item2}.\nShape Name: {hitShape.Name}\n\nContinue to next instance?", "Alt‑Find", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.No) break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error during AltFind navigation: {ex.Message}");
                    MessageBox.Show($"Error processing find result: {ex.Message}", "Alt-Find Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    break;
                }
                finally
                {
                    // Release the shape COM object from the 'hits' list after it's been processed in this iteration
                    if (hitShape != null && System.Runtime.InteropServices.Marshal.IsComObject(hitShape))
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(hitShape);
                }
                idx++;
            }
            // Clean up any remaining shapes in 'hits' if the loop was exited early
            for (int i = idx; i < hits.Count; i++)
            {
                if (hits[i].Item1 != null && System.Runtime.InteropServices.Marshal.IsComObject(hits[i].Item1))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(hits[i].Item1);
            }
        }

        public static void ChangeParagraphSpacing(PowerPoint.Application app, float deltaPts)
        {
            PowerPoint.Selection sel = null;
            PowerPoint.TextRange textRange = null;
            PowerPoint.ShapeRange shapeRange = null;
            try
            {
                sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    textRange = sel.TextRange;
                    if (textRange != null) AdjustRange(textRange, deltaPts);
                }
                else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    shapeRange = sel.ShapeRange;
                    if (shapeRange != null)
                    {
                        foreach (PowerPoint.Shape shp in shapeRange)
                        {
                            PowerPoint.TextFrame shpTextFrame = null;
                            PowerPoint.TextRange shpTextRange = null;
                            try
                            {
                                if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                                {
                                    shpTextFrame = shp.TextFrame;
                                    if (shpTextFrame.HasText == Office.MsoTriState.msoTrue)
                                    {
                                        shpTextRange = shpTextFrame.TextRange;
                                        AdjustRange(shpTextRange, deltaPts);
                                    }
                                }
                            }
                            finally
                            {
                                if (shpTextRange != null && System.Runtime.InteropServices.Marshal.IsComObject(shpTextRange)) System.Runtime.InteropServices.Marshal.ReleaseComObject(shpTextRange);
                                if (shpTextFrame != null && System.Runtime.InteropServices.Marshal.IsComObject(shpTextFrame)) System.Runtime.InteropServices.Marshal.ReleaseComObject(shpTextFrame);
                                if (shp != null && System.Runtime.InteropServices.Marshal.IsComObject(shp)) System.Runtime.InteropServices.Marshal.ReleaseComObject(shp);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ChangeParagraphSpacing: {ex.Message}");
            }
            finally
            {
                if (textRange != null && System.Runtime.InteropServices.Marshal.IsComObject(textRange)) System.Runtime.InteropServices.Marshal.ReleaseComObject(textRange);
                if (shapeRange != null && System.Runtime.InteropServices.Marshal.IsComObject(shapeRange)) System.Runtime.InteropServices.Marshal.ReleaseComObject(shapeRange);
                if (sel != null && System.Runtime.InteropServices.Marshal.IsComObject(sel)) System.Runtime.InteropServices.Marshal.ReleaseComObject(sel);
            }
        }

        private static void AdjustRange(PowerPoint.TextRange rng, float deltaPts)
        {
            PowerPoint.ParagraphFormat pf = null;
            try
            {
                pf = rng.ParagraphFormat;
                float before = pf.SpaceBefore + deltaPts;
                pf.SpaceBefore = before > 0 ? before : 0;
                float after = pf.SpaceAfter + deltaPts;
                pf.SpaceAfter = after > 0 ? after : 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adjusting paragraph spacing: {ex.Message}");
            }
            finally
            {
                if (pf != null && System.Runtime.InteropServices.Marshal.IsComObject(pf)) System.Runtime.InteropServices.Marshal.ReleaseComObject(pf);
            }
        }

public static void RemoveMargin(PowerPoint.ShapeRange shapeRange)
{
    if (shapeRange == null) return;

    try
    {
        // Using foreach is often cleaner and less error-prone for COM collections
        foreach (PowerPoint.Shape shp in shapeRange)
        {
            PowerPoint.TextFrame2 tf2 = null;
            try
            {
                if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                {
                    tf2 = shp.TextFrame2; 

                    if (tf2 != null)
                    {
                        tf2.MarginLeft = 0f;
                        tf2.MarginRight = 0f;
                        tf2.MarginTop = 0f;
                        tf2.MarginBottom = 0f;
                    }
                }
            }
            catch (Exception exShape)
            {
                System.Diagnostics.Debug.WriteLine($"RemoveMargin Error on shape {shp?.Name}: {exShape.Message}");
            }
            finally
            {
                // It's important to release the COM objects you've used in the loop
                if (tf2 != null && System.Runtime.InteropServices.Marshal.IsComObject(tf2))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(tf2);
                if (shp != null && System.Runtime.InteropServices.Marshal.IsComObject(shp))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(shp);
            }
        }
    }
    catch (Exception ex) 
    { 
        System.Diagnostics.Debug.WriteLine($"RemoveMargin Main Error: {ex.Message}"); 
    }
    finally
    {
        // Release the ShapeRange itself after the loop is done
        if (shapeRange != null && System.Runtime.InteropServices.Marshal.IsComObject(shapeRange))
            System.Runtime.InteropServices.Marshal.ReleaseComObject(shapeRange);
    }
}
    }
}