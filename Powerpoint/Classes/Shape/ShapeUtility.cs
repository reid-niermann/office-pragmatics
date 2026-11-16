// ShapeUtility.cs
using Microsoft.VisualBasic;
using OfficePragmatics_PowerPoint.Settings;   // For SettingsForm.CurrentSettings
using OfficePragmatics_PowerPoint.Utilities;  // For Utility
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_PowerPoint.Shape
{
    public static partial class ShapeUtility
    {
        public static void SnapTogether(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count != 2) { MessageBox.Show("Invalid selection: choose two shapes"); return; }
                var s1 = sel.ShapeRange[1]; var s2 = sel.ShapeRange[2];
                float line1 = 0, line2 = 0;
                bool alignToBorder = SettingsForm.CurrentSettings.AlignToBorder;
                if (alignToBorder)
                {
                    if (s1.Type != Office.MsoShapeType.msoTable && s1.Line.Visible == Office.MsoTriState.msoTrue) line1 = Math.Max(0, s1.Line.Weight);
                    if (s2.Type != Office.MsoShapeType.msoTable && s2.Line.Visible == Office.MsoTriState.msoTrue) line2 = Math.Max(0, s2.Line.Weight);
                }
                float w1 = s1.Width + line1, h1 = s1.Height + line1, l1 = s1.Left - line1 / 2, t1 = s1.Top - line1 / 2, r1 = l1 + w1, b1 = t1 + h1;
                float w2 = s2.Width + line2, h2 = s2.Height + line2, l2 = s2.Left - line2 / 2, t2 = s2.Top - line2 / 2, r2 = l2 + w2, b2 = t2 + h2;

                if (r2 > l1 && r1 > l2 && t1 > b2) s1.Top = b2 + line1 / 2; // s2 above s1
                else if (l1 > r2 && b2 > t1 && b1 > t2) s1.Left = r2 + line1 / 2; // s2 left of s1
                else if (r2 > l1 && r1 > l2 && t2 > b1) s1.Top = t2 - h1 + line1 / 2; // s2 below s1
                else if (l2 > r1 && b2 > t1 && b1 > t2) s1.Left = l2 - w1 + line1 / 2; // s2 right of s1
                else MessageBox.Show("Invalid shape placement");
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"SnapTogether: {ex.Message}"); }
        }

        public static void PositionSwap(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count < 2) { MessageBox.Show("Invalid selection: choose at least two shapes"); return; }
                if (!int.TryParse(Interaction.InputBox("Pivot point (1=TL, 2=TR, 3=C, 4=BL, 5=BR)", "Pivot", "3"), out int pos) || pos < 1 || pos > 5)
                { MessageBox.Show("Pivot must be 1-5"); return; }

                int count = sel.ShapeRange.Count; int aIndex = 1;
                for (int i = 2; i <= count; i++)
                {
                    var A = sel.ShapeRange[aIndex]; var B = sel.ShapeRange[i];
                    float w1 = A.Width, h1 = A.Height, l1 = A.Left, t1 = A.Top;
                    float w2 = B.Width, h2 = B.Height, l2 = B.Left, t2 = B.Top;
                    switch (pos)
                    {
                        case 1: A.Left = l2; A.Top = t2; B.Left = l1; B.Top = t1; break;
                        case 2: A.Left = l2 + w2 - w1; A.Top = t2; B.Left = l1 + w1 - w2; B.Top = t1; break;
                        case 3: A.Left = l2 + (w2 - w1) / 2; A.Top = t2 + (h2 - h1) / 2; B.Left = l1 + (w1 - w2) / 2; B.Top = t1 + (h1 - h2) / 2; break;
                        case 4: A.Left = l2; A.Top = t2 + h2 - h1; B.Left = l1; B.Top = t1 + h1 - h2; break;
                        case 5: A.Left = l2 + w2 - w1; A.Top = t2 + h2 - h1; B.Left = l1 + w1 - w2; B.Top = t1 + h1 - h2; break;
                    }
                    aIndex = i;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"PositionSwap: {ex.Message}"); }
        }

        public static void AutoGrid(PowerPoint.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;

            bool slideIsAnchor = sr.Count == 1;
            PowerPoint.Shape anchorShape = null;
            PowerPoint.Slide slide = null;

            double anchorLeft, anchorTop, anchorWidth, anchorHeight, anchorLw;
            int shapesToPlace;
            int firstShapeIndex = 1;

            try
            {
                if (slideIsAnchor)
                {
                    PowerPoint.Shape shp = sr[1];
                    slide = (PowerPoint.Slide)shp.Parent;
                    anchorLeft = 0;
                    anchorTop = 0;
                    anchorWidth = slide.Parent.PageSetup.SlideWidth;
                    anchorHeight = slide.Parent.PageSetup.SlideHeight;
                    anchorLw = 0;
                    shapesToPlace = 1;
                }
                else
                {
                    shapesToPlace = sr.Count - 1;
                    anchorShape = sr[shapesToPlace + 1];
                    anchorLw = Utility.GetLineWeight(anchorShape);
                    anchorLeft = anchorShape.Left;
                    anchorTop = anchorShape.Top;
                    anchorWidth = anchorShape.Width;
                    anchorHeight = anchorShape.Height;
                }

                if (!int.TryParse(Interaction.InputBox("Rows", "AutoGrid", "1"), out int rows) || rows < 1)
                    rows = 1;

                int cols = (int)Math.Ceiling(shapesToPlace / (double)rows);

                string padDefault = ((int)(Math.Sqrt(anchorWidth * anchorHeight) * 0.05) + 1).ToString();
                if (!double.TryParse(Interaction.InputBox("Padding (points)", "AutoGrid", padDefault), out double pad))
                    pad = Utility.ConvertToPoints(padDefault);
                pad += anchorLw / 2.0;

                double availH = anchorHeight - pad * 2;
                double availW = anchorWidth - pad * 2;
                double cellH = availH / rows;
                double cellW = availW / cols;

                double startL = anchorLeft + pad;
                double startT = anchorTop + pad;

                int idx = firstShapeIndex;
                for (int r = 0; r < rows && idx <= shapesToPlace; r++)
                {
                    for (int c = 0; c < cols && idx <= shapesToPlace; c++)
                    {
                        PowerPoint.Shape shp = sr[idx++];
                        try
                        {
                            double shpLw = Utility.GetLineWeight(shp);
                            shp.Left = Utility.F(startL + c * cellW + (cellW - shp.Width) / 2);
                            shp.Top = Utility.F(startT + r * cellH + (cellH - shp.Height) / 2);
                        }
                        finally
                        {
                            if (System.Runtime.InteropServices.Marshal.IsComObject(shp))
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(shp);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoGrid error: {ex.Message}");
            }
            finally
            {
                if (anchorShape != null && System.Runtime.InteropServices.Marshal.IsComObject(anchorShape))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(anchorShape);
                if (slide != null && System.Runtime.InteropServices.Marshal.IsComObject(slide))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(slide);
            }
        }

        public static void AutoSelect(PowerPoint.Application app)
        {
            var sel = app.ActiveWindow.Selection;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 1)
            { MessageBox.Show("Please select a shape first.", "Auto-Select Shapes"); return; }

            var template = sel.ShapeRange[1];
            float targetW = template.Width, targetH = template.Height; Office.MsoShapeType targetType = template.Type;
            bool isAutoShape = targetType == Office.MsoShapeType.msoAutoShape;
            Office.MsoAutoShapeType targetAutoType = isAutoShape ? template.AutoShapeType : Office.MsoAutoShapeType.msoShapeRectangle;
            var slide = app.ActiveWindow.View.Slide as PowerPoint.Slide;
            bool firstMatch = true;

            foreach (PowerPoint.Shape shp in slide.Shapes)
            {
                bool sizeMatch = Math.Abs(shp.Width - targetW) < 0.1 && Math.Abs(shp.Height - targetH) < 0.1;
                bool typeMatch = shp.Type == targetType;
                bool autoMatch = !isAutoShape || shp.AutoShapeType == targetAutoType;
                if (sizeMatch && typeMatch && autoMatch)
                {
                    shp.Select(firstMatch ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse);
                    firstMatch = false;
                }
            }
        }

        public static void CellCenter(PowerPoint.Application app)
        {
            try
            {
                var win = app.ActiveWindow;
                if (win == null) return;

                var sel = win.Selection;
                if (sel?.ShapeRange == null || sel.ShapeRange.Count == 0) return;

                var tblShape = sel.ShapeRange[1];
                if (tblShape.HasTable != Microsoft.Office.Core.MsoTriState.msoTrue) return;

                var tbl = tblShape.Table;

                var targets = new List<(float L, float T, float W, float H)>();
                for (int c = 1; c <= tbl.Columns.Count; c++)
                {
                    for (int r = 1; r <= tbl.Rows.Count; r++)
                    {
                        var cell = tbl.Cell(r, c);
                        if (cell.Selected)
                        {
                            var s = cell.Shape;
                            targets.Add((s.Left, s.Top, s.Width, s.Height));
                        }
                    }
                }
                if (targets.Count == 0) return;

                win.View.Paste();
                var pasted = win.Selection.ShapeRange;
                int pairs = Math.Min(pasted.Count, targets.Count);

                for (int i = 1; i <= pairs; i++)
                {
                    var shp = pasted[i];
                    var tgt = targets[i - 1];

                    shp.Left = tgt.L + (tgt.W - shp.Width) / 2f;
                    shp.Top = tgt.T + (tgt.H - shp.Height) / 2f;
                }
            }
            catch
            {
                /* swallow any add-in-runtime errors */
            }
        }

        public static void ToggleAutoFit(PowerPoint.Selection sel)
        {
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 1) { MessageBox.Show("Please select one or more shapes."); return; }
            foreach (PowerPoint.Shape shp in sel.ShapeRange)
            {
                if (shp.HasTextFrame == Office.MsoTriState.msoTrue)
                    shp.TextFrame.AutoSize = (shp.TextFrame.AutoSize == PowerPoint.PpAutoSize.ppAutoSizeNone)
                        ? PowerPoint.PpAutoSize.ppAutoSizeShapeToFitText
                        : PowerPoint.PpAutoSize.ppAutoSizeNone;
            }
        }

        public static void ConformAutoShapes(PowerPoint.Selection sel)
        {
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count < 2) { MessageBox.Show("Select at least two AutoShapes (last one is the template)."); return; }
            var sr = sel.ShapeRange; var template = sr[sr.Count];
            if (template.Type != Office.MsoShapeType.msoAutoShape) { MessageBox.Show("The last selected shape must be an AutoShape."); return; }
            var shapeType = template.AutoShapeType; float w = template.Width, h = template.Height;
            int adjCount = template.Adjustments.Count; var adjValues = new float[adjCount];
            for (int i = 1; i <= adjCount; i++) adjValues[i - 1] = template.Adjustments[i];
            for (int i = 1; i < sr.Count; i++)
            {
                var shp = sr[i];
                if (shp.Type == Office.MsoShapeType.msoAutoShape)
                {
                    shp.AutoShapeType = shapeType; shp.Width = w; shp.Height = h;
                    for (int j = 1; j <= adjCount && j <= shp.Adjustments.Count; j++) shp.Adjustments[j] = adjValues[j - 1];
                }
            }
        }

        public static void duplicateAcrossSlides(PowerPoint.Application app)
        {
            if (app == null) { MessageBox.Show("PowerPoint application is not running."); return; }
            try
            {
                PowerPoint.Presentation presentation = app.ActivePresentation;
                if (presentation == null) { MessageBox.Show("No active presentation found."); return; }
                PowerPoint.DocumentWindow activeWindow = app.ActiveWindow;
                if (activeWindow == null) { MessageBox.Show("No active window found."); return; }
                PowerPoint.Slide activeSlide = activeWindow.View.Slide;
                if (activeSlide == null) { MessageBox.Show("No active slide found."); return; }
                PowerPoint.Selection selection = activeWindow.Selection;
                if (selection == null || selection.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                { MessageBox.Show("No object selected. Please select a chart, table, image, or shape."); return; }

                selection.Copy();
                if (!System.Windows.Forms.Clipboard.GetDataObject().GetFormats().Any())
                { MessageBox.Show("Failed to copy the selected object."); return; }

                int currentSlideIndex = activeSlide.SlideIndex;
                foreach (PowerPoint.Slide slide in presentation.Slides)
                {
                    if (slide.SlideIndex != currentSlideIndex)
                    {
                        slide.Select();
                        PowerPoint.ShapeRange pastedShapeRange = slide.Shapes.Paste();
                        if (pastedShapeRange == null || pastedShapeRange.Count == 0) continue;
                        PowerPoint.Shape pastedShape = pastedShapeRange[1];

                        if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes || selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                        {
                            PowerPoint.ShapeRange originalShapeRange = selection.ShapeRange;
                            if (originalShapeRange.Count > 0)
                            {
                                PowerPoint.Shape originalShape = originalShapeRange[1];
                                pastedShape.Left = originalShape.Left;
                                pastedShape.Top = originalShape.Top;
                            }
                        }
                    }
                }
                presentation.Slides[currentSlideIndex].Select();
                MessageBox.Show("The selected object has been duplicated across all slides.");
            }
            catch (Exception ex) { MessageBox.Show($"An error occurred: {ex.Message}"); }
        }
    }
}