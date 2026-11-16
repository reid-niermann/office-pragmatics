using Microsoft.VisualBasic;
using OfficePragmatics_Word.Settings;   // shared settings
using OfficePragmatics_Word.Utilities;  // shared utilities (ParseDistance, GetLineWeight, F, etc.)
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficePragmatics_Word.Shape
{
    public static class ShapeUtility
    {
        /* ─────────────────────────────────────────────────────────────────────────
         * SNAP TOGETHER (Word)
         * ───────────────────────────────────────────────────────────────────────── */
        public static void SnapTogether(Word.Selection sel)
        {
            try
            {
                if (sel == null || sel.Type != Word.WdSelectionType.wdSelectionShape || sel.ShapeRange == null || sel.ShapeRange.Count != 2)
                { MessageBox.Show("Invalid selection: choose two shapes"); return; }

                var s1 = sel.ShapeRange[1];
                var s2 = sel.ShapeRange[2];

                float line1 = 0f, line2 = 0f;
                bool alignToBorder = SettingsForm.CurrentSettings.AlignToBorder;
                if (alignToBorder)
                {
                    if (s1.Line.Visible == Office.MsoTriState.msoTrue) line1 = (float)Math.Max(0, s1.Line.Weight);
                    if (s2.Line.Visible == Office.MsoTriState.msoTrue) line2 = (float)Math.Max(0, s2.Line.Weight);
                }

                float w1 = s1.Width + line1, h1 = s1.Height + line1, l1 = s1.Left - line1 / 2f, t1 = s1.Top - line1 / 2f, r1 = l1 + w1, b1 = t1 + h1;
                float w2 = s2.Width + line2, h2 = s2.Height + line2, l2 = s2.Left - line2 / 2f, t2 = s2.Top - line2 / 2f, r2 = l2 + w2, b2 = t2 + h2;

                if (r2 > l1 && r1 > l2 && t1 > b2) s1.Top = b2 + line1 / 2f;                    // s2 above s1
                else if (l1 > r2 && b2 > t1 && b1 > t2) s1.Left = r2 + line1 / 2f;             // s2 left of s1
                else if (r2 > l1 && r1 > l2 && t2 > b1) s1.Top = t2 - h1 + line1 / 2f;         // s2 below s1
                else if (l2 > r1 && b2 > t1 && b1 > t2) s1.Left = l2 - w1 + line1 / 2f;        // s2 right of s1
                else MessageBox.Show("Invalid shape placement");
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"SnapTogether (Word): {ex.Message}"); }
        }

        /* ─────────────────────────────────────────────────────────────────────────
         * POSITION SWAP (Word)
         * ───────────────────────────────────────────────────────────────────────── */
        public static void PositionSwap(Word.Selection sel)
        {
            try
            {
                if (sel == null || sel.Type != Word.WdSelectionType.wdSelectionShape || sel.ShapeRange == null || sel.ShapeRange.Count < 2)
                { MessageBox.Show("Invalid selection: choose at least two shapes"); return; }

                if (!int.TryParse(Interaction.InputBox("Pivot point (1=TL, 2=TR, 3=C, 4=BL, 5=BR)", "Pivot", "3"), out int pos) || pos < 1 || pos > 5)
                { MessageBox.Show("Pivot must be 1-5"); return; }

                var sr = sel.ShapeRange;
                int count = sr.Count; int aIndex = 1;

                for (int i = 2; i <= count; i++)
                {
                    var A = sr[aIndex]; var B = sr[i];
                    float w1 = A.Width, h1 = A.Height, l1 = A.Left, t1 = A.Top;
                    float w2 = B.Width, h2 = B.Height, l2 = B.Left, t2 = B.Top;

                    switch (pos)
                    {
                        case 1: A.Left = l2; A.Top = t2; B.Left = l1; B.Top = t1; break;
                        case 2: A.Left = l2 + w2 - w1; A.Top = t2; B.Left = l1 + w1 - w2; B.Top = t1; break;
                        case 3: A.Left = l2 + (w2 - w1) / 2f; A.Top = t2 + (h2 - h1) / 2f; B.Left = l1 + (w1 - w2) / 2f; B.Top = t1 + (h1 - h2) / 2f; break;
                        case 4: A.Left = l2; A.Top = t2 + h2 - h1; B.Left = l1; B.Top = t1 + h1 - h2; break;
                        case 5: A.Left = l2 + w2 - w1; A.Top = t2 + h2 - h1; B.Left = l1 + w1 - w2; B.Top = t1 + h1 - h2; break;
                    }
                    aIndex = i;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"PositionSwap (Word): {ex.Message}"); }
        }

        /* ─────────────────────────────────────────────────────────────────────────
         * AUTO GRID (Word)
         * ───────────────────────────────────────────────────────────────────────── */
        public static void AutoGrid(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;

            bool pageIsAnchor = (sr.Count == 1);
            Word.Shape anchorShape = null;
            Word.Document doc = null;

            double anchorLeft, anchorTop, anchorWidth, anchorHeight, anchorLw = 0;
            int shapesToPlace;
            int firstShapeIndex = 1;

            try
            {
                if (pageIsAnchor)
                {
                    var shp = sr[1];
                    doc = shp.Application.ActiveDocument;

                    // Page area inside margins
                    var ps = doc.PageSetup;
                    anchorLeft = ps.LeftMargin;
                    anchorTop = ps.TopMargin;
                    anchorWidth = ps.PageWidth - ps.LeftMargin - ps.RightMargin;
                    anchorHeight = ps.PageHeight - ps.TopMargin - ps.BottomMargin;
                    shapesToPlace = 1; // the single selected shape will be positioned
                }
                else
                {
                    shapesToPlace = sr.Count - 1;       // all but the last shape
                    anchorShape = sr[shapesToPlace + 1];
                    anchorLw = Utility.GetLineWeight(anchorShape);
                    anchorLeft = anchorShape.Left;
                    anchorTop = anchorShape.Top;
                    anchorWidth = anchorShape.Width;
                    anchorHeight = anchorShape.Height;
                }

                // Rows
                int rows;
                if (!int.TryParse(Interaction.InputBox("Rows", "AutoGrid", "1"), out rows) || rows < 1) rows = 1;

                // Columns
                int cols = (int)Math.Ceiling(shapesToPlace / (double)rows);

                // Padding (points)
                string padDefault = ((int)(Math.Sqrt(anchorWidth * anchorHeight) * 0.05) + 1).ToString();
                double pad;
                if (!double.TryParse(Interaction.InputBox("Padding (points)", "AutoGrid", padDefault), out pad))
                    pad = Utility.ConvertToPoints(padDefault);

                pad += anchorLw / 2.0;

                double availH = anchorHeight - pad * 2;
                double availW = anchorWidth - pad * 2;
                double cellH = (rows > 0) ? (availH / rows) : availH;
                double cellW = (cols > 0) ? (availW / cols) : availW;

                double startL = anchorLeft + pad;
                double startT = anchorTop + pad;

                int idx = firstShapeIndex;
                for (int r = 0; r < rows && idx <= shapesToPlace; r++)
                {
                    for (int c = 0; c < cols && idx <= shapesToPlace; c++)
                    {
                        Word.Shape shp = sr[idx++];
                        try
                        {
                            // Ensure positioning is relative to page for consistency
                            shp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                            shp.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;

                            shp.Left = Utility.F(startL + c * cellW + (cellW - shp.Width) / 2.0);
                            shp.Top = Utility.F(startT + r * cellH + (cellH - shp.Height) / 2.0);
                        }
                        catch { /* ignore per-shape issues */ }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"AutoGrid (Word) error: {ex.Message}");
            }
        }

        /* ─────────────────────────────────────────────────────────────────────────
         * CELL CENTER (Word)
         *  - Paste and center into each selected table cell
         * ───────────────────────────────────────────────────────────────────────── */
        private struct CellBox
        {
            public float L, T, W, H;
            public CellBox(float l, float t, float w, float h) { L = l; T = t; W = w; H = h; }
        }

        public static void CellCenter(Word.Application app)
        {
            try
            {
                var sel = app.Selection;
                if (sel == null || sel.Range == null) return;

                if (sel.Tables == null || sel.Tables.Count == 0)
                { MessageBox.Show("Please select one or more table cells.", "Cell Center"); return; }

                var boxes = new List<CellBox>();
                var cells = sel.Cells;
                if (cells == null || cells.Count == 0) return;

                for (int i = 1; i <= cells.Count; i++)
                {
                    var cell = cells[i];
                    var r = cell.Range;

                    float left = 0f, top = 0f, w = 0f, h = 0f;
                    try
                    {
                        left = Convert.ToSingle(r.get_Information(Word.WdInformation.wdHorizontalPositionRelativeToPage));
                        top = Convert.ToSingle(r.get_Information(Word.WdInformation.wdVerticalPositionRelativeToPage));
                        w = Convert.ToSingle(cell.Width);
                        float ch = Convert.ToSingle(cell.Height);
                        if (ch <= 0) ch = Convert.ToSingle(cell.Row.Height > 0 ? cell.Row.Height : 0);
                        if (ch <= 0) ch = 18f; // fallback
                        h = ch;
                    }
                    catch { /* ignore problematic cells */ }

                    if (w > 0 && h > 0) boxes.Add(new CellBox(left, top, w, h));
                }

                if (boxes.Count == 0) return;

                // Paste using Selection (Word has no Shapes.Paste)
                sel.Paste();

                // Gather pasted shapes (could be ShapeRange or InlineShapes)
                var pastedShapes = CollectPastedShapes(app);

                int pairs = Math.Min(pastedShapes.Count, boxes.Count);
                for (int i = 0; i < pairs; i++)
                {
                    var shp = pastedShapes[i];
                    var box = boxes[i];

                    shp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                    shp.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;

                    shp.Left = box.L + (box.W - shp.Width) / 2f;
                    shp.Top = box.T + (box.H - shp.Height) / 2f;
                }
            }
            catch
            {
                /* swallow runtime add-in errors */
            }
        }

        /// <summary>
        /// Collects shapes from the current Selection after a paste.
        /// Handles both ShapeRange and InlineShapes by converting inline to floating.
        /// </summary>
        private static List<Word.Shape> CollectPastedShapes(Word.Application app)
        {
            var list = new List<Word.Shape>();
            var sel = app.Selection;
            if (sel == null) return list;

            try
            {
                if (sel.Type == Word.WdSelectionType.wdSelectionShape && sel.ShapeRange != null && sel.ShapeRange.Count > 0)
                {
                    for (int i = 1; i <= sel.ShapeRange.Count; i++)
                        list.Add(sel.ShapeRange[i]);
                    return list;
                }
            }
            catch { /* ignore */ }

            try
            {
                if (sel.InlineShapes != null && sel.InlineShapes.Count > 0)
                {
                    for (int i = 1; i <= sel.InlineShapes.Count; i++)
                    {
                        try
                        {
                            var s = sel.InlineShapes[i].ConvertToShape();
                            if (s != null) list.Add(s);
                        }
                        catch { /* ignore single inline */ }
                    }
                }
            }
            catch { /* ignore */ }

            return list;
        }
    }
}