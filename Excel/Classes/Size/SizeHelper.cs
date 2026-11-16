using System;
using System.Linq;
using Microsoft.VisualBasic;                        // Interaction.InputBox
using Excel = Microsoft.Office.Interop.Excel;
using OfficePragmatics_Excel.Utilities;            // Utility.F, Utility.GetLineWeight

namespace OfficePragmatics_Excel.Size
{
    /// <summary>Logical rectangle bounded by the nearest page breaks,
    /// limited to the worksheet’s Print Area (if any).</summary>
    internal sealed class PageBreakBounds
    {
        public readonly int TopRow, BottomRow, LeftCol, RightCol;
        public PageBreakBounds(int t, int b, int l, int r)
        { TopRow = t; BottomRow = b; LeftCol = l; RightCol = r; }
    }

    public static class SizeHelper
    {
        /*──────────────────────────────────────────────────────────────────────*/
        /*  PAGE-BREAK HELPERS (print-area aware, TopLeftCell based)           */
        /*──────────────────────────────────────────────────────────────────────*/
        private static PageBreakBounds GetPageBreakBounds(Excel.Worksheet ws,
                                                          Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return null;

            // — 1. row/col bounds of the whole selection --------------------
            int rMin = int.MaxValue, cMin = int.MaxValue;
            int rMax = 0, cMax = 0;

            foreach (Excel.Shape shp in sr)
            {
                var tl = shp.TopLeftCell;
                var br = shp.BottomRightCell;
                rMin = Math.Min(rMin, tl.Row);
                cMin = Math.Min(cMin, tl.Column);
                rMax = Math.Max(rMax, br.Row);
                cMax = Math.Max(cMax, br.Column);
            }

            // — 2. clamp to Print Area or UsedRange -------------------------
            Excel.Range pa;
            if (!string.IsNullOrEmpty(ws.PageSetup.PrintArea))
                pa = ws.Range[ws.PageSetup.PrintArea];
            else
                pa = ws.UsedRange;

            int areaT = pa.Row,
                areaL = pa.Column,
                areaB = areaT + pa.Rows.Count - 1,
                areaR = areaL + pa.Columns.Count - 1;

            int top = areaT, bottom = areaB;
            int left = areaL, right = areaR;

            // — 3. narrow to the slab bounded by nearest manual page breaks -
            foreach (Excel.HPageBreak br in ws.HPageBreaks)
            {
                int rw = br.Location.Row;
                if (rw < areaT || rw > areaB) continue;

                if (rw <= rMin) top = Math.Max(top, rw);
                else if (rw > rMax) { bottom = Math.Min(bottom, rw - 1); break; }
            }

            foreach (Excel.VPageBreak br in ws.VPageBreaks)
            {
                int cl = br.Location.Column;
                if (cl < areaL || cl > areaR) continue;

                if (cl <= cMin) left = Math.Max(left, cl);
                else if (cl > cMax) { right = Math.Min(right, cl - 1); break; }
            }

            return new PageBreakBounds(top, bottom, left, right);
        }

        /*──────────────────────────── Extend (left / right / top / bottom) ───*/
        public static void ExtendLeft(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr); if (pb == null) return;

            double pageLeft = ws.Columns[pb.LeftCol].Left;

            // single shape
            if (sr.Count == 1)
            {
                Excel.Shape shp = sr.Item(1);
                double lw = Utility.GetLineWeight(shp);
                double delta = shp.Left - lw / 2.0 - pageLeft;
                if (delta > 0)
                {
                    shp.Width = Utility.F(shp.Width + delta);
                    shp.Left = Utility.F(pageLeft + lw / 2.0);
                }
                return;
            }

            // multiple shapes (anchor = last selected)
            Excel.Shape anchor = sr.Item(sr.Count);
            double anchorOuterLeft = anchor.Left - Utility.GetLineWeight(anchor) / 2.0;

            for (int i = 1; i < sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                double shpOuterLeft = shp.Left - lw / 2.0;
                double delta = shpOuterLeft - anchorOuterLeft;

                shp.Width = Utility.F(shp.Width + delta);
                shp.Left = Utility.F(anchorOuterLeft + lw / 2.0);
            }
        }

        public static void ExtendRight(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr); if (pb == null) return;

            double pageRight = ws.Columns[pb.RightCol].Left +
                               ws.Columns[pb.RightCol].Width;

            if (sr.Count == 1)
            {
                Excel.Shape shp = sr.Item(1);
                double lw = Utility.GetLineWeight(shp);
                double shpOuterRight = shp.Left + shp.Width + lw / 2.0;
                double delta = pageRight - shpOuterRight;
                if (delta > 0)
                    shp.Width = Utility.F(shp.Width + delta);
                return;
            }

            Excel.Shape anchor = sr.Item(sr.Count);
            double anchorOuterRight = anchor.Left + anchor.Width +
                                      Utility.GetLineWeight(anchor) / 2.0;

            for (int i = 1; i < sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                double shpOuterRight = shp.Left + shp.Width + lw / 2.0;
                double delta = anchorOuterRight - shpOuterRight;
                shp.Width = Utility.F(shp.Width + delta);
            }
        }

        public static void ExtendTop(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr); if (pb == null) return;

            double pageTop = ws.Rows[pb.TopRow].Top;

            if (sr.Count == 1)
            {
                Excel.Shape shp = sr.Item(1);
                double lw = Utility.GetLineWeight(shp);
                double delta = shp.Top - lw / 2.0 - pageTop;
                if (delta > 0)
                {
                    shp.Height = Utility.F(shp.Height + delta);
                    shp.Top = Utility.F(pageTop + lw / 2.0);
                }
                return;
            }

            Excel.Shape anchor = sr.Item(sr.Count);
            double anchorOuterTop = anchor.Top - Utility.GetLineWeight(anchor) / 2.0;

            for (int i = 1; i < sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                double shpOuterTop = shp.Top - lw / 2.0;
                double delta = shpOuterTop - anchorOuterTop;

                shp.Height = Utility.F(shp.Height + delta);
                shp.Top = Utility.F(anchorOuterTop + lw / 2.0);
            }
        }

        public static void ExtendBottom(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr); if (pb == null) return;

            double pageBottom = ws.Rows[pb.BottomRow].Top +
                                ws.Rows[pb.BottomRow].Height;

            if (sr.Count == 1)
            {
                Excel.Shape shp = sr.Item(1);
                double lw = Utility.GetLineWeight(shp);
                double shpOuterBottom = shp.Top + shp.Height + lw / 2.0;
                double delta = pageBottom - shpOuterBottom;
                if (delta > 0)
                    shp.Height = Utility.F(shp.Height + delta);
                return;
            }

            Excel.Shape anchor = sr.Item(sr.Count);
            double anchorOuterBottom = anchor.Top + anchor.Height +
                                       Utility.GetLineWeight(anchor) / 2.0;

            for (int i = 1; i < sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                double shpOuterBottom = shp.Top + shp.Height + lw / 2.0;
                double delta = anchorOuterBottom - shpOuterBottom;
                shp.Height = Utility.F(shp.Height + delta);
            }
        }

        /*────────────────────────────────────────────  Fit  ──────────────────*/
        public static void FitToHorizontal(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;

            // single shape – fit to worksheet drawing width
            if (sr.Count == 1)
            {
                Excel.Shape shp = sr.Item(1);
                string m = Interaction.InputBox("Enter horizontal margin (pts):",
                                                 "Fit Horizontally", "10");
                if (!double.TryParse(m, out double margin)) margin = 10;

                double wsWidth = shp.Parent.Width;
                double lw = Utility.GetLineWeight(shp);
                double avail = wsWidth - 2 * margin;
                if (avail <= 0) return;

                shp.Width = Utility.F(avail - lw);
                shp.Left = Utility.F(margin + lw / 2.0);
                return;
            }

            // multi-shape – anchor = last selected (container)
            Excel.Shape anchor = sr.Item(sr.Count);
            double anchorOuterLeft = anchor.Left - Utility.GetLineWeight(anchor) / 2.0;
            double anchorOuterWidth = anchor.Width + Utility.GetLineWeight(anchor);

            string g = Interaction.InputBox("Enter horizontal gap (pts):",
                                            "Fit Horizontally", "10");
            if (!double.TryParse(g, out double gap)) gap = 10;

            int n = sr.Count - 1; if (n <= 0) return;
            double totalGap = (n - 1) * gap;
            double availForShapes = anchorOuterWidth - totalGap;
            if (availForShapes <= 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Not enough space in the anchor shape to fit other shapes " +
                    "with the specified gap.",
                    "Fit Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            double widthEach = availForShapes / n;

            double x = anchorOuterLeft;
            for (int i = 1; i <= n; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                shp.Width = Utility.F(widthEach - lw);
                shp.Left = Utility.F(x + lw / 2.0);
                x += widthEach + gap;
            }
        }

        public static void FitToVertical(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;

            // single shape – fit to worksheet drawing height
            if (sr.Count == 1)
            {
                Excel.Shape shp = sr.Item(1);
                string m = Interaction.InputBox("Enter vertical margin (pts):",
                                                 "Fit Vertically", "10");
                if (!double.TryParse(m, out double margin)) margin = 10;

                double wsHeight = shp.Parent.Height;
                double lw = Utility.GetLineWeight(shp);
                double avail = wsHeight - 2 * margin;
                if (avail <= 0) return;

                shp.Height = Utility.F(avail - lw);
                shp.Top = Utility.F(margin + lw / 2.0);
                return;
            }

            // multi-shape – anchor = last selected (container)
            Excel.Shape anchor = sr.Item(sr.Count);
            double anchorOuterTop = anchor.Top - Utility.GetLineWeight(anchor) / 2.0;
            double anchorOuterHeight = anchor.Height + Utility.GetLineWeight(anchor);

            string g = Interaction.InputBox("Enter vertical gap (pts):",
                                            "Fit Vertically", "10");
            if (!double.TryParse(g, out double gap)) gap = 10;

            int n = sr.Count - 1; if (n <= 0) return;
            double totalGap = (n - 1) * gap;
            double availForShapes = anchorOuterHeight - totalGap;
            if (availForShapes <= 0)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Not enough space in the anchor shape to fit other shapes " +
                    "with the specified gap.",
                    "Fit Error",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
                return;
            }
            double heightEach = availForShapes / n;

            double y = anchorOuterTop;
            for (int i = 1; i <= n; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                shp.Height = Utility.F(heightEach - lw);
                shp.Top = Utility.F(y + lw / 2.0);
                y += heightEach + gap;
            }
        }

        /*────────────────────────────────────────────  Same  ─────────────────*/
        public static void SameWidth(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            Excel.Shape anchor = sr.Item(sr.Count);
            double targetOuterWidth = anchor.Width + Utility.GetLineWeight(anchor);

            for (int i = 1; i <= sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                shp.Width = Utility.F(targetOuterWidth - Utility.GetLineWeight(shp));
            }
        }

        public static void SameHeight(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            Excel.Shape anchor = sr.Item(sr.Count);
            double targetOuterHeight = anchor.Height + Utility.GetLineWeight(anchor);

            for (int i = 1; i <= sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                shp.Height = Utility.F(targetOuterHeight - Utility.GetLineWeight(shp));
            }
        }

        public static void SameSize(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            Excel.Shape anchor = sr.Item(sr.Count);
            double targetOuterWidth = anchor.Width + Utility.GetLineWeight(anchor);
            double targetOuterHeight = anchor.Height + Utility.GetLineWeight(anchor);

            for (int i = 1; i <= sr.Count; i++)
            {
                Excel.Shape shp = sr.Item(i);
                double lw = Utility.GetLineWeight(shp);
                shp.Width = Utility.F(targetOuterWidth - lw);
                shp.Height = Utility.F(targetOuterHeight - lw);
            }
        }
    }
}