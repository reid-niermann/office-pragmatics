using System;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Core = Microsoft.Office.Core;
using OfficePragmatics_Excel.Settings;     // SettingsForm.CurrentSettings
using OfficePragmatics_Excel.Utilities;    // Utility.GetLineWeight

namespace OfficePragmatics_Excel.Alignment
{
    /*──────────────────────────────────────────────────────────────────────────*/
    /*  PAGE-BREAK HELPERS                                                     */
    /*──────────────────────────────────────────────────────────────────────────*/
    internal sealed class PageBreakBounds
    {
        public readonly int TopRow, BottomRow, LeftCol, RightCol;
        public PageBreakBounds(int t, int b, int l, int r)
        { TopRow = t; BottomRow = b; LeftCol = l; RightCol = r; }
    }

    public static class AlignmentHelper
    {
        /// <summary>
        /// Returns the row/column rectangle of the *page* that contains the
        /// supplied shapes, limited to the worksheet’s Print Area if one exists.
        /// </summary>
        private static PageBreakBounds GetPageBreakBounds(Excel.Worksheet ws,
                                                          Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return null;

            /* 1️⃣  BOUNDING ROW/COL OF ENTIRE SELECTION -------------------- */
            int rowMin = int.MaxValue, colMin = int.MaxValue;
            int rowMax = 0, colMax = 0;

            foreach (Excel.Shape shp in sr)
            {
                var tl = shp.TopLeftCell;
                var br = shp.BottomRightCell;
                rowMin = Math.Min(rowMin, tl.Row);
                colMin = Math.Min(colMin, tl.Column);
                rowMax = Math.Max(rowMax, br.Row);
                colMax = Math.Max(colMax, br.Column);
            }

            /* 2️⃣  PRINT-AREA LIMITS --------------------------------------- */
            Excel.Range printRange;

            if (!string.IsNullOrEmpty(ws.PageSetup.PrintArea))
                printRange = ws.Range[ws.PageSetup.PrintArea];
            else
                printRange = ws.UsedRange;

            int areaTop = printRange.Row;
            int areaLeft = printRange.Column;
            int areaBottom = areaTop + printRange.Rows.Count - 1;
            int areaRight = areaLeft + printRange.Columns.Count - 1;

            /* 3️⃣  INITIAL PAGE SLAB IS WHOLE PRINT-AREA ------------------- */
            int top = areaTop;
            int bottom = areaBottom;
            int left = areaLeft;
            int right = areaRight;

            /* 4️⃣  CLAMP TO NEAREST PAGE BREAKS INSIDE PRINT-AREA ---------- */
            foreach (Excel.HPageBreak br in ws.HPageBreaks)
            {
                int rw = br.Location.Row;
                if (rw < areaTop || rw > areaBottom) continue;     // ignore breaks outside PA

                if (rw <= rowMin) top = Math.Max(top, rw);
                else if (rw > rowMax) { bottom = Math.Min(bottom, rw - 1); break; }
            }

            foreach (Excel.VPageBreak br in ws.VPageBreaks)
            {
                int cl = br.Location.Column;
                if (cl < areaLeft || cl > areaRight) continue;     // ignore breaks outside PA

                if (cl <= colMin) left = Math.Max(left, cl);
                else if (cl > colMax) { right = Math.Min(right, cl - 1); break; }
            }

            return new PageBreakBounds(top, bottom, left, right);
        }

        /*──────────────────────────────────────────────────────────────────────*/
        /*  SNAP TO GRID + DISTRIBUTE (unchanged)                              */
        /*──────────────────────────────────────────────────────────────────────*/
        public static void SnapSelectionToCellGrid(Excel.Worksheet ws,
                                                   Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr); if (pb == null) return;

            foreach (Excel.Shape shp in sr)
            {
                var tl = shp.TopLeftCell;
                var br = shp.BottomRightCell;

                int r1 = Math.Max(tl.Row, pb.TopRow);
                int c1 = Math.Max(tl.Column, pb.LeftCol);
                int r2 = Math.Min(br.Row, pb.BottomRow);
                int c2 = Math.Min(br.Column, pb.RightCol);

                Excel.Range rng = ws.Range[ws.Cells[r1, c1], ws.Cells[r2, c2]];
                shp.Left = (float)rng.Left;
                shp.Top = (float)rng.Top;
                shp.Width = (float)rng.Width;
                shp.Height = (float)rng.Height;
            }
        }

        public static void DistributeInPage(Excel.Worksheet ws,
                                            Excel.ShapeRange sr, bool horizontal)
        {
            if (sr == null || sr.Count < 3) return;
            var pb = GetPageBreakBounds(ws, sr); if (pb == null) return;

            if (horizontal)
                sr.Distribute(Core.MsoDistributeCmd.msoDistributeHorizontally,
                              Core.MsoTriState.msoFalse);
            else
                sr.Distribute(Core.MsoDistributeCmd.msoDistributeVertically,
                              Core.MsoTriState.msoFalse);
        }

        /*──────────────────────────────────────────────────────────────────────*/
        /*  ALIGN LEFT / RIGHT / CENTRE                                        */
        /*──────────────────────────────────────────────────────────────────────*/
        public static void AlignLeft(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            bool border = SettingsForm.CurrentSettings.AlignToBorder;
            var pb = GetPageBreakBounds(ws, sr);

            double pageLeft = ws.Columns[pb?.LeftCol ?? 1].Left;

            if (sr.Count == 1)
            {
                var s = sr.Item(1);
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Left = (float)(pageLeft + offset);
                return;
            }

            var anchor = sr.Item(sr.Count);
            double anchorOuterLeft = anchor.Left -
                                     (border && anchor.Line.Visible == Core.MsoTriState.msoTrue
                                      ? anchor.Line.Weight / 2.0 : 0);

            foreach (Excel.Shape s in sr)
            {
                if (s == anchor) continue;
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Left = (float)(anchorOuterLeft + offset);
            }
        }

        public static void AlignRight(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            bool border = SettingsForm.CurrentSettings.AlignToBorder;
            var pb = GetPageBreakBounds(ws, sr);

            double pageRight = ws.Columns[pb.RightCol].Left +
                               ws.Columns[pb.RightCol].Width;

            if (sr.Count == 1)
            {
                var s = sr.Item(1);
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Left = (float)(pageRight - s.Width - offset);
                return;
            }

            var anchor = sr.Item(sr.Count);
            double anchorOuterRight = anchor.Left + anchor.Width +
                                      (border && anchor.Line.Visible == Core.MsoTriState.msoTrue
                                       ? anchor.Line.Weight / 2.0 : 0);

            foreach (Excel.Shape s in sr)
            {
                if (s == anchor) continue;
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Left = (float)(anchorOuterRight - offset - s.Width);
            }
        }

        public static void AlignCenter(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr);

            if (sr.Count == 1)
            {
                var s = sr.Item(1);
                double pageLeft = ws.Columns[pb.LeftCol].Left;
                double pageRight = ws.Columns[pb.RightCol].Left +
                                   ws.Columns[pb.RightCol].Width;
                s.Left = (float)((pageRight - pageLeft - s.Width) / 2.0 + pageLeft);
                return;
            }

            var anchor = sr.Item(sr.Count);
            double anchorMid = anchor.Left + anchor.Width / 2.0;

            foreach (Excel.Shape s in sr)
            {
                if (s == anchor) continue;
                s.Left = (float)(anchorMid - s.Width / 2.0);
            }
        }

        /*──────────────────────────────────────────────────────────────────────*/
        /*  ALIGN TOP / BOTTOM / MIDDLE                                        */
        /*──────────────────────────────────────────────────────────────────────*/
        public static void AlignTop(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            bool border = SettingsForm.CurrentSettings.AlignToBorder;
            var pb = GetPageBreakBounds(ws, sr);

            double pageTop = ws.Rows[pb.TopRow].Top;

            if (sr.Count == 1)
            {
                var s = sr.Item(1);
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Top = (float)(pageTop + offset);
                return;
            }

            var anchor = sr.Item(sr.Count);
            double anchorOuterTop = anchor.Top -
                                    (border && anchor.Line.Visible == Core.MsoTriState.msoTrue
                                     ? anchor.Line.Weight / 2.0 : 0);

            foreach (Excel.Shape s in sr)
            {
                if (s == anchor) continue;
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Top = (float)(anchorOuterTop + offset);
            }
        }

        public static void AlignBottom(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            bool border = SettingsForm.CurrentSettings.AlignToBorder;
            var pb = GetPageBreakBounds(ws, sr);

            double pageBottom = ws.Rows[pb.BottomRow].Top +
                                ws.Rows[pb.BottomRow].Height;

            if (sr.Count == 1)
            {
                var s = sr.Item(1);
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Top = (float)(pageBottom - s.Height - offset);
                return;
            }

            var anchor = sr.Item(sr.Count);
            double anchorOuterBottom = anchor.Top + anchor.Height +
                                       (border && anchor.Line.Visible == Core.MsoTriState.msoTrue
                                        ? anchor.Line.Weight / 2.0 : 0);

            foreach (Excel.Shape s in sr)
            {
                if (s == anchor) continue;
                double offset = border && s.Line.Visible == Core.MsoTriState.msoTrue
                              ? Math.Max(0, s.Line.Weight / 2.0) : 0;
                s.Top = (float)(anchorOuterBottom - offset - s.Height);
            }
        }

        public static void AlignMiddle(Excel.Worksheet ws, Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            var pb = GetPageBreakBounds(ws, sr);

            if (sr.Count == 1)
            {
                var s = sr.Item(1);
                double pageTop = ws.Rows[pb.TopRow].Top;
                double pageBottom = ws.Rows[pb.BottomRow].Top +
                                    ws.Rows[pb.BottomRow].Height;
                s.Top = (float)((pageBottom - pageTop - s.Height) / 2.0 + pageTop);
                return;
            }

            var anchor = sr.Item(sr.Count);
            double anchorMid = anchor.Top + anchor.Height / 2.0;

            foreach (Excel.Shape s in sr)
            {
                if (s == anchor) continue;
                s.Top = (float)(anchorMid - s.Height / 2.0);
            }
        }

        /*──────────────────────────────────────────────────────────────────────*/
        /*  SNAP TOGETHER / CYCLE SWAP (unchanged)                              */
        /*──────────────────────────────────────────────────────────────────────*/
        public static void SnapTogether(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count < 2) return;

            var shapes = sr.Cast<Excel.Shape>()
                           .OrderBy(s => s.Left)
                           .ToArray();

            double x = shapes[0].Left;
            for (int i = 1; i < shapes.Length; i++)
            {
                var prev = shapes[i - 1];
                var shp = shapes[i];
                double lw = Utility.GetLineWeight(shp);
                x = prev.Left + prev.Width;
                shp.Left = (float)(x + lw / 2.0);
            }
        }

        public static void CycleSwap(Excel.ShapeRange sr)
        {
            if (sr == null || sr.Count < 2) return;

            var pos = sr.Cast<Excel.Shape>()
                        .Select(s => new { s.Left, s.Top })
                        .ToArray();

            for (int i = 1; i <= sr.Count; i++)
            {
                var shp = sr.Item(i);
                var tgt = pos[i % sr.Count];
                shp.Left = (float)tgt.Left;
                shp.Top = (float)tgt.Top;
            }
        }
    }
}
