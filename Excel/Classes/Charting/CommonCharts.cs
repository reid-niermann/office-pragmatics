// WaterfallChartHelper.cs – anchored-row version (no ROW()-2 magic)
// C# 7.3 / .NET 4.8, VSTO
//
// - One range prompt
// - Chart 2 rows below the written range
// - No Excel table
// - Header styled (#F2F2F2, bold, bottom border, center except Label)
// - Simple waterfall: now indexes as ROW()-ROW($LabelHeader$)
// - Stacked waterfall:
//     * Subtotal/end total = row where COUNTBLANK(rowCats) = COLUMNS(headerCats)
//     * No header-name lookups
//     * Only references the rows we created
//     * All “ROW math” is done vs. anchored header / anchored first-data-row,
//       so inserting rows above the table won’t break anything.

using System;
using System.Runtime.InteropServices;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_Excel
{
    public static class WaterfallChartHelper
    {

        public static void InsertStackedWaterfall()
        {
            var xl = Globals.ThisAddIn.Application;
            var tblTL = (Excel.Range)xl.InputBox(
                "Pick the top-left cell for the stacked waterfall table:", Type: 8);
            if (tblTL == null) return;
            BuildStacked(tblTL);
        }

        // ────────────────────────── HELPERS ──────────────────────────

        private static string ColLetter(int col)
        {
            int dividend = col;
            string colLetter = "";
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colLetter = Convert.ToChar(65 + modulo) + colLetter;
                dividend = (dividend - modulo) / 26;
            }
            return colLetter;
        }

        private static void ApplyHeaderFormat(Excel.Worksheet ws, int row, int colStart, int colEnd, int labelCol)
        {
            Excel.Range header = ws.Range[ws.Cells[row, colStart], ws.Cells[row, colEnd]];
            header.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(242, 242, 242)); // #F2F2F2
            header.Font.Bold = true;
            header.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            header.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
            header.Borders[Excel.XlBordersIndex.xlEdgeBottom].Color = ColorTranslator.ToOle(Color.FromArgb(200, 200, 200));
            header.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = Excel.XlBorderWeight.xlThin;

            Excel.Range labelCell = ws.Cells[row, labelCol];
            labelCell.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

            Marshal.ReleaseComObject(labelCell);
            Marshal.ReleaseComObject(header);
        }

        // ────────────────────────── STACKED WATERFALL ──────────────────────────
        private static void BuildStacked(Excel.Range tl)
        {
            Excel.Worksheet ws = tl.Worksheet;
            int r0 = tl.Row;
            int c0 = tl.Column;

            int colLabel = c0 + 0;
            int colCat1 = c0 + 1;
            int colCat2 = c0 + 2;
            int colCat3 = c0 + 3;
            int colCat4 = c0 + 4;
            int colSeg1 = c0 + 5;
            int colSeg2 = c0 + 6;
            int colSeg3 = c0 + 7;
            int colSeg4 = c0 + 8;
            int colBlank = c0 + 9;
            int colTotal = c0 + 10;
            int colCum = c0 + 11;
            int colReset = c0 + 12;
            int colIdx = c0 + 13;
            int colPrevRst = c0 + 14;

            // headers
            ws.Cells[r0, colLabel].Value = "Label";
            ws.Cells[r0, colCat1].Value = "Cat 1";
            ws.Cells[r0, colCat2].Value = "Cat 2";
            ws.Cells[r0, colCat3].Value = "Cat 3";
            ws.Cells[r0, colCat4].Value = "Cat 4";
            ws.Cells[r0, colSeg1].Value = "Seg1";
            ws.Cells[r0, colSeg2].Value = "Seg2";
            ws.Cells[r0, colSeg3].Value = "Seg3";
            ws.Cells[r0, colSeg4].Value = "Seg4";
            ws.Cells[r0, colBlank].Value = "Blank";
            ws.Cells[r0, colTotal].Value = "Total";
            ws.Cells[r0, colCum].Value = "Cum.";
            ws.Cells[r0, colReset].Value = "Reset";
            ws.Cells[r0, colIdx].Value = "Idx";
            ws.Cells[r0, colPrevRst].Value = "PrevRst";

            ApplyHeaderFormat(ws, r0, colLabel, colPrevRst, colLabel);

            // seed (7 rows)
            string[] labels = { "Part 1", "Part 2", "Part 3", "Part 4", "Subtotal", "Part 5", "Total" };
            double?[,] cat = {
                {1000, 300, 100, null},
                { 500, 200,  80, null},
                { 300, 300, 800, null},
                {-600,-200,-400, null},
                { null,null,null, null},
                { null,null,null, 500},
                { null,null,null, null}
            };
            int nRows = labels.Length;
            int dataStartRow = r0 + 1;
            int dataEndRow = dataStartRow + nRows - 1;

            for (int i = 0; i < nRows; i++)
            {
                int r = dataStartRow + i;
                ws.Cells[r, colLabel].Value = labels[i];
                for (int k = 0; k < 4; k++)
                    if (cat[i, k].HasValue)
                        ws.Cells[r, colCat1 + k].Value = cat[i, k].Value;
            }

            // handy addresses
            string labelHeaderAddr = $"${ColLetter(colLabel)}${r0}";
            string firstDataLabelAddr = $"${ColLetter(colLabel)}${dataStartRow}";
            string catHeaderRange = $"${ColLetter(colCat1)}${r0}:${ColLetter(colCat4)}${r0}";
            string catsDataRange = $"${ColLetter(colCat1)}${dataStartRow}:${ColLetter(colCat4)}${dataEndRow}";
            string idxFullRange = $"${ColLetter(colIdx)}${dataStartRow}:${ColLetter(colIdx)}${dataEndRow}";
            string colCountExpr = $"COLUMNS({catHeaderRange})";

            string colCat1L = ColLetter(colCat1);
            string colCat4L = ColLetter(colCat4);
            string colSeg1L = ColLetter(colSeg1);
            string colSeg2L = ColLetter(colSeg2);
            string colSeg3L = ColLetter(colSeg3);
            string colSeg4L = ColLetter(colSeg4);
            string colBlankL = ColLetter(colBlank);
            string colTotalL = ColLetter(colTotal);
            string colCumL = ColLetter(colCum);
            string colResetL = ColLetter(colReset);
            string colIdxL = ColLetter(colIdx);
            string colPrevRstL = ColLetter(colPrevRst);

            for (int i = 0; i < nRows; i++)
            {
                int r = dataStartRow + i;

                // Index: ROW(this) - ROW(header)
                ws.Cells[r, colIdx].Formula = $"=ROW()-ROW({labelHeaderAddr})";

                // Reset: first data row OR "all cat cols blank" (using dynamic column count)
                ws.Cells[r, colReset].Formula =
                    $"=IF(OR(ROW()-ROW({labelHeaderAddr})=1,COUNTBLANK(${colCat1L}{r}:${colCat4L}{r})={colCountExpr}),1,0)";

                // PrevRst: look up in the used range above, but anchor to firstDataLabelAddr
                // height = ROW() - ROW(firstDataLabelAddr) + 1
                string heightExpr = $"ROW()-ROW({firstDataLabelAddr})+1";
                string resetRange = $"OFFSET(${colResetL}${dataStartRow},0,0,{heightExpr},1)";
                string idxRange = $"OFFSET(${colIdxL}${dataStartRow},0,0,{heightExpr},1)";
                ws.Cells[r, colPrevRst].Formula =
                    $"=IF(ROW()-ROW({labelHeaderAddr})=1,0,IFERROR(LOOKUP(2,1/({resetRange}=1),{idxRange}),0))";

                // structural subtotal
                string isSubtotal = $"(COUNTBLANK(${colCat1L}{r}:${colCat4L}{r})={colCountExpr})";

                // rowInside for INDEX(...)
                string rowInside = $"ROW()-ROW({firstDataLabelAddr})+1";

                // build 4 guarded pieces
                string[] segSubtotal = new string[4];
                string[] segCurrent = new string[4];
                for (int c = 1; c <= 4; c++)
                {
                    segSubtotal[c - 1] =
                        $"IF({colCountExpr}>={c},IFERROR(SUMIFS(INDEX({catsDataRange},0,{c}),{idxFullRange},\">=\"&${colPrevRstL}{r},{idxFullRange},\"<=\"&(${colIdxL}{r}-1)),0),0)";
                    segCurrent[c - 1] =
                        $"IF({colCountExpr}>={c},IFERROR(INDEX({catsDataRange},{rowInside},{c}),0),0)";
                }

                // Seg 1..4
                ws.Cells[r, colSeg1].Formula =
                    $"=IF({isSubtotal},{segSubtotal[0]},IF({segCurrent[0]}=0,0,ABS({segCurrent[0]})))";
                ws.Cells[r, colSeg2].Formula =
                    $"=IF({isSubtotal},{segSubtotal[1]},IF({segCurrent[1]}=0,0,ABS({segCurrent[1]})))";
                ws.Cells[r, colSeg3].Formula =
                    $"=IF({isSubtotal},{segSubtotal[2]},IF({segCurrent[2]}=0,0,ABS({segCurrent[2]})))";
                ws.Cells[r, colSeg4].Formula =
                    $"=IF({isSubtotal},{segSubtotal[3]},IF({segCurrent[3]}=0,0,ABS({segCurrent[3]})))";

                // Total from same 4 pieces
                string subtotalSum = $"{segSubtotal[0]}+{segSubtotal[1]}+{segSubtotal[2]}+{segSubtotal[3]}";
                string currentSum = $"{segCurrent[0]}+{segCurrent[1]}+{segCurrent[2]}+{segCurrent[3]}";
                ws.Cells[r, colTotal].Formula =
                    $"=IF({isSubtotal},{subtotalSum},{currentSum})";

                // prev cum using anchored firstDataLabelAddr
                string prevCum = $"OFFSET(${colCumL}${dataStartRow},MAX(ROW()-ROW({firstDataLabelAddr})-1,0),0,1,1)";

                // Cum.
                ws.Cells[r, colCum].Formula =
                    $"=IF(ROW()-ROW({labelHeaderAddr})=1,${colTotalL}{r},IF(${colResetL}{r}=1,${colTotalL}{r},${colTotalL}{r}+{prevCum}))";

                // Blank
                ws.Cells[r, colBlank].Formula =
                    $"=IF(ROW()-ROW({labelHeaderAddr})=1,0,(IF(${colTotalL}{r}>0,{prevCum},{prevCum}+${colTotalL}{r}))*IF(${colResetL}{r}=1,0,1))";
            }

            // number format
            Excel.Range rngNums = ws.Range[
                ws.Cells[dataStartRow, colCat1],
                ws.Cells[dataEndRow, colPrevRst]
            ];
            rngNums.NumberFormat = @"#,##0_);(#,##0);""–""_);_)@";

            // chart 2 rows below used range
            int chartTopRow = dataEndRow + 2;
            int chartLeftCol = c0;

            var chartObjs = (Excel.ChartObjects)ws.ChartObjects();
            Excel.Range chartCell = ws.Cells[chartTopRow, chartLeftCol];
            Excel.ChartObject chObj = chartObjs.Add(chartCell.Left, chartCell.Top, 680, 380);
            Excel.Chart ch = chObj.Chart;
            ch.ChartType = Excel.XlChartType.xlColumnStacked;
            ch.PlotVisibleOnly = false;

            Excel.Range xCats = ws.Range[
                ws.Cells[dataStartRow, colLabel],
                ws.Cells[dataEndRow, colLabel]
            ];

            void AddColSeries(string name, Excel.Range values, bool hide = false)
            {
                var s = ch.SeriesCollection().NewSeries();
                s.Name = name;
                s.Values = values;
                s.XValues = xCats;
                s.ChartType = Excel.XlChartType.xlColumnStacked;
                if (hide)
                {
                    s.Format.Fill.Visible = Office.MsoTriState.msoFalse;
                    s.Format.Line.Visible = Office.MsoTriState.msoFalse;
                }
                Marshal.ReleaseComObject(s);
            }

            Excel.Range rngBlank = ws.Range[
                ws.Cells[dataStartRow, colBlank],
                ws.Cells[dataEndRow, colBlank]
            ];
            Excel.Range rngSeg1 = ws.Range[
                ws.Cells[dataStartRow, colSeg1],
                ws.Cells[dataEndRow, colSeg1]
            ];
            Excel.Range rngSeg2 = ws.Range[
                ws.Cells[dataStartRow, colSeg2],
                ws.Cells[dataEndRow, colSeg2]
            ];
            Excel.Range rngSeg3 = ws.Range[
                ws.Cells[dataStartRow, colSeg3],
                ws.Cells[dataEndRow, colSeg3]
            ];
            Excel.Range rngSeg4 = ws.Range[
                ws.Cells[dataStartRow, colSeg4],
                ws.Cells[dataEndRow, colSeg4]
            ];

            AddColSeries("Blank", rngBlank, hide: true);
            AddColSeries("Cat 1", rngSeg1);
            AddColSeries("Cat 2", rngSeg2);
            AddColSeries("Cat 3", rngSeg3);
            AddColSeries("Cat 4", rngSeg4);

            try
            {
                var cg = (Excel.ChartGroup)ch.ChartGroups(1);
                cg.GapWidth = 50;
                try { cg.Overlap = 100; } catch { }
                Marshal.ReleaseComObject(cg);
            }
            catch { }

            try
            {
                var yAx = ch.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                yAx.TickLabelPosition = Excel.XlTickLabelPosition.xlTickLabelPositionNone;
                yAx.HasMajorGridlines = false;
                yAx.HasMinorGridlines = false;
                yAx.MajorTickMark = Excel.XlTickMark.xlTickMarkNone;
                yAx.MinorTickMark = Excel.XlTickMark.xlTickMarkNone;
                yAx.Format.Line.Visible = Office.MsoTriState.msoFalse;
                Marshal.ReleaseComObject(yAx);
            }
            catch { }

            ch.HasLegend = true;
            ch.Legend.Position = Excel.XlLegendPosition.xlLegendPositionBottom;

            try
            {
                var les = ch.Legend.LegendEntries();
                for (int i = les.Count; i >= 1; i--)
                {
                    var le = les.Item(i);
                    var ser = ch.SeriesCollection(le.Index);
                    string nm = (ser.Name ?? "").Trim();
                    if (nm == "Blank")
                        le.Delete();
                    Marshal.ReleaseComObject(ser);
                    Marshal.ReleaseComObject(le);
                }
                Marshal.ReleaseComObject(les);
            }
            catch { }

            // cleanup
            Marshal.ReleaseComObject(xCats);
            Marshal.ReleaseComObject(rngNums);
            Marshal.ReleaseComObject(ch);
            Marshal.ReleaseComObject(chObj);
            Marshal.ReleaseComObject(chartObjs);
            Marshal.ReleaseComObject(chartCell);
        }
    }
}