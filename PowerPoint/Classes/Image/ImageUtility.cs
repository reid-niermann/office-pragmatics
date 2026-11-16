using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using Microsoft.VisualBasic; // For Interaction.InputBox
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace OfficePragmatics_PowerPoint.Image
{
    public static class ImageUtility
    {

        public static void SortLogos(PowerPoint.Selection selection)
        {
            try
            {
                if (selection.ShapeRange.Count <= 1) return;
                int count = selection.ShapeRange.Count;
                var list = new List<(string alt, int idx)>();
                for (int z = 1; z <= count; z++)
                {
                    var shp = selection.ShapeRange[z];
                    var alt = shp.AlternativeText?.ToLower().Replace("-", " ") ?? "";
                    shp.AlternativeText = alt;
                    list.Add((alt, z));
                }

                for (int i = 0; i < list.Count - 1; i++)
                {
                    for (int j = i + 1; j < list.Count; j++)
                    {
                        if (string.Compare(list[i].alt, list[j].alt, StringComparison.Ordinal) > 0)
                        {
                            int a = list[i].idx, b = list[j].idx;
                            var A = selection.ShapeRange[a];
                            var B = selection.ShapeRange[b];
                            float w1 = A.Width, h1 = A.Height, l1 = A.Left, t1 = A.Top;
                            float w2 = B.Width, h2 = B.Height, l2 = B.Left, t2 = B.Top;

                            A.Left = l2 + (w2 / 2) - (w1 / 2);
                            A.Top = t2 + (h2 / 2) - (h1 / 2);
                            B.Left = l1 + (w1 / 2) - (w2 / 2);
                            B.Top = t1 + (h1 / 2) - (h2 / 2);

                            var temp = list[i];
                            list[i] = list[j];
                            list[j] = temp;
                        }
                    }
                }
            }
            catch { }
        }

        public static void RemoveWhiteBackground(PowerPoint.ShapeRange shapes)
        {
            try
            {
                foreach (PowerPoint.Shape shp in shapes)
                {
                    shp.PictureFormat.TransparentBackground = Office.MsoTriState.msoTrue;
                    shp.PictureFormat.TransparencyColor = 0xFFFFFF; // White
                    shp.Fill.Visible = Office.MsoTriState.msoFalse;
                }
            }
            catch { }
        }

        public static void HarmonizeLogos(PowerPoint.Selection sel)
        {
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes || sel.ShapeRange.Count == 0) return;
            if (!double.TryParse(Interaction.InputBox("Base width (pt):", "Resize Logos", "35"), out double baseWidth)) return;
            const double scale = 0.525;
            foreach (PowerPoint.Shape shp in sel.ShapeRange)
            {
                if (shp.Type == Office.MsoShapeType.msoPicture || shp.Type == Office.MsoShapeType.msoLinkedPicture)
                {
                    double ratio = shp.Width / shp.Height;
                    shp.Width = (float)(Math.Pow(ratio, scale) * baseWidth);
                }
            }
        }
    }
}