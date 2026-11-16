using System;
using Microsoft.VisualBasic; // For Interaction.InputBox
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using System.Windows.Forms; // For MessageBox

namespace OfficePragmatics_Word.Size
{
    public static class SizeHelper
    {
        #region Core Helpers

        private enum ExtendDirection { Left, Right, Top, Bottom }

        private struct VisualBounds
        {
            public float MinX;
            public float MinY;
            public float MaxX;
            public float MaxY;
            public float CenterX => MinX + (MaxX - MinX) / 2;
            public float CenterY => MinY + (MaxY - MinY) / 2;
            public float VisualWidth => MaxX - MinX;
            public float VisualHeight => MaxY - MinY;
        }

        /// <summary>
        /// Gets the page width for a given shape's section
        /// </summary>
        private static float GetPageWidth(Word.Shape shape)
        {
            try
            {
                Word.Document doc = shape.Application.ActiveDocument;
                Word.Range shapeRange = shape.Anchor;
                return shapeRange.PageSetup.PageWidth;
            }
            catch
            {
                return 612f; // 8.5 inches in points (US Letter)
            }
        }

        /// <summary>
        /// Gets the page height for a given shape's section
        /// </summary>
        private static float GetPageHeight(Word.Shape shape)
        {
            try
            {
                Word.Document doc = shape.Application.ActiveDocument;
                Word.Range shapeRange = shape.Anchor;
                return shapeRange.PageSetup.PageHeight;
            }
            catch
            {
                return 792f; // 11 inches in points (US Letter)
            }
        }

        private static float GetShapeLineWeight(Word.Shape shape)
        {
            bool hasStandardLine = shape.Type != Office.MsoShapeType.msoChart &&
                                   shape.Type != Office.MsoShapeType.msoTable;
            if (hasStandardLine && shape.Line.Visible == Office.MsoTriState.msoTrue)
            {
                return shape.Line.Weight;
            }
            return 0f;
        }

        private static VisualBounds GetVisualBoundingBox(Word.Shape shape, bool includeLineWeight = false)
        {
            float lineHalfWeight = includeLineWeight ? GetShapeLineWeight(shape) / 2 : 0;

            if (shape.Rotation == 0f)
            {
                return new VisualBounds
                {
                    MinX = shape.Left - lineHalfWeight,
                    MinY = shape.Top - lineHalfWeight,
                    MaxX = shape.Left + shape.Width + lineHalfWeight,
                    MaxY = shape.Top + shape.Height + lineHalfWeight
                };
            }

            float width = shape.Width;
            float height = shape.Height;
            double angleRad = -shape.Rotation * Math.PI / 180.0;
            double cosA = Math.Cos(angleRad);
            double sinA = Math.Sin(angleRad);
            float centerX = shape.Left + width / 2;
            float centerY = shape.Top + height / 2;
            var corners = new[] { new { X = -width / 2, Y = -height / 2 }, new { X = width / 2, Y = -height / 2 }, new { X = width / 2, Y = height / 2 }, new { X = -width / 2, Y = height / 2 } };
            float minX = float.MaxValue, minY = float.MaxValue, maxX = float.MinValue, maxY = float.MinValue;
            foreach (var corner in corners)
            {
                float rotatedX = (float)(corner.X * cosA - corner.Y * sinA) + centerX;
                float rotatedY = (float)(corner.X * sinA + corner.Y * cosA) + centerY;
                minX = Math.Min(minX, rotatedX);
                minY = Math.Min(minY, rotatedY);
                maxX = Math.Max(maxX, rotatedX);
                maxY = Math.Max(maxY, rotatedY);
            }
            return new VisualBounds { MinX = minX - lineHalfWeight, MinY = minY - lineHalfWeight, MaxX = maxX + lineHalfWeight, MaxY = maxY + lineHalfWeight };
        }

        private static void SetVisualSize(Word.Shape shp, float targetGeometricVisualWidth, float targetGeometricVisualHeight)
        {
            float centerX = shp.Left + shp.Width / 2;
            float centerY = shp.Top + shp.Height / 2;
            if (shp.Rotation == 0f)
            {
                shp.Width = targetGeometricVisualWidth > 0 ? targetGeometricVisualWidth : 1;
                shp.Height = targetGeometricVisualHeight > 0 ? targetGeometricVisualHeight : 1;
            }
            else
            {
                double angleRad = -shp.Rotation * Math.PI / 180.0;
                double cosA = Math.Abs(Math.Cos(angleRad));
                double sinA = Math.Abs(Math.Sin(angleRad));
                double determinant = cosA * cosA - sinA * sinA;
                if (Math.Abs(determinant) < 1E-6) return;
                float newWidth = (float)((targetGeometricVisualWidth * cosA - targetGeometricVisualHeight * sinA) / determinant);
                float newHeight = (float)((targetGeometricVisualHeight * cosA - targetGeometricVisualWidth * sinA) / determinant);
                if (newWidth > 0) shp.Width = newWidth;
                if (newHeight > 0) shp.Height = newHeight;
            }
            shp.Left = centerX - shp.Width / 2;
            shp.Top = centerY - shp.Height / 2;
        }

        #endregion

        #region Public Sizing Functions

        public static void RatioResize(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            string methodInput = Interaction.InputBox("Choose resize method:\n1 – Width only\n2 – Height only\n3 – Both width & height", "Ratio Resize", "3");
            if (!int.TryParse(methodInput, out int method) || method < 1 || method > 3) method = 3;
            string ratioInput = Interaction.InputBox("Enter resize ratio (e.g. 0.75 for 75 %):", "Ratio Resize", "0.75");
            if (!double.TryParse(ratioInput, out double ratio) || ratio <= 0) ratio = 1.0;

            for (int i = 1; i <= sr.Count; i++)
            {
                Word.Shape shp = sr[i];
                // Scale relative to current size (msoFalse)
                if (method == 1 || method == 3) shp.ScaleWidth((float)ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
                if (method == 2 || method == 3) shp.ScaleHeight((float)ratio, Office.MsoTriState.msoFalse, Office.MsoScaleFrom.msoScaleFromMiddle);
            }
        }

        /// <summary>
        /// Master function to extend a shape, calculating the new target visual dimension
        /// while keeping the other constant, then resizing and repositioning.
        /// </summary>
        private static void ExtendShapeTo(Word.Shape shape, ExtendDirection direction, float targetCoordinate)
        {
            var outerBoundsBefore = GetVisualBoundingBox(shape, true);
            var geometricBoundsBefore = GetVisualBoundingBox(shape, false);
            float shapeLw = GetShapeLineWeight(shape);

            float targetGeometricVisualWidth = geometricBoundsBefore.VisualWidth;
            float targetGeometricVisualHeight = geometricBoundsBefore.VisualHeight;
            float anchorX = 0, anchorY = 0;

            switch (direction)
            {
                case ExtendDirection.Left:
                    if (outerBoundsBefore.MinX <= targetCoordinate) return;
                    targetGeometricVisualWidth = (outerBoundsBefore.MaxX - targetCoordinate) - shapeLw;
                    anchorX = outerBoundsBefore.MaxX; anchorY = outerBoundsBefore.CenterY;
                    break;
                case ExtendDirection.Right:
                    if (outerBoundsBefore.MaxX >= targetCoordinate) return;
                    targetGeometricVisualWidth = (targetCoordinate - outerBoundsBefore.MinX) - shapeLw;
                    anchorX = outerBoundsBefore.MinX; anchorY = outerBoundsBefore.CenterY;
                    break;
                case ExtendDirection.Top:
                    if (outerBoundsBefore.MinY <= targetCoordinate) return;
                    targetGeometricVisualHeight = (outerBoundsBefore.MaxY - targetCoordinate) - shapeLw;
                    anchorX = outerBoundsBefore.CenterX; anchorY = outerBoundsBefore.MaxY;
                    break;
                case ExtendDirection.Bottom:
                    if (outerBoundsBefore.MaxY >= targetCoordinate) return;
                    targetGeometricVisualHeight = (targetCoordinate - outerBoundsBefore.MinY) - shapeLw;
                    anchorX = outerBoundsBefore.CenterX; anchorY = outerBoundsBefore.MinY;
                    break;
            }

            if (targetGeometricVisualWidth <= 0 || targetGeometricVisualHeight <= 0) return;

            SetVisualSize(shape, targetGeometricVisualWidth, targetGeometricVisualHeight);

            var outerBoundsAfter = GetVisualBoundingBox(shape, true);
            float newAnchorX = 0, newAnchorY = 0;

            switch (direction)
            {
                case ExtendDirection.Left: newAnchorX = outerBoundsAfter.MaxX; newAnchorY = outerBoundsAfter.CenterY; break;
                case ExtendDirection.Right: newAnchorX = outerBoundsAfter.MinX; newAnchorY = outerBoundsAfter.CenterY; break;
                case ExtendDirection.Top: newAnchorX = outerBoundsAfter.CenterX; newAnchorY = outerBoundsAfter.MaxY; break;
                case ExtendDirection.Bottom: newAnchorX = outerBoundsAfter.CenterX; newAnchorY = outerBoundsAfter.MinY; break;
            }

            shape.Left += (anchorX - newAnchorX);
            shape.Top += (anchorY - newAnchorY);
        }

        public static void ExtendLeft(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            float targetMinX = (sr.Count > 1) ? GetVisualBoundingBox(sr[sr.Count], true).MinX : 0;
            for (int i = 1; i <= sr.Count - (sr.Count > 1 ? 1 : 0); i++) { ExtendShapeTo(sr[i], ExtendDirection.Left, targetMinX); }
        }

        public static void ExtendRight(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            float targetMaxX = (sr.Count > 1) ? GetVisualBoundingBox(sr[sr.Count], true).MaxX : GetPageWidth(sr[1]);
            for (int i = 1; i <= sr.Count - (sr.Count > 1 ? 1 : 0); i++) { ExtendShapeTo(sr[i], ExtendDirection.Right, targetMaxX); }
        }

        public static void ExtendTop(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            float targetMinY = (sr.Count > 1) ? GetVisualBoundingBox(sr[sr.Count], true).MinY : 0;
            for (int i = 1; i <= sr.Count - (sr.Count > 1 ? 1 : 0); i++) { ExtendShapeTo(sr[i], ExtendDirection.Top, targetMinY); }
        }

        public static void ExtendBottom(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            float targetMaxY = (sr.Count > 1) ? GetVisualBoundingBox(sr[sr.Count], true).MaxY : GetPageHeight(sr[1]);
            for (int i = 1; i <= sr.Count - (sr.Count > 1 ? 1 : 0); i++) { ExtendShapeTo(sr[i], ExtendDirection.Bottom, targetMaxY); }
        }

        public static void FitToHorizontal(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            if (sr.Count == 1)
            {
                Word.Shape shp = sr[1];
                string gapInput = Interaction.InputBox("Enter horizontal margin (pts):", "Fit Horizontally", "10");
                if (!double.TryParse(gapInput, out double margin)) margin = 10;
                float targetOuterVisualWidth = GetPageWidth(shp) - (float)(2 * margin);
                if (targetOuterVisualWidth <= 0) return;
                float shpLw = GetShapeLineWeight(shp);
                float targetGeometricVisualWidth = targetOuterVisualWidth - shpLw;
                SetVisualSize(shp, targetGeometricVisualWidth, GetVisualBoundingBox(shp, false).VisualHeight);
                var newBounds = GetVisualBoundingBox(shp, true);
                shp.Left += ((float)margin + shpLw / 2 - newBounds.MinX);
            }
            else
            {
                Word.Shape anchor = sr[sr.Count];
                var anchorBounds = GetVisualBoundingBox(anchor, true);
                string gapInput = Interaction.InputBox("Enter horizontal gap (pts):", "Fit Horizontally", "10");
                if (!double.TryParse(gapInput, out double gap)) gap = 10;
                int numShapesToFit = sr.Count - 1;
                if (numShapesToFit <= 0) return;
                float availableWidth = anchorBounds.VisualWidth - (float)((numShapesToFit - 1) * gap);
                if (availableWidth <= 0) { MessageBox.Show("Not enough space.", "Fit Error"); return; }
                float widthPerShape = availableWidth / numShapesToFit;
                float currentX = anchorBounds.MinX;
                for (int i = 1; i <= numShapesToFit; i++)
                {
                    Word.Shape shp = sr[i];
                    float shpLw = GetShapeLineWeight(shp);

                    // Correctly calculate the target geometric width
                    float targetGeometricVisualWidth = widthPerShape - shpLw;
                    SetVisualSize(shp, targetGeometricVisualWidth, GetVisualBoundingBox(shp, false).VisualHeight);

                    // Correctly position the shape
                    var newBounds = GetVisualBoundingBox(shp, true);
                    shp.Left += (currentX - newBounds.MinX);

                    currentX += widthPerShape + (float)gap;
                }
            }
        }

        public static void FitToVertical(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count == 0) return;
            if (sr.Count == 1)
            {
                Word.Shape shp = sr[1];
                string gapInput = Interaction.InputBox("Enter vertical margin (pts):", "Fit Vertically", "10");
                if (!double.TryParse(gapInput, out double margin)) margin = 10;
                float targetOuterVisualHeight = GetPageHeight(shp) - (float)(2 * margin);
                if (targetOuterVisualHeight <= 0) return;
                float shpLw = GetShapeLineWeight(shp);
                float targetGeometricVisualHeight = targetOuterVisualHeight - shpLw;
                SetVisualSize(shp, GetVisualBoundingBox(shp, false).VisualWidth, targetGeometricVisualHeight);
                var newBounds = GetVisualBoundingBox(shp, true);
                shp.Top += ((float)margin + shpLw / 2 - newBounds.MinY);
            }
            else
            {
                Word.Shape anchor = sr[sr.Count];
                var anchorBounds = GetVisualBoundingBox(anchor, true);
                string gapInput = Interaction.InputBox("Enter vertical gap (pts):", "Fit Vertically", "10");
                if (!double.TryParse(gapInput, out double gap)) gap = 10;
                int numShapesToFit = sr.Count - 1;
                if (numShapesToFit <= 0) return;
                float availableHeight = anchorBounds.VisualHeight - (float)((numShapesToFit - 1) * gap);
                if (availableHeight <= 0) { MessageBox.Show("Not enough space.", "Fit Error"); return; }
                float heightPerShape = availableHeight / numShapesToFit;
                float currentY = anchorBounds.MinY;
                for (int i = 1; i <= numShapesToFit; i++)
                {
                    Word.Shape shp = sr[i];
                    float shpLw = GetShapeLineWeight(shp);

                    // Correctly calculate the target geometric height
                    float targetGeometricVisualHeight = heightPerShape - shpLw;
                    SetVisualSize(shp, GetVisualBoundingBox(shp, false).VisualWidth, targetGeometricVisualHeight);

                    // Correctly position the shape
                    var newBounds = GetVisualBoundingBox(shp, true);
                    shp.Top += (currentY - newBounds.MinY);

                    currentY += heightPerShape + (float)gap;
                }
            }
        }

        public static void SameWidth(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count < 2) return;
            float targetOuterVisualWidth = GetVisualBoundingBox(sr[sr.Count], true).VisualWidth;
            for (int i = 1; i < sr.Count; i++)
            {
                var shp = sr[i];
                float shpLw = GetShapeLineWeight(shp);
                float targetGeometricVisualWidth = targetOuterVisualWidth - shpLw;
                if (targetGeometricVisualWidth > 0)
                {
                    SetVisualSize(shp, targetGeometricVisualWidth, GetVisualBoundingBox(shp, false).VisualHeight);
                }
            }
        }

        public static void SameHeight(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count < 2) return;
            float targetOuterVisualHeight = GetVisualBoundingBox(sr[sr.Count], true).VisualHeight;
            for (int i = 1; i < sr.Count; i++)
            {
                var shp = sr[i];
                float shpLw = GetShapeLineWeight(shp);
                float targetGeometricVisualHeight = targetOuterVisualHeight - shpLw;
                if (targetGeometricVisualHeight > 0)
                {
                    SetVisualSize(shp, GetVisualBoundingBox(shp, false).VisualWidth, targetGeometricVisualHeight);
                }
            }
        }

        public static void SameSize(Word.ShapeRange sr)
        {
            if (sr == null || sr.Count < 2) return;
            var anchorBounds = GetVisualBoundingBox(sr[sr.Count], true);
            float targetOuterVisualWidth = anchorBounds.VisualWidth;
            float targetOuterVisualHeight = anchorBounds.VisualHeight;
            for (int i = 1; i < sr.Count; i++)
            {
                var shp = sr[i];
                float shpLw = GetShapeLineWeight(shp);
                float targetGeometricVisualWidth = targetOuterVisualWidth - shpLw;
                float targetGeometricVisualHeight = targetOuterVisualHeight - shpLw;
                if (targetGeometricVisualWidth > 0 && targetGeometricVisualHeight > 0)
                {
                    SetVisualSize(shp, targetGeometricVisualWidth, targetGeometricVisualHeight);
                }
            }
        }

        #endregion
    }
}