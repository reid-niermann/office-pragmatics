using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using OfficePragmatics_PowerPoint.Settings; // For SettingsManager

namespace OfficePragmatics_PowerPoint.Alignment
{
    public static class AlignmentHelper
    {
        /// <summary>
        /// A struct to hold the calculated visual bounds of a shape.
        /// </summary>
        private struct VisualBounds
        {
            public float MinX;
            public float MinY;
            public float MaxX;
            public float MaxY;
            public float CenterX => MinX + (MaxX - MinX) / 2;
            public float CenterY => MinY + (MaxY - MinY) / 2;
        }

        /// <summary>
        /// Calculates the visual bounding box of a shape, accounting for rotation.
        /// </summary>
        /// <param name="shape">The shape to measure.</param>
        /// <param name="includeLineWeight">If true, the bounding box will include the shape's outline.</param>
        /// <returns>A VisualBounds struct with the min/max coordinates.</returns>
        private static VisualBounds GetVisualBoundingBox(PowerPoint.Shape shape, bool includeLineWeight = false)
        {
            // A shape is not considered to have a standard line if it's a Chart, Table, or certain other object types.
            bool hasStandardLine = shape.Type != Office.MsoShapeType.msoChart &&
                                   shape.Type != Office.MsoShapeType.msoTable;

            // If there's no rotation, the calculation is simple.
            if (shape.Rotation == 0f)
            {
                var simpleBounds = new VisualBounds
                {
                    MinX = shape.Left,
                    MinY = shape.Top,
                    MaxX = shape.Left + shape.Width,
                    MaxY = shape.Top + shape.Height
                };

                // Only access .Line property if the shape type supports it.
                if (includeLineWeight && hasStandardLine && shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    float weight = shape.Line.Weight / 2;
                    simpleBounds.MinX -= weight;
                    simpleBounds.MinY -= weight;
                    simpleBounds.MaxX += weight;
                    simpleBounds.MaxY += weight;
                }
                return simpleBounds;
            }

            // Get shape properties
            float width = shape.Width;
            float height = shape.Height;
            float left = shape.Left;
            float top = shape.Top;

            // PowerPoint's rotation is clockwise, so we use a negative angle for standard trigonometric functions.
            double angleRad = -shape.Rotation * Math.PI / 180.0;
            double cosA = Math.Cos(angleRad);
            double sinA = Math.Sin(angleRad);

            // Calculate center point (pivot for rotation)
            float centerX = left + width / 2;
            float centerY = top + height / 2;

            // Define the 4 corners of the unrotated box relative to its center
            var corners = new[]
            {
                new { X = -width / 2, Y = -height / 2 },
                new { X =  width / 2, Y = -height / 2 },
                new { X =  width / 2, Y =  height / 2 },
                new { X = -width / 2, Y =  height / 2 }
            };

            float minX = float.MaxValue, minY = float.MaxValue, maxX = float.MinValue, maxY = float.MinValue;

            // Rotate each corner and find the min/max extents
            foreach (var corner in corners)
            {
                float rotatedX = (float)(corner.X * cosA - corner.Y * sinA) + centerX;
                float rotatedY = (float)(corner.X * sinA + corner.Y * cosA) + centerY;

                minX = Math.Min(minX, rotatedX);
                minY = Math.Min(minY, rotatedY);
                maxX = Math.Max(maxX, rotatedX);
                maxY = Math.Max(maxY, rotatedY);
            }

            var bounds = new VisualBounds { MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY };

            // Only access .Line property if the shape type supports it.
            if (includeLineWeight && hasStandardLine && shape.Line.Visible == Office.MsoTriState.msoTrue)
            {
                float weight = shape.Line.Weight / 2;
                bounds.MinX -= weight;
                bounds.MinY -= weight;
                bounds.MaxX += weight;
                bounds.MaxY += weight;
            }

            return bounds;
        }


        public static void AlignLeft(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count == 0) return;

                bool alignToBorder = SettingsForm.CurrentSettings.AlignToBorder;

                // Single shape - align to slide left edge
                if (sel.ShapeRange.Count == 1)
                {
                    var shp = sel.ShapeRange[1];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    shp.Left += (0 - visualBounds.MinX); // Target X is 0
                    return;
                }

                // Multiple shapes - align to last selected shape
                int count = sel.ShapeRange.Count;
                var refShp = sel.ShapeRange[count];
                var refVisualBounds = GetVisualBoundingBox(refShp, alignToBorder);
                float referencePointX = refVisualBounds.MinX;

                for (int i = 1; i < count; i++)
                {
                    var shp = sel.ShapeRange[i];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    float moveDistanceX = referencePointX - visualBounds.MinX;
                    shp.Left += moveDistanceX;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AlignLeft: {ex.Message}"); }
        }

        public static void AlignRight(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count == 0) return;

                bool alignToBorder = SettingsForm.CurrentSettings.AlignToBorder;
                var slide = (PowerPoint.Slide)sel.ShapeRange[1].Parent;
                float slideWidth = slide.Parent.PageSetup.SlideWidth;

                // Single shape - align to slide right edge
                if (sel.ShapeRange.Count == 1)
                {
                    var shp = sel.ShapeRange[1];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    shp.Left += (slideWidth - visualBounds.MaxX);
                    return;
                }

                // Multiple shapes - align to last selected shape
                int count = sel.ShapeRange.Count;
                var refShp = sel.ShapeRange[count];
                var refVisualBounds = GetVisualBoundingBox(refShp, alignToBorder);
                float referencePointX = refVisualBounds.MaxX;

                for (int i = 1; i < count; i++)
                {
                    var shp = sel.ShapeRange[i];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    float moveDistanceX = referencePointX - visualBounds.MaxX;
                    shp.Left += moveDistanceX;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AlignRight: {ex.Message}"); }
        }

        public static void AlignCenter(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count == 0) return;
                // Align to border is ignored for center/middle alignment

                var slide = (PowerPoint.Slide)sel.ShapeRange[1].Parent;
                float slideCenterX = slide.Parent.PageSetup.SlideWidth / 2;

                // Single shape - align to slide center
                if (sel.ShapeRange.Count == 1)
                {
                    var shp = sel.ShapeRange[1];
                    var visualBounds = GetVisualBoundingBox(shp);
                    shp.Left += (slideCenterX - visualBounds.CenterX);
                    return;
                }

                // Multiple shapes - align to last selected shape's center
                int count = sel.ShapeRange.Count;
                var refShp = sel.ShapeRange[count];
                var refVisualBounds = GetVisualBoundingBox(refShp);
                float referencePointX = refVisualBounds.CenterX;

                for (int i = 1; i < count; i++)
                {
                    var shp = sel.ShapeRange[i];
                    var visualBounds = GetVisualBoundingBox(shp);
                    float moveDistanceX = referencePointX - visualBounds.CenterX;
                    shp.Left += moveDistanceX;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AlignCenter: {ex.Message}"); }
        }

        public static void AlignTop(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count == 0) return;

                bool alignToBorder = SettingsForm.CurrentSettings.AlignToBorder;

                // Single shape - align to slide top edge
                if (sel.ShapeRange.Count == 1)
                {
                    var shp = sel.ShapeRange[1];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    shp.Top += (0 - visualBounds.MinY); // Target Y is 0
                    return;
                }

                // Multiple shapes - align to last selected shape
                int count = sel.ShapeRange.Count;
                var refShp = sel.ShapeRange[count];
                var refVisualBounds = GetVisualBoundingBox(refShp, alignToBorder);
                float referencePointY = refVisualBounds.MinY;

                for (int i = 1; i < count; i++)
                {
                    var shp = sel.ShapeRange[i];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    float moveDistanceY = referencePointY - visualBounds.MinY;
                    shp.Top += moveDistanceY;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AlignTop: {ex.Message}"); }
        }

        public static void AlignBottom(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count == 0) return;

                bool alignToBorder = SettingsForm.CurrentSettings.AlignToBorder;
                var slide = (PowerPoint.Slide)sel.ShapeRange[1].Parent;
                float slideHeight = slide.Parent.PageSetup.SlideHeight;

                // Single shape - align to slide bottom edge
                if (sel.ShapeRange.Count == 1)
                {
                    var shp = sel.ShapeRange[1];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    shp.Top += (slideHeight - visualBounds.MaxY);
                    return;
                }

                // Multiple shapes - align to last selected shape
                int count = sel.ShapeRange.Count;
                var refShp = sel.ShapeRange[count];
                var refVisualBounds = GetVisualBoundingBox(refShp, alignToBorder);
                float referencePointY = refVisualBounds.MaxY;

                for (int i = 1; i < count; i++)
                {
                    var shp = sel.ShapeRange[i];
                    var visualBounds = GetVisualBoundingBox(shp, alignToBorder);
                    float moveDistanceY = referencePointY - visualBounds.MaxY;
                    shp.Top += moveDistanceY;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AlignBottom: {ex.Message}"); }
        }

        public static void AlignMiddle(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.ShapeRange.Count == 0) return;
                // Align to border is ignored for center/middle alignment

                var slide = (PowerPoint.Slide)sel.ShapeRange[1].Parent;
                float slideCenterY = slide.Parent.PageSetup.SlideHeight / 2;

                // Single shape - align to slide middle
                if (sel.ShapeRange.Count == 1)
                {
                    var shp = sel.ShapeRange[1];
                    var visualBounds = GetVisualBoundingBox(shp);
                    shp.Top += (slideCenterY - visualBounds.CenterY);
                    return;
                }

                // Multiple shapes - align to last selected shape's middle
                int count = sel.ShapeRange.Count;
                var refShp = sel.ShapeRange[count];
                var refVisualBounds = GetVisualBoundingBox(refShp);
                float referencePointY = refVisualBounds.CenterY;

                for (int i = 1; i < count; i++)
                {
                    var shp = sel.ShapeRange[i];
                    var visualBounds = GetVisualBoundingBox(shp);
                    float moveDistanceY = referencePointY - visualBounds.CenterY;
                    shp.Top += moveDistanceY;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"AlignMiddle: {ex.Message}"); }
        }

        public static void SameRotation(PowerPoint.Selection sel)
        {
            try
            {
                // Must have more than one shape to match rotation
                if (sel.ShapeRange.Count < 2) return;

                // Get the rotation of the last selected shape (the reference shape)
                int count = sel.ShapeRange.Count;
                float referenceRotation = sel.ShapeRange[count].Rotation;

                // Apply the reference rotation to all other shapes in the selection
                for (int i = 1; i < count; i++)
                {
                    sel.ShapeRange[i].Rotation = referenceRotation;
                }
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"SameRotation: {ex.Message}"); }
        }

        public static void DistributeHorizontal(PowerPoint.ShapeRange sr)
        {
            try
            {
                // PowerPoint's built-in Distribute handles rotated objects correctly. No change needed.
                sr.Distribute(Office.MsoDistributeCmd.msoDistributeHorizontally, Office.MsoTriState.msoFalse);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"DistributeHorizontal: {ex.Message}"); }
        }

        public static void DistributeVertical(PowerPoint.ShapeRange sr)
        {
            try
            {
                // PowerPoint's built-in Distribute handles rotated objects correctly. No change needed.
                sr.Distribute(Office.MsoDistributeCmd.msoDistributeVertically, Office.MsoTriState.msoFalse);
            }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"DistributeVertical: {ex.Message}"); }
        }
    }
}