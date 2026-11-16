using System;
using System.IO;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using OfficePragmatics_Excel.Settings;

namespace OfficePragmatics_Excel.Utilities
{
    public static class Utility
    {
        public static float F(double d) => (float)d;

        public static double GetLineWeight(Excel.Shape shp)
        {
            if (shp == null) return 0;

            // Excel.Shape also has Line property like PowerPoint
            try
            {
                // Check if Line property itself is accessible and if it's visible
                if (shp.Line == null || shp.Line.Visible != Office.MsoTriState.msoTrue)
                    return 0;

                if (!SettingsForm.CurrentSettings.AlignToBorder || IsTableShape(shp))
                    return 0;

                return Math.Max(shp.Line.Weight, 0);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                // This can happen if .Line or .Line.Visible or .Line.Weight is not accessible for the shape type
                System.Diagnostics.Debug.WriteLine($"GetLineWeight COMException for shape {shp.Name}: {ex.Message}");
                return 0;
            }
            catch (Exception ex) // Catch other potential errors
            {
                System.Diagnostics.Debug.WriteLine($"GetLineWeight Exception for shape {shp.Name}: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// Determines if the shape is a table or contains table-like content in Excel context
        /// </summary>
        private static bool IsTableShape(Excel.Shape shp)
        {
            try
            {
                // In Excel, check if the shape type indicates it's a table-related shape
                // Excel doesn't have HasTable like PowerPoint, so we check the shape type instead
                switch (shp.Type)
                {
                    case Office.MsoShapeType.msoTable:
                        return true;
                    case Office.MsoShapeType.msoOLEControlObject:
                        // Could be an embedded table object
                        return true;
                    case Office.MsoShapeType.msoChart:
                        // Charts often contain tabular data
                        return true;
                    default:
                        return false;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"IsTableShape Exception for shape {shp.Name}: {ex.Message}");
                return false;
            }
        }
    }
}