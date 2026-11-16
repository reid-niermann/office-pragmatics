using System;
using System.IO;
using System.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using OfficePragmatics_PowerPoint.Settings;

namespace OfficePragmatics_PowerPoint.Utilities
{
    public static class Utility
    {
        public static float F(double d) => (float)d;

        public static double GetLineWeight(PowerPoint.Shape shp)
        {
            if (shp == null) return 0;

            // Directly use shp.Line which is PowerPoint.LineFormat
            try
            {
                // Check if Line property itself is accessible and if it's visible
                if (shp.Line == null || shp.Line.Visible != Office.MsoTriState.msoTrue) // Office.MsoTriState is fine
                    return 0;

                if (!SettingsForm.CurrentSettings.AlignToBorder || shp.HasTable == Office.MsoTriState.msoTrue)
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

        public static double ConvertToPoints(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0;
            s = s.Trim();
            double val;
            if (s.EndsWith("\""))
            {
                if (double.TryParse(s.Substring(0, s.Length - 1), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out val)) return val * 72;
            }
            else if (s.EndsWith("in", StringComparison.OrdinalIgnoreCase))
            {
                if (double.TryParse(s.Substring(0, s.Length - 2), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out val)) return val * 72;
            }
            else
            {
                if (double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out val)) return val;
            }
            return 0;
        }

        public static bool IsValidFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            char[] invalid = Path.GetInvalidFileNameChars();
            return name.All(c => !invalid.Contains(c));
        }
    }
}