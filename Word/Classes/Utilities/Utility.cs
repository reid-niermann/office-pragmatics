using System;
using System.IO;
using System.Linq;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficePragmatics_Word.Utilities
{
    public static class Utility
    {

        public static bool IsValidFileName(string name)
        {
            if (string.IsNullOrEmpty(name)) return false;
            char[] invalid = Path.GetInvalidFileNameChars();
            return name.All(c => !invalid.Contains(c));
        }

        /// <summary>
        /// Gets the line weight of a shape, returning 0 if the shape has no visible line.
        /// </summary>
        public static float GetLineWeight(Word.Shape shape)
        {
            try
            {
                // Some shape types don't have a standard line property
                bool hasStandardLine = shape.Type != Office.MsoShapeType.msoChart &&
                                       shape.Type != Office.MsoShapeType.msoTable;

                if (hasStandardLine && shape.Line.Visible == Office.MsoTriState.msoTrue)
                {
                    return shape.Line.Weight;
                }
            }
            catch
            {
                // If accessing Line property fails, return 0
            }

            return 0f;
        }

        /// <summary>
        /// Parses a distance string that may include units (pt, in, cm, mm).
        /// Returns the value in points.
        /// </summary>
        public static double ParseDistance(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return 0;

            input = input.Trim().ToLower();

            // Try to extract numeric part and unit
            string numPart = "";
            string unitPart = "";

            foreach (char c in input)
            {
                if (char.IsDigit(c) || c == '.' || c == '-' || c == '+')
                    numPart += c;
                else if (char.IsLetter(c))
                    unitPart += c;
            }

            if (!double.TryParse(numPart, out double value))
                return 0;

            // Convert to points based on unit
            switch (unitPart)
            {
                case "in":
                case "\"":
                    return value * 72;  // 72 points per inch
                case "cm":
                    return value * 28.35; // approximately 28.35 points per cm
                case "mm":
                    return value * 2.835; // approximately 2.835 points per mm
                case "pt":
                case "":
                default:
                    return value;  // already in points
            }
        }

        /// <summary>
        /// Converts a distance string to points. Alias for ParseDistance.
        /// </summary>
        public static double ConvertToPoints(string input)
        {
            return ParseDistance(input);
        }

        /// <summary>
        /// Converts a double to float, useful for shape positioning.
        /// </summary>
        public static float F(double value)
        {
            return (float)value;
        }
    }
}