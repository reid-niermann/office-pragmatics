using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace OfficePragmatics_Word.Text
{
    public static class TextUtility
    {

        /* ─────────────────────────────────────────────────────────────────────────
         * CONVERT NUMBER TO WORDS (Word)
         * - If text is selected: converts all numbers in selection to "words (number)"
         * - If no selection: converts the number at/near the cursor position
         * ───────────────────────────────────────────────────────────────────────── */
        public static void ConvertNumberToWords(Word.Application app)
        {
            Word.Selection sel = null;
            try
            {
                sel = app.ActiveWindow?.Selection;
                if (sel == null) return;

                // If we have selected text, process all numbers in the selection
                if (sel.Type == Word.WdSelectionType.wdSelectionNormal && sel.Range.Text.Length > 0)
                {
                    ConvertNumbersInRange(sel.Range);
                }
                // If insertion point, try to find and convert number at cursor
                else if (sel.Type == Word.WdSelectionType.wdSelectionIP)
                {
                    ConvertNumberAtCursor(sel);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ConvertNumberToWords: {ex.Message}");
                MessageBox.Show($"Error converting numbers: {ex.Message}", "Convert Number to Words",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ConvertNumbersInRange(Word.Range range)
        {
            if (range == null) return;

            try
            {
                string text = range.Text;
                if (string.IsNullOrEmpty(text)) return;

                // Find all numbers in the text (including decimals and commas)
                var regex = new System.Text.RegularExpressions.Regex(@"\b(\d{1,3}(,\d{3})*(\.\d+)?|\d+(\.\d+)?)\b");
                var matches = regex.Matches(text);

                if (matches.Count == 0)
                {
                    MessageBox.Show("No numbers found in selection.", "Convert Number to Words");
                    return;
                }

                // Process matches in reverse order to maintain position accuracy
                for (int i = matches.Count - 1; i >= 0; i--)
                {
                    var match = matches[i];
                    string numStr = match.Value;

                    // Parse the number (remove commas for parsing)
                    string cleanNum = numStr.Replace(",", "");
                    if (double.TryParse(cleanNum, out double number))
                    {
                        string words = NumberToWords(number);
                        string replacement = $"{words} ({numStr})";

                        // Create a range for this specific number
                        Word.Range numRange = range.Duplicate;
                        numRange.SetRange(range.Start + match.Index, range.Start + match.Index + match.Length);
                        numRange.Text = replacement;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ConvertNumbersInRange: {ex.Message}");
                throw;
            }
        }

        private static void ConvertNumberAtCursor(Word.Selection sel)
        {
            try
            {
                int cursorPos = sel.Start;
                Word.Range range = sel.Range;

                // Get surrounding text to search for the complete number
                Word.Range searchRange = range.Duplicate;
                searchRange.MoveStart(Word.WdUnits.wdCharacter, -50);
                searchRange.MoveEnd(Word.WdUnits.wdCharacter, 50);

                string surroundingText = searchRange.Text;
                if (string.IsNullOrEmpty(surroundingText)) return;

                int relativePos = cursorPos - searchRange.Start;

                // Find all numbers in the surrounding text
                var regex = new System.Text.RegularExpressions.Regex(@"\b(\d{1,3}(,\d{3})*(\.\d+)?|\d+(\.\d+)?)\b");
                var matches = regex.Matches(surroundingText);

                // Find which number contains the cursor
                foreach (System.Text.RegularExpressions.Match match in matches)
                {
                    int matchStart = match.Index;
                    int matchEnd = match.Index + match.Length;

                    // Check if cursor is within this number
                    if (relativePos >= matchStart && relativePos <= matchEnd)
                    {
                        string numStr = match.Value;
                        string cleanNum = numStr.Replace(",", "");

                        if (double.TryParse(cleanNum, out double number))
                        {
                            // Select the entire number in the document
                            Word.Range numRange = range.Duplicate;
                            numRange.SetRange(searchRange.Start + matchStart, searchRange.Start + matchEnd);
                            numRange.Select();

                            string words = NumberToWords(number);
                            string replacement = $"{words} ({numStr})";
                            numRange.Text = replacement;
                            return;
                        }
                    }
                }

                MessageBox.Show("Cursor is not positioned at a number.", "Convert Number to Words");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ConvertNumberAtCursor: {ex.Message}");
                throw;
            }
        }

        private static string NumberToWords(double number)
        {
            try
            {
                // Handle decimals
                if (number != Math.Floor(number))
                {
                    string[] parts = number.ToString("F2").Split('.');
                    long wholePart = long.Parse(parts[0].Replace(",", ""));
                    int decimalPart = int.Parse(parts[1]);

                    string wholeWords = ConvertWholeNumberToWords(wholePart);
                    string decimalWords = ConvertWholeNumberToWords(decimalPart);
                    return $"{wholeWords} and {decimalPart}/100";
                }

                // Handle whole numbers
                long wholeNumber = (long)number;
                return ConvertWholeNumberToWords(wholeNumber);
            }
            catch
            {
                return number.ToString();
            }
        }

        private static string ConvertWholeNumberToWords(long number)
        {
            if (number == 0) return "zero";
            if (number < 0) return "negative " + ConvertWholeNumberToWords(Math.Abs(number));

            string[] ones = { "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
                      "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
                      "seventeen", "eighteen", "nineteen" };

            string[] tens = { "", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety" };

            string[] thousands = { "", "thousand", "million", "billion", "trillion" };

            if (number < 20)
                return ones[number];

            if (number < 100)
            {
                int tensPart = (int)(number / 10);
                int onesPart = (int)(number % 10);
                return tens[tensPart] + (onesPart > 0 ? "-" + ones[onesPart] : "");
            }

            if (number < 1000)
            {
                int hundreds = (int)(number / 100);
                int remainder = (int)(number % 100);
                string result = ones[hundreds] + " hundred";
                if (remainder > 0)
                    result += " " + ConvertWholeNumberToWords(remainder);
                return result;
            }

            // Handle thousands, millions, billions, etc.
            for (int i = thousands.Length - 1; i >= 0; i--)
            {
                long divisor = (long)Math.Pow(1000, i);
                if (number >= divisor)
                {
                    long quotient = number / divisor;
                    long remainder = number % divisor;
                    string result = ConvertWholeNumberToWords(quotient) + " " + thousands[i];
                    if (remainder > 0)
                        result += " " + ConvertWholeNumberToWords(remainder);
                    return result.Trim();
                }
            }

            return number.ToString();
        }

        private static string SafeName(Word.Shape s)
        {
            try { return s?.Name ?? "(unnamed)"; } catch { return "(unnamed)"; }
        }

        /* ─────────────────────────────────────────────────────────────────────────
         * CHANGE PARAGRAPH SPACING (Word)
         * - If text is selected: modify Selection.Range.ParagraphFormat
         * - If shapes selected: modify Shape.TextFrame.TextRange.ParagraphFormat
         * ───────────────────────────────────────────────────────────────────────── */
        public static void ChangeParagraphSpacing(Word.Application app, float deltaPts)
        {
            Word.Selection sel = null;
            try
            {
                sel = app.ActiveWindow?.Selection;
                if (sel == null) return;

                // Text selection (normal or insertion point)
                if (sel.Type == Word.WdSelectionType.wdSelectionNormal ||
                    sel.Type == Word.WdSelectionType.wdSelectionIP ||
                    sel.Type == Word.WdSelectionType.wdSelectionBlock)
                {
                    AdjustRange(sel.Range, deltaPts);
                    return;
                }

                // Shapes (floating)
                if (sel.Type == Word.WdSelectionType.wdSelectionShape && sel.ShapeRange != null)
                {
                    var sr = sel.ShapeRange;
                    for (int i = 1; i <= sr.Count; i++)
                    {
                        var shp = sr[i];
                        try
                        {
                            var tf = shp.TextFrame;
                            if (tf != null && tf.HasText != 0)  // Word returns -1 for True
                            {
                                Word.Range tr = null;
                                try
                                {
                                    tr = tf.TextRange;
                                    if (tr != null) AdjustRange(tr, deltaPts);
                                }
                                finally
                                {
                                    if (tr != null && System.Runtime.InteropServices.Marshal.IsComObject(tr))
                                        System.Runtime.InteropServices.Marshal.ReleaseComObject(tr);
                                }
                            }
                        }
                        catch (Exception exShape)
                        {
                            System.Diagnostics.Debug.WriteLine($"ChangeParagraphSpacing: shape '{SafeName(shp)}' error: {exShape.Message}");
                        }
                    }
                    return;
                }

                // Inline shape selected: nothing to adjust unless it contains text (rare)
                if (sel.Type == Word.WdSelectionType.wdSelectionInlineShape)
                {
                    // Best effort: adjust the containing paragraph(s)
                    AdjustRange(sel.Range, deltaPts);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ChangeParagraphSpacing: {ex.Message}");
            }
        }

        private static void AdjustRange(Word.Range rng, float deltaPts)
        {
            if (rng == null) return;
            Word.ParagraphFormat pf = null;
            try
            {
                pf = rng.ParagraphFormat;
                float before = (float)pf.SpaceBefore + deltaPts;
                pf.SpaceBefore = before > 0 ? before : 0;

                float after = (float)pf.SpaceAfter + deltaPts;
                pf.SpaceAfter = after > 0 ? after : 0;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adjusting paragraph spacing: {ex.Message}");
            }
            finally
            {
                if (pf != null && System.Runtime.InteropServices.Marshal.IsComObject(pf))
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pf);
            }
        }
    }
}
