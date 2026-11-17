// FormulaTools.cs – full version for C# 7.2 / VSTO Excel
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficePragmatics_Excel.Classes
{
    internal static class FormulaTools
    {
        /* ─────────────────────────────  PUBLIC API  ───────────────────────────── */

        // Magnitude now moves in thousands (10^3) each click
        public static void Magnitude_Plus() => ApplyToSelection(s => AdjustPowerOfTen(s, +3));
        public static void Magnitude_Minus() => ApplyToSelection(s => AdjustPowerOfTen(s, -3));

        public static void FlipSign() => ApplyToSelection(ToggleTopLevelNegation);
        public static void ErrorWrapNA() => ApplyToSelection(s => WrapInIfError(s, "\"NA\""));
        public static void ErrorWrapZero() => ApplyToSelection(s => WrapInIfError(s, "0"));
        public static void ConvertToLet() => ApplyToSelection(ConvertFormulaToLet);
        public static void SolveFormula() => ApplyToSelection(ResolveCommonFunctions);
        public static void AnchorCycle() => ApplyToSelection(CycleAnchors);

        /* ───────────────────────────  CORE UTILITIES  ────────────────────────── */

        public static void FormulaWrap()
        {
            using (var dlg = new FormulaWrapDialog())
            {
                if (dlg.ShowDialog() != DialogResult.OK) return;

                string begin = dlg.BeginText ?? "";
                string end = dlg.EndText ?? "";

                if (begin.Length == 0 && end.Length == 0)
                {
                    MessageBox.Show("Both beginning and ending strings are blank.", "Formula Wrap",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ApplyToSelection(f => WrapFormulaBody(f, begin, end));
            }
        }

        /* ───────────── helper : wrap a formula string ───────────── */
        private static string WrapFormulaBody(string fullFormula, string begin, string end)
        {
            if (!fullFormula.StartsWith("=")) return fullFormula;          // not a formula
            string core = fullFormula.Substring(1);

            // simple guard against double-wrapping
            if (core.StartsWith(begin, StringComparison.Ordinal) &&
                core.EndsWith(end, StringComparison.Ordinal))
                return fullFormula;

            return "=" + begin + core + end;
        }

        private static void ApplyToSelection(Func<string, string> transformer)
        {
            var app = Globals.ThisAddIn.Application;
            var sel = app.Selection as Excel.Range;
            if (sel == null) { MessageBox.Show("Select at least one cell first."); return; }

            foreach (Excel.Range cell in sel.Cells)
            {
                string f = cell.Formula?.ToString() ?? string.Empty;
                if (!f.StartsWith("=")) f = "=" + f;          // normalise
                string updated = transformer(f);
                updated = NormalizeTopLevelMultipliers(updated);

                if (!updated.Equals(f, StringComparison.Ordinal))
                    cell.Formula = updated;
            }
        }

        /* ───────────── 1.  CONVERT DUPLICATES → LET( )  ───────────── */
        private static string ConvertFormulaToLet(string formula)
        {
            // ❶ Remove leading “=”
            string body = formula.StartsWith("=") ? formula.Substring(1) : formula;

            // ❷ Collect every candidate sub-expression --------------------------------
            //     • all cell / range references
            //     • every balanced function call (with full nesting)
            // --------------------------------------------------------------------------
            var counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            // ——— a) cell / range tokens
            const string cellPat =
                @"\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?";            // A1 / $A$1 / A1:B2
            foreach (Match m in Regex.Matches(body, cellPat, RegexOptions.IgnoreCase))
                counts[m.Value] = counts.TryGetValue(m.Value, out int c) ? c + 1 : 1;

            // ——— b) function calls (walk the string and record each balanced chunk)
            var stack = new Stack<int>();           // positions of '('
            for (int i = 0; i < body.Length; i++)
            {
                char ch = body[i];
                if (ch == '(') stack.Push(i);
                else if (ch == ')' && stack.Count > 0)
                {
                    int open = stack.Pop();
                    int close = i;

                    // look left from 'open' to capture the function name
                    int start = open;
                    while (start > 0 && char.IsLetter(body[start - 1])) start--;

                    string sub = body.Substring(start, close - start + 1);
                    // Ignore one-character “(…)” groups that aren’t functions
                    if (Regex.IsMatch(sub, @"^[A-Z][A-Z0-9_]*\(", RegexOptions.IgnoreCase))
                        counts[sub] = counts.TryGetValue(sub, out int c) ? c + 1 : 1;
                }
            }

            // ❸ Pick only items that occur more than once, longest first
            var duplicates = counts.Where(kv => kv.Value > 1)
                                   .Select(kv => kv.Key)
                                   .OrderByDescending(s => s.Length)
                                   .ToList();
            if (duplicates.Count == 0) return "=" + body;   // nothing to optimise

            // ❹ Assign variable names a, b, c, …
            var replacements = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            char varName = 'a';
            foreach (string d in duplicates)
            {
                if (varName > 'z') break;                   // 26 vars max
                replacements[d] = varName.ToString();
                varName++;
            }

            // ❺ Replace occurrences – longer strings first avoids nested collisions
            var sb = new StringBuilder(body);
            foreach (string target in replacements.Keys.OrderByDescending(s => s.Length))
            {
                sb.Replace(target, replacements[target]);
            }

            // ❻ Build the final LET( … ) wrapper
            var let = new StringBuilder("=LET(");
            foreach (var kv in replacements)
                let.Append(kv.Value).Append(',').Append(kv.Key).Append(',');

            let.Append(sb).Append(')');
            return let.ToString();
        }

        /* ───────────── 2.  “SOLVE” COMMON FUNCTIONS  ───────────── */
        private static string ResolveCommonFunctions(string formula)
        {
            var app = Globals.ThisAddIn.Application;

            // Replace leading =+ with =
            if (formula.StartsWith("=+"))
                formula = "=" + formula.Substring(2);

            string body = formula.Substring(1).TrimStart();
            // identify the outer-most function token
            var m = Regex.Match(body, @"^([A-Z][A-Z0-9_]*)\s*\(", RegexOptions.IgnoreCase);
            if (!m.Success) return formula;

            string fn = m.Groups[1].Value.ToUpperInvariant();

            switch (fn)
            {
                /* LOOKUPS → resolved cell address */
                case "XLOOKUP":
                case "VLOOKUP":
                case "HLOOKUP":
                case "INDEX":
                    {
                        object result = app.Evaluate(formula);

                        // INDEX (incl. INDEX+MATCH) yields either a cell or a literal
                        if (result is Excel.Range rng && rng.Count == 1)
                            return "=" + rng.Address[false, false, Excel.XlReferenceStyle.xlA1];

                        return ToFormulaLiteral(result) ?? formula;
                    }

                /* MIN / MAX → address of first extreme value */
                case "MIN":
                case "MAX":
                    {
                        string inside = body.Substring(fn.Length).Trim();
                        int firstComma = FindTopLevelComma(inside);
                        string rangeText = firstComma < 0
                            ? inside.Trim('(', ')')
                            : inside.Substring(1, firstComma - 1);

                        try
                        {
                            var rng = app.Range[rangeText];
                            double extreme = (double)app.Evaluate(fn + "(" + rangeText + ")");
                            var found = rng.Find(extreme, Type.Missing,
                                                 Excel.XlFindLookIn.xlValues,
                                                 Excel.XlLookAt.xlWhole);
                            return found != null
                                ? "=" + found.Address[false, false, Excel.XlReferenceStyle.xlA1]
                                : formula;
                        }
                        catch { return formula; }
                    }

                /* SUMIF / SUMIFS → explicit SUM(range1,range2,…) */
                case "SUMIF":
                case "SUMIFS":
                    {
                        try
                        {
                            var pieces = SplitTopLevel(body.Substring(fn.Length));
                            if (pieces.Count < 3) return formula;

                            var critRange = app.Range[pieces[0]];
                            var criteria = pieces[1];
                            var sumRange = pieces.Count >= 3 ? app.Range[pieces[2]] : critRange;

                            var addrs = new List<string>();
                            int n = Math.Min(critRange.Count, sumRange.Count);

                            for (int i = 1; i <= n; i++)
                            {
                                string critExpr = "=" + critRange[i].Value2 + criteria;
                                if (Convert.ToBoolean(app.Evaluate(critExpr)))
                                    addrs.Add(sumRange[i].Address[false, false, Excel.XlReferenceStyle.xlA1]);
                            }
                            return addrs.Count > 0 ? "=SUM(" + string.Join(",", addrs) + ")" : "=0";
                        }
                        catch { return formula; }
                    }

                /* CHOOSE / OFFSET / INDIRECT / TRANSPOSE (scalar) */
                case "CHOOSE":
                case "OFFSET":
                case "INDIRECT":
                case "TRANSPOSE":
                    {
                        object result = app.Evaluate(formula);
                        return ToFormulaLiteral(result) ?? formula;
                    }

                default:
                    return formula;
            }
        }

        /* ───────────── helper: convert evaluation result to literal/address ───────────── */
        private static string ToFormulaLiteral(object result)
        {
            switch (result)
            {
                case Excel.Range rng:
                    return "=" + rng.Address[true, true, Excel.XlReferenceStyle.xlA1];

                case double d:
                case float f:
                case int i:
                case long l:
                case decimal dec:
                    return "=" + Convert.ToString(result, CultureInfo.InvariantCulture);

                case string s:
                    return "=\"" + s.Replace("\"", "\"\"") + "\"";

                case bool b:
                    return "=" + (b ? "TRUE" : "FALSE");

                default:
                    return null;
            }
        }

        /* naive comma splitter – ignores deeper nesting */
        private static List<string> SplitTopLevel(string argListWithParens)
        {
            var list = new List<string>();
            var sb = new StringBuilder();
            int depth = 0;
            foreach (char c in argListWithParens.Trim('(', ')'))
            {
                if (c == ',' && depth == 0)
                {
                    list.Add(sb.ToString().Trim());
                    sb.Clear();
                }
                else
                {
                    if (c == '(') depth++;
                    if (c == ')') depth--;
                    sb.Append(c);
                }
            }
            if (sb.Length > 0) list.Add(sb.ToString().Trim());
            return list;
        }

        private static int FindTopLevelComma(string s)
        {
            int depth = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '(') depth++;
                else if (c == ')') depth--;
                else if (c == ',' && depth == 0) return i;
            }
            return -1;
        }

        /* ───────────── 3.  ANCHOR-CYCLE (F4-style)  ───────────── */
        private static string CycleAnchors(string formula)
        {
            return Regex.Replace(formula,
                @"(\$?)([A-Za-z]{1,3})(\$?)(\d+)",
                m =>
                {
                    string colAnchor = m.Groups[1].Value;
                    string col = m.Groups[2].Value;
                    string rowAnchor = m.Groups[3].Value;
                    string row = m.Groups[4].Value;

                    int state = (colAnchor != "" ? 2 : 0) + (rowAnchor != "" ? 1 : 0);
                    switch (state)            // none → both → row → col → none
                    {
                        case 0: return $"${col}${row}";
                        case 3: return $"{col}${row}";
                        case 1: return $"${col}{row}";
                        default: return $"{col}{row}";
                    }
                },
                RegexOptions.IgnoreCase);
        }

        /* ─────────────────────────  SHARED  HELPERS  ───────────────────────── */

        /* MAGNITUDE ±10^n  (now n = ±3) */
        public static string AdjustPowerOfTen(string formula, int delta)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') return formula;

            string body = formula.Substring(1);
            body = NormalizeTopLevel(body);

            if (delta == 0) return "=" + body;

            if (TryFindTopLevelPower(body, out int starIdx, out int caretIdx, out int expStart, out int expEnd, out int currentExp))
            {
                int newExp = currentExp + delta;
                if (newExp == 0)
                {
                    string left = (starIdx > 0) ? body.Substring(0, starIdx) : string.Empty;
                    string right = (expEnd < body.Length) ? body.Substring(expEnd) : string.Empty;
                    string combined = NormalizeTopLevel(left + right);
                    return "=" + combined;
                }

                var sb = new StringBuilder(body);
                sb.Remove(expStart, expEnd - expStart).Insert(expStart, newExp.ToString());
                return "=" + sb.ToString();
            }
            else
            {
                string inner = NormalizeTopLevel(body);
                bool needsParens = NeedsParensForTopLevelMultiply(inner);
                string lhs = needsParens ? "(" + inner + ")" : inner;
                return "=" + lhs + "*10^" + delta;
            }
        }

        public static string ToggleTopLevelNegation(string formula)
        {
            if (string.IsNullOrEmpty(formula) || formula[0] != '=') return formula;

            string body = NormalizeTopLevel(formula.Substring(1));

            int depth = 0;
            for (int j = 0; j <= body.Length - 3; j++)
            {
                char c = body[j];
                if (c == '(') { depth++; continue; }
                if (c == ')') { depth--; continue; }
                if (c == '*' && depth == 0 && body[j + 1] == '-' && body[j + 2] == '1')
                {
                    string left = body.Remove(j, 3);
                    return "=" + NormalizeTopLevel(left);
                }
            }

            bool needsParens = HasTopLevelPlusMinusOrConcatOrCompare(body);
            string expr = needsParens ? "(" + body + ")" : body;
            return "=" + expr + "*-1";
        }

        private static bool TryFindTopLevelPower(
            string text,
            out int starIdx, out int caretIdx, out int expStart, out int expEnd, out int exponent)
        {
            int depth = 0;
            starIdx = caretIdx = expStart = expEnd = -1;
            exponent = 0;

            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '(') { depth++; continue; }
                if (c == ')') { depth--; continue; }
                if (depth != 0) continue;

                if (c == '*')
                {
                    int pos = i + 1;
                    while (pos < text.Length && char.IsWhiteSpace(text[pos])) pos++;
                    if (pos + 1 >= text.Length || text[pos] != '1' || text[pos + 1] != '0') continue;
                    pos += 2;
                    while (pos < text.Length && char.IsWhiteSpace(text[pos])) pos++;
                    if (pos >= text.Length || text[pos] != '^') continue;
                    caretIdx = pos;
                    pos++;
                    while (pos < text.Length && char.IsWhiteSpace(text[pos])) pos++;

                    int s = pos;
                    if (pos < text.Length && (text[pos] == '+' || text[pos] == '-')) pos++;
                    int digitsStart = pos;
                    while (pos < text.Length && char.IsDigit(text[pos])) pos++;

                    if (pos > digitsStart)
                    {
                        int e = pos;
                        if (int.TryParse(text.Substring(s, e - s), out int exp))
                        {
                            starIdx = i;
                            expStart = s;
                            expEnd = e;
                            exponent = exp;
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        private static bool NeedsParensForTopLevelMultiply(string text)
        {
            return HasTopLevelPlusMinusOrConcatOrCompare(text);
        }

        private static bool HasTopLevelPlusMinusOrConcatOrCompare(string text)
        {
            int depth = 0;
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '(') { depth++; continue; }
                if (c == ')') { depth--; continue; }
                if (depth != 0) continue;

                if (c == '+' || c == '-')
                {
                    int k = i - 1;
                    while (k >= 0 && char.IsWhiteSpace(text[k])) k--;
                    bool unary = (k < 0) || "=+*/^&,(<>".IndexOf(text[k]) >= 0;
                    if (!unary)
                    {
                        if (k >= 0 && (text[k] == 'E' || text[k] == 'e')) continue;
                        return true;
                    }
                }

                if (c == '&' || c == '<' || c == '>') return true;
            }
            return false;
        }

        private static string NormalizeTopLevel(string body)
        {
            body = TrimOuterWhitespace(body);
            body = StripUnaryPlus(body);
            body = StripAllRedundantOuterParens(body);
            return body;
        }

        private static string TrimOuterWhitespace(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            int i = 0, j = s.Length - 1;
            while (i <= j && char.IsWhiteSpace(s[i])) i++;
            while (j >= i && char.IsWhiteSpace(s[j])) j--;
            return (i == 0 && j == s.Length - 1) ? s : s.Substring(i, j - i + 1);
        }

        private static string StripUnaryPlus(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            int i = 0;
            while (i < s.Length && char.IsWhiteSpace(s[i])) i++;
            if (i < s.Length && s[i] == '+') return s.Remove(i, 1);
            return s;
        }

        private static string StripAllRedundantOuterParens(string s)
        {
            string prev;
            do { prev = s; s = StripOneRedundantOuterParens(s); } while (s.Length < prev.Length);
            return s;
        }

        private static string StripOneRedundantOuterParens(string s)
        {
            if (string.IsNullOrEmpty(s)) return s;
            s = TrimOuterWhitespace(s);
            if (s.Length < 2 || s[0] != '(' || s[s.Length - 1] != ')') return s;

            int depth = 0;
            for (int i = 0; i < s.Length; i++)
            {
                char c = s[i];
                if (c == '(') depth++;
                else if (c == ')') depth--;
                if (i < s.Length - 1 && depth == 0) return s;
            }
            return s.Substring(1, s.Length - 2);
        }

        /* IFERROR wrap */
        private static string WrapInIfError(string formula, string alt)
        {
            if (formula.StartsWith("=IFERROR(", StringComparison.OrdinalIgnoreCase))
            {
                // Remove IFERROR wrapper
                return UnwrapIfError(formula);
            }

            // Add IFERROR wrapper
            return $"=IFERROR({formula.Substring(1)},{alt})";
        }

        private static string UnwrapIfError(string formula)
        {
            // Remove "=IFERROR(" prefix
            string content = formula.Substring(9); // "=IFERROR(".Length = 9

            int depth = 0;
            int commaIndex = -1;

            // Find the comma that separates the two IFERROR parameters
            for (int i = 0; i < content.Length; i++)
            {
                char c = content[i];
                if (c == '(')
                    depth++;
                else if (c == ')')
                    depth--;
                else if (c == ',' && depth == 0)
                {
                    commaIndex = i;
                    break;
                }
            }

            if (commaIndex == -1)
                return formula; // Malformed IFERROR, return as-is

            // Extract the original formula and add back the "=" prefix
            string originalFormula = content.Substring(0, commaIndex);
            return "=" + originalFormula;
        }

        /* MULTIPLIER NORMALISATION */
        private static string NormalizeTopLevelMultipliers(string fullFormula)
        {
            if (!fullFormula.StartsWith("=")) return fullFormula;
            string work = fullFormula.Substring(1);

            int signParity = 0; int expSum = 0;
            while (FindLastTopLevelStar(work, out int starIdx))
            {
                string factor = work.Substring(starIdx + 1).Trim();

                if (factor.StartsWith("-1"))
                {
                    signParity ^= 1; work = work.Substring(0, starIdx); continue;
                }
                if (factor.StartsWith("10^") &&
                    int.TryParse(factor.Substring(3), out int exp))
                {
                    expSum += exp; work = work.Substring(0, starIdx); continue;
                }
                break; // stop on unknown factor
            }

            string core = TrimOuterParens(work);
            bool needsParens = ContainsTopLevelAddSub(core);
            if (needsParens) core = "(" + core + ")";

            var sb = new StringBuilder("=" + core);
            if (expSum != 0) sb.Append("*10^").Append(expSum);
            if (signParity == 1) sb.Append("*-1");
            return sb.ToString();
        }

        private static bool FindLastTopLevelStar(string text, out int idx)
        {
            int depth = 0;
            for (int i = text.Length - 1; i >= 0; i--)
            {
                char c = text[i];
                if (c == ')') depth++;
                else if (c == '(') depth--;
                else if (c == '*' && depth == 0) { idx = i; return true; }
            }
            idx = -1; return false;
        }

        private static string TrimOuterParens(string s)
        {
            while (s.StartsWith("(") && s.EndsWith(")"))
            {
                int depth = 0; bool balanced = true;
                for (int i = 0; i < s.Length - 1; i++)
                {
                    if (s[i] == '(') depth++;
                    else if (s[i] == ')')
                    {
                        depth--;
                        if (depth == 0 && i != s.Length - 2)
                        { balanced = false; break; }
                    }
                }
                if (balanced) s = s.Substring(1, s.Length - 2); else break;
            }
            return s;
        }

        private static bool ContainsTopLevelAddSub(string text)
        {
            int depth = 0;
            foreach (char c in text)
            {
                if (c == '(') depth++;
                else if (c == ')') depth--;
                else if ((c == '+' || c == '-') && depth == 0) return true;
            }
            return false;
        }
    }
}

internal sealed class FormulaWrapDialog : Form
{
    private readonly TextBox txtBegin;
    private readonly TextBox txtEnd;

    public string BeginText => txtBegin.Text;
    public string EndText => txtEnd.Text;

    public FormulaWrapDialog()
    {
        Text = "Formula Wrap";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = MinimizeBox = false;
        StartPosition = FormStartPosition.CenterParent;
        Width = 300; Height = 150;

        var lbl1 = new Label { Left = 12, Top = 15, Width = 60, Text = "Beginning:" };
        var lbl2 = new Label { Left = 12, Top = 50, Width = 60, Text = "Ending:" };

        txtBegin = new TextBox { Left = 80, Top = 12, Width = 190 };
        txtEnd = new TextBox { Left = 80, Top = 47, Width = 190 };

        var btnOK = new Button { Text = "OK", Left = 115, Width = 75, Top = 80, DialogResult = DialogResult.OK };
        var btnCL = new Button { Text = "Cancel", Left = 195, Width = 75, Top = 80, DialogResult = DialogResult.Cancel };
        AcceptButton = btnOK; CancelButton = btnCL;

        Controls.AddRange(new Control[] { lbl1, lbl2, txtBegin, txtEnd, btnOK, btnCL });
    }
}
