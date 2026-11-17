using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using SD = System.Drawing;
using System.Drawing.Text;

namespace OfficePragmatics_Excel
{
    public static class TracePrecedent
    {
        /* ───────── WM_SETREDRAW for batched UI updates ───────── */
        [DllImport("user32.dll")]
        private static extern int SendMessage(IntPtr hWnd, int wMsg, bool wParam, int lParam);
        private const int WM_SETREDRAW = 0x000B;

        private static void SuspendDrawing(Control control)
        {
            SendMessage(control.Handle, WM_SETREDRAW, false, 0);
        }

        private static void ResumeDrawing(Control control)
        {
            SendMessage(control.Handle, WM_SETREDRAW, true, 0);
            control.Refresh();
        }

        /* ───────── static, compiled regex – parsed once ───────── */
        private static readonly Regex FuncRx = new Regex(
            @"\b[A-Za-z_][A-Za-z0-9_]*\s*\(",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex RefRx = new Regex(
            @"(?:(?:'[^']+'|[^!,\s()+\-*/^<>]+)!)?" +
            "(" +
                @"(?:\b\w+\[[^\]]*\]|@\[[^\]]*\]|\[\[[^\]]*\]\])" +
                "|" +
                @"\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?" +
                "|" +
                @"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}" +
                "|" +
                @"\$?\d+:\$?\d+" +
            ")",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly Regex ExternalPathRx = new Regex(
            @"'(?<path>[A-Za-z]:\\[^']*?)\[(?<book>[^\]]+)\](?<sheet>[^']+)'(?<bang>!)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant);

        private static readonly FieldInfo BfStart = typeof(PrecedentInfo).GetField(
            "<Start>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic);
        private static readonly FieldInfo BfLen = typeof(PrecedentInfo).GetField(
            "<Length>k__BackingField", BindingFlags.Instance | BindingFlags.NonPublic);

        private static readonly Dictionary<string, string> _fmtCache = new Dictionary<string, string>(256);
        private static readonly Dictionary<string, string> _evalCache = new Dictionary<string, string>(128);

        private static void SafeReleaseCom(object obj)
        {
            if (obj != null)
            {
                try { Marshal.ReleaseComObject(obj); }
                catch { }
            }
        }

        public static void Run(Excel.Application app)
        {
            var activeCell = app.ActiveCell as Excel.Range;
            RevertBackwards.SaveCurrentCell(activeCell);
            if (activeCell == null)
            {
                MessageBox.Show("No cell is selected.", "Trace Precedents", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var formula = activeCell.Formula as string;
            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("="))
            {
                MessageBox.Show("The selected cell does not contain a formula.", "Trace Precedents", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            bool scrn = app.ScreenUpdating;
            var calc = app.Calculation;
            app.ScreenUpdating = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;

            var precedents = new List<PrecedentInfo>();
            try
            {
                precedents.Capacity = 50;
                ParseFormula(app, formula, activeCell, precedents);
            }
            finally
            {
                app.Calculation = calc;
                app.ScreenUpdating = scrn;
            }

            precedents.Sort((a, b) => a.Start.CompareTo(b.Start));
            var form = new PrecedentsForm(formula, precedents, activeCell);
            form.Show();
        }

        // ─────────────────────────────────────────────────────────────────────
        // Parsing Logic
        // ─────────────────────────────────────────────────────────────────────
        private struct Replacement
        {
            public int OrigStart;
            public int OrigLen;
            public int NewLen;
        }

        private static string SimplifyFormulaForDisplay(string original, out List<Replacement> reps)
        {
            reps = new List<Replacement>(8);
            if (string.IsNullOrEmpty(original) || original[0] != '=') return original;

            var sb = new StringBuilder(original.Length);
            int last = 0;

            foreach (Match m in ExternalPathRx.Matches(original))
            {
                if (m.Index > last) sb.Append(original, last, m.Index - last);
                string repl = $"[{m.Groups["book"].Value}]{m.Groups["sheet"].Value}{m.Groups["bang"].Value}";
                sb.Append(repl);
                reps.Add(new Replacement { OrigStart = m.Index, OrigLen = m.Length, NewLen = repl.Length });
                last = m.Index + m.Length;
            }

            if (last < original.Length) sb.Append(original, last, original.Length - last);
            return sb.ToString();
        }

        private static void RetargetPrecedentSpansForDisplay(List<PrecedentInfo> precs, List<Replacement> reps)
        {
            if (reps == null || reps.Count == 0 || precs == null || precs.Count == 0) return;

            foreach (var p in precs)
            {
                int start = p.Start;
                int len = p.Length;
                foreach (var r in reps)
                {
                    int delta = r.OrigLen - r.NewLen;
                    int rEnd = r.OrigStart + r.OrigLen;
                    int pEnd = start + len;

                    if (start >= rEnd) start -= delta;
                    else if (pEnd > r.OrigStart)
                    {
                        int overlap = Math.Max(0, Math.Min(pEnd, rEnd) - Math.Max(start, r.OrigStart));
                        if (overlap > 0)
                        {
                            int shrink = (int)Math.Round((double)overlap * (r.OrigLen - r.NewLen) / r.OrigLen);
                            len -= Math.Max(0, shrink);
                        }
                    }
                }
                if (p.IsReference && !string.IsNullOrEmpty(p.DisplayText)) len = p.DisplayText.Length;
                BfStart?.SetValue(p, start);
                BfLen?.SetValue(p, len);
            }
        }

        private static void FixNonReferenceLengthsUsingDisplayedFormula(List<PrecedentInfo> precs, string displayed)
        {
            if (precs == null || precs.Count == 0 || string.IsNullOrEmpty(displayed)) return;

            foreach (var p in precs)
            {
                if (p.IsReference) continue;
                int s = p.Start;
                if (s < 0 || s >= displayed.Length) continue;
                int open = (displayed[s] == '(') ? s : displayed.IndexOf('(', s);
                if (open < 0) continue;
                int depth = 1, i = open + 1;
                while (i < displayed.Length && depth > 0)
                {
                    char c = displayed[i];
                    if (c == '(') depth++; else if (c == ')') depth--;
                    i++;
                }
                if (depth != 0) continue;
                int newLen = i - s;
                if (newLen > 0) BfLen?.SetValue(p, newLen);
            }
        }

        private static string InjectZeroWidthSpacesAndRemap(string text, List<PrecedentInfo> precs)
        {
            int[] indexMap = new int[text.Length + 1];
            StringBuilder sb = new StringBuilder(text.Length + text.Length / 4);
            int shift = 0;

            for (int i = 0; i < text.Length; i++)
            {
                indexMap[i] = i + shift;
                char c = text[i];
                sb.Append(c);
                if ("+-*/^&=,()".IndexOf(c) >= 0) { sb.Append('\u200B'); shift++; }
            }
            indexMap[text.Length] = text.Length + shift;

            foreach (var p in precs)
            {
                int oldStart = Math.Max(0, Math.Min(p.Start, text.Length));
                int oldEnd = Math.Max(0, Math.Min(p.Start + p.Length, text.Length));
                int newStart = indexMap[oldStart];
                int newLen = indexMap[oldEnd] - newStart;
                BfStart?.SetValue(p, newStart);
                BfLen?.SetValue(p, newLen);
            }
            return sb.ToString();
        }

        private static string SimplifyExternalReference(string refText)
        {
            return ExternalPathRx.Replace(refText, m => $"[{m.Groups["book"].Value}]{m.Groups["sheet"].Value}{m.Groups["bang"].Value}");
        }

        private static void ParseFormula(Excel.Application app, string formula, Excel.Range activeCell, List<PrecedentInfo> precedents)
        {
            var funcMatches = FuncRx.Matches(formula);
            var refMatches = RefRx.Matches(formula);

            var rangeLookupCache = new Dictionary<string, Excel.Range>(StringComparer.OrdinalIgnoreCase);
            var processedRanges = new HashSet<int>(refMatches.Count);

            foreach (Match m in funcMatches)
            {
                int funcStart = m.Index;
                int paren = m.Index + m.Length - 1;
                int depth = 1, pos = paren + 1;
                while (pos < formula.Length && depth > 0)
                {
                    char c = formula[pos];
                    if (c == '(') depth++; else if (c == ')') depth--;
                    pos++;
                }
                if (depth != 0) continue;
                int funcLen = pos - funcStart;
                string fText = formula.Substring(funcStart, funcLen);
                string val = GetCachedEvaluation(app, fText, activeCell);
                precedents.Add(new PrecedentInfo(fText, funcStart, funcLen, val));
            }

            ParseBareParentheses(app, formula, activeCell, precedents);

            foreach (Match m in refMatches)
            {
                if (processedRanges.Contains(m.Index)) continue;
                processedRanges.Add(m.Index);
                string refText = m.Value;
                string simplified = SimplifyExternalReference(refText);

                Excel.Range rng;
                if (!rangeLookupCache.TryGetValue(refText, out rng))
                {
                    rng = TryGetOpenRange(app, refText);
                    rangeLookupCache[refText] = rng;
                }

                precedents.Add(new PrecedentInfo(rng, simplified, m.Index, m.Length));
            }
        }

        private static Excel.Range TryGetOpenRange(Excel.Application app, string refText)
        {
            try { return app.Range[refText]; } catch { return null; }
        }

        private static void ParseBareParentheses(Excel.Application app, string formula, Excel.Range activeCell, List<PrecedentInfo> precedents)
        {
            var parenStack = new Stack<int>(16);
            var funcMatches = FuncRx.Matches(formula);
            var funcStarts = new HashSet<int>(funcMatches.Count);
            foreach (Match m in funcMatches) funcStarts.Add(m.Index + m.Length - 1);

            for (int i = 0; i < formula.Length; i++)
            {
                if (formula[i] == '(')
                {
                    if (!funcStarts.Contains(i))
                    {
                        int k = i - 1;
                        while (k >= 0 && char.IsWhiteSpace(formula[k])) k--;
                        bool preceded = k >= 0 && (char.IsLetterOrDigit(formula[k]) || formula[k] == '_');
                        if (!preceded) parenStack.Push(i);
                    }
                }
                else if (formula[i] == ')' && parenStack.Count > 0)
                {
                    int start = parenStack.Pop();
                    string grpText = formula.Substring(start, i - start + 1);
                    string val = GetCachedEvaluation(app, grpText, activeCell);
                    precedents.Add(new PrecedentInfo("( )", start, i - start + 1, val));
                }
            }
        }

        private static string GetCachedEvaluation(Excel.Application app, string funcText, Excel.Range originCell)
        {
            if (_evalCache.TryGetValue(funcText, out string cached)) return cached;
            string result = EvaluateFunctionFormatted(app, funcText, originCell);
            if (_evalCache.Count < 1000) _evalCache[funcText] = result;
            return result;
        }

        private static string EvaluateFunctionFormatted(Excel.Application app, string funcText, Excel.Range originCell)
        {
            object raw = null;
            try { raw = app.Evaluate("=" + funcText); } catch { }

            if (raw == null && funcText.StartsWith("XLOOKUP(", StringComparison.OrdinalIgnoreCase))
            {
                string converted = ConvertXLookupToIndexMatch(funcText);
                if (!string.IsNullOrEmpty(converted)) try { raw = app.Evaluate("=" + converted); } catch { }
            }

            if (IsExcelError(raw)) return ConvertExcelErrorToText(raw);
            string fmt = DetermineNumberFormat(app, funcText, originCell);
            return FormatWithExcel(app, raw, fmt);
        }

        private static string DetermineNumberFormat(Excel.Application app, string funcText, Excel.Range originCell)
        {
            string commonFmt = null;
            var matches = RefRx.Matches(funcText);
            if (matches.Count == 0) return (originCell.NumberFormat as string ?? "General").Replace("*", "");

            foreach (Match m in matches)
            {
                Excel.Range rng = null;
                Excel.Worksheet sheet = null;
                try
                {
                    rng = app.Range[m.Value];
                    sheet = (Excel.Worksheet)rng.Worksheet;
                    string key = $"{sheet.Name}|{rng.Address[false, false]}";
                    if (!_fmtCache.TryGetValue(key, out string fmt))
                    {
                        fmt = (((Excel.Range)rng.Cells[1, 1]).NumberFormat as string ?? "").Replace("*", "");
                        _fmtCache[key] = fmt;
                    }
                    if (commonFmt == null) commonFmt = fmt;
                    else if (fmt != commonFmt) { commonFmt = null; break; }
                }
                catch { }
                finally
                {
                    SafeReleaseCom(sheet);
                    SafeReleaseCom(rng);
                }
            }
            return !string.IsNullOrEmpty(commonFmt) ? commonFmt : (originCell.NumberFormat as string ?? "General").Replace("*", "");
        }

        private static object UnwrapIfArray(object val)
        {
            if (val is object[,] arr2D && arr2D.Length == 1) return arr2D[arr2D.GetLowerBound(0), arr2D.GetLowerBound(1)];
            if (val is object[] arr1D && arr1D.Length == 1) return arr1D[0];
            if (val is object[,] || val is object[]) return "[Array]";
            return val;
        }

        private static string FormatWithExcel(Excel.Application app, object val, string fmt)
        {
            val = UnwrapIfArray(val);
            if (val == null) return "";
            if (IsExcelError(val)) return ConvertExcelErrorToText(val);
            try { return app.WorksheetFunction.Text(val, fmt); } catch { return val.ToString(); }
        }

        private static string ConvertXLookupToIndexMatch(string funcText)
        {
            int first = funcText.IndexOf('(');
            if (first < 0) return null;
            string inner = funcText.Substring(first + 1, funcText.Length - first - 2);
            var args = new List<string>(6);
            var cur = new StringBuilder(32);
            int depth = 0;
            foreach (char c in inner)
            {
                if (c == ',' && depth == 0) { args.Add(cur.ToString().Trim()); cur.Clear(); continue; }
                if (c == '(') depth++; else if (c == ')') depth--;
                cur.Append(c);
            }
            args.Add(cur.ToString().Trim());
            if (args.Count < 3 || (args.Count >= 4 && args[3] != "0" && args[3] != "")) return null;
            return $"INDEX({args[2]}, MATCH({args[0]}, {args[1]}, 0))";
        }

        private static bool IsExcelError(object value)
        {
            if (value == null) return false;
            if (value is int i && i < -2000000000) return true;
            if (value.GetType().Name == "CVErr") return true;
            string s = value.ToString();
            return s.StartsWith("#") && (s.EndsWith("!") || s == "#N/A" || s == "#NAME?");
        }

        private static string ConvertExcelErrorToText(object errorValue)
        {
            if (errorValue is int i)
            {
                if (i == -2146826281) return "#DIV/0!";
                if (i == -2146826246) return "#N/A";
                if (i == -2146826259) return "#NAME?";
                if (i == -2146826288) return "#NULL!";
                if (i == -2146826252) return "#NUM!";
                if (i == -2146826265) return "#REF!";
                if (i == -2146826284) return "#VALUE!";
            }
            try { return errorValue.ToString(); } catch { return "#ERROR!"; }
        }

        public class PrecedentInfo
        {
            public bool IsReference { get; }
            public bool IsOutput { get; set; } = false;
            public Excel.Range Range { get; private set; }
            public string DisplayText { get; }
            public string FunctionVal { get; }
            public int Start { get; }
            public int Length { get; }
            public int Level { get; set; }
            public int ColorIndex { get; set; } = -1;

            public PrecedentInfo(Excel.Range rng, string disp, int start, int len) { IsReference = true; Range = rng; DisplayText = disp; Start = start; Length = len; }
            public PrecedentInfo(string disp, int start, int len, string val) { IsReference = false; DisplayText = disp; FunctionVal = val; Start = start; Length = len; }
            public void UpdateRange(Excel.Range newRange) { this.Range = newRange; }
        }

        /* ────────────────────────────────────────────────────────────────────────── */
        /* ─────────────── MODERN CUSTOM FORM (Deep Navy Style) ───────────────────── */
        /* ────────────────────────────────────────────────────────────────────────── */
        public class PrecedentsForm : Form
        {
            [DllImport("user32.dll")] public static extern bool ReleaseCapture();
            [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            private struct UIColors
            {
                public static readonly SD.Color Surface = SD.Color.FromArgb(255, 255, 255);
                public static readonly SD.Color SurfaceAlt = SD.Color.FromArgb(248, 248, 248);
                public static readonly SD.Color TitleBg = SD.Color.FromArgb(32, 42, 102);
                public static readonly SD.Color TitleText = SD.Color.White;
                public static readonly SD.Color Border = SD.Color.FromArgb(220, 220, 220);
                public static readonly SD.Color TextPrimary = SD.Color.FromArgb(32, 31, 30);
                public static readonly SD.Color TextSecondary = SD.Color.FromArgb(96, 94, 92);
                public static readonly SD.Color Accent = SD.Color.FromArgb(110, 125, 220);
                public static readonly SD.Color AccentHover = SD.Color.FromArgb(85, 95, 180);
                public static readonly SD.Color Danger = SD.Color.FromArgb(232, 17, 35);
            }

            private class DrillDownState
            {
                public List<PrecedentInfo> Precedents { get; }
                public Excel.Range OriginCell { get; }
                public int SelectedIndex { get; }
                public int FormulaPrefixLength { get; }
                public Color PrefixColor { get; }
                public Color PrefixHighlightColor { get; }
                public int ColorCycleStartIndex { get; }

                public DrillDownState(List<PrecedentInfo> precedents, Excel.Range origin, int selectedIndex, int prefixLength, Color prefixColor, Color prefixHighlightColor, int colorCycleStartIndex)
                {
                    Precedents = precedents; OriginCell = origin; SelectedIndex = selectedIndex; FormulaPrefixLength = prefixLength; PrefixColor = prefixColor; PrefixHighlightColor = prefixHighlightColor; ColorCycleStartIndex = colorCycleStartIndex;
                }
            }

            private readonly Excel.Application _cachedApp;

            private readonly Stack<DrillDownState> _drillDownStack = new Stack<DrillDownState>(8);
            private readonly Panel _formulaPanel;
            private RichTextBox _currentFormulaLabel;
            private readonly Excel.Range _initialOriginCell;
            private List<PrecedentInfo> _precedents;
            private Excel.Range _origin;
            private int _formulaPrefixLength;
            private Color _currentPrefixColor;
            private Color _currentPrefixHighlightColor;
            private int _colorCycleStartIndex;
            private readonly List<PrecedentInfo> _shown = new List<PrecedentInfo>(64);
            private int _cur = -1;

            private readonly ListView _list;
            private readonly Color[] _colors = { Color.FromArgb(236, 0, 140), Color.FromArgb(0, 176, 80), Color.FromArgb(112, 48, 160), Color.FromArgb(192, 0, 0), Color.FromArgb(0, 176, 240), Color.FromArgb(255, 192, 0), Color.FromArgb(148, 112, 53), Color.FromArgb(127, 127, 127), Color.FromArgb(247, 150, 70), Color.FromArgb(0, 128, 128), Color.FromArgb(128, 0, 0), Color.FromArgb(0, 0, 128) };
            private readonly Color _outputFontColor = Color.FromArgb(0, 112, 192);
            private readonly Color _outputHighlightColor = Color.FromArgb(176, 196, 222);
            private readonly Color _groupHL = Color.FromArgb(240, 248, 255);

            private const int INDENT = 16, FXCOL = 26;

            private readonly HashSet<Excel.Range> _unhiddenRows = new HashSet<Excel.Range>(16);
            private readonly HashSet<Excel.Range> _unhiddenCols = new HashSet<Excel.Range>(16);
            private bool _keepExposed = false;

            private readonly List<Excel.Workbook> _workbooksToClose;
            private int _prevStart = -1, _prevLen = 0;
            private Point _originalPosition = Point.Empty;
            private bool _hasMovedAway = false;
            private readonly Dictionary<Excel.Range, string> _valueCache = new Dictionary<Excel.Range, string>(64);
            private bool _updating = false;
            private bool _isClosing = false;

            private Font _italicFont;
            private Font _underlineFont;
            private Font _titleFont;
            private Font _closeButtonFont;

            public PrecedentsForm(string formula, List<PrecedentInfo> precs, Excel.Range origin)
            {
                this.SuspendLayout();

                this._initialOriginCell = origin;
                this._cachedApp = (Excel.Application)origin.Application;

                this._precedents = precs;
                this._origin = origin;

                _workbooksToClose = new List<Excel.Workbook>(4);

                this._currentPrefixColor = _outputFontColor;
                this._currentPrefixHighlightColor = _outputHighlightColor;
                this._colorCycleStartIndex = 0;

                string prefix = origin.Address[false, false] + "=";
                int baseOffset = prefix.Length - 1;

                List<Replacement> reps;
                string simplifiedFormula = SimplifyFormulaForDisplay(formula, out reps);
                RetargetPrecedentSpansForDisplay(_precedents, reps);
                FixNonReferenceLengthsUsingDisplayedFormula(_precedents, simplifiedFormula);

                foreach (var p in _precedents)
                {
                    int currentStart = (int)BfStart.GetValue(p);
                    BfStart.SetValue(p, currentStart + baseOffset);
                }

                string displayedFormulaBase = prefix + simplifiedFormula.Substring(1);
                string displayedFormulaFinal = InjectZeroWidthSpacesAndRemap(displayedFormulaBase, _precedents);

                this._formulaPrefixLength = prefix.Length;

                this.FormBorderStyle = FormBorderStyle.None;
                this.DoubleBuffered = true;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.BackColor = UIColors.Border;
                this.Padding = new Padding(1);
                this.ShowIcon = false;
                this.Text = "Trace Precedents";
                this.TopMost = true;
                this.ResizeRedraw = true;

                float scaleFactor = 1.0f;
                using (Graphics g = CreateGraphics())
                {
                    float dpiX = g.DpiX;
                    float dpiY = g.DpiY;
                    scaleFactor = dpiX / 96f;
                    Size = new System.Drawing.Size((int)(3.5f * dpiX), (int)(3.5f * dpiY));
                }

                var mainPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = UIColors.Surface
                };
                mainPanel.Paint += (s, e) => { using (var p = new Pen(UIColors.Border)) e.Graphics.DrawRectangle(p, 0, 0, mainPanel.Width - 1, mainPanel.Height - 1); };
                this.Controls.Add(mainPanel);

                _titleFont = new Font("Segoe UI Semibold", 10f);
                _closeButtonFont = new Font("Segoe UI", 11f);

                int titleHeight = (int)(32 * scaleFactor);

                var pnlTitle = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = titleHeight,
                    BackColor = UIColors.TitleBg
                };
                pnlTitle.MouseDown += (s, e) => {
                    if (e.Button == MouseButtons.Left) { ReleaseCapture(); SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); }
                };

                pnlTitle.Paint += (s, e) =>
                {
                    e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                    using (var brush = new SolidBrush(UIColors.TitleText))
                    {
                        float y = (pnlTitle.Height - _titleFont.Height) / 2f;
                        e.Graphics.DrawString("Trace Precedents", _titleFont, brush, 10, y);
                    }
                };

                var btnClose = new Label
                {
                    Text = "✕",
                    Font = _closeButtonFont,
                    ForeColor = UIColors.TitleText,
                    Dock = DockStyle.Right,
                    Width = (int)(40 * scaleFactor),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = UIColors.TitleBg
                };
                btnClose.MouseEnter += (s, e) => { btnClose.BackColor = UIColors.Danger; btnClose.ForeColor = SD.Color.White; };
                btnClose.MouseLeave += (s, e) => { btnClose.BackColor = UIColors.TitleBg; btnClose.ForeColor = UIColors.TitleText; };
                btnClose.Click += (s, e) => { Close(); };
                pnlTitle.Controls.Add(btnClose);

                mainPanel.Controls.Add(pnlTitle);

                KeyPreview = true;
                KeyDown += OnKeyDown;
                Resize += (s, e) => LayoutControls();
                this.FormClosing += (s, e) => { _isClosing = true; };
                FormClosed += OnFormClosed;

                _formulaPanel = new Panel
                {
                    Location = new Point(10, pnlTitle.Bottom + 10),
                    Width = mainPanel.Width - 20
                };
                mainPanel.Controls.Add(_formulaPanel);

                _currentFormulaLabel = new RichTextBox
                {
                    Dock = DockStyle.None,
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    BackColor = Color.White,
                    Font = new Font("Segoe UI", 10),
                    Text = displayedFormulaFinal,
                    HideSelection = true,
                    TabStop = false,
                    WordWrap = true,
                    ScrollBars = RichTextBoxScrollBars.None
                };
                _formulaPanel.Controls.Add(_currentFormulaLabel);
                _formulaPanel.Paint += (s, e) => {
                    using (var p = new Pen(UIColors.Border))
                    {
                        e.Graphics.DrawRectangle(p, 0, 0, _formulaPanel.Width - 1, _formulaPanel.Height - 1);

                        var rtbControls = _formulaPanel.Controls.Cast<Control>()
                            .Where(c => c is RichTextBox)
                            .OrderBy(c => c.Top)
                            .ToList();

                        for (int i = 1; i < rtbControls.Count; i++)
                        {
                            int y = rtbControls[i].Top - 2;
                            e.Graphics.DrawLine(p, 1, y, _formulaPanel.Width - 2, y);
                        }
                    }
                };
                _formulaPanel.Padding = new Padding(1);

                _list = new ListView
                {
                    View = View.Details,
                    FullRowSelect = true,
                    HideSelection = false,
                    MultiSelect = false,
                    Font = new Font("Segoe UI", 10),
                    OwnerDraw = true,
                    VirtualMode = false,
                    BorderStyle = BorderStyle.FixedSingle
                };

                _italicFont = new Font(_list.Font, FontStyle.Italic);
                _underlineFont = new Font(_currentFormulaLabel.Font, FontStyle.Underline);

                var doubleBufferProperty = typeof(Control).GetProperty("DoubleBuffered", BindingFlags.NonPublic | BindingFlags.Instance);
                doubleBufferProperty?.SetValue(_list, true, null);

                _list.Columns.Add("Reference");
                _list.Columns.Add("*");
                _list.Columns.Add("Value");

                _list.DrawColumnHeader += (s, e) =>
                {
                    e.Graphics.FillRectangle(new SolidBrush(UIColors.SurfaceAlt), e.Bounds);
                    TextRenderer.DrawText(e.Graphics, e.Header.Text, e.Font, e.Bounds, UIColors.TextPrimary, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
                    using (var p = new Pen(UIColors.Border)) e.Graphics.DrawLine(p, e.Bounds.Right - 1, e.Bounds.Top, e.Bounds.Right - 1, e.Bounds.Bottom);
                    using (var p = new Pen(UIColors.Border)) e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
                };
                _list.DrawItem += OnDrawItem;
                _list.DrawSubItem += OnDrawSub;
                _list.SelectedIndexChanged += OnSelectedIndexChanged;

                mainPanel.Controls.Add(_list);

                ColorizeFormula();
                PopulateList();
                LayoutControls();
                if (_list.Items.Count > 0) { _list.Items[0].Selected = true; _cur = 0; Highlight(); }
                _list.Focus();

                this.ResumeLayout(false);
            }

            private void LayoutControls()
            {
                // Make the panel slightly wider to ensure text doesn't wrap prematurely
                _formulaPanel.Width = ClientSize.Width - 20;

                // Start very close to the top (was 4)
                int currentY = 0;

                for (int i = _formulaPanel.Controls.Count - 1; i >= 0; i--)
                {
                    var rtb = _formulaPanel.Controls[i] as RichTextBox;
                    if (rtb != null)
                    {
                        // Maximizing width to prevent "ghost" wrapping
                        rtb.Width = _formulaPanel.Width;
                        rtb.Location = new Point(0, currentY);

                        // ─── Autofit Logic ───
                        // Get the Y position of the last character
                        int lastCharIndex = Math.Max(0, rtb.TextLength - 1);
                        Point lastCharPos = rtb.GetPositionFromCharIndex(lastCharIndex);

                        // Calculate exact height required by the font lines
                        int textHeight = lastCharPos.Y + rtb.Font.Height;

                        // Use the Font's actual height + 2px buffer as the minimum, 
                        // replacing the hardcoded '26' which caused gaps on small zooms.
                        int minHeight = rtb.Font.Height + 2;

                        // Add minimal buffer (1px) for the bottom edge
                        int neededHeight = textHeight + 1;

                        rtb.Height = Math.Max(minHeight, Math.Min(neededHeight, 150));

                        // Stack the next control immediately below (was +4)
                        currentY += rtb.Height;
                    }
                }

                // Set panel height to exactly what is needed (no extra buffer)
                _formulaPanel.Height = currentY;

                // Adjust List position
                _list.Location = new Point(10, _formulaPanel.Bottom + 10);
                _list.Size = new System.Drawing.Size(ClientSize.Width - 20, ClientSize.Height - _formulaPanel.Bottom - 20);

                // Column resizing logic remains same
                int total = _list.ClientSize.Width;
                _list.Columns[1].Width = FXCOL;
                _list.Columns[0].Width = (int)(total * 0.65);
                _list.Columns[2].Width = total - _list.Columns[0].Width - FXCOL;

                _formulaPanel.Invalidate();
            }

            private bool RangesAreEqual(Excel.Range r1, Excel.Range r2)
            {
                if (r1 == null || r2 == null) return r1 == r2;
                try { return (string)r1.Address[External: true] == (string)r2.Address[External: true]; } catch { return false; }
            }

            private bool IsDrillable(PrecedentInfo p)
            {
                if (p == null || !p.IsReference || p.Range == null) return false;
                if (RangesAreEqual(p.Range, _origin)) return false;
                foreach (var state in _drillDownStack) if (RangesAreEqual(p.Range, state.OriginCell)) return false;
                try { string formula = p.Range.Formula as string; return !string.IsNullOrEmpty(formula) && formula.StartsWith("="); } catch { return false; }
            }

            private void OnKeyDown(object s, KeyEventArgs e)
            {
                if (_shown.Count == 0 && e.KeyCode != Keys.Escape) return;
                switch (e.KeyCode)
                {
                    case Keys.Escape: _keepExposed = false; Close(); break;
                    case Keys.Enter: if (_cur >= 0 && _shown[_cur].IsReference && _shown[_cur].Range != null) _keepExposed = true; Close(); break;
                    case Keys.Down: _cur = NextRef(_cur); Highlight(); break;
                    case Keys.Up: _cur = PrevRef(_cur); Highlight(); break;
                    case Keys.Right: DrillDown(); break;
                    case Keys.Left: DrillUp(); break;
                    default: return;
                }
                e.Handled = true;
            }

            private void DrillDown()
            {
                if (_cur < 0 || _cur >= _shown.Count) return;
                var precedentToDrill = _shown[_cur];
                if (!IsDrillable(precedentToDrill)) return;
                var targetCell = precedentToDrill.Range;
                string formula; try { formula = targetCell.Formula as string; } catch { return; }
                _drillDownStack.Push(new DrillDownState(_precedents, _origin, _cur, _formulaPrefixLength, _currentPrefixColor, _currentPrefixHighlightColor, _colorCycleStartIndex));
                this.SuspendLayout(); _list.BeginUpdate();
                try
                {
                    if (precedentToDrill.IsReference && precedentToDrill.ColorIndex >= 0) { var parentColor = _colors[precedentToDrill.ColorIndex % _colors.Length]; _currentPrefixColor = parentColor; _currentPrefixHighlightColor = Lighten(parentColor, 0.85); _colorCycleStartIndex = (precedentToDrill.ColorIndex + 1) % _colors.Length; }
                    else { _currentPrefixColor = Color.Black; _currentPrefixHighlightColor = _groupHL; _colorCycleStartIndex = 0; }
                    _currentFormulaLabel.ReadOnly = true;
                    _currentFormulaLabel.Select(precedentToDrill.Start, precedentToDrill.Length);
                    Color hl = precedentToDrill.IsReference && precedentToDrill.ColorIndex >= 0 ? Lighten(_colors[precedentToDrill.ColorIndex % _colors.Length]) : _groupHL;
                    _currentFormulaLabel.SelectionBackColor = hl; _currentFormulaLabel.DeselectAll();

                    bool scrn = _cachedApp.ScreenUpdating;
                    var calc = _cachedApp.Calculation;
                    _cachedApp.ScreenUpdating = false;
                    _cachedApp.Calculation = Excel.XlCalculation.xlCalculationManual;

                    var newPrecedents = new List<PrecedentInfo>();
                    try { ParseFormula(_cachedApp, formula, targetCell, newPrecedents); } finally { _cachedApp.Calculation = calc; _cachedApp.ScreenUpdating = scrn; }

                    newPrecedents.Sort((a, b) => a.Start.CompareTo(b.Start)); _precedents = newPrecedents; _origin = targetCell;
                    string prefix = _origin.Address[false, false] + "="; int baseOffset = prefix.Length - 1;
                    foreach (var p in newPrecedents) { int currentStart = (int)BfStart.GetValue(p); BfStart.SetValue(p, currentStart + baseOffset); }
                    List<Replacement> reps; string simplified = SimplifyFormulaForDisplay(formula, out reps); RetargetPrecedentSpansForDisplay(_precedents, reps); FixNonReferenceLengthsUsingDisplayedFormula(_precedents, simplified); string displayed = prefix + simplified.Substring(1); string newFormulaText = InjectZeroWidthSpacesAndRemap(displayed, _precedents); var newLabel = new RichTextBox { Dock = DockStyle.None, ReadOnly = true, BorderStyle = BorderStyle.None, BackColor = Color.White, Font = new Font("Segoe UI", 10), Text = newFormulaText, HideSelection = true, TabStop = false, WordWrap = true, ScrollBars = RichTextBoxScrollBars.None }; _formulaPanel.Controls.Add(newLabel); newLabel.BringToFront(); _currentFormulaLabel = newLabel; _shown.Clear(); _list.Items.Clear(); ColorizeFormula(); PopulateList(); _cur = _list.Items.Count > 0 ? 0 : -1; if (_cur != -1) { _list.Items[_cur].Selected = true; Highlight(); }
                    _list.Focus();
                }
                finally { _list.EndUpdate(); LayoutControls(); this.ResumeLayout(true); }
            }

            private void DrillUp()
            {
                if (_drillDownStack.Count == 0) return;
                this.SuspendLayout(); _list.BeginUpdate();
                try
                {
                    var prevState = _drillDownStack.Pop();
                    _formulaPanel.Controls.Remove(_currentFormulaLabel); _currentFormulaLabel.Dispose();
                    _precedents = prevState.Precedents; _origin = prevState.OriginCell; _formulaPrefixLength = prevState.FormulaPrefixLength; _currentPrefixColor = prevState.PrefixColor; _currentPrefixHighlightColor = prevState.PrefixHighlightColor; _colorCycleStartIndex = prevState.ColorCycleStartIndex;
                    _currentFormulaLabel = (RichTextBox)_formulaPanel.Controls[0];
                    _shown.Clear(); _list.Items.Clear(); ColorizeFormula(); PopulateList();
                    _cur = prevState.SelectedIndex; if (_cur >= 0 && _cur < _list.Items.Count) { _list.Items[_cur].Selected = true; Highlight(); }
                    _list.Focus();
                }
                finally { _list.EndUpdate(); LayoutControls(); this.ResumeLayout(true); }
            }

            private void OnFormClosed(object sender, FormClosedEventArgs e)
            {
                _italicFont?.Dispose();
                _underlineFont?.Dispose();
                _titleFont?.Dispose();
                _closeButtonFont?.Dispose();

                if (!_keepExposed) { RestoreHidden(); try { SafeSelect(_initialOriginCell); } catch { } foreach (var wb in _workbooksToClose) { CloseTracedWorkbook(wb); } }
                else
                {
                    if (_cur >= 0 && _cur < _shown.Count)
                    {
                        var currentRef = _shown[_cur];
                        if (currentRef.IsReference && currentRef.Range != null)
                        {
                            try
                            {
                                var currentWorkbook = (Excel.Workbook)currentRef.Range.Worksheet.Parent;
                                var initialWorkbook = (Excel.Workbook)_initialOriginCell.Worksheet.Parent;
                                if (currentWorkbook.Name != initialWorkbook.Name)
                                {
                                    foreach (var wb in _workbooksToClose.ToList()) { if (wb.Name != currentWorkbook.Name) { CloseTracedWorkbook(wb); _workbooksToClose.Remove(wb); } }
                                }
                                else { foreach (var wb in _workbooksToClose) { CloseTracedWorkbook(wb); } }
                            }
                            catch { foreach (var wb in _workbooksToClose) { CloseTracedWorkbook(wb); } }
                        }
                        else { foreach (var wb in _workbooksToClose) { CloseTracedWorkbook(wb); } }
                    }
                    else { foreach (var wb in _workbooksToClose) { CloseTracedWorkbook(wb); } }
                }
                foreach (var wb in _workbooksToClose) { try { Marshal.ReleaseComObject(wb); } catch { } }
                _workbooksToClose.Clear();
            }

            private string ExtractWorkbookPath(string refText)
            {
                int bookEnd = refText.IndexOf(']'); if (bookEnd < 0) return null;
                int bookStart = refText.LastIndexOf('[', bookEnd); if (bookStart < 0) return null;
                string bookName = refText.Substring(bookStart + 1, bookEnd - bookStart - 1);
                string fullPath = null;
                if (refText.StartsWith("'") && refText.Substring(1).Contains(":\\")) { int pathEnd = bookStart; string path = refText.Substring(1, pathEnd - 1); fullPath = Path.Combine(path, bookName); }
                else { try { string activeWorkbookPath = _origin.Worksheet.Parent.Path; if (!string.IsNullOrEmpty(activeWorkbookPath)) { fullPath = Path.Combine(activeWorkbookPath, bookName); } } catch { } }
                return fullPath;
            }

            // Workbooks that Trace opens or closes in the background do so with Excel events
            // paused, so code inside them (Workbook_Open, Workbook_BeforeClose) doesn't run
            // just because a formula points at them. Reading cells and formulas is unaffected.
            private Excel.Workbook OpenTracedWorkbook(string bookPath)
            {
                bool events = true;
                try { events = _cachedApp.EnableEvents; } catch { }
                try { _cachedApp.EnableEvents = false; } catch { }
                try { return _cachedApp.Workbooks.Open(bookPath, ReadOnly: true, UpdateLinks: 0); }
                finally { try { _cachedApp.EnableEvents = events; } catch { } }
            }

            private void CloseTracedWorkbook(Excel.Workbook wb)
            {
                bool events = true;
                try { events = _cachedApp.EnableEvents; } catch { }
                try { _cachedApp.EnableEvents = false; } catch { }
                try { wb.Close(SaveChanges: false); } catch { }
                finally { try { _cachedApp.EnableEvents = events; } catch { } }
            }

            private Excel.Range ResolveClosedReference(string refText)
            {
                string bookPath = ExtractWorkbookPath(refText);
                if (string.IsNullOrEmpty(bookPath) || !File.Exists(bookPath)) return null;
                foreach (Excel.Workbook wbCheck in _cachedApp.Workbooks) { if (wbCheck.FullName.Equals(bookPath, StringComparison.OrdinalIgnoreCase)) { try { return _cachedApp.Range[refText]; } catch { return null; } } }
                try
                {
                    this.TopMost = true;
                    try { var wb = OpenTracedWorkbook(bookPath); _workbooksToClose.Add(wb); }
                    finally { this.TopMost = true; this.Activate(); }
                    return _cachedApp.Range[refText];
                }
                catch { return null; }
            }

            private void OnSelectedIndexChanged(object sender, EventArgs e) { if (_updating) return; if (_list.SelectedIndices.Count > 0) { _cur = _list.SelectedIndices[0]; Highlight(); } }

            private string GetCanonicalReferenceKey(PrecedentInfo p)
            {
                if (p.Range != null)
                {
                    Excel.Worksheet sheet = null;
                    Excel.Workbook book = null;
                    try
                    {
                        sheet = (Excel.Worksheet)p.Range.Worksheet;
                        book = (Excel.Workbook)sheet.Parent;
                        return $"{book.Name}|{sheet.Name}|{p.Range.Address[false, false]}";
                    }
                    catch { return p.DisplayText.ToUpperInvariant(); }
                }
                else { return p.DisplayText.ToUpperInvariant(); }
            }

            private void ColorizeFormula()
            {
                SuspendDrawing(_currentFormulaLabel);
                try
                {
                    _currentFormulaLabel.SelectAll();
                    _currentFormulaLabel.SelectionColor = Color.Black;
                    _currentFormulaLabel.SelectionBackColor = _currentFormulaLabel.BackColor;
                    _currentFormulaLabel.SelectionFont = _currentFormulaLabel.Font;
                    if (_formulaPrefixLength > 1)
                    {
                        _currentFormulaLabel.SelectionStart = 0;
                        _currentFormulaLabel.SelectionLength = _formulaPrefixLength - 1;
                        _currentFormulaLabel.SelectionFont = _underlineFont;
                        _currentFormulaLabel.SelectionColor = Color.Black;
                    }
                    int refCount = _precedents.Count(p => p.IsReference);
                    var assignedColors = new Dictionary<string, int>(refCount);
                    int colorCounter = 0;
                    foreach (var p in _precedents.Where(p => p.IsReference))
                    {
                        string key = GetCanonicalReferenceKey(p);
                        if (assignedColors.TryGetValue(key, out int existingColorIndex)) { p.ColorIndex = existingColorIndex; }
                        else { int newColorIndex = (_colorCycleStartIndex + colorCounter) % _colors.Length; p.ColorIndex = newColorIndex; assignedColors[key] = newColorIndex; colorCounter++; }
                        _currentFormulaLabel.SelectionStart = p.Start;
                        _currentFormulaLabel.SelectionLength = p.Length;
                        _currentFormulaLabel.SelectionColor = _colors[p.ColorIndex];
                        _currentFormulaLabel.SelectionFont = _currentFormulaLabel.Font;
                    }
                    _currentFormulaLabel.SelectionLength = 0;
                }
                finally { ResumeDrawing(_currentFormulaLabel); }
            }

            private string GetFormattedValue(Excel.Range rng)
            {
                if (rng == null) return "";
                if (_valueCache.TryGetValue(rng, out string cached)) return cached;

                object rawValue = null;
                string rawFmt = "";
                try
                {
                    if (rng.CountLarge > 1)
                    {
                        _valueCache[rng] = "[Array]";
                        return "[Array]";
                    }

                    rawFmt = rng.NumberFormat as string ?? "";
                    rawValue = rng.Value2;
                    if (rawValue == null)
                    {
                        try
                        {
                            var formula = rng.Formula as string;
                            if (!string.IsNullOrEmpty(formula) && formula.Contains("[") && formula.Contains("]"))
                            {
                                _valueCache[rng] = "#REF!";
                                return "#REF!";
                            }
                        }
                        catch { }
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx)
                {
                    if (comEx.HResult == unchecked((int)0x800A03EC)) { _valueCache[rng] = "#REF!"; return "#REF!"; }
                    throw;
                }
                catch { }

                if (rawValue == null) { _valueCache[rng] = ""; return ""; }
                if (IsExcelError(rawValue)) { string errorText = ConvertExcelErrorToText(rawValue); if (_valueCache.Count < 500) _valueCache[rng] = errorText; return errorText; }

                string cleanFmt = rawFmt.Replace("*", "");

                // OPTIMIZATION: Return Trimmed .ToString() for General
                if (string.Equals(cleanFmt, "General", StringComparison.OrdinalIgnoreCase))
                {
                    string fastResult = rawValue.ToString().Trim();
                    if (_valueCache.Count < 500) _valueCache[rng] = fastResult;
                    return fastResult;
                }

                // Trim removes accounting format padding (e.g. _(* )
                try { string result = _cachedApp.WorksheetFunction.Text(rawValue, cleanFmt).Trim(); if (_valueCache.Count < 500) _valueCache[rng] = result; return result; }
                catch { try { string result = rawValue.ToString().Trim(); if (_valueCache.Count < 500) _valueCache[rng] = result; return result; } catch { return ""; } }
            }

            private void PopulateList()
            {
                _list.BeginUpdate();
                try
                {
                    var functionSpans = _precedents
                        .Where(f => !f.IsReference)
                        .Select(f => (Start: f.Start, End: f.Start + f.Length))
                        .OrderBy(s => s.Start)
                        .ToList();

                    foreach (var p in _precedents)
                    {
                        int level = 0;
                        foreach (var span in functionSpans)
                        {
                            if (p.Start > span.Start && p.Start < span.End)
                                level++;
                        }
                        p.Level = level;
                    }

                    _shown.Capacity = Math.Max(_shown.Capacity, _precedents.Count + 1);
                    var outItem = BuildListViewItemForOutput();
                    _list.Items.Add(outItem);
                    _shown.Add((PrecedentInfo)outItem.Tag);
                    foreach (var p in _precedents)
                    {
                        var itm = BuildListViewItem(p);
                        _list.Items.Add(itm);
                        _shown.Add(p);
                    }
                }
                finally { _list.EndUpdate(); }
            }

            private ListViewItem BuildListViewItemForOutput()
            {
                var outPrec = new PrecedentInfo(_origin, _origin.Address[false, false] ?? "Output", 0, 1) { IsOutput = true, Level = 0 };
                string addr = outPrec.DisplayText;
                string val = GetFormattedValue(_origin);
                var item = new ListViewItem(addr) { Tag = outPrec, ForeColor = SystemColors.WindowText };
                var fSub = new ListViewItem.ListViewSubItem(item, "=") { Font = _italicFont, ForeColor = Color.Gray };
                item.SubItems.Add(fSub);
                var valSub = new ListViewItem.ListViewSubItem(item, val);
                item.SubItems.Add(valSub);
                return item;
            }

            private ListViewItem BuildListViewItem(PrecedentInfo p)
            {
                string name;
                if (p.IsReference) name = p.DisplayText;
                else { var parts = p.DisplayText.Split('('); name = string.IsNullOrWhiteSpace(parts[0]) ? "( )" : parts[0]; }
                var item = new ListViewItem(name) { Tag = p };
                if (!p.IsReference) item.ForeColor = Color.Gray;
                string fText = IsDrillable(p) ? "↪" : (!p.IsReference ? "fx" : "");
                var fSub = new ListViewItem.ListViewSubItem(item, fText) { ForeColor = Color.Gray };
                if (fText == "fx") fSub.Font = _italicFont;
                item.SubItems.Add(fSub);
                string val;
                if (p.IsReference)
                {
                    if (p.Range != null) val = GetFormattedValue(p.Range);
                    else
                    {
                        object rawVal = null;
                        try { rawVal = _cachedApp.Evaluate(p.DisplayText); } catch { }
                        if (rawVal == null) val = "#REF!";
                        else if (IsExcelError(rawVal)) val = ConvertExcelErrorToText(rawVal);
                        else val = rawVal.ToString().Trim();
                    }
                }
                else val = p.FunctionVal.Trim();
                var valSub = new ListViewItem.ListViewSubItem(item, val);
                if (!p.IsReference) valSub.ForeColor = Color.Gray;
                item.SubItems.Add(valSub);
                return item;
            }

            private void OnDrawItem(object s, DrawListViewItemEventArgs e)
            {
                Color backColor = e.Item.BackColor;
                if (e.Item.Selected)
                {
                    var info = e.Item.Tag as PrecedentInfo;
                    if (info != null)
                    {
                        if (info.IsOutput) backColor = UIColors.Accent;
                        else if (info.IsReference && info.ColorIndex >= 0) backColor = _colors[info.ColorIndex % _colors.Length];
                        else backColor = UIColors.Accent;
                    }
                    else backColor = UIColors.Accent;
                }
                using (var brush = new SolidBrush(backColor)) e.Graphics.FillRectangle(brush, e.Bounds);
            }

            private void OnDrawSub(object s, DrawListViewSubItemEventArgs e)
            {
                var info = (PrecedentInfo)e.Item.Tag;
                if (info == null) { e.DrawDefault = true; return; }
                Color txt = e.SubItem.ForeColor;
                if (e.Item.Selected) txt = Color.White;
                switch (e.ColumnIndex)
                {
                    case 0:
                        int startLevel = info.IsOutput ? 1 : 0;
                        using (var pen = new Pen(Color.Silver))
                        {
                            for (int l = startLevel; l <= info.Level; l++)
                            {
                                int x = e.Bounds.Left + 6 + l * INDENT;
                                e.Graphics.DrawLine(pen, x, e.Bounds.Top, x, e.Bounds.Bottom - 1);
                            }
                        }
                        int offset = 6 + info.Level * INDENT;
                        var rect = new Rectangle(e.Bounds.Left + offset, e.Bounds.Top, e.Bounds.Width - offset, e.Bounds.Height);
                        // EndEllipsis adds "..." to long text
                        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.Item.Font, rect, txt, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                        break;
                    case 1:
                        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, e.Bounds, txt, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);
                        break;
                    case 2:
                        var valRect = new Rectangle(e.Bounds.Left + 4, e.Bounds.Top, e.Bounds.Width - 4, e.Bounds.Height);
                        // EndEllipsis adds "..." to long text
                        TextRenderer.DrawText(e.Graphics, e.SubItem.Text, e.SubItem.Font, valRect, txt, TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);
                        break;
                    default:
                        e.DrawDefault = true;
                        break;
                }
            }

            private int NextRef(int i) { int n = _shown.Count, j = i; do { j = (j + 1) % n; } while (!_shown[j].IsReference && j != i); return j; }
            private int PrevRef(int i) { int n = _shown.Count, j = i; do { j = (j - 1 + n) % n; } while (!_shown[j].IsReference && j != i); return j; }
            private void RestoreHidden() { foreach (var r in _unhiddenRows) { try { r.Hidden = true; } catch { } } foreach (var c in _unhiddenCols) { try { c.Hidden = true; } catch { } } _unhiddenRows.Clear(); _unhiddenCols.Clear(); }

            private void SafeSelect(Excel.Range rng)
            {
                bool scrn = _cachedApp.ScreenUpdating;
                bool evts = _cachedApp.EnableEvents;
                try
                {
                    _cachedApp.ScreenUpdating = false;
                    _cachedApp.EnableEvents = false;

                    Excel.Range target = rng;
                    target.Worksheet.Activate();

                    // 1. Select the range (Scroll=true puts it top-left initially)
                    _cachedApp.Goto(target, true);

                    // 2. Restore Old Logic: Calculate center of screen and adjust scroll
                    try
                    {
                        var window = _cachedApp.ActiveWindow;
                        Excel.Range firstCell = (Excel.Range)target.Cells[1, 1];

                        int targetRow = firstCell.Row;
                        int targetCol = firstCell.Column;

                        // Get visible area dimensions
                        int visibleRows = window.VisibleRange.Rows.Count;
                        int visibleCols = window.VisibleRange.Columns.Count;

                        // Calculate the row/col to scroll to so target is in the middle
                        int scrollRow = Math.Max(1, targetRow - visibleRows / 2);
                        int scrollCol = Math.Max(1, targetCol - visibleCols / 2);

                        window.ScrollRow = scrollRow;
                        window.ScrollColumn = scrollCol;

                        // Cleanup
                        Marshal.ReleaseComObject(firstCell);
                    }
                    catch
                    {
                        // If window sizing fails (e.g. minimized), strictly standard Goto applies
                    }
                }
                catch { }
                finally
                {
                    _cachedApp.EnableEvents = evts;
                    _cachedApp.ScreenUpdating = scrn;
                }
            }

            private void UnhideForRange(Excel.Range rng)
            {
                if (rng == null) return;
                try
                {
                    var entireCols = rng.EntireColumn;
                    dynamic colHidden = entireCols.Hidden;
                    if (!(colHidden is bool && (bool)colHidden == false))
                    {
                        entireCols.Hidden = false;
                        _unhiddenCols.Add(entireCols);
                    }

                    var entireRows = rng.EntireRow;
                    dynamic rowHidden = entireRows.Hidden;
                    if (!(rowHidden is bool && (bool)rowHidden == false))
                    {
                        entireRows.Hidden = false;
                        _unhiddenRows.Add(entireRows);
                    }
                }
                catch { }
            }

            private static Color Lighten(Color c, double f = 0.7)
            {
                int r = c.R + (int)((255 - c.R) * f);
                int g = c.G + (int)((255 - c.G) * f);
                int b = c.B + (int)((255 - c.B) * f);
                return Color.FromArgb(r, g, b);
            }

            private Rectangle GetRangeScreenCoordinates(Excel.Range range)
            {
                if (range == null) return Rectangle.Empty;
                try
                {
                    var app = _cachedApp; // Use cached app
                    var window = app.ActiveWindow;
                    double left = (double)range.Left, top = (double)range.Top, width = (double)range.Width, height = (double)range.Height;
                    int screenLeft = window.ActivePane.PointsToScreenPixelsX((int)left);
                    int screenTop = window.ActivePane.PointsToScreenPixelsY((int)top);
                    int screenRight = window.ActivePane.PointsToScreenPixelsX((int)(left + width));
                    int screenBottom = window.ActivePane.PointsToScreenPixelsY((int)(top + height));
                    return new Rectangle(screenLeft, screenTop, screenRight - screenLeft, screenBottom - screenTop);
                }
                catch { return Rectangle.Empty; }
            }

            private void AutoMoveFormAwayFromRange(Excel.Range selectedRange)
            {
                Rectangle formBounds = this.Bounds;
                Rectangle rangeScreenCoords = GetRangeScreenCoordinates(selectedRange);
                if (!formBounds.IntersectsWith(rangeScreenCoords))
                {
                    if (_hasMovedAway && _originalPosition != Point.Empty)
                    {
                        Rectangle originalBounds = new Rectangle(_originalPosition, this.Size);
                        if (!originalBounds.IntersectsWith(rangeScreenCoords)) { this.Location = _originalPosition; _hasMovedAway = false; }
                    }
                    return;
                }
                if (!_hasMovedAway) _originalPosition = this.Location;
                Rectangle workingArea = Screen.FromControl(this).WorkingArea;
                Point[] candidatePositions = new Point[]
                {
                    new Point(rangeScreenCoords.Right + 10, this.Location.Y),
                    new Point(rangeScreenCoords.Left - this.Width - 10, this.Location.Y),
                    new Point(this.Location.X, rangeScreenCoords.Bottom + 10),
                    new Point(this.Location.X, rangeScreenCoords.Top - this.Height - 10),
                    new Point(rangeScreenCoords.Right + 10, rangeScreenCoords.Top),
                    new Point(rangeScreenCoords.Left - this.Width - 10, rangeScreenCoords.Top)
                };
                Point bestPosition = this.Location;
                double minDistance = double.MaxValue;
                foreach (Point candidatePos in candidatePositions)
                {
                    Rectangle potentialBounds = new Rectangle(candidatePos, this.Size);
                    if (!workingArea.Contains(potentialBounds) || potentialBounds.IntersectsWith(rangeScreenCoords)) continue;
                    double distance = Math.Sqrt(Math.Pow(candidatePos.X - this.Location.X, 2) + Math.Pow(candidatePos.Y - this.Location.Y, 2));
                    if (distance < minDistance) { minDistance = distance; bestPosition = candidatePos; }
                }
                if (bestPosition != this.Location) { this.Location = bestPosition; _hasMovedAway = true; }
            }

            private void Highlight()
            {
                if (_cur < 0 || _cur >= _shown.Count) return;
                var sel = _shown[_cur];
                if (sel.IsReference && sel.Range == null) { Excel.Range newRange = ResolveClosedReference(sel.DisplayText); if (newRange != null) { sel.UpdateRange(newRange); _list.Items[_cur].SubItems[2].Text = GetFormattedValue(newRange); } }
                _updating = true; RestoreHidden();
                if (_prevStart >= 0) { _currentFormulaLabel.Select(_prevStart, _prevLen); _currentFormulaLabel.SelectionBackColor = _currentFormulaLabel.BackColor; }
                foreach (ListViewItem li in _list.Items) li.BackColor = _list.BackColor;

                PrecedentInfo grp = null;
                if (!sel.IsOutput && !sel.IsReference) grp = sel;
                else if (!sel.IsOutput && sel.IsReference)
                {
                    foreach (var p in _precedents.Where(p => !p.IsReference))
                    {
                        if (sel.Start >= p.Start && (sel.Start + sel.Length) <= (p.Start + p.Length))
                        {
                            grp = p;
                            break;
                        }
                    }
                }

                if (grp != null)
                {
                    Color groupBackColor = _groupHL;
                    if (sel.IsReference && sel.ColorIndex >= 0) { Color baseColor = _colors[sel.ColorIndex % _colors.Length]; groupBackColor = Lighten(baseColor, 0.9); }
                    for (int i = 0; i < _shown.Count; i++)
                    {
                        var p = _shown[i];
                        bool inside = (!p.IsReference && p == grp) || (p.Start >= grp.Start && (p.Start + p.Length) <= (grp.Start + grp.Length));
                        if (inside) _list.Items[i].BackColor = groupBackColor;
                    }
                }

                if (!sel.IsOutput) { _currentFormulaLabel.Select(sel.Start, sel.Length); Color hl = sel.IsReference && sel.ColorIndex >= 0 ? Lighten(_colors[sel.ColorIndex % _colors.Length]) : _groupHL; _currentFormulaLabel.SelectionBackColor = hl; _currentFormulaLabel.SelectionLength = 0; _prevStart = sel.Start; _prevLen = sel.Length; }
                else { _prevStart = -2; _prevLen = 0; }
                _list.Items[_cur].Selected = true; _list.EnsureVisible(_cur); _updating = false;

                if (sel.IsReference && sel.Range != null)
                {
                    BeginInvoke((MethodInvoker)(() => {
                        if (_isClosing) return;
                        try
                        {
                            UnhideForRange(sel.Range);
                            SafeSelect(sel.Range);
                            AutoMoveFormAwayFromRange(sel.Range);
                            this.Activate();
                        }
                        catch { }
                    }));
                }
            }
        }
    }
}