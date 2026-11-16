// ShapeUtility.MeasureMove.cs
// .NET Framework 4.8 / C# 7.3 (VSTO PowerPoint)
//
// Uses PragmaticsUI (Dpi/Theme/PragmaticsForm):
//   - The form derives from PragmaticsForm
//   - DPI-aware layout using Theme paddings and a TableLayoutPanel
//
// Features:
// - Orange dashed lines (0.85f) between selected shapes and:
//      * other shapes (when X/Y projections overlap)
//      * slide edges (top/bottom/left/right) when projections overlap
//    measuring the perpendicular gap (no line if zero/overlap).
// - Inch labels sit on the target edge and autofit to text.
// - Modeless TopMost form with two inputs (Horizontal/Vertical inches).
// - Native titlebar with Close (X). Press Enter to apply move, Esc to close.
// - Bottom live preview (no brackets): 0.00" Right, 0.00" Down updates as you type.
// - On close: timer stops/disposes; temp shapes are removed.
// - Live move: Shapes move in real-time based on the initial state of the session. Escape reverts all changes.
//
// Usage: ShapeUtility.MeasureMove(Globals.ThisAddIn.Application.ActiveWindow.Selection);

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using SD = System.Drawing;
using Mso = Microsoft.Office.Core;
using Ppt = Microsoft.Office.Interop.PowerPoint;
using OfficePragmatics_Excel.Settings; // PragmaticsUI (Dpi/Theme/PragmaticsForm)

namespace OfficePragmatics_PowerPoint.Shape
{
    public static partial class ShapeUtility
    {
        private static MeasureMoveForm _form;

        public static void MeasureMove(Ppt.Selection sel)
        {
            try
            {
                var app = sel?.Application;
                var wnd = app?.ActiveWindow;
                if (app == null || wnd == null) return;

                if (_form != null && !_form.IsDisposed)
                {
                    _form.TryRebindToCurrentSlide();
                    _form.BringToFront();
                    _form.TopMost = true; // re-assert
                    return;
                }

                _form = new MeasureMoveForm(app);
                _form.FormClosed += (s, e) => { _form = null; };
                _form.Show(); // modeless
            }
            catch
            {
                // silent fail
            }
        }

        // --------------------------- Inner Form ---------------------------

        private sealed class MeasureMoveForm : PragmaticsForm
        {
            private readonly Ppt.Application _app;

            private TextBox _txH;
            private TextBox _txV;
            private Label _live;
            private Panel _mainPanel;
            private Panel _bottomBar;
            private TableLayoutPanel _grid;
            private ToolTip _tips;

            private Timer _tick;

            private Ppt.Slide _slide;
            private bool _isClosing = false;
            private bool _isUpdatingText = false; // Prevents re-entrant TextChanged events

            // -- State Management for Move Sessions --
            private Dictionary<int, SD.PointF> _originalPositions;
            private string _lastPositionFingerprint = "";
            private string _lastSelectionIdFingerprint = "";

            private const string TAG_PREFIX = "_OP_MEASURE_";
            private const double PT_PER_IN = 72.0;
            private const float EPS = 0.5f;

            private const float LINE_WEIGHT_DEFAULT = 0.85f;
            private const float LBL_BORDER_DEFAULT = 0.64f;

            public MeasureMoveForm(Ppt.Application app)
            {
                _app = app;

                // CRITICAL: Handle 4K/High DPI properly
                this.AutoScaleMode = AutoScaleMode.Font;

                Text = "Measure & Move";
                FormBorderStyle = FormBorderStyle.SizableToolWindow;
                ShowInTaskbar = false;
                TopMost = true;
                KeyPreview = true;

                ControlBox = true;
                MinimizeBox = false;
                MaximizeBox = false;

                // Use slightly wider default to accommodate font scaling
                this.ClientSize = new SD.Size(Theme.Px(320), Theme.Px(190));
                this.MinimumSize = new SD.Size(Theme.Px(300), Theme.Px(190));

                _bottomBar = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = Theme.RowHeight,
                    Padding = new Padding(Theme.PadMd, Theme.PadSm, Theme.PadMd, Theme.PadSm),
                    BackColor = Theme.SurfaceAlt
                };
                _live = new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = SD.ContentAlignment.MiddleLeft,
                    AutoEllipsis = true,
                    Font = new SD.Font("Segoe UI", 9f, SD.FontStyle.Bold)
                };
                _bottomBar.Controls.Add(_live);
                Controls.Add(_bottomBar);

                _mainPanel = new Panel { Dock = DockStyle.Fill, Padding = new Padding(Theme.PadMd), BackColor = Theme.Surface };
                Theme.DoubleBuffer(_mainPanel);
                Theme.Cardize(_mainPanel);
                Controls.Add(_mainPanel);

                _grid = new TableLayoutPanel
                {
                    Dock = DockStyle.Top,
                    AutoSize = true,
                    AutoSizeMode = AutoSizeMode.GrowAndShrink,
                    ColumnCount = 2,
                    RowCount = 2,
                    Padding = new Padding(Theme.PadSm)
                };

                // FLUID LAYOUT: Auto-size label column, fluid 100% width for inputs
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
                _grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

                _grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                _grid.RowStyles.Add(new RowStyle(SizeType.AutoSize));
                _mainPanel.Controls.Add(_grid);

                var lblH = new Label { Text = "Horizontal (in):", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(Theme.PadSm, Theme.PadSm + 4, Theme.PadSm, Theme.PadSm) };
                var lblV = new Label { Text = "Vertical (in):", AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(Theme.PadSm, Theme.PadSm + 4, Theme.PadSm, Theme.PadSm) };

                // DOCK.FILL ensures textboxes stretch on 4K/Low Res rather than being fixed width
                _txH = new TextBox { Dock = DockStyle.Fill, Text = "0", Margin = new Padding(Theme.PadSm) };
                _txV = new TextBox { Dock = DockStyle.Fill, Text = "0", Margin = new Padding(Theme.PadSm) };

                _grid.Controls.Add(lblH, 0, 0);
                _grid.Controls.Add(_txH, 1, 0);
                _grid.Controls.Add(lblV, 0, 1);
                _grid.Controls.Add(_txV, 1, 1);

                _tips = new ToolTip { ShowAlways = true };
                _tips.SetToolTip(_txH, "Inches: negative = Left, positive = Right");
                _tips.SetToolTip(_txV, "Inches: negative = Up, positive = Down");

                _txH.KeyDown += OnMoveEnter;
                _txV.KeyDown += OnMoveEnter;
                _txH.TextChanged += OnInputTextChanged;
                _txV.TextChanged += OnInputTextChanged;

                _tick = new Timer { Interval = 100 };
                _tick.Tick += (s, e) => RefreshAll();
                _tick.Start();

                FormClosing += OnFormClosingStopEverything;
                FormClosed += OnFormClosedCleanup;

                RefreshAll();
            }

            protected override void ApplyScaledMetrics()
            {
                base.ApplyScaledMetrics();
                try
                {
                    TopMost = true;

                    if (_bottomBar != null)
                    {
                        _bottomBar.Height = Theme.RowHeight;
                        _bottomBar.Padding = new Padding(Theme.PadMd, Theme.PadSm, Theme.PadMd, Theme.PadSm);
                        _bottomBar.BackColor = Theme.SurfaceAlt;
                    }
                    if (_mainPanel != null)
                    {
                        _mainPanel.Padding = new Padding(Theme.PadMd);
                        _mainPanel.BackColor = Theme.Surface;
                    }
                    if (_grid != null)
                    {
                        _grid.Padding = new Padding(Theme.PadSm);
                        // Removed manual column/textbox resizing logic;
                        // DockStyle.Fill and SizeType.Percent handle DPI scaling automatically now.
                    }
                }
                catch { }
            }

            private void OnFormClosingStopEverything(object sender, FormClosingEventArgs e)
            {
                _isClosing = true;
                try { _tick?.Stop(); } catch { }
                try { _tick?.Dispose(); } catch { }
                _tick = null;
                RevertMove();
            }

            private void OnFormClosedCleanup(object sender, FormClosedEventArgs e)
            {
                try { DeleteMeasureShapesOnSlide(_slide); } catch { }
            }

            protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
            {
                if (keyData == Keys.Escape)
                {
                    RevertMove();
                    Close();
                    return true;
                }
                if (keyData == Keys.Enter)
                {
                    ApplyMove();
                    return true;
                }
                return base.ProcessCmdKey(ref msg, keyData);
            }

            public void TryRebindToCurrentSlide()
            {
                if (_isClosing || IsDisposed) return;
                try
                {
                    var wnd = _app.ActiveWindow;
                    if (wnd?.View?.Slide is Ppt.Slide s)
                    {
                        if (!ReferenceEquals(_slide, s))
                        {
                            DeleteMeasureShapesOnSlide(_slide);
                            _slide = s;
                            _lastPositionFingerprint = "";
                        }
                    }
                    else
                    {
                        DeleteMeasureShapesOnSlide(_slide);
                        _slide = null;
                    }
                }
                catch { }
            }

            private void OnMoveEnter(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    ApplyMove();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }

            private void ApplyMove()
            {
                if (_isClosing || IsDisposed) return;
                try
                {
                    FinalizeMove();
                    _lastSelectionIdFingerprint = "";
                    _lastPositionFingerprint = "";
                    RefreshAll();
                }
                catch { /* silent */ }
            }

            private void UpdateLivePreview()
            {
                if (_isClosing || IsDisposed) return;

                double dxIn = ParseInches(_txH.Text);
                double dyIn = ParseInches(_txV.Text);

                string lr = dxIn > 0 ? "Right" : (dxIn < 0 ? "Left" : "Right");
                string ud = dyIn > 0 ? "Down" : (dyIn < 0 ? "Up" : "Down");
                string txt = $"{Math.Abs(dxIn):0.00}\" {lr}, {Math.Abs(dyIn):0.00}\" {ud}";

                try { _live.Text = txt; } catch { }
            }

            private void RefreshAll()
            {
                if (_isClosing || IsDisposed) return;
                try
                {
                    TryRebindToCurrentSlide();
                    if (_slide == null) return;

                    var sel = _app.ActiveWindow?.Selection;
                    bool validSel = (sel != null &&
                                     sel.Type == Ppt.PpSelectionType.ppSelectionShapes &&
                                     sel.ShapeRange != null &&
                                     sel.ShapeRange.Count > 0);

                    var shapeRange = validSel ? sel.ShapeRange : null;
                    string newIdFp = BuildSelectionIdFingerprint(shapeRange);

                    if (newIdFp != _lastSelectionIdFingerprint)
                    {
                        _lastSelectionIdFingerprint = newIdFp;
                        FinalizeMove();

                        if (validSel)
                        {
                            CaptureOriginalPositions(shapeRange);
                            SafeSetText("0");
                        }
                    }

                    string newPositionFp = BuildSelectionFingerprint(shapeRange);
                    if (newPositionFp != _lastPositionFingerprint)
                    {
                        _lastPositionFingerprint = newPositionFp;
                        DeleteMeasureShapesOnSlide(_slide);
                        var lines = ComputeMeasureLines(_slide, shapeRange);
                        DrawMeasureShapes(_slide, lines);
                    }
                }
                catch { /* swallow */ }
            }

            // ---------- Live Move Logic ----------

            private void OnInputTextChanged(object sender, EventArgs e)
            {
                if (_isClosing || IsDisposed || _isUpdatingText) return;
                UpdateLivePreview();

                try
                {
                    if (_originalPositions == null) return;

                    var sel = _app.ActiveWindow?.Selection?.ShapeRange;
                    if (sel == null) return;

                    double dxIn = ParseInches(_txH.Text);
                    double dyIn = ParseInches(_txV.Text);
                    float dxPt = (float)(dxIn * PT_PER_IN);
                    float dyPt = (float)(dyIn * PT_PER_IN);

                    foreach (var entry in _originalPositions)
                    {
                        try
                        {
                            var shape = sel.Cast<Ppt.Shape>().FirstOrDefault(s => s.Id == entry.Key);
                            if (shape != null)
                            {
                                shape.Left = entry.Value.X + dxPt;
                                shape.Top = entry.Value.Y + dyPt;
                            }
                        }
                        catch { /* Shape might have been deleted; ignore */ }
                    }
                }
                catch
                {
                    FinalizeMove();
                }
            }

            private void CaptureOriginalPositions(Ppt.ShapeRange shapeRange)
            {
                _originalPositions = new Dictionary<int, SD.PointF>();
                try
                {
                    foreach (Ppt.Shape shape in shapeRange)
                    {
                        _originalPositions[shape.Id] = new SD.PointF(shape.Left, shape.Top);
                    }
                }
                catch { _originalPositions = null; }
            }

            private void RevertMove()
            {
                if (_originalPositions == null || _originalPositions.Count == 0) return;

                try
                {
                    var shapes = _slide?.Shapes;
                    if (shapes == null) return;

                    foreach (var entry in _originalPositions)
                    {
                        try
                        {
                            var shape = shapes.Cast<Ppt.Shape>().FirstOrDefault(s => s.Id == entry.Key);
                            if (shape != null)
                            {
                                shape.Left = entry.Value.X;
                                shape.Top = entry.Value.Y;
                            }
                        }
                        catch { /* Ignore if shape is gone */ }
                    }
                }
                finally
                {
                    FinalizeMove();
                }
            }

            private void FinalizeMove()
            {
                _originalPositions = null;
            }

            private void SafeSetText(string text)
            {
                if (_isUpdatingText) return;
                _isUpdatingText = true;
                try
                {
                    _txH.Text = text;
                    _txV.Text = text;
                }
                finally
                {
                    _isUpdatingText = false;
                    UpdateLivePreview();
                }
            }

            // ---------- Measure logic ----------

            private struct ShapeBounds
            {
                public float L, T, R, B, W, H;
            }

            private static ShapeBounds B(Ppt.Shape s)
            {
                return new ShapeBounds
                {
                    L = s.Left,
                    T = s.Top,
                    W = s.Width,
                    H = s.Height,
                    R = s.Left + s.Width,
                    B = s.Top + s.Height
                };
            }

            private static bool Overlap1D(float a1, float a2, float b1, float b2)
            {
                return Math.Max(a1, b1) < Math.Min(a2, b2);
            }

            private static float Mid(float a, float b) => (a + b) * 0.5f;

            private sealed class MeasureLine
            {
                public float X1, Y1, X2, Y2;   // slide points (line ends)
                public float LengthPt;         // length in points
                public float LabelX, LabelY;   // label on target edge (other shape or slide edge)
            }

            private List<MeasureLine> ComputeMeasureLines(Ppt.Slide slide, Ppt.ShapeRange selected)
            {
                var output = new List<MeasureLine>();
                try
                {
                    var all = slide.Shapes;
                    if (selected == null || selected.Count == 0) return output;

                    float sw = slide.Parent.PageSetup.SlideWidth;
                    float sh = slide.Parent.PageSetup.SlideHeight;

                    var selList = new List<Ppt.Shape>();
                    foreach (Ppt.Shape s in selected)
                        selList.Add(s);

                    var others = new List<Ppt.Shape>();
                    foreach (Ppt.Shape s in all)
                    {
                        if (!IsShapeVisible(s)) continue;
                        if (s.Name != null && s.Name.StartsWith(TAG_PREFIX, StringComparison.Ordinal)) continue;
                        if (selList.Any(ss => ReferenceEquals(ss, s))) continue;
                        others.Add(s);
                    }

                    var seen = new HashSet<string>();

                    foreach (var a in selList)
                    {
                        var ba = B(a);

                        // ---- 1) To other shapes ----
                        foreach (var b in others)
                        {
                            var bb = B(b);

                            // vertical gaps (X overlap)
                            if (Overlap1D(ba.L, ba.R, bb.L, bb.R))
                            {
                                if (ba.B <= bb.T) // a above b
                                {
                                    float overlapL = Math.Max(ba.L, bb.L);
                                    float overlapR = Math.Min(ba.R, bb.R);
                                    if (overlapR - overlapL > EPS)
                                    {
                                        float midX = Mid(overlapL, overlapR);
                                        float y1 = ba.B, y2 = bb.T;
                                        if (y2 - y1 > EPS)
                                            AddUnique(output, seen, midX, y1, midX, y2, midX, bb.T);
                                    }
                                }
                                else if (bb.B <= ba.T) // b above a
                                {
                                    float overlapL = Math.Max(ba.L, bb.L);
                                    float overlapR = Math.Min(ba.R, bb.R);
                                    if (overlapR - overlapL > EPS)
                                    {
                                        float midX = Mid(overlapL, overlapR);
                                        float y1 = bb.B, y2 = ba.T;
                                        if (y2 - y1 > EPS)
                                            AddUnique(output, seen, midX, y1, midX, y2, midX, bb.B);
                                    }
                                }
                            }

                            // horizontal gaps (Y overlap)
                            if (Overlap1D(ba.T, ba.B, bb.T, bb.B))
                            {
                                if (ba.R <= bb.L) // a left of b
                                {
                                    float overlapT = Math.Max(ba.T, bb.T);
                                    float overlapB = Math.Min(ba.B, bb.B);
                                    if (overlapB - overlapT > EPS)
                                    {
                                        float midY = Mid(overlapT, overlapB);
                                        float x1 = ba.R, x2 = bb.L;
                                        if (x2 - x1 > EPS)
                                            AddUnique(output, seen, x1, midY, x2, midY, bb.L, midY);
                                    }
                                }
                                else if (bb.R <= ba.L) // b left of a
                                {
                                    float overlapT = Math.Max(ba.T, bb.T);
                                    float overlapB = Math.Min(ba.B, bb.B);
                                    if (overlapB - overlapT > EPS)
                                    {
                                        float midY = Mid(overlapT, overlapB);
                                        float x1 = bb.R, x2 = ba.L;
                                        if (x2 - x1 > EPS)
                                            AddUnique(output, seen, x1, midY, x2, midY, bb.R, midY);
                                    }
                                }
                            }
                        }

                        // ---- 2) To slide edges ----
                        float overlapL_slide = Math.Max(ba.L, 0);
                        float overlapR_slide = Math.Min(ba.R, sw);
                        if ((overlapR_slide - overlapL_slide) > EPS)
                        {
                            float midX = Mid(overlapL_slide, overlapR_slide);
                            if (ba.T > EPS) AddUnique(output, seen, midX, 0f, midX, ba.T, midX, 0f);
                            if (sh - ba.B > EPS) AddUnique(output, seen, midX, ba.B, midX, sh, midX, sh);
                        }

                        float overlapT_slide = Math.Max(ba.T, 0);
                        float overlapB_slide = Math.Min(ba.B, sh);
                        if ((overlapB_slide - overlapT_slide) > EPS)
                        {
                            float midY = Mid(overlapT_slide, overlapB_slide);
                            if (ba.L > EPS) AddUnique(output, seen, 0f, midY, ba.L, midY, 0f, midY);
                            if (sw - ba.R > EPS) AddUnique(output, seen, ba.R, midY, sw, midY, sw, midY);
                        }
                    }
                }
                catch { }

                return output;
            }

            private static void AddUnique(List<MeasureLine> list, HashSet<string> seen, float x1, float y1, float x2, float y2, float labelX, float labelY)
            {
                float len = (float)Math.Sqrt(Math.Pow(x2 - x1, 2) + Math.Pow(y2 - y1, 2));
                if (len < EPS) return;

                bool vertical = Math.Abs(x1 - x2) < EPS;
                string key = vertical
                    ? $"V|{(int)x1}|{(int)Math.Min(y1, y2)}|{(int)Math.Max(y1, y2)}"
                    : $"H|{(int)y1}|{(int)Math.Min(x1, x2)}|{(int)Math.Max(x1, x2)}";

                if (seen.Add(key))
                {
                    list.Add(new MeasureLine { X1 = x1, Y1 = y1, X2 = x2, Y2 = y2, LengthPt = len, LabelX = labelX, LabelY = labelY });
                }
            }

            private static bool IsShapeVisible(Ppt.Shape s)
            {
                try
                {
                    return s.Visible == Mso.MsoTriState.msoTrue &&
                           s.Width > 0.1f && s.Height > 0.1f;
                }
                catch { return false; }
            }

            private void DrawMeasureShapes(Ppt.Slide slide, List<MeasureLine> lines)
            {
                if (_isClosing || IsDisposed || slide == null || lines == null) return;
                try
                {
                    foreach (var ml in lines)
                    {
                        var shp = slide.Shapes.AddLine(ml.X1, ml.Y1, ml.X2, ml.Y2);
                        shp.Name = TAG_PREFIX + "LINE_" + Guid.NewGuid().ToString("N");
                        shp.Line.ForeColor.RGB = SD.ColorTranslator.ToOle(SD.Color.Orange);
                        shp.Line.DashStyle = Mso.MsoLineDashStyle.msoLineDash;
                        shp.Line.Weight = LINE_WEIGHT_DEFAULT;
                        shp.ZOrder(Mso.MsoZOrderCmd.msoBringToFront);

                        string txt = $"{(ml.LengthPt / PT_PER_IN):0.###}\"";
                        var label = slide.Shapes.AddTextbox(Mso.MsoTextOrientation.msoTextOrientationHorizontal, ml.LabelX - 5f, ml.LabelY - 7f, 10f, 14f);
                        label.Name = TAG_PREFIX + "LBL_" + Guid.NewGuid().ToString("N");

                        try
                        {
                            var tf2 = label.TextFrame2;
                            tf2.TextRange.Text = txt;
                            tf2.TextRange.Font.Size = 9;
                            tf2.TextRange.Font.Fill.ForeColor.RGB = SD.ColorTranslator.ToOle(SD.Color.Black);
                            tf2.WordWrap = Mso.MsoTriState.msoFalse;
                            tf2.MarginLeft = 1.0f; tf2.MarginRight = 1.0f; tf2.MarginTop = 0.5f; tf2.MarginBottom = 0.5f;
                            tf2.AutoSize = Mso.MsoAutoSize.msoAutoSizeShapeToFitText;
                        }
                        catch
                        {
                            var tf = label.TextFrame;
                            tf.TextRange.Text = txt;
                            tf.WordWrap = Mso.MsoTriState.msoFalse;
                            tf.MarginLeft = 1.0f; tf.MarginRight = 1.0f; tf.MarginTop = 0.5f; tf.MarginBottom = 0.5f;
                            tf.AutoSize = Ppt.PpAutoSize.ppAutoSizeShapeToFitText;
                        }

                        label.Fill.ForeColor.RGB = SD.ColorTranslator.ToOle(SD.Color.White);
                        label.Fill.Transparency = 0.25f;
                        label.Line.ForeColor.RGB = SD.ColorTranslator.ToOle(SD.Color.Orange);
                        label.Line.Weight = LBL_BORDER_DEFAULT;
                        label.ZOrder(Mso.MsoZOrderCmd.msoBringToFront);

                        try
                        {
                            label.Left = ml.LabelX - label.Width * 0.5f;
                            label.Top = ml.LabelY - label.Height * 0.5f;
                        }
                        catch { }
                    }
                }
                catch { }
            }

            private void DeleteMeasureShapesOnSlide(Ppt.Slide slide)
            {
                if (slide == null) return;
                try
                {
                    var shapesToDelete = new List<Ppt.Shape>();
                    foreach (Ppt.Shape s in slide.Shapes)
                    {
                        try { if ((s.Name ?? "").StartsWith(TAG_PREFIX, StringComparison.Ordinal)) shapesToDelete.Add(s); } catch { }
                    }
                    foreach (var s in shapesToDelete)
                    {
                        try { s.Delete(); } catch { }
                    }
                }
                catch { }
            }

            // ---------- Fingerprinting ----------
            private string BuildSelectionFingerprint(Ppt.ShapeRange range)
            {
                if (range == null || range.Count == 0) return "NONE";
                try
                {
                    var parts = new List<string>(range.Count * 5);
                    foreach (Ppt.Shape s in range)
                    {
                        parts.Add($"{s.Id}:{(int)s.Left}:{(int)s.Top}:{(int)s.Width}:{(int)s.Height}");
                    }
                    var slideIndex = _app.ActiveWindow?.View?.Slide?.SlideIndex ?? 0;
                    parts.Add($"S{slideIndex}");
                    return string.Join("|", parts);
                }
                catch { return Guid.NewGuid().ToString("N"); }
            }

            private string BuildSelectionIdFingerprint(Ppt.ShapeRange range)
            {
                if (range == null || range.Count == 0) return "NONE";
                try
                {
                    var ids = new List<int>(range.Count);
                    foreach (Ppt.Shape s in range)
                    {
                        ids.Add(s.Id);
                    }
                    ids.Sort();
                    var slideIndex = _app.ActiveWindow?.View?.Slide?.SlideIndex ?? 0;
                    return $"S{slideIndex}|" + string.Join(",", ids);
                }
                catch { return Guid.NewGuid().ToString("N"); }
            }

            // ---------- Unit parsing (inches) ----------
            private static double ParseInches(string s)
            {
                if (string.IsNullOrWhiteSpace(s)) return 0;
                s = s.Trim().ToLowerInvariant();

                double factorToIn = 1.0;
                if (s.EndsWith("in") || s.EndsWith("\"")) { s = s.Substring(0, s.Length - (s.EndsWith("\"") ? 1 : 2)); }
                else if (s.EndsWith("pt")) { factorToIn = 1.0 / PT_PER_IN; s = s.Substring(0, s.Length - 2); }
                else if (s.EndsWith("cm")) { factorToIn = 1.0 / 2.54; s = s.Substring(0, s.Length - 2); }
                else if (s.EndsWith("mm")) { factorToIn = 1.0 / 25.4; s = s.Substring(0, s.Length - 2); }

                if (!double.TryParse(s.Trim(), System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out double val))
                {
                    double.TryParse(s.Trim(), out val);
                }
                return val * factorToIn;
            }
        }
    }
}