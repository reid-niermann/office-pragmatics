// ColorDropperUtility_Word.cs
// Word VSTO version (C# 7.3 / .NET Framework 4.8)

using System;
using System.Diagnostics;
using System.Drawing;              // ColorTranslator.ToOle
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace Office_Pragmatics_Word
{
    public static class ColorDropperUtility
    {
        #region Win32 / Hook Definitions
        private enum PickMode { Fill, Outline, Font }
        private static PickMode _currentMode = PickMode.Fill;
        private static readonly LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookId = IntPtr.Zero;

        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int VK_O = (int)Keys.O;
        private const int VK_F = (int)Keys.F;
        private const int VK_H = (int)Keys.H;
        private const int VK_ESCAPE = (int)Keys.Escape;
        private const int VK_LEFT = (int)Keys.Left;
        private const int VK_RIGHT = (int)Keys.Right;
        private const int VK_LBUTTON = 0x01; // left mouse button

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        [DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        private const uint MOUSEEVENTF_LEFTDOWN = 0x02;
        private const uint MOUSEEVENTF_LEFTUP = 0x04;

        [StructLayout(LayoutKind.Sequential)]
        private struct POINT { public int X; public int Y; }
        #endregion

        private static bool _continuePickingColor = true;
        private static ColorDropperModeForm _indicatorForm = null;
        private static ColorPreviewForm _previewForm = null;

        private class SelectionInfo
        {
            public Word.ShapeRange Shapes { get; set; }        // floating shapes
            public Word.Range TextRange { get; set; }          // document/shape/frame/header text
            public bool IsTextSelection { get; set; }
            public bool IsTableCellSelection { get; set; }     // Selection.Cells.Count > 0
            public bool IsChartSelection { get; set; }         // selected Shape/InlineShape with Chart
            public bool IsPureDocumentText { get; set; }       // text not in table
        }

        public static void SmartColorPicker()
        {
            var app = Globals.ThisAddIn.Application;
            var selection = app.Selection;

            if (selection == null)
            {
                MessageBox.Show("Please select a shape, text, table, or chart element first.");
                return;
            }

            var selectionInfo = GetSelectionInfo(selection);
            if (selectionInfo == null)
            {
                // With the new logic this should be practically unreachable,
                // but keep the guard just in case.
                MessageBox.Show("The selected item is not supported. Please select a shape, text, table, or chart element.");
                return;
            }

            _currentMode = PickMode.Fill;
            _continuePickingColor = true;
            _hookId = SetHook(_proc);
            Cursor oldCur = Cursor.Current;
            POINT pt;

            try
            {
                Cursor.Current = Cursors.Cross;
                ShowModeIndicator();
                ShowPreviewBox();

                while (_continuePickingColor && (GetAsyncKeyState(VK_LBUTTON) & 0x8000) == 0)
                {
                    Application.DoEvents();
                    HandleModeSwitchKeys();

                    GetCursorPos(out pt);
                    var cursorPoint = new System.Drawing.Point(pt.X, pt.Y);
                    var currentColor = GetScreenColor(pt.X, pt.Y);

                    UpdateUIPositionsAndColor(cursorPoint, currentColor);
                    Thread.Sleep(10);
                }

                if (!_continuePickingColor) return;

                while ((GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0)
                {
                    Application.DoEvents();
                    Thread.Sleep(10);
                }

                Application.DoEvents();

                GetCursorPos(out pt);
                var picked = GetScreenColor(pt.X, pt.Y);
                int oleColor = ColorTranslator.ToOle(picked); // OLE/BGR expected by Office

                ApplyColorToSelection(selectionInfo, oleColor, pt);
            }
            finally
            {
                Cursor.Current = oldCur;
                UnhookWindowsHookEx(_hookId);
                HideModeIndicator();
                HidePreviewBox();
            }
        }

        // ===== Robust selection detection: never returns null for valid text selections =====
        private static SelectionInfo GetSelectionInfo(Word.Selection sel)
        {
            var info = new SelectionInfo();
            Word.WdSelectionType stype;
            try { stype = sel.Type; } catch { stype = Word.WdSelectionType.wdSelectionNormal; }

            // Are we in a table?
            bool inTable = false;
            try { inTable = (sel.Cells != null && sel.Cells.Count > 0) || (sel.Range?.Tables?.Count > 0); } catch { }

            // 1) Floating shapes
            try
            {
                var sr = sel.ShapeRange;
                if (sr != null && sr.Count > 0)
                {
                    info.Shapes = sr;
                    info.IsChartSelection = ContainsChart(sr);
                    return info;
                }
            }
            catch { /* ShapeRange can throw when not applicable */ }

            // 2) Inline shapes (e.g., charts in text flow)
            try
            {
                var ilsCount = sel.InlineShapes?.Count ?? 0;
                if (ilsCount > 0)
                {
                    for (int i = 1; i <= ilsCount; i++)
                    {
                        var ils = sel.InlineShapes[i];
                        if (ils.HasChart == Office.MsoTriState.msoTrue)
                        {
                            info.IsChartSelection = true;
                            return info;
                        }
                    }
                }
            }
            catch { }

            // 3) Tables
            if (inTable)
            {
                info.IsTableCellSelection = true;
                // also keep text range for Font/Underline/Highlight in tables
                try { info.TextRange = sel.Range; info.IsTextSelection = true; } catch { }
                return info;
            }

            // 4) Text of any kind (normal story, header/footer, text frame, content control, caret-only)
            try
            {
                var r = sel.Range;
                if (r != null)
                {
                    info.TextRange = r;
                    info.IsTextSelection = true;
                    info.IsPureDocumentText = true;
                    return info;
                }
            }
            catch { }

            // Final fallback: treat as text if we can grab Selection.Range again
            try
            {
                info.TextRange = sel.Application?.Selection?.Range;
                if (info.TextRange != null)
                {
                    info.IsTextSelection = true;
                    info.IsPureDocumentText = !inTable;
                    return info;
                }
            }
            catch { }

            return null;
        }

        private static bool ContainsChart(Word.ShapeRange sr)
        {
            try
            {
                for (int i = 1; i <= sr.Count; i++)
                    if (sr[i].HasChart == Office.MsoTriState.msoTrue) return true;
            }
            catch { }
            return false;
        }

        private static void ApplyColorToSelection(SelectionInfo s, int oleColor, POINT cursorPosition)
        {
            var app = Globals.ThisAddIn.Application;

            // Charts: use native MSO eyedropper (context-aware)
            if (s.IsChartSelection)
            {
                string msoCommand = _currentMode == PickMode.Fill ? "ShapeFillColorPicker"
                                   : _currentMode == PickMode.Outline ? "ShapeOutlineColorPicker"
                                   : "FontColorPicker";

                try
                {
                    app.Selection.Select();
                    app.CommandBars.ExecuteMso(msoCommand);
                    Thread.Sleep(50);
                    mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)cursorPosition.X, (uint)cursorPosition.Y, 0, 0);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"Failed to execute native eyedropper: {ex.Message}");
                    MessageBox.Show("Could not activate the native eyedropper for the selected chart element.");
                }
                return;
            }

            // Pure document/shape/frame story text (not table cells)
            if (s.IsTextSelection && s.IsPureDocumentText && s.TextRange != null)
            {
                try
                {
                    var r = s.TextRange.Duplicate;

                    // If the selection is a caret (IP), operate on the next word as a sensible default for highlight/underline.
                    if (r.Start == r.End)
                    {
                        try { r = r.Words[1]; } catch { /* keep collapsed if words[] fails */ }
                    }

                    switch (_currentMode)
                    {
                        case PickMode.Fill:   // Highlight
                            try { r.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight; } catch { }
                            r.Shading.Texture = Word.WdTextureIndex.wdTextureNone;
                            r.Shading.ForegroundPatternColor = Word.WdColor.wdColorAutomatic;
                            r.Shading.BackgroundPatternColor = (Word.WdColor)oleColor;
                            break;

                        case PickMode.Outline: // Underline + colored underline
                            r.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                            try { r.Font.UnderlineColor = (Word.WdColor)oleColor; } catch { }
                            break;

                        case PickMode.Font:    // Font color
                            r.Font.Color = (Word.WdColor)oleColor;
                            break;
                    }
                }
                catch { /* keep tool smooth */ }
                return;
            }

            // Table cells
            if (s.IsTableCellSelection)
            {
                try
                {
                    var cells = app.Selection?.Cells;
                    if (cells != null && cells.Count > 0)
                    {
                        for (int i = 1; i <= cells.Count; i++)
                        {
                            var cell = cells[i];
                            switch (_currentMode)
                            {
                                case PickMode.Fill:
                                    cell.Shading.BackgroundPatternColor = (Word.WdColor)oleColor;
                                    break;

                                case PickMode.Outline:
                                    var btypes = new[]
                                    {
                                        Word.WdBorderType.wdBorderTop,
                                        Word.WdBorderType.wdBorderBottom,
                                        Word.WdBorderType.wdBorderLeft,
                                        Word.WdBorderType.wdBorderRight
                                    };
                                    foreach (var bt in btypes)
                                    {
                                        var b = cell.Borders[bt];
                                        b.LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                        b.Color = (Word.WdColor)oleColor;
                                    }
                                    break;

                                case PickMode.Font:
                                    cell.Range.Font.Color = (Word.WdColor)oleColor;
                                    break;
                            }
                        }
                    }
                }
                catch { }
                return;
            }

            // Floating shapes
            if (s.Shapes != null)
            {
                for (int i = 1; i <= s.Shapes.Count; i++)
                {
                    var shp = s.Shapes[i];
                    try
                    {
                        switch (_currentMode)
                        {
                            case PickMode.Fill:
                                if (shp.Fill.Type != Office.MsoFillType.msoFillBackground)
                                    shp.Fill.ForeColor.RGB = oleColor;
                                break;

                            case PickMode.Outline:
                                shp.Line.ForeColor.RGB = oleColor;
                                break;

                            case PickMode.Font:
                                try { shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = oleColor; }
                                catch
                                {
                                    try
                                    {
                                        if (shp.TextFrame != null)
                                        {
                                            var tr = shp.TextFrame.TextRange;
                                            if (tr != null) tr.Font.Color = (Word.WdColor)oleColor;
                                        }
                                    }
                                    catch { }
                                }
                                break;
                        }
                    }
                    catch { }
                }
            }
        }

        private static void HandleModeSwitchKeys()
        {
            if ((GetAsyncKeyState(VK_F) & 0x8000) != 0) { _currentMode = PickMode.Font; UpdateModeIndicatorText(); Thread.Sleep(150); }
            else if ((GetAsyncKeyState(VK_O) & 0x8000) != 0) { _currentMode = PickMode.Outline; UpdateModeIndicatorText(); Thread.Sleep(150); }
            else if ((GetAsyncKeyState(VK_H) & 0x8000) != 0) { _currentMode = PickMode.Fill; UpdateModeIndicatorText(); Thread.Sleep(150); }
            else if ((GetAsyncKeyState(VK_RIGHT) & 0x8000) != 0) { CycleModeNext(); UpdateModeIndicatorText(); Thread.Sleep(150); }
            else if ((GetAsyncKeyState(VK_LEFT) & 0x8000) != 0) { CycleModePrevious(); UpdateModeIndicatorText(); Thread.Sleep(150); }
            else if ((GetAsyncKeyState(VK_ESCAPE) & 0x8000) != 0) { _continuePickingColor = false; }
        }

        private static IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam == (IntPtr)WM_KEYDOWN)
            {
                int vk = Marshal.ReadInt32(lParam);
                bool handled = false;

                if (vk == VK_O) { _currentMode = PickMode.Outline; handled = true; }
                else if (vk == VK_F) { _currentMode = PickMode.Font; handled = true; }
                else if (vk == VK_H) { _currentMode = PickMode.Fill; handled = true; }
                else if (vk == VK_ESCAPE) { _continuePickingColor = false; handled = true; }
                else if (vk == VK_RIGHT) { CycleModeNext(); handled = true; }
                else if (vk == VK_LEFT) { CycleModePrevious(); handled = true; }

                if (handled)
                {
                    UpdateModeIndicatorText();
                    return (IntPtr)1;
                }
            }
            return CallNextHookEx(_hookId, nCode, wParam, lParam);
        }

        private static void CycleModeNext()
        {
            int nextMode = ((int)_currentMode + 1) % 3;
            _currentMode = (PickMode)nextMode;
        }

        private static void CycleModePrevious()
        {
            int prevMode = ((int)_currentMode - 1 + 3) % 3;
            _currentMode = (PickMode)prevMode;
        }

        private static System.Drawing.Color GetScreenColor(int x, int y)
        {
            using (var bmp = new Bitmap(1, 1))
            using (var g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(1, 1));
                return bmp.GetPixel(0, 0);
            }
        }

        private static void UpdateUIPositionsAndColor(System.Drawing.Point cursorPosition, System.Drawing.Color color)
        {
            if (_indicatorForm != null && !_indicatorForm.IsDisposed)
            {
                _indicatorForm.UpdateColor(color);

                int baseX = cursorPosition.X + 20;
                int baseY = cursorPosition.Y + 20;

                var screen = Screen.FromPoint(cursorPosition).WorkingArea;
                var size = _indicatorForm.Size;

                if (baseX + size.Width > screen.Right) baseX = screen.Right - size.Width - 4;
                if (baseY + size.Height > screen.Bottom) baseY = screen.Bottom - size.Height - 4;

                _indicatorForm.Location = new System.Drawing.Point(baseX, baseY);
                if (!_indicatorForm.Visible) _indicatorForm.Show();
                _indicatorForm.Update();
            }
        }

        private static void ShowModeIndicator()
        {
            if (_indicatorForm == null || _indicatorForm.IsDisposed)
                _indicatorForm = new ColorDropperModeForm();

            if (!_indicatorForm.IsHandleCreated) _indicatorForm.CreateControl();
            _indicatorForm.Show();
            _indicatorForm.BringToFront();
            _indicatorForm.UpdateModeText(_currentMode.ToString());
            _indicatorForm.Update();
        }

        private static void UpdateModeIndicatorText()
        {
            if (_indicatorForm != null && !_indicatorForm.IsDisposed && _indicatorForm.IsHandleCreated)
            {
                _indicatorForm.Invoke((MethodInvoker)delegate
                {
                    _indicatorForm.UpdateModeText(_currentMode.ToString());
                });
            }
        }

        private static void HideModeIndicator()
        {
            if (_indicatorForm != null && !_indicatorForm.IsDisposed)
            {
                _indicatorForm.Close();
                _indicatorForm.Dispose();
                _indicatorForm = null;
            }
        }

        private static void ShowPreviewBox()
        {
            if (_previewForm == null || _previewForm.IsDisposed)
            {
                _previewForm = new ColorPreviewForm();
            }
        }

        private static void HidePreviewBox()
        {
            if (_previewForm != null && !_previewForm.IsDisposed)
            {
                _previewForm.Close();
                _previewForm.Dispose();
                _previewForm = null;
            }
        }

        private class ColorPreviewForm : Form
        {
            public ColorPreviewForm()
            {
                this.FormBorderStyle = FormBorderStyle.None;
                this.ShowInTaskbar = false;
                this.TopMost = true;
                this.StartPosition = FormStartPosition.Manual;
                this.Size = new System.Drawing.Size(35, 35);

                this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
                this.Paint += ColorPreviewForm_Paint;
            }

            private void ColorPreviewForm_Paint(object sender, PaintEventArgs e)
            {
                using (var backBrush = new SolidBrush(this.BackColor))
                    e.Graphics.FillRectangle(backBrush, this.ClientRectangle);

                using (var borderPen = new Pen(System.Drawing.Color.FromArgb(255, 45, 45, 48), 1))
                    e.Graphics.DrawRectangle(borderPen, 0, 0, this.ClientSize.Width - 1, this.ClientSize.Height - 1);
            }

            public void UpdateColor(System.Drawing.Color color)
            {
                if (this.BackColor != color)
                {
                    this.BackColor = color;
                    this.Invalidate();
                }
            }
        }
    }

    public class ColorDropperModeForm : Form
    {
        private string _currentModeDisplay = "Fill";
        private System.Drawing.Color _swatchColor = System.Drawing.Color.White;

        private const int PadX = 10;
        private const int PadY = 6;
        private const int SwatchSize = 18;
        private const int Gap = 8;

        public ColorDropperModeForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            ShowInTaskbar = false;
            BackColor = System.Drawing.Color.FromArgb(255, 45, 45, 48);
            DoubleBuffered = true;

            AutoScaleMode = AutoScaleMode.None;
            Size = new System.Drawing.Size(150, 30);

            SetStyle(ControlStyles.UserPaint |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            EnsureSizeForText($"Mode: {_currentModeDisplay}");
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            var swatchY = (ClientSize.Height - SwatchSize) / 2;
            var swatchRect = new Rectangle(PadX, swatchY, SwatchSize, SwatchSize);

            using (var swatchBrush = new SolidBrush(_swatchColor))
                e.Graphics.FillRectangle(swatchBrush, swatchRect);

            using (var swatchBorder = new Pen(System.Drawing.Color.FromArgb(200, 90, 90, 95)))
                e.Graphics.DrawRectangle(swatchBorder, swatchRect);

            var modeText = $"Mode: {_currentModeDisplay}";
            using (var font = new Font("Segoe UI", 10f, FontStyle.Regular))
            {
                var textRect = new Rectangle(
                    swatchRect.Right + Gap,
                    0,
                    ClientSize.Width - (swatchRect.Right + Gap) - PadX,
                    ClientSize.Height
                );

                TextRenderer.DrawText(
                    e.Graphics,
                    modeText,
                    font,
                    textRect,
                    System.Drawing.Color.White,
                    TextFormatFlags.Left |
                    TextFormatFlags.VerticalCenter |
                    TextFormatFlags.NoPadding |
                    TextFormatFlags.NoClipping
                );
            }

            using (var outer = new Pen(System.Drawing.Color.FromArgb(255, 45, 45, 48)))
                e.Graphics.DrawRectangle(outer, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        public void UpdateModeText(string modeName)
        {
            _currentModeDisplay = modeName;
            EnsureSizeForText($"Mode: {_currentModeDisplay}");
            if (IsHandleCreated && !IsDisposed)
            {
                Invalidate();
                Update();
            }
        }

        public void UpdateColor(System.Drawing.Color color)
        {
            if (_swatchColor != color)
            {
                _swatchColor = color;
                if (IsHandleCreated && !IsDisposed) Invalidate();
            }
        }

        private void EnsureSizeForText(string text)
        {
            using (var font = new Font("Segoe UI", 10f, FontStyle.Regular))
            {
                var textSz = TextRenderer.MeasureText(
                    text,
                    font,
                    new System.Drawing.Size(int.MaxValue, int.MaxValue),
                    TextFormatFlags.NoPadding
                );

                int neededW = PadX + SwatchSize + Gap + textSz.Width + PadX;
                int neededH = PadY * 2 + Math.Max(SwatchSize, textSz.Height);

                if (ClientSize.Width != neededW || ClientSize.Height != neededH)
                    ClientSize = new System.Drawing.Size(neededW, neededH);
            }
        }
    }
}