using OfficePragmatics_Powerpoint;
using System;
using System.Collections;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_PowerPoint.Shape
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
        private const int VK_LBUTTON = 0x01;
        private const int VK_LEFT = (int)Keys.Left;
        private const int VK_RIGHT = (int)Keys.Right;


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

        // NEW: Win32 API for simulating a mouse click
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
            public PowerPoint.ShapeRange Shapes { get; set; }
            public PowerPoint.TextRange TextRange { get; set; }
            public PowerPoint.Table Table { get; set; }
            public PowerPoint.Chart Chart { get; set; }
            public object ChartSelection { get; set; }
            public int StartRow { get; set; }
            public int StartColumn { get; set; }
            public int EndRow { get; set; }
            public int EndColumn { get; set; }
            public bool IsTextSelection { get; set; }
            public bool IsTableCellSelection { get; set; }
            public bool IsChartSelection { get; set; }
        }

        public static void SmartColorPicker()
        {
            var app = Globals.ThisAddIn.Application;
            var initialSel = app.ActiveWindow?.Selection;

            if (initialSel == null || initialSel.Type == PowerPoint.PpSelectionType.ppSelectionNone)
            {
                MessageBox.Show("Please select a shape, text, table, or chart element first.");
                return;
            }

            var selectionInfo = GetSelectionInfo(initialSel);
            if (selectionInfo == null)
            {
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
                    System.Drawing.Color currentColor = GetScreenColor(pt.X, pt.Y);

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

                // MODIFIED: Capture final cursor position and pass it to the apply function
                GetCursorPos(out pt);
                System.Drawing.Color picked = GetScreenColor(pt.X, pt.Y);
                int oleColor = picked.R | (picked.G << 8) | (picked.B << 16);

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

        private static SelectionInfo GetSelectionInfo(PowerPoint.Selection selection)
        {
            var info = new SelectionInfo();

            if (selection.Type == PowerPoint.PpSelectionType.ppSelectionShapes && selection.ShapeRange.HasChart == Office.MsoTriState.msoTrue)
            {
                try
                {
                    info.Chart = selection.ShapeRange[1].Chart;
                    // This direct access can be unreliable, which is why the new method is better
                    info.ChartSelection = info.Chart.Application.ActiveWindow.Selection.Parent;
                    info.IsChartSelection = true;
                    info.Shapes = selection.ShapeRange;
                    return info;
                }
                catch { }
            }

            if (selection.ShapeRange != null && selection.ShapeRange.Count == 1)
            {
                var shape = selection.ShapeRange[1];
                if (shape.HasTable == Office.MsoTriState.msoTrue)
                {
                    var tableInfo = GetSelectedTableCells(selection, shape);
                    if (tableInfo != null)
                    {
                        return tableInfo;
                    }
                }
            }

            switch (selection.Type)
            {
                case PowerPoint.PpSelectionType.ppSelectionShapes:
                    info.Shapes = selection.ShapeRange;
                    return info;

                case PowerPoint.PpSelectionType.ppSelectionText:
                    info.TextRange = selection.TextRange;
                    info.IsTextSelection = true;
                    info.Shapes = selection.ShapeRange;
                    return info;
            }

            return null;
        }

        private static SelectionInfo GetSelectedTableCells(PowerPoint.Selection selection, PowerPoint.Shape tableShape)
        {
            try
            {
                var table = tableShape.Table;

                var info = new SelectionInfo
                {
                    Table = table,
                    IsTableCellSelection = true,
                    IsTextSelection = false,
                    Shapes = selection.ShapeRange
                };

                if (selection.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    var textRange = selection.TextRange;
                    if (textRange != null)
                    {
                        for (int row = 1; row <= table.Rows.Count; row++)
                        {
                            for (int col = 1; col <= table.Columns.Count; col++)
                            {
                                try
                                {
                                    var cell = table.Cell(row, col);
                                    if (cell.Shape.HasTextFrame == Office.MsoTriState.msoTrue)
                                    {
                                        var cellText = cell.Shape.TextFrame.TextRange;
                                        if (cellText.Text == textRange.Text ||
                                            (cellText.Text.Contains(textRange.Text) && textRange.Text.Length > 0))
                                        {
                                            info.StartRow = row; info.StartColumn = col; info.EndRow = row; info.EndColumn = col;
                                            return info;
                                        }
                                    }
                                }
                                catch { }
                            }
                        }
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        // MODIFIED: Accepts cursor position to handle chart clicks
        private static void ApplyColorToSelection(SelectionInfo selectionInfo, int oleColor, POINT cursorPosition)
        {
            // MODIFIED: New method for charts using the native eyedropper
            if (selectionInfo.IsChartSelection)
            {
                string msoCommand = "";
                switch (_currentMode)
                {
                    case PickMode.Fill:
                        msoCommand = "ShapeFillColorPicker";
                        break;
                    case PickMode.Outline:
                        msoCommand = "ShapeOutlineColorPicker";
                        break;
                    case PickMode.Font:
                        msoCommand = "FontColorPicker";
                        break;
                }

                if (!string.IsNullOrEmpty(msoCommand))
                {
                    var app = Globals.ThisAddIn.Application;
                    try
                    {
                        // Re-select the chart shape to ensure CommandBars context is correct
                        selectionInfo.Shapes.Select();

                        // Execute the MSO command to activate the native eyedropper
                        app.CommandBars.ExecuteMso(msoCommand);

                        // Pause briefly to allow PowerPoint to activate the eyedropper cursor
                        Thread.Sleep(50);

                        // Simulate an immediate click at the saved cursor position
                        mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)cursorPosition.X, (uint)cursorPosition.Y, 0, 0);
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"Failed to execute native eyedropper: {ex.Message}");
                        MessageBox.Show("Could not activate the native eyedropper. The selected chart part might not support this operation.");
                    }
                }
                return; // Exit after handling the chart
            }

            if (selectionInfo.IsTextSelection && _currentMode == PickMode.Font && selectionInfo.TextRange != null)
            {
                selectionInfo.TextRange.Font.Color.RGB = oleColor;
                return;
            }

            if (selectionInfo.Shapes != null)
            {
                foreach (PowerPoint.Shape shp in selectionInfo.Shapes)
                {
                    if (shp.HasTable == Office.MsoTriState.msoTrue)
                    {
                        var table = shp.Table;
                        for (int row = 1; row <= table.Rows.Count; row++)
                        {
                            for (int col = 1; col <= table.Columns.Count; col++)
                            {
                                var cell = table.Cell(row, col);
                                if (cell.Selected)
                                {
                                    var cellShape = cell.Shape;
                                    switch (_currentMode)
                                    {
                                        case PickMode.Fill:
                                            cellShape.Fill.ForeColor.RGB = oleColor;
                                            break;
                                        case PickMode.Outline:
                                            var borderTypes = new[] { PowerPoint.PpBorderType.ppBorderTop, PowerPoint.PpBorderType.ppBorderBottom, PowerPoint.PpBorderType.ppBorderLeft, PowerPoint.PpBorderType.ppBorderRight };
                                            foreach (var borderType in borderTypes)
                                            {
                                                var border = cell.Borders[borderType];
                                                border.ForeColor.RGB = oleColor;
                                                border.Visible = Office.MsoTriState.msoTrue;
                                            }
                                            break;
                                        case PickMode.Font:
                                            if (cellShape.HasTextFrame == Office.MsoTriState.msoTrue)
                                            {
                                                cellShape.TextFrame.TextRange.Font.Color.RGB = oleColor;
                                            }
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        switch (_currentMode)
                        {
                            case PickMode.Fill:
                                if (shp.Fill.Type != Office.MsoFillType.msoFillBackground) shp.Fill.ForeColor.RGB = oleColor;
                                break;
                            case PickMode.Outline:
                                shp.Line.ForeColor.RGB = oleColor;
                                break;
                            case PickMode.Font:
                                if (shp.HasTextFrame == Office.MsoTriState.msoTrue) shp.TextFrame.TextRange.Font.Color.RGB = oleColor;
                                break;
                        }
                    }
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

                // keep the pill fully on-screen if near edges
                var screen = Screen.FromPoint(cursorPosition).WorkingArea;
                var size = _indicatorForm.Size;

                if (baseX + size.Width > screen.Right) baseX = screen.Right - size.Width - 4;
                if (baseY + size.Height > screen.Bottom) baseY = screen.Bottom - size.Height - 4;

                _indicatorForm.Location = new System.Drawing.Point(baseX, baseY);
                if (!_indicatorForm.Visible) _indicatorForm.Show();
                _indicatorForm.Update();                 // ensure it draws immediately
            }
        }

        private static void ShowModeIndicator()
        {
            if (_indicatorForm == null || _indicatorForm.IsDisposed)
                _indicatorForm = new ColorDropperModeForm();

            if (!_indicatorForm.IsHandleCreated) _indicatorForm.CreateControl();
            _indicatorForm.Show();               // make it visible now
            _indicatorForm.BringToFront();
            _indicatorForm.UpdateModeText(_currentMode.ToString());
            _indicatorForm.Update();             // flush first paint
        }


        private static void UpdateModeIndicatorText()
        {
            if (_indicatorForm != null && !_indicatorForm.IsDisposed && _indicatorForm.IsHandleCreated)
            {
                _indicatorForm.Invoke((MethodInvoker)delegate {
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
                {
                    e.Graphics.FillRectangle(backBrush, this.ClientRectangle);
                }

                using (var borderPen = new Pen(System.Drawing.Color.FromArgb(255, 45, 45, 48), 1))
                {
                    e.Graphics.DrawRectangle(borderPen, 0, 0, this.ClientSize.Width - 1, this.ClientSize.Height - 1);
                }
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

            AutoScaleMode = AutoScaleMode.None;   // avoid DPI autosize jitter
            Size = new System.Drawing.Size(150, 30);

            SetStyle(ControlStyles.UserPaint |
                     ControlStyles.AllPaintingInWmPaint |
                     ControlStyles.OptimizedDoubleBuffer, true);

            // size correctly for initial text
            EnsureSizeForText($"Mode: {_currentModeDisplay}");
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // swatch
            var swatchY = (ClientSize.Height - SwatchSize) / 2;
            var swatchRect = new Rectangle(PadX, swatchY, SwatchSize, SwatchSize);

            using (var swatchBrush = new SolidBrush(_swatchColor))
                e.Graphics.FillRectangle(swatchBrush, swatchRect);

            using (var swatchBorder = new Pen(System.Drawing.Color.FromArgb(200, 90, 90, 95)))
                e.Graphics.DrawRectangle(swatchBorder, swatchRect);

            // text
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
                    TextFormatFlags.NoPadding |    // tighter + consistent with MeasureText below
                    TextFormatFlags.NoClipping     // avoid overhang clipping edge-cases
                );
            }

            using (var outer = new Pen(System.Drawing.Color.FromArgb(255, 45, 45, 48)))
                e.Graphics.DrawRectangle(outer, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }

        public void UpdateModeText(string modeName)
        {
            _currentModeDisplay = modeName;

            // size BEFORE painting to avoid the one-frame clip
            EnsureSizeForText($"Mode: {_currentModeDisplay}");

            if (IsHandleCreated && !IsDisposed)
            {
                Invalidate();
                Update(); // paint now so user sees the change immediately
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
                int neededH = PadY * 2 + System.Math.Max(SwatchSize, textSz.Height);

                if (ClientSize.Width != neededW || ClientSize.Height != neededH)
                    ClientSize = new System.Drawing.Size(neededW, neededH);
            }
        }
    }
}