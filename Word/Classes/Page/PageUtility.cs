using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace OfficePragmatics_Word.Page
{
    public static class PageUtility
    {

        public static void GoToPage(Word.Application app)
        {
            try
            {
                var doc = app?.ActiveDocument;
                var win = app?.ActiveWindow;
                if (doc == null || win == null) return;

                int currentPage = 1;
                try { currentPage = Convert.ToInt32(app.Selection.get_Information(Word.WdInformation.wdActiveEndPageNumber)); } catch { }

                int totalPages = 1;
                try { totalPages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages, false); } catch { }

                // Snapshot current view/zoom + nav pane
                var view = win.View;
                var origType = view.Type;
                var origZoomPct = view.Zoom.Percentage;
                var origCols = view.Zoom.PageColumns;
                var origRows = view.Zoom.PageRows;
                var origFit = view.Zoom.PageFit;
                bool origNavVisible = GetNavVisible(app);

                // “Slide sorter”-ish grid
                try
                {
                    view.Type = Word.WdViewType.wdPrintView;

                    int cols = (int)Math.Ceiling(Math.Sqrt(totalPages));
                    int rows = (int)Math.Ceiling(totalPages / (double)cols);
                    cols = Math.Min(cols, 6);
                    rows = Math.Min(rows, 6);

                    view.Zoom.PageFit = Word.WdPageFit.wdPageFitFullPage;
                    view.Zoom.PageColumns = cols;
                    view.Zoom.PageRows = rows;

                    // show nav pane thumbnails if it wasn't visible
                    if (!origNavVisible) ToggleNavigationPane(app, show: true);
                }
                catch { /* non-fatal */ }

                using (var form = new GoToPageForm(currentPage, totalPages))
                {
                    form.ShowDialog();

                    if (!form.WasCancelled && int.TryParse(form.PageNumber, out int n))
                    {
                        if (n >= 1 && n <= totalPages)
                        {
                            object what = Word.WdGoToItem.wdGoToPage;
                            object which = Word.WdGoToDirection.wdGoToAbsolute;
                            object pageNum = n;
                            object missing = Type.Missing;

                            var target = doc.GoTo(ref what, ref which, ref pageNum, ref missing);
                            try { target.Select(); } catch { }
                        }
                        else
                        {
                            MessageBox.Show($"Page {n} is out of range (1–{totalPages}).",
                                "Invalid Page", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }

                // ---- Restore in the right order ----
                try
                {
                    // 1) restore view type FIRST (changing this later can blow away zoom)
                    view.Type = origType;

                    // 2) exit any page-fit mode BEFORE setting percentage
                    view.Zoom.PageFit = Word.WdPageFit.wdPageFitNone;

                    // 3) restore multi-page grid (usually 1x1)
                    view.Zoom.PageColumns = origCols;
                    view.Zoom.PageRows = origRows;

                    // 4) restore zoom percentage
                    view.Zoom.Percentage = Math.Max(10, Math.Min(500, origZoomPct));

                    // 5) if the original session used a PageFit (rare), put it back last
                    if (origFit != Word.WdPageFit.wdPageFitNone)
                        view.Zoom.PageFit = origFit;

                    // restore nav pane to original visibility
                    if (GetNavVisible(app) != origNavVisible)
                        ToggleNavigationPane(app, show: origNavVisible);
                }
                catch { /* ignore */ }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GoToPage Error: {ex.Message}");
            }
        }

        private static bool GetNavVisible(Word.Application app)
        {
            try { return app.CommandBars["Navigation"].Visible; } catch { return false; }
        }
        private static void ToggleNavigationPane(Word.Application app, bool show)
        {
            try
            {
                bool visible = GetNavVisible(app);
                if (show && !visible) app.CommandBars.ExecuteMso("NavigationPane");
                if (!show && visible) app.CommandBars.ExecuteMso("NavigationPane");
            }
            catch { /* safe to ignore */ }
        }

        // Lightweight input UI, scaled for DPI
        private class GoToPageForm : Form
        {
            private string _pageInfo;
            private string _currentInput = "";
            private float _scale = 1f; // Scaling factor for DPI
            public string PageNumber { get; private set; } = "";
            public bool WasCancelled { get; private set; } = true;

            public GoToPageForm(int currentPage, int totalPages)
            {
                FormBorderStyle = FormBorderStyle.None;
                StartPosition = FormStartPosition.Manual;

                // 1. Calculate the current DPI scale (Standard is 96 DPI)
                using (Graphics g = this.CreateGraphics())
                {
                    _scale = Math.Max(1.0f, g.DpiX / 96.0f);
                }

                // 2. Scale the form size
                int baseWidth = 350;
                int baseHeight = 180;
                Size = new System.Drawing.Size((int)(baseWidth * _scale), (int)(baseHeight * _scale));

                BackColor = Color.FromArgb(45, 45, 48);
                Opacity = 0.95;
                TopMost = true;
                ShowInTaskbar = false;
                Name = "GoToPageForm";
                KeyPreview = true;

                // Center on Primary Screen
                var bounds = Screen.PrimaryScreen.Bounds;
                Location = new Point((bounds.Width - Width) / 2, (bounds.Height - Height) / 2);

                _pageInfo = $"{currentPage} of {totalPages}";

                SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
                Paint += GoToPageForm_Paint;
                KeyDown += GoToPageForm_KeyDown;
                KeyPress += GoToPageForm_KeyPress;
            }

            private void GoToPageForm_Paint(object sender, PaintEventArgs e)
            {
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                g.Clear(BackColor);
                using (var borderPen = new Pen(Color.FromArgb(100, 100, 100), 1))
                    g.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);

                var titleColor = Color.FromArgb(220, 220, 220);
                var textColor = Color.FromArgb(200, 200, 200);
                var accentColor = Color.FromArgb(0, 120, 215);
                var inputBgColor = Color.FromArgb(60, 60, 63);
                var inputBorderColor = Color.FromArgb(120, 120, 120);

                // 3. Scale Fonts
                using (var titleFont = new Font("Segoe UI", 13 * _scale, FontStyle.Bold))
                using (var labelFont = new Font("Segoe UI", 10 * _scale))
                using (var inputFont = new Font("Segoe UI", 12 * _scale, FontStyle.Bold))
                using (var helpFont = new Font("Segoe UI", 9 * _scale))
                using (var titleBrush = new SolidBrush(titleColor))
                using (var textBrush = new SolidBrush(textColor))
                using (var inputBgBrush = new SolidBrush(inputBgColor))
                using (var inputBorderPen = new Pen(inputBorderColor, 1))
                {
                    // 4. Scale all coordinates (x, y, width, height)
                    int pad20 = (int)(20 * _scale);

                    g.DrawString("Go To Page", titleFont, titleBrush, new Point(pad20, pad20));

                    var infoText = $"Current: {_pageInfo}";
                    var infoSize = g.MeasureString(infoText, labelFont);
                    var infoPoint = new PointF(Width - infoSize.Width - pad20, 23 * _scale);
                    g.DrawString(infoText, labelFont, textBrush, infoPoint);

                    g.DrawString("Enter page number:", labelFont, textBrush, new PointF(pad20, 60 * _scale));

                    var inputBoxRect = new Rectangle(pad20, (int)(95 * _scale), (int)(200 * _scale), (int)(35 * _scale));
                    g.FillRectangle(inputBgBrush, inputBoxRect);
                    g.DrawRectangle(inputBorderPen, inputBoxRect);

                    int textOffsetX = (int)(10 * _scale);
                    if (!string.IsNullOrEmpty(_currentInput))
                        g.DrawString(_currentInput, inputFont, titleBrush, new PointF(inputBoxRect.X + textOffsetX, inputBoxRect.Y));

                    var cursorX = inputBoxRect.X + textOffsetX;
                    if (!string.IsNullOrEmpty(_currentInput))
                    {
                        var textSize = g.MeasureString(_currentInput, inputFont);
                        // MeasureString adds some padding, trim slightly for cursor visual if needed, 
                        // but strictly using width is safer for variable fonts.
                        cursorX += (int)textSize.Width - (int)(4 * _scale);
                    }

                    // Cursor logic
                    var cursorHeight = (int)(15 * _scale);
                    var cursorYOffset = (int)(11 * _scale);
                    var cursorRect = new Rectangle(cursorX, inputBoxRect.Y + cursorYOffset, Math.Max(1, (int)(1 * _scale)), cursorHeight);
                    g.FillRectangle(titleBrush, cursorRect);

                    g.DrawString("Press ENTER to navigate, ESC to cancel", helpFont, textBrush, new PointF(pad20, 145 * _scale));

                    using (var accentPen = new Pen(accentColor, Math.Max(2, 2 * _scale)))
                        g.DrawLine(accentPen, 0, Height - (int)(2 * _scale), Width, Height - (int)(2 * _scale));
                }
            }

            // ... KeyPress and KeyDown methods remain exactly the same ...
            private void GoToPageForm_KeyPress(object sender, KeyPressEventArgs e)
            {
                if (char.IsDigit(e.KeyChar))
                {
                    _currentInput += e.KeyChar;
                    Invalidate();
                    e.Handled = true;
                }
                else if (e.KeyChar == (char)Keys.Back && _currentInput.Length > 0)
                {
                    _currentInput = _currentInput.Substring(0, _currentInput.Length - 1);
                    Invalidate();
                    e.Handled = true;
                }
            }

            private void GoToPageForm_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                {
                    if (!string.IsNullOrEmpty(_currentInput))
                    {
                        PageNumber = _currentInput;
                        WasCancelled = false;
                    }
                    Close();
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    WasCancelled = true;
                    Close();
                    e.Handled = true;
                }
            }
        }

    }
}