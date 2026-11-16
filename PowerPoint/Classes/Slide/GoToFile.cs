using System;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficePragmatics_PowerPoint.Slide
{
    public static class GoToSlideUtility
    {
        public static void GoToSlide(PowerPoint.Application app)
        {
            try
            {
                if (app.ActivePresentation == null || app.ActivePresentation.Slides == null) return;

                // Get current slide info before switching views
                int currentSlide = app.ActiveWindow.View.Slide.SlideIndex;
                int totalSlides = app.ActivePresentation.Slides.Count;

                // Store the current view type and zoom
                var originalViewType = app.ActiveWindow.ViewType;
                var originalZoom = app.ActiveWindow.View.Zoom;

                // Switch to slide sorter view and set zoom to 40%
                app.ActiveWindow.ViewType = PowerPoint.PpViewType.ppViewSlideSorter;
                app.ActiveWindow.View.Zoom = 40;

                using (var form = new GoToSlideForm(currentSlide, totalSlides))
                {
                    form.ShowDialog();

                    if (!form.WasCancelled && int.TryParse(form.SlideNumber, out int n))
                    {
                        if (n >= 1 && n <= totalSlides)
                        {
                            app.ActiveWindow.View.GotoSlide(n);
                        }
                        else
                        {
                            MessageBox.Show($"Slide {n} is out of range (1–{totalSlides}).",
                                            "Invalid Slide",
                                            MessageBoxButtons.OK,
                                            MessageBoxIcon.Warning);
                        }
                    }
                }

                // Restore the original view type and zoom
                app.ActiveWindow.ViewType = originalViewType;
                app.ActiveWindow.View.Zoom = originalZoom;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"GoToSlide Error: {ex.Message}");
            }
        }
    }

    public class GoToSlideForm : Form
    {
        private string _slideInfo = "";
        private string _currentInput = "";

        public string SlideNumber { get; private set; } = "";
        public bool WasCancelled { get; private set; } = true;

        public GoToSlideForm(int currentSlide, int totalSlides)
        {
            // Form Properties
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            // Increased size slightly to prevent cramping
            this.Size = new System.Drawing.Size(400, 220);
            this.BackColor = System.Drawing.Color.FromArgb(45, 45, 48);
            this.Opacity = 0.95;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.Name = "GoToSlideForm";
            this.KeyPreview = true;

            // Center the form on screen
            var bounds = Screen.PrimaryScreen.Bounds;
            this.Location = new System.Drawing.Point(
                (bounds.Width - this.Width) / 2,
                (bounds.Height - this.Height) / 2);

            _slideInfo = $"{currentSlide} of {totalSlides}";

            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            this.Paint += GoToSlideForm_Paint;
            this.KeyDown += GoToSlideForm_KeyDown;
            this.KeyPress += GoToSlideForm_KeyPress;
        }

        private void GoToSlideForm_Paint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

            // Clear background
            g.Clear(this.BackColor);

            // Draw subtle border
            using (var borderPen = new System.Drawing.Pen(System.Drawing.Color.FromArgb(100, 100, 100), 1))
            {
                g.DrawRectangle(borderPen, 0, 0, this.Width - 1, this.Height - 1);
            }

            // Colors
            var titleColor = System.Drawing.Color.FromArgb(220, 220, 220);
            var textColor = System.Drawing.Color.FromArgb(200, 200, 200);
            var accentColor = System.Drawing.Color.FromArgb(0, 120, 215);
            var inputBgColor = System.Drawing.Color.FromArgb(60, 60, 63);
            var inputBorderColor = System.Drawing.Color.FromArgb(120, 120, 120);

            // Fonts
            using (var titleFont = new System.Drawing.Font("Segoe UI", 13, System.Drawing.FontStyle.Bold))
            using (var labelFont = new System.Drawing.Font("Segoe UI", 10))
            using (var inputFont = new System.Drawing.Font("Segoe UI", 12, System.Drawing.FontStyle.Bold))
            using (var helpFont = new System.Drawing.Font("Segoe UI", 9))
            using (var titleBrush = new System.Drawing.SolidBrush(titleColor))
            using (var textBrush = new System.Drawing.SolidBrush(textColor))
            using (var accentBrush = new System.Drawing.SolidBrush(accentColor))
            using (var inputBgBrush = new System.Drawing.SolidBrush(inputBgColor))
            using (var inputBorderPen = new System.Drawing.Pen(inputBorderColor, 1))
            {
                // Dynamic Layout Calculation
                float currentY = 25;
                float leftMargin = 25;
                float rightMargin = 25;

                // 1. Title section
                string titleText = "Go To Slide";
                var titleSize = g.MeasureString(titleText, titleFont);
                g.DrawString(titleText, titleFont, titleBrush, new System.Drawing.PointF(leftMargin, currentY));

                // 2. Current slide info (right aligned)
                // We assume same Y baseline roughly, or centered relative to title height
                var infoText = $"Current: {_slideInfo}";
                var infoSize = g.MeasureString(infoText, labelFont);
                // Align vertically with title center
                float infoY = currentY + (titleSize.Height - infoSize.Height) / 2;
                var infoPoint = new System.Drawing.PointF(this.Width - rightMargin - infoSize.Width, infoY);
                g.DrawString(infoText, labelFont, textBrush, infoPoint);

                // Update Y position for next element (Title height + spacing)
                currentY += titleSize.Height + 20;

                // 3. Input Label
                var inputLabel = "Enter slide number:";
                var labelSize = g.MeasureString(inputLabel, labelFont);
                g.DrawString(inputLabel, labelFont, textBrush, new System.Drawing.PointF(leftMargin, currentY));

                // Update Y position (Label height + spacing)
                currentY += labelSize.Height + 8;

                // 4. Input Box
                float boxHeight = 38;
                float boxWidth = this.Width - (leftMargin + rightMargin);
                var inputBoxRect = new System.Drawing.Rectangle((int)leftMargin, (int)currentY, (int)boxWidth, (int)boxHeight);
                g.FillRectangle(inputBgBrush, inputBoxRect);
                g.DrawRectangle(inputBorderPen, inputBoxRect);

                // Input Text Drawing
                float textPadding = 10;
                var inputText = _currentInput;

                // Calculate vertical center for text inside the box
                float inputFontHeight = inputFont.GetHeight(g);
                float textY = inputBoxRect.Y + (boxHeight - inputFontHeight) / 2;

                if (!string.IsNullOrEmpty(inputText))
                {
                    var inputTextPoint = new System.Drawing.PointF(inputBoxRect.X + textPadding, textY);
                    g.DrawString(inputText, inputFont, titleBrush, inputTextPoint);
                }

                // Cursor Drawing
                var cursorX = inputBoxRect.X + textPadding;
                if (!string.IsNullOrEmpty(inputText))
                {
                    var textSize = g.MeasureString(inputText, inputFont);
                    // MeasureString often adds padding, trim slightly for visual cursor tightness
                    cursorX += (int)textSize.Width - 4;
                }

                var cursorRect = new System.Drawing.Rectangle((int)cursorX, (int)textY + 2, 2, (int)inputFontHeight - 4);
                g.FillRectangle(titleBrush, cursorRect);

                // Update Y position (Box height + spacing)
                currentY += boxHeight + 15;

                // 5. Help text
                var helpText = "Press ENTER to navigate, ESC to cancel";
                g.DrawString(helpText, helpFont, textBrush, new System.Drawing.PointF(leftMargin, currentY));

                // Accent line at bottom
                using (var accentPen = new System.Drawing.Pen(accentColor, 2))
                {
                    g.DrawLine(accentPen, 0, this.Height - 2, this.Width, this.Height - 2);
                }
            }
        }

        private void GoToSlideForm_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (char.IsDigit(e.KeyChar))
            {
                _currentInput += e.KeyChar;
                this.Invalidate();
                e.Handled = true;
            }
            else if (e.KeyChar == (char)Keys.Back && _currentInput.Length > 0)
            {
                _currentInput = _currentInput.Substring(0, _currentInput.Length - 1);
                this.Invalidate();
                e.Handled = true;
            }
        }

        private void GoToSlideForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!string.IsNullOrEmpty(_currentInput))
                {
                    SlideNumber = _currentInput;
                    WasCancelled = false;
                }
                this.Close();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                WasCancelled = true;
                this.Close();
                e.Handled = true;
            }
        }
    }
}