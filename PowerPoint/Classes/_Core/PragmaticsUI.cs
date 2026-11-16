// PragmaticsUI.cs  (C# 7.3 / .NET 4.8)
// Consolidated: DpiHelpers, Theme, and PragmaticsForm base class

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace OfficePragmatics_Excel.Settings
{
    // -------------------------
    // DPI Helper
    // -------------------------
    internal static class Dpi
    {
        // Design-time baseline (Visual Studio designer default)
        public const float DesignDpi = 96f;

        // Current effective DPI and init flag
        private static int _dpi = (int)DesignDpi;
        private static bool _initialized;

        /// <summary>Current DPI value.</summary>
        public static int Current => _dpi;

        /// <summary>Scale factor relative to design DPI.</summary>
        public static float Scale => _dpi / DesignDpi;

        /// <summary>
        /// Convert logical (design) pixels to device pixels.
        /// - Before InitForWindow is called, returns design units (no scale) to avoid double-scaling with WinForms AutoScale.
        /// - After Init/WM_DPICHANGED, returns scaled values.
        /// </summary>
        public static int Px(int logicalPx)
        {
            var scale = _initialized ? Scale : 1f;
            return (int)Math.Round(logicalPx * scale);
        }

        public static float PxF(float logicalPx)
        {
            var scale = _initialized ? Scale : 1f;
            return logicalPx * scale;
        }

        /// <summary>
        /// Bind the helper to a real window; from now on scaling uses the actual window DPI.
        /// </summary>
        public static void InitForWindow(IntPtr hwnd)
        {
            try
            {
                _dpi = (hwnd != IntPtr.Zero) ? GetDpiForWindow(hwnd) : GetDpiForSystemSafe();
            }
            catch
            {
                _dpi = GetDpiForSystemSafe();
            }
            _initialized = true;
        }

        /// <summary>
        /// Update DPI from WM_DPICHANGED (wParam LOWORD). Marks initialized=true.
        /// </summary>
        public static void UpdateFromWmDpiChanged(IntPtr wParam)
        {
            try
            {
                // LOWORD contains the new DPI value
                int newDpi = (int)(wParam.ToInt64() & 0xFFFF);
                if (newDpi < 50 || newDpi > 1000) // sanity bounds
                    newDpi = GetDpiForSystemSafe();
                _dpi = newDpi;
            }
            catch
            {
                _dpi = GetDpiForSystemSafe();
            }
            _initialized = true;
        }

        /// <summary>
        /// Return to design-time (96 DPI) and mark as not initialized.
        /// Call this at the very start of a form ctor if you want Theme.Px to
        /// emit design units until the handle is created.
        /// </summary>
        public static void ResetToDesign()
        {
            _dpi = (int)DesignDpi;
            _initialized = false;
        }

        /// <summary>
        /// Use Per-Monitor V2 awareness on the current thread while the scope is alive.
        /// Safe no-op on older Windows.
        /// </summary>
        public static DpiScope UsePerMonitorV2()
            => new DpiScope(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2);

        // -------------------------
        // Internal helpers & P/Invoke
        // -------------------------

        private static int GetDpiForSystemSafe()
        {
            try
            {
                // Available on Win10+. P/Invoke will throw on older systems; we catch and fallback.
                int sys = GetDpiForSystem();
                if (sys < 50 || sys > 1000) return (int)DesignDpi;
                return sys;
            }
            catch
            {
                // Conservative fallback
                return (int)DesignDpi;
            }
        }

        // Scope for SetThreadDpiAwarenessContext
        public sealed class DpiScope : IDisposable
        {
            private readonly IntPtr _prev;
            private readonly bool _ok;

            internal DpiScope(IntPtr ctx)
            {
                try
                {
                    _prev = SetThreadDpiAwarenessContext(ctx);
                    _ok = _prev != IntPtr.Zero;
                }
                catch
                {
                    _prev = IntPtr.Zero;
                    _ok = false;
                }
            }

            public void Dispose()
            {
                if (_ok)
                {
                    try { SetThreadDpiAwarenessContext(_prev); } catch { }
                }
            }
        }

        // P/Invoke declarations (Win10+; guarded by try/catch)
        [DllImport("user32.dll")] private static extern int GetDpiForSystem();
        [DllImport("user32.dll")] private static extern int GetDpiForWindow(IntPtr hwnd);
        [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);

        // DPI awareness context constants
        // https://learn.microsoft.com/windows/win32/api/winuser/nf-winuser-setthreaddpiawarenesscontext
        private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new IntPtr(-4);
    }

    // -------------------------
    // Theme
    // -------------------------
    internal static class Theme
    {
        public static readonly Color Accent = ColorTranslator.FromHtml("#00000");
        public static readonly Color Surface = Color.White;
        public static readonly Color SurfaceAlt = Color.FromArgb(249, 249, 249);
        public static readonly Color Border = Color.FromArgb(210, 210, 210);
        public static readonly Color NavBg = Color.FromArgb(240, 240, 240);
        public static readonly Color NavText = Color.Black;
        public static readonly Color HeaderBg = Color.FromArgb(245, 245, 245);

        // Title bar colors
        public static readonly Color TitleBarBg = Color.Black;
        public static readonly Color TitleBarText = Color.White;

        public static int ButtonHeight => Dpi.Px(36);
        public static int RowHeight => Dpi.Px(36);
        public static int PadSm => Dpi.Px(4);
        public static int PadMd => Dpi.Px(12);
        public static int PadLg => Dpi.Px(16);

        public static int Px(int logicalPx) => Dpi.Px(logicalPx);

        public static void PrimaryButton(Button b)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 0;
            b.BackColor = Color.Black;
            b.ForeColor = Color.White;
            b.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
            b.Height = ButtonHeight;
            b.MinimumSize = new System.Drawing.Size(Px(100), ButtonHeight);
        }

        public static void GhostButton(Button b)
        {
            b.FlatStyle = FlatStyle.Flat;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.BorderColor = Border;
            b.BackColor = Surface;
            b.ForeColor = NavText;
            b.Height = ButtonHeight;
            b.MinimumSize = new System.Drawing.Size(Px(100), ButtonHeight);
        }

        public static void Cardize(Control c)
        {
            c.BackColor = Surface;
            c.Padding = new Padding(PadMd);
            c.Paint += (s, e) =>
            {
                using (var p = new Pen(Border, Math.Max(1f, Dpi.PxF(1f)))) // crisp on HiDPI
                    e.Graphics.DrawRectangle(p, new Rectangle(0, 0, c.Width - 1, c.Height - 1));
            };
        }

        public static void DoubleBuffer(Control c)
        {
            try
            {
                var prop = typeof(Control).GetProperty("DoubleBuffered",
                    System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
                prop?.SetValue(c, true, null);
            }
            catch { }
        }
    }

    // -------------------------
    // PragmaticsForm Base Class
    // -------------------------
    public class PragmaticsForm : Form
    {
        // Remember design-time autoscale (so WinForms also helps us)
        public PragmaticsForm()
        {
            AutoScaleMode = AutoScaleMode.Dpi;
            AutoScaleDimensions = new SizeF(Dpi.DesignDpi, Dpi.DesignDpi);
            Font = new Font("Segoe UI", 9f, GraphicsUnit.Point); // tie font to DPI
            StartPosition = FormStartPosition.CenterParent;
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Dpi.InitForWindow(this.Handle);
            ApplyTitleBarStyle(this.Handle);
            ApplyScaledMetrics();
        }

        protected override void WndProc(ref Message m)
        {
            const int WM_DPICHANGED = 0x02E0;
            if (m.Msg == WM_DPICHANGED)
            {
                // 1) Update current DPI value
                Dpi.UpdateFromWmDpiChanged(m.WParam);
                // 2) Resize to suggested rect (avoids fuzzy non-client areas)
                var suggested = (RECT)Marshal.PtrToStructure(m.LParam, typeof(RECT));
                SetBounds(suggested.Left, suggested.Top, suggested.Right - suggested.Left, suggested.Bottom - suggested.Top);
                // 3) Reapply metrics (fonts / row heights / paddings)
                ApplyScaledMetrics();
            }
            base.WndProc(ref m);
        }

        // Override in derived forms if you need custom tweaks; otherwise we do sane defaults here.
        protected virtual void ApplyScaledMetrics()
        {
            // Update base font against current DPI so standard controls follow along
            Font = new Font(Font.FontFamily, 9f, Font.Style, GraphicsUnit.Point);
            // Walk common controls and fix OwnerDraw heights that WinForms doesn't auto-scale
            FixOwnerDrawHeights(this);
        }

        private void FixOwnerDrawHeights(Control root)
        {
            foreach (Control c in root.Controls)
            {
                if (c is ListBox lb && lb.DrawMode != DrawMode.Normal)
                    lb.ItemHeight = Dpi.Px(36); // your design row height
                if (c is ListView lv)
                {
                    // Ensure checkboxes / row spacing look right when you use OwnerDraw
                    lv.Font = new Font(lv.Font.FontFamily, 9f, lv.Font.Style);
                }
                // Recursively fix children
                if (c.HasChildren) FixOwnerDrawHeights(c);
            }
        }

        // Apply black title bar with white text (Windows 10 build 17763+)
        private static void ApplyTitleBarStyle(IntPtr hwnd)
        {
            try
            {
                // Convert color to COLORREF format (0x00BBGGRR)
                int bgColor = ColorRef(Theme.TitleBarBg);
                int textColor = ColorRef(Theme.TitleBarText);

                // DWMWA_CAPTION_COLOR = 35, DWMWA_TEXT_COLOR = 36
                DwmSetWindowAttribute(hwnd, 35, ref bgColor, 4);
                DwmSetWindowAttribute(hwnd, 36, ref textColor, 4);
            }
            catch
            {
                // Silently fail on older Windows versions that don't support this
            }
        }

        private static int ColorRef(Color c)
        {
            return c.R | (c.G << 8) | (c.B << 16);
        }

        // Native struct for WM_DPICHANGED suggested rect
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        // DWM API for title bar customization
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);
    }
}