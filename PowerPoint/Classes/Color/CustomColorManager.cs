using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing.Drawing2D;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SD = System.Drawing;
using IOFile = System.IO.File;

namespace OfficePragmatics_PowerPoint
{
    public static class CustomColorManager
    {
        public static void ShowDialog(PowerPoint.Application app)
        {
            var pres = app.ActivePresentation;
            if (pres == null)
            {
                MessageBox.Show("Open a presentation first.", "Custom Colors",
                                MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (string.IsNullOrEmpty(pres.FullName))
            {
                MessageBox.Show("Please save your presentation before applying custom colors. This ensures your changes are saved correctly.",
                                "Save Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var existingColors = GetExistingCustomColors(pres);

            using (var frm = new CustomColorForm(existingColors))
            {
                if (frm.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var colors = frm.GetColors();
                        ApplyCustomColors(app, pres, colors);
                        MessageBox.Show("Custom color palette applied!", "Custom Colors",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error applying colors: {ex.Message}",
                                        "Custom Colors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else if (frm.ClearRequested)
                {
                    try
                    {
                        ClearCustomColors(app, pres);
                        MessageBox.Show("Custom colors removed.", "Custom Colors",
                                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error clearing colors: {ex.Message}",
                                        "Custom Colors", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

#region Open-XML helpers

        private static void ApplyCustomColors(PowerPoint.Application app, PowerPoint.Presentation pres, List<CustomColor> colors)
        {
            string srcPath = pres.FullName;
            string modifiedCopyPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try
            {
                pres.SaveCopyAs(modifiedCopyPath);
                bool wasFinal = pres.Final;

                pres.Saved = Office.MsoTriState.msoTrue;
                pres.Close();

                PatchThemeXml(modifiedCopyPath, colors, remove: false);

                var newPres = app.Presentations.Open(modifiedCopyPath, WithWindow: Office.MsoTriState.msoTrue);
                newPres.SaveAs(srcPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation);
                newPres.Final = wasFinal;
            }
            finally
            {
                if (IOFile.Exists(modifiedCopyPath)) IOFile.Delete(modifiedCopyPath);
            }
        }

        private static List<CustomColor> GetExistingCustomColors(PowerPoint.Presentation pres)
        {
            var colors = new List<CustomColor>();
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

                try
                {
                    pres.SaveCopyAs(tempPath);

                    using (Package pkg = Package.Open(tempPath, FileMode.Open, FileAccess.Read))
                    {
                        var themeUris = new List<Uri>();

                        for (int i = 1; i <= 10; i++)
                        {
                            Uri themeUri = new Uri($"/ppt/theme/theme{i}.xml", UriKind.Relative);
                            if (pkg.PartExists(themeUri)) themeUris.Add(themeUri);
                        }

                        Uri defaultThemeUri = new Uri("/ppt/theme/theme.xml", UriKind.Relative);
                        if (pkg.PartExists(defaultThemeUri)) themeUris.Add(defaultThemeUri);

                        if (themeUris.Count == 0) return colors;

                        foreach (Uri themeUri in themeUris)
                        {
                            PackagePart part = pkg.GetPart(themeUri);
                            XDocument xDoc;
                            using (Stream s = part.GetStream(FileMode.Open, FileAccess.Read))
                            {
                                xDoc = XDocument.Load(s);
                            }

                            XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                            XElement custClrLstNode = xDoc.Root.Element(a + "custClrLst");

                            if (custClrLstNode != null)
                            {
                                foreach (XElement custClr in custClrLstNode.Elements(a + "custClr"))
                                {
                                    string name = custClr.Attribute("name")?.Value;
                                    string hex = custClr.Element(a + "srgbClr")?.Attribute("val")?.Value;

                                    if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(hex))
                                    {
                                        if (!colors.Any(c => c.Name == name && c.Hex == hex))
                                            colors.Add(new CustomColor { Name = name, Hex = hex });
                                    }
                                }
                            }
                        }
                    }
                }
                finally
                {
                    if (IOFile.Exists(tempPath)) { try { IOFile.Delete(tempPath); } catch { } }
                }
            }
            catch { }

            return colors;
        }

        private static void ClearCustomColors(PowerPoint.Application app, PowerPoint.Presentation pres)
        {
            string srcPath = pres.FullName;
            string modifiedCopyPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".pptx");

            try
            {
                pres.SaveCopyAs(modifiedCopyPath);
                bool wasFinal = pres.Final;

                pres.Saved = Office.MsoTriState.msoTrue;
                pres.Close();

                PatchThemeXml(modifiedCopyPath, null, remove: true);

                var newPres = app.Presentations.Open(modifiedCopyPath, WithWindow: Office.MsoTriState.msoTrue);
                newPres.SaveAs(srcPath);
                newPres.Final = wasFinal;
            }
            finally
            {
                if (IOFile.Exists(modifiedCopyPath)) IOFile.Delete(modifiedCopyPath);
            }
        }

        private static void PatchThemeXml(string pptxPath, List<CustomColor> colors, bool remove)
        {
            using (Package pkg = Package.Open(pptxPath, FileMode.Open, FileAccess.ReadWrite))
            {
                Uri themeUri = new Uri("/ppt/theme/theme1.xml", UriKind.Relative);
                if (!pkg.PartExists(themeUri)) throw new InvalidOperationException("theme1.xml not found in the package.");

                PackagePart part = pkg.GetPart(themeUri);
                XDocument xDoc;
                using (Stream s = part.GetStream(FileMode.Open, FileAccess.Read))
                {
                    xDoc = XDocument.Load(s);
                }

                XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
                XElement themeRoot = xDoc.Root;
                if (themeRoot == null || themeRoot.Name != a + "theme")
                    throw new InvalidOperationException("theme1.xml does not have a valid <a:theme> root element.");

                XElement themeElements = themeRoot.Element(a + "themeElements");
                if (themeElements == null) throw new InvalidOperationException("Cannot find <a:themeElements> in theme XML.");

                XElement custClrLstNode = themeRoot.Element(a + "custClrLst");

                if (remove)
                {
                    custClrLstNode?.Remove();
                }
                else
                {
                    if (custClrLstNode == null)
                    {
                        custClrLstNode = new XElement(a + "custClrLst");
                        themeElements.AddAfterSelf(custClrLstNode);
                    }

                    custClrLstNode.RemoveAll();

                    if (colors != null)
                    {
                        foreach (var c in colors)
                        {
                            custClrLstNode.Add(new XElement(a + "custClr",
                                    new XAttribute("name", c.Name),
                                    new XElement(a + "srgbClr",
                                        new XAttribute("val", c.Hex.ToUpper()))));
                        }
                    }
                }

                using (Stream s = part.GetStream(FileMode.Create, FileAccess.Write))
                {
                    xDoc.Save(s);
                }
            }
        }
        #endregion
    }

    internal class CustomColor
    {
        public string Name;
        public string Hex;
    }

    #region Enhanced WinForms UI
    internal sealed class CustomColorForm : Form
    {
        private static class Caption
        {
            private const int DWMWA_CAPTION_COLOR = 35, DWMWA_TEXT_COLOR = 36;
            [DllImport("dwmapi.dll")]
            private static extern int DwmSetWindowAttribute(IntPtr h, int attr, ref int val, int size);
            private static int Ref(SD.Color c) => c.R | (c.G << 8) | (c.B << 16);
            public static void Apply(IntPtr h)
            {
                int b = Ref(SD.ColorTranslator.FromHtml("#e1966d"));
                int t = Ref(SD.Color.White);
                DwmSetWindowAttribute(h, DWMWA_CAPTION_COLOR, ref b, 4);
                DwmSetWindowAttribute(h, DWMWA_TEXT_COLOR, ref t, 4);
            }
        }
        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); Caption.Apply(Handle); }

        private readonly ColorSwatch[] _swatches = new ColorSwatch[50];
        public bool ClearRequested { get; private set; }

        public CustomColorForm(List<CustomColor> existingColors)
        {
            Text = "Custom Color Palette";
            AutoScaleMode = AutoScaleMode.Dpi;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            Width = 720;
            MinimumSize = new SD.Size(720, 0);
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            Padding = new Padding(20);
            Font = new SD.Font("Segoe UI", 9F);
            BackColor = SD.Color.White;

            var main = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 4
            };
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(main);

            var instructions = new Label
            {
                AutoSize = true,
                MaximumSize = new SD.Size(650, 0),
                Margin = new Padding(0, 0, 0, 10),
                ForeColor = SD.Color.FromArgb(90, 97, 105),
                Text = "Click any swatch to edit. Use the 'Pick Color' button in the dialog for visual selection with automatic color naming. Define primary colors in the top row, then use 'Generate Shades' to create variations below."
            };
            main.Controls.Add(instructions, 0, 0);

            var card = new Panel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                Padding = new Padding(12),
                Margin = new Padding(0, 0, 0, 8),
                BackColor = SD.Color.FromArgb(250, 250, 251)
            };
            card.Paint += (s, e) =>
            {
                var g = e.Graphics;
                var r = card.ClientRectangle; r.Width -= 1; r.Height -= 1;
                using (var pen = new SD.Pen(SD.Color.FromArgb(225, 228, 232)))
                using (var bg = new SD.SolidBrush(card.BackColor))
                { g.FillRectangle(bg, 0, 0, r.Width, r.Height); g.DrawRectangle(pen, 0, 0, r.Width, r.Height); }
            };
            main.Controls.Add(card, 0, 1);

            var grid = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                ColumnCount = 10,
                RowCount = 5,
                Padding = new Padding(4),
                BackColor = SD.Color.Transparent
            };
            for (int c = 0; c < 10; c++) grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 10f));
            for (int r = 0; r < 5; r++) grid.RowStyles.Add(new RowStyle(SizeType.Absolute, 46f));

            for (int i = 0; i < 50; i++)
            {
                var sw = new ColorSwatch(i) { TabStop = true, TabIndex = i, Anchor = AnchorStyles.None };
                sw.Click += Sw_Click;
                _swatches[i] = sw;
                grid.Controls.Add(sw, i % 10, i / 10);
            }
            card.Controls.Add(grid);

            if (existingColors != null)
            {
                for (int i = 0; i < existingColors.Count && i < _swatches.Length; i++)
                {
                    var c = existingColors[i];
                    if (c != null && !string.Equals(c.Hex, "FFFFFF", StringComparison.OrdinalIgnoreCase))
                        _swatches[i].SetColor(c.Name, c.Hex);
                }
            }

            var btnGenerate = new ModernButton
            {
                Text = "&Generate Shades",
                ButtonStyle = ModernButton.Style.Secondary,
                Dock = DockStyle.Fill,
                Height = 38,
                Margin = new Padding(0, 4, 0, 6)
            };
            btnGenerate.Click += (s, e) => FillShades();
            main.Controls.Add(btnGenerate, 0, 2);

            var bar = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 16, 0, 0)
            };

            var btnApply = new ModernButton { Text = "&Apply", Width = 92, Height = 36, ButtonStyle = ModernButton.Style.Primary, Margin = new Padding(8, 0, 0, 0) };
            var btnCancel = new ModernButton { Text = "C&ancel", Width = 92, Height = 36, ButtonStyle = ModernButton.Style.Secondary, Margin = new Padding(8, 0, 0, 0) };
            var btnClear = new ModernButton { Text = "&Clear", Width = 92, Height = 36, ButtonStyle = ModernButton.Style.Secondary, Margin = new Padding(8, 0, 0, 0) };

            btnApply.Click += BtnApply_Click;
            btnCancel.Click += (s, e) => DialogResult = DialogResult.Cancel;
            btnClear.Click += BtnClear_Click;

            bar.Controls.AddRange(new Control[] { btnApply, btnCancel, btnClear });
            main.Controls.Add(bar, 0, 3);

            AcceptButton = btnApply;
            CancelButton = btnCancel;
        }

        private void Sw_Click(object sender, EventArgs e)
        {
            var sw = (ColorSwatch)sender;

            using (var dlg = new ColorInputDialog(sw.ColorName, sw.HexString))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                    sw.SetColor(dlg.ColorName, dlg.Hex);
            }
        }

        private void BtnApply_Click(object sender, EventArgs e) => DialogResult = DialogResult.OK;

        private void BtnClear_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Are you sure you want to remove ALL custom colors from this presentation?",
                                "Confirm Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                ClearRequested = true;
                DialogResult = DialogResult.Cancel;
            }
        }

        public List<CustomColor> GetColors()
        {
            var list = new List<CustomColor>(50);
            for (int i = 0; i < 50; i++)
            {
                var sw = _swatches[i];
                list.Add(new CustomColor
                {
                    Name = sw.IsDefault ? $"Color {i + 1}" : sw.ColorName,
                    Hex = sw.HexString.ToUpper()
                });
            }
            return list;
        }

        private void FillShades()
        {
            for (int col = 0; col < 10; col++)
            {
                var top = _swatches[col];
                if (top.IsDefault) continue;

                SD.Color baseCol = top.BackColor;
                double[] factors = { 1.80, 1.40, 0.60, 0.40 };

                for (int row = 1; row < 5; row++)
                {
                    int idx = row * 10 + col;
                    if (idx >= _swatches.Length) break;

                    string shadeHex = ShadeHex(baseCol, factors[row - 1]);
                    string shadeName = $"{top.ColorName} {GetShadeLabel(row)}";
                    _swatches[idx].SetColor(shadeName, shadeHex);
                }
            }
        }

        private static string GetShadeLabel(int row)
            => row == 1 ? "Lightest" : row == 2 ? "Lighter" : row == 3 ? "Darker" : "Darkest";

        private static string ShadeHex(SD.Color c, double f)
        {
            int r = Math.Max(0, Math.Min(255, (int)(c.R * f)));
            int g = Math.Max(0, Math.Min(255, (int)(c.G * f)));
            int b = Math.Max(0, Math.Min(255, (int)(c.B * f)));
            return $"{r:X2}{g:X2}{b:X2}";
        }
    }

    internal sealed class ColorSwatch : Panel
    {
        private readonly ToolTip _tip = new ToolTip();
        private bool _hover, _pressed;

        public string ColorName { get; private set; } = "White";
        public string HexString { get; private set; } = "FFFFFF";
        public bool IsDefault => HexString.Equals("FFFFFF", StringComparison.OrdinalIgnoreCase);

        public ColorSwatch(int index)
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer | ControlStyles.ResizeRedraw, true);

            Width = 40; Height = 40;

            int col = index % 10;
            Margin = new Padding(col == 0 ? 0 : 6, 6, col == 9 ? 0 : 6, 6);

            BackColor = SD.Color.White;
            Cursor = Cursors.Hand;
            Anchor = AnchorStyles.None;

            AccessibleName = $"Color {index + 1}";
            AccessibleDescription = "Click to edit";

            Paint += OnPaint;
            MouseEnter += (s, e) => { _hover = true; Invalidate(); };
            MouseLeave += (s, e) => { _hover = false; _pressed = false; Invalidate(); };
            MouseDown += (s, e) => { _pressed = true; Invalidate(); };
            MouseUp += (s, e) => { _pressed = false; Invalidate(); };

            _tip.ToolTipTitle = "Color Swatch";
            UpdateTip($"Click to edit #{index + 1}");
        }

        private void OnPaint(object sender, PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.None;

            float bw = (_pressed || _hover) ? 2f : 1f;
            var rect = new SD.Rectangle(0, 0, Width - 1, Height - 1);
            if (bw > 1f) rect = SD.Rectangle.Inflate(rect, -1, -1);

            bool nearWhite = BackColor.R > 245 && BackColor.G > 245 && BackColor.B > 245;
            if (nearWhite)
            {
                using (var hb = new HatchBrush(HatchStyle.LargeCheckerBoard,
                           SD.Color.FromArgb(242, 243, 244), SD.Color.White))
                    g.FillRectangle(hb, rect);
            }

            using (var fill = new SD.SolidBrush(BackColor))
                g.FillRectangle(fill, rect);

            SD.Color border = _pressed || _hover
                ? SD.Color.FromArgb(79, 70, 229)
                : (IsDefault ? SD.Color.FromArgb(214, 218, 222) : SD.Color.FromArgb(160, 167, 174));

            using (var pen = new SD.Pen(border, bw))
                g.DrawRectangle(pen, rect);

            if (IsDefault && _hover)
            {
                int t = 2;
                int cx = rect.Left + rect.Width / 2;
                int cy = rect.Top + rect.Height / 2;
                int arm = Math.Min(rect.Width, rect.Height) / 4;

                var plusColor = SD.Color.FromArgb(156, 163, 175);
                using (var br = new SD.SolidBrush(plusColor))
                {
                    var h = new SD.Rectangle(cx - arm, cy - t / 2, arm * 2, t);
                    var v = new SD.Rectangle(cx - t / 2, cy - arm, t, arm * 2);
                    g.FillRectangle(br, h);
                    g.FillRectangle(br, v);
                }
            }
        }

        public void SetColor(string name, string hex)
        {
            ColorName = name;
            HexString = hex;
            try { BackColor = SD.ColorTranslator.FromHtml("#" + hex); }
            catch { BackColor = SD.Color.White; HexString = "FFFFFF"; }
            UpdateTip($"{ColorName}\n#{HexString.ToUpper()}");
            Invalidate();
        }

        private void UpdateTip(string t) => _tip.SetToolTip(this, t);
    }

    internal sealed class ModernButton : Button
    {
        public enum Style { Primary, Secondary }
        private bool _hover, _down;
        public Style ButtonStyle { get; set; } = Style.Primary;

        private readonly SD.Color _bg = SD.Color.White;
        private readonly SD.Color _fg = SD.Color.FromArgb(55, 65, 81);
        private readonly SD.Color _border = SD.Color.FromArgb(209, 213, 219);

        public ModernButton()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint |
                     ControlStyles.DoubleBuffer | ControlStyles.ResizeRedraw, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
            UseMnemonic = true;

            MouseEnter += (s, e) => { _hover = true; Invalidate(); };
            MouseLeave += (s, e) => { _hover = false; _down = false; Invalidate(); };
            MouseDown += (s, e) => { _down = true; Invalidate(); };
            MouseUp += (s, e) => { _down = false; Invalidate(); };
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics; g.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = ClientRectangle;
            var draw = new SD.Rectangle(rect.X, rect.Y, rect.Width - 1, rect.Height - 1);

            SD.Color bg = _bg;
            if (_down) bg = SD.Color.FromArgb(_bg.A, (int)(_bg.R * 0.95), (int)(_bg.G * 0.95), (int)(_bg.B * 0.95));
            else if (_hover) bg = SD.Color.FromArgb(249, 250, 251);

            g.Clear(Parent?.BackColor ?? SD.SystemColors.Control);

            using (var fill = new SD.SolidBrush(bg))
            using (var pen = new SD.Pen(_border))
            {
                g.FillRectangle(fill, draw);
                g.DrawRectangle(pen, draw);
                TextRenderer.DrawText(g, Text, Font, rect, _fg,
                    TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPadding);
            }
        }
    }

    internal sealed class ColorInputDialog : Form
    {
        private static class Caption
        {
            private const int DWMWA_CAPTION_COLOR = 35, DWMWA_TEXT_COLOR = 36;
            [DllImport("dwmapi.dll")]
            private static extern int DwmSetWindowAttribute(IntPtr h, int attr, ref int val, int size);
            private static int Ref(SD.Color c) => c.R | (c.G << 8) | (c.B << 16);
            public static void Apply(IntPtr h)
            {
                int b = Ref(SD.ColorTranslator.FromHtml("#e1966d"));
                int t = Ref(SD.Color.White);
                DwmSetWindowAttribute(h, DWMWA_CAPTION_COLOR, ref b, 4);
                DwmSetWindowAttribute(h, DWMWA_TEXT_COLOR, ref t, 4);
            }
        }
        protected override void OnHandleCreated(EventArgs e) { base.OnHandleCreated(e); Caption.Apply(Handle); }

        private readonly TextBox _txtName, _txtHex;
        private readonly Panel _preview;
        private static readonly Regex HexRe = new Regex("^[0-9A-Fa-f]{6}$");

        public string ColorName => _txtName.Text.Trim();
        public string Hex => _txtHex.Text.ToUpper();

        public ColorInputDialog(string currentName, string currentHex)
        {
            Text = "Edit Color";
            AutoScaleMode = AutoScaleMode.Dpi;
            BackColor = SD.Color.White;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;

            Width = 400;
            Height = 380;
            Font = new SD.Font("Segoe UI", 9F);
            Padding = new Padding(20);

            var main = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 5 };
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110f));
            main.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            main.RowStyles.Add(new RowStyle(SizeType.Absolute, 100F));
            main.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            Controls.Add(main);

            var lblName = new Label { Text = "Label:", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SD.Color.FromArgb(75, 85, 99) };
            _txtName = new TextBox { Text = currentName, Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 4, 0, 8) };
            main.Controls.Add(lblName, 0, 0);
            main.Controls.Add(_txtName, 1, 0);

            var lblHex = new Label { Text = "HEX:", AutoSize = true, Anchor = AnchorStyles.Left, ForeColor = SD.Color.FromArgb(75, 85, 99) };
            _txtHex = new TextBox { Text = currentHex, Dock = DockStyle.Fill, MaxLength = 6, BorderStyle = BorderStyle.FixedSingle, Font = new SD.Font("Consolas", 10f), Margin = new Padding(0, 4, 0, 10) };
            main.Controls.Add(lblHex, 0, 1);
            main.Controls.Add(_txtHex, 1, 1);

            var btnPicker = new ModernButton
            {
                Text = "&Pick Color...",
                Dock = DockStyle.Left,
                Width = 120,
                Height = 34,
                Margin = new Padding(0, 0, 0, 10)
            };
            btnPicker.Click += BtnPicker_Click;
            main.Controls.Add(btnPicker, 1, 2);

            _preview = new Panel
            {
                Dock = DockStyle.Fill,
                MinimumSize = new SD.Size(0, 96),
                Margin = new Padding(0, 8, 0, 12)
            };
            _preview.Paint += (s, e) =>
            {
                var g = e.Graphics;
                var r = _preview.ClientRectangle; r.Width -= 1; r.Height -= 1;
                using (var bg = new SD.SolidBrush(_preview.BackColor))
                using (var pen = new SD.Pen(SD.Color.FromArgb(209, 213, 219)))
                { g.FillRectangle(bg, 0, 0, r.Width, r.Height); g.DrawRectangle(pen, 0, 0, r.Width, r.Height); }
            };
            main.Controls.Add(_preview, 0, 3);
            main.SetColumnSpan(_preview, 2);

            var bar = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                Dock = DockStyle.Fill,
                WrapContents = false,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Margin = new Padding(0, 10, 0, 0)
            };
            var ok = new ModernButton
            {
                Text = "&Save",
                DialogResult = DialogResult.OK,
                Width = 86,
                Height = 34,
                Margin = new Padding(12, 0, 0, 0)
            };

            var cancel = new ModernButton
            {
                Text = "C&ancel",
                DialogResult = DialogResult.Cancel,
                Width = 86,
                Height = 34,
                Margin = new Padding(12, 0, 0, 0)
            };

            bar.Controls.AddRange(new Control[] { ok, cancel });
            main.Controls.Add(bar, 0, 4);
            main.SetColumnSpan(bar, 2);

            ok.Click += OnOk;
            _txtHex.TextChanged += (s, e) => UpdatePreview();
            AcceptButton = ok;
            CancelButton = cancel;
            UpdatePreview();
        }

        private void BtnPicker_Click(object sender, EventArgs e)
        {
            using (var colorDialog = new ColorDialog())
            {
                colorDialog.AllowFullOpen = true;
                colorDialog.FullOpen = true;
                colorDialog.AnyColor = true;

                try
                {
                    colorDialog.Color = SD.ColorTranslator.FromHtml("#" + _txtHex.Text);
                }
                catch
                {
                    colorDialog.Color = SD.Color.White;
                }

                if (colorDialog.ShowDialog(this) == DialogResult.OK)
                {
                    var selectedColor = colorDialog.Color;
                    _txtHex.Text = selectedColor.R.ToString("X2") +
                                  selectedColor.G.ToString("X2") +
                                  selectedColor.B.ToString("X2");

                    if (string.IsNullOrWhiteSpace(_txtName.Text) ||
                        _txtName.Text.StartsWith("Color ") ||
                        _txtName.Text == "White")
                    {
                        _txtName.Text = GetColorName(selectedColor);
                    }
                }
            }
        }

        private static string GetColorName(SD.Color color)
        {
            var knownColor = SD.Color.FromArgb(color.ToArgb());
            foreach (SD.KnownColor known in Enum.GetValues(typeof(SD.KnownColor)))
            {
                var kc = SD.Color.FromKnownColor(known);
                if (kc.R == color.R && kc.G == color.G && kc.B == color.B &&
                    !kc.IsSystemColor && kc.IsNamedColor)
                {
                    return kc.Name;
                }
            }

            float hue = color.GetHue();
            float sat = color.GetSaturation();
            float bri = color.GetBrightness();

            string shade = "";
            if (bri < 0.2f) shade = "Dark ";
            else if (bri < 0.4f) shade = "Deep ";
            else if (bri > 0.8f) shade = "Light ";
            else if (bri > 0.9f) shade = "Pale ";

            string intensity = "";
            if (sat < 0.2f) return bri > 0.9f ? "White" : bri < 0.1f ? "Black" : "Gray";
            else if (sat < 0.4f) intensity = "Muted ";
            else if (sat > 0.8f) intensity = "Vivid ";

            string baseName;
            if (hue < 15 || hue >= 345) baseName = "Red";
            else if (hue < 45) baseName = "Orange";
            else if (hue < 75) baseName = "Yellow";
            else if (hue < 150) baseName = "Green";
            else if (hue < 200) baseName = "Cyan";
            else if (hue < 260) baseName = "Blue";
            else if (hue < 290) baseName = "Purple";
            else baseName = "Magenta";

            return $"{shade}{intensity}{baseName}".Trim();
        }

        private void UpdatePreview()
        {
            try
            {
                _preview.BackColor = HexRe.IsMatch(_txtHex.Text)
                    ? SD.ColorTranslator.FromHtml("#" + _txtHex.Text)
                    : SD.Color.White;
            }
            catch { _preview.BackColor = SD.Color.White; }
            _preview.Invalidate();
        }

        private void OnOk(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(ColorName))
            {
                MessageBox.Show("Please enter a name for the color.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtName.Focus();
                DialogResult = DialogResult.None;
                return;
            }
            if (!HexRe.IsMatch(_txtHex.Text))
            {
                MessageBox.Show("Enter a valid 6-digit hex color (e.g., FF5733).", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _txtHex.Focus();
                DialogResult = DialogResult.None;
            }
        }
    }
    #endregion
}