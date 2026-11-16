using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SDColor = System.Drawing.Color;
using SDPoint = System.Drawing.Point;
using SDPointF = System.Drawing.PointF;
using SDRect = System.Drawing.Rectangle;
using SDSize = System.Drawing.Size;
using SDStringFormat = System.Drawing.StringFormat;
using OfficePragmatics_Excel.Settings;
using ApiKeys = OfficePragmatics_PowerPoint.Settings.ApiKeys;

namespace OfficePragmatics_PowerPoint.Map
{
    public static class MapUtility
    {
        // ── session limit ────────────────────────────────────────────────────
        private static int _sessionGeocodingCount = 0;
        private const int MAX_GEOCODING_PER_SESSION = 1000;

        // ── embedded SVG resource name ───────────────────────────────────────
        private const string MAP_SVG_RESOURCE_NAME =
            "OfficePragmatics_Powerpoint.Resources.Continental_US_Map.svg";

        // ── Mapbox geocoding, called with the user's own access token ─────────
        private const string MAPBOX_GEOCODE_BASE = "https://api.mapbox.com/geocoding/v5/mapbox.places/";

        // ── progress-form bookkeeping ───────────────────────────────────────
        private static MapProgressForm _progressForm;
        private static int _progressCurrent;
        private static int _progressTotal;

        public static void InsertMap(PowerPoint.Application app)
        {
            PowerPoint.Slide slide = null;
            PowerPoint.Shape markerTemplate = null;
            PowerPoint.Shape mapShape = null;
            string tempSvgPath = null;
            var addedMarkers = new List<PowerPoint.Shape>();

            if (!ApiKeys.EnsureMapsKey()) return;
            string mapboxToken = ApiKeys.MapboxAccessToken;

            try
            {
                // 1) Resolve target slide
                if (app.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewNormal ||
                    app.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewSlide)
                {
                    slide = app.ActiveWindow.View.Slide;
                }
                else if (app.ActiveWindow.ViewType == PowerPoint.PpViewType.ppViewSlideSorter &&
                         app.ActiveWindow.Selection.SlideRange.Count > 0)
                {
                    slide = app.ActiveWindow.Selection.SlideRange[1];
                }
                else if (app.ActivePresentation.SlideShowWindow != null)
                {
                    slide = app.ActivePresentation.SlideShowWindow.View.Slide;
                }

                if (slide == null)
                {
                    MessageBox.Show("Please select a slide in Normal or Slide-Sorter view.",
                                    "No slide selected",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // 2) Marker template (optional)
                var sel = app.ActiveWindow.Selection;
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes &&
                    sel.ShapeRange.Count > 0 &&
                    sel.ShapeRange[1].Parent == slide)
                {
                    markerTemplate = sel.ShapeRange[1];
                }

                // 3) Prompt user for locations
                List<string> addresses;
                using (var f = new LocationInputForm())
                {
                    if (f.ShowDialog() != DialogResult.OK) return;
                    addresses = f.GetLocations();
                }

                if (addresses.Count == 0)
                {
                    MessageBox.Show("No locations were entered.",
                                    "Input Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Dedup + session guard
                var uniqueKeys = new HashSet<string>(addresses.Select(NormalizeAddress));
                int neededRequests = uniqueKeys.Count;
                int remaining = MAX_GEOCODING_PER_SESSION - _sessionGeocodingCount;

                if (remaining <= 0)
                {
                    MessageBox.Show($"The session limit of {MAX_GEOCODING_PER_SESSION:N0} geocodes has been reached.\nClose PowerPoint to reset.",
                                    "Geocoding Limit Reached",
                                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                if (neededRequests > remaining)
                {
                    MessageBox.Show($"You requested {addresses.Count:N0} locations but only {remaining:N0} geocodes remain " +
                                    $"in this session (after removing duplicates, {neededRequests:N0} unique look-ups are required).\n\n" +
                                    $"Please reduce the number of locations and try again or restart PowerPoint.",
                                    "Geocoding Limit Exceeded",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // 4) Extract SVG to temp
                var asm = Assembly.GetExecutingAssembly();
                using (var s = asm.GetManifestResourceStream(MAP_SVG_RESOURCE_NAME))
                {
                    if (s == null)
                    {
                        MessageBox.Show($"Resource '{MAP_SVG_RESOURCE_NAME}' not found.",
                                        "Resource Error",
                                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    tempSvgPath = System.IO.Path.Combine(
                        System.IO.Path.GetTempPath(),
                        Guid.NewGuid() + ".svg");

                    using (var fs = new System.IO.FileStream(
                        tempSvgPath,
                        System.IO.FileMode.Create,
                        System.IO.FileAccess.Write))
                    {
                        s.CopyTo(fs);
                    }
                }

                // 5) Insert SVG on slide
                mapShape = slide.Shapes.AddPicture(
                    tempSvgPath,
                    Office.MsoTriState.msoFalse,
                    Office.MsoTriState.msoTrue,
                    100, 100, 600, 350);
                mapShape.Name = "US_Base_Map";

                // 6) Albers Equal-Area (CONUS)
                double φ1 = 29.5 * Math.PI / 180,
                       φ2 = 45.5 * Math.PI / 180,
                       φ0 = 37.5 * Math.PI / 180,
                       λ0 = -96 * Math.PI / 180;

                double n = (Math.Sin(φ1) + Math.Sin(φ2)) / 2.0;
                double C = Math.Pow(Math.Cos(φ1), 2) + 2 * n * Math.Sin(φ1);
                double ρ0 = Math.Sqrt(C - 2 * n * Math.Sin(φ0)) / n;

                Func<double, double, PointF> Project = (lat, lon) =>
                {
                    double φ = lat * Math.PI / 180,
                           λ = lon * Math.PI / 180;

                    double ρ = Math.Sqrt(C - 2 * n * Math.Sin(φ)) / n;
                    double θ = n * (λ - λ0);

                    double px = ρ * Math.Sin(θ);
                    double py = ρ0 - ρ * Math.Cos(θ);

                    double x = mapShape.Left + mapShape.Width / 2 + px * (mapShape.Width / 0.76);
                    double y = mapShape.Top + mapShape.Height / 2 - py * (mapShape.Height / 0.46) + 10;

                    return new PointF((float)x, (float)y);
                };

                // 7) Progress UI
                ShowProgress(addresses.Count);

                ServicePointManager.SecurityProtocol =
                    SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

                var failures = new List<string>();
                var notFound = new List<string>();
                string tokenRejected = null;
                var geoCache = new Dictionary<string, PointF?>();

                using (var wc = new WebClient())
                {
                    wc.Headers.Add("User-Agent", "PowerPointMapAddin/1.0");

                    foreach (string raw in addresses)
                    {
                        _progressCurrent++;
                        UpdateProgress();
                        Application.DoEvents();

                        string key = NormalizeAddress(raw);

                        if (geoCache.TryGetValue(key, out PointF? cached))
                        {
                            if (cached != null)
                                AddMarker(slide, markerTemplate, cached.Value, raw, addedMarkers);
                            else
                                notFound.Add(raw);
                            continue;
                        }

                        try
                        {
                            // Mapbox geocoding (US only, 1 result)
                            var url = MAPBOX_GEOCODE_BASE + Uri.EscapeDataString(raw) + ".json" +
                                      "?access_token=" + Uri.EscapeDataString(mapboxToken) +
                                      "&country=US&limit=1";

                            var jsonText = wc.DownloadString(url);
                            var json = JObject.Parse(jsonText);

                            var cen = json["features"] is JArray fArr &&
                                      fArr.Count > 0 &&
                                      fArr[0]["center"] is JArray cArr &&
                                      cArr.Count >= 2
                                      ? cArr
                                      : null;

                            if (cen != null)
                            {
                                double lon = (double)cen[0];
                                double lat = (double)cen[1];
                                var pt = Project(lat, lon);

                                geoCache[key] = pt;
                                AddMarker(slide, markerTemplate, pt, raw, addedMarkers);
                            }
                            else
                            {
                                geoCache[key] = null;
                                notFound.Add(raw);
                            }

                            _sessionGeocodingCount++;
                        }
                        catch (WebException wex) when (
                            wex.Response is HttpWebResponse r &&
                            (r.StatusCode == HttpStatusCode.Unauthorized || r.StatusCode == HttpStatusCode.Forbidden))
                        {
                            // Every remaining lookup would fail the same way.
                            tokenRejected = "Mapbox rejected your access token (HTTP " + (int)r.StatusCode + "). " +
                                            "Check it in Settings > API Keys.";
                            break;
                        }
                        catch (Exception ex)
                        {
                            geoCache[key] = null;
                            failures.Add(raw + " (" + ex.Message + ")");
                        }

                        if (_progressCurrent < addresses.Count) Thread.Sleep(100);
                    }
                }

                // 8) Summary & grouping
                HideProgress();

                int remain = MAX_GEOCODING_PER_SESSION - _sessionGeocodingCount;
                string msgBox =
                    (tokenRejected != null ? tokenRejected + "\n\n" : "") +
                    "Map created with " + addedMarkers.Count + " of " + addresses.Count + " locations.\n" +
                    "Session usage: " + _sessionGeocodingCount.ToString("N0") + " / " +
                    MAX_GEOCODING_PER_SESSION.ToString("N0") + " (" + remain.ToString("N0") + " remaining)\n\n";

                if (notFound.Count > 0)
                    msgBox += "Not found (" + notFound.Count + "): " + string.Join("; ", notFound.Take(5)) + "\n";
                if (failures.Count > 0)
                    msgBox += "Failed (" + failures.Count + "): " + string.Join("; ", failures.Take(5));

                MessageBox.Show(msgBox, "Map Creation Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (mapShape != null && addedMarkers.Count > 0)
                {
                    var names = new List<string> { mapShape.Name };
                    names.AddRange(addedMarkers.Select(m => m.Name));

                    try
                    {
                        slide.Shapes.Range(names.ToArray()).Group()
                             .Name = "US_Map_With_Coordinates";
                    }
                    catch { /* ignore grouping errors */ }
                }
            }
            finally
            {
                if (!string.IsNullOrEmpty(tempSvgPath) &&
                    System.IO.File.Exists(tempSvgPath))
                {
                    try { System.IO.File.Delete(tempSvgPath); } catch { }
                }

                Release(markerTemplate);
                addedMarkers.ForEach(Release);
                Release(mapShape);
                Release(slide);
            }
        }

        /* ─────────── helpers ─────────── */

        private static void AddMarker(PowerPoint.Slide slide,
                                      PowerPoint.Shape markerTemplate,
                                      PointF pt,
                                      string sourceAddress,
                                      List<PowerPoint.Shape> added)
        {
            PowerPoint.Shape marker =
                markerTemplate != null
                    ? markerTemplate.Duplicate()[1]
                    : slide.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, 0, 0, 10f, 10f);

            if (markerTemplate == null) StyleDefaultMarker(marker);

            string safe = sourceAddress.Replace(" ", "_").Replace(",", "").Replace(".", "");
            if (safe.Length > 20) safe = safe.Substring(0, 20);

            marker.Name = "Marker_" + safe + "_" + added.Count;
            marker.Left = pt.X - marker.Width / 2;
            marker.Top = pt.Y - marker.Height / 2;

            added.Add(marker);
        }

        private static string NormalizeAddress(string addr) =>
            (addr ?? string.Empty).Trim().ToLowerInvariant();

        private static void StyleDefaultMarker(PowerPoint.Shape marker)
        {
            marker.Fill.ForeColor.RGB =
                System.Drawing.ColorTranslator.ToOle(SDColor.Red);
            marker.Line.Visible = Office.MsoTriState.msoTrue;
            marker.Line.ForeColor.RGB =
                System.Drawing.ColorTranslator.ToOle(SDColor.Black);
            marker.Line.Weight = 0.75f;
        }

        private static void ShowProgress(int total)
        {
            _progressTotal = total;
            _progressCurrent = 0;

            _progressForm = (_progressForm == null || _progressForm.IsDisposed)
                ? new MapProgressForm()
                : _progressForm;

            _progressForm.UpdateProgress(0, total);
            if (!Application.OpenForms.OfType<MapProgressForm>().Any())
                _progressForm.Show(Form.ActiveForm ?? null);

            Application.DoEvents();
        }

        private static void UpdateProgress()
        {
            if (_progressForm != null && !_progressForm.IsDisposed)
                _progressForm.UpdateProgress(_progressCurrent, _progressTotal);
        }

        private static void HideProgress()
        {
            if (_progressForm != null && !_progressForm.IsDisposed)
            {
                _progressForm.Close();
                _progressForm.Dispose();
                _progressForm = null;
            }
        }

        private static void Release(object o)
        {
            if (o != null && Marshal.IsComObject(o))
                Marshal.ReleaseComObject(o);
        }

        public static (int used, int limit, int remaining) GetSessionGeocodingUsage()
        {
            int remaining = MAX_GEOCODING_PER_SESSION - _sessionGeocodingCount;
            return (_sessionGeocodingCount, MAX_GEOCODING_PER_SESSION, remaining);
        }

        /* ─────────── LocationInputForm (Responsive + Enter Fix + Data Commit) ─────────── */
        internal class LocationInputForm : Form
        {
            // ── NATIVE INTEROP ──
            [DllImport("user32.dll")] public static extern bool ReleaseCapture();
            [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            // ── THEME COLORS ──
            private struct UIColors
            {
                public static readonly SDColor Surface = SDColor.FromArgb(255, 255, 255);
                public static readonly SDColor SurfaceAlt = SDColor.FromArgb(248, 248, 248);
                public static readonly SDColor TitleBg = SDColor.FromArgb(32, 42, 102);
                public static readonly SDColor TitleText = SDColor.White;
                public static readonly SDColor Border = SDColor.FromArgb(220, 220, 220);
                public static readonly SDColor TextPrimary = SDColor.FromArgb(32, 31, 30);
                public static readonly SDColor TextSecondary = SDColor.FromArgb(96, 94, 92);
                public static readonly SDColor Accent = SDColor.FromArgb(110, 125, 220);
                public static readonly SDColor AccentHover = SDColor.FromArgb(85, 95, 180);
                public static readonly SDColor Danger = SDColor.FromArgb(232, 17, 35);
            }

            /* ── CUSTOM GRID: Stops "Row Down" on Enter ── */
            private class EnterAwareDataGridView : DataGridView
            {
                // 1. Handle Enter when a cell is selected (Navigation Mode)
                protected override bool ProcessDialogKey(Keys keyData)
                {
                    if (keyData == Keys.Enter)
                    {
                        // Find parent form and click OK
                        Form f = this.FindForm();
                        if (f != null && f.AcceptButton != null)
                        {
                            f.AcceptButton.PerformClick();
                            return true; // We handled it, do NOT move down
                        }
                    }
                    return base.ProcessDialogKey(keyData);
                }

                // 2. Handle Enter when typing inside a cell (Edit Mode)
                protected override bool ProcessDataGridViewKey(KeyEventArgs e)
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        // Find parent form and click OK
                        Form f = this.FindForm();
                        if (f != null && f.AcceptButton != null)
                        {
                            f.AcceptButton.PerformClick();
                            return true; // We handled it, do NOT move down
                        }
                    }
                    return base.ProcessDataGridViewKey(e);
                }
            }

            private readonly Label instructionLabel;
            private readonly EnterAwareDataGridView grid; // Using custom class
            private readonly Button okBtn;
            private readonly Button cancelBtn;
            private readonly Panel buttonPanel;
            private readonly Panel mainPanel;
            private readonly Panel pnlTitle;
            private readonly Label hintLabel;
            private readonly Label btnClose;

            // Scaling fields
            private float _scaleFactor = 1.0f;
            private Font _titleFont;
            private Font _closeButtonFont;

            public LocationInputForm()
            {
                // 1. Calculate DPI Scale
                using (Graphics g = CreateGraphics())
                {
                    _scaleFactor = g.DpiX / 96f;
                }

                // 2. Form Settings
                this.FormBorderStyle = FormBorderStyle.None;
                this.DoubleBuffered = true;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.BackColor = UIColors.Border;
                this.Padding = new Padding(1);
                this.ResizeRedraw = true;
                this.KeyPreview = true;

                // Scale Initial Size
                this.Size = new SDSize((int)(550 * _scaleFactor), (int)(550 * _scaleFactor));

                // 3. Main Surface
                mainPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = UIColors.Surface
                };
                this.Controls.Add(mainPanel);

                // 4. Fonts
                _titleFont = new Font("Segoe UI Semibold", 11f);
                _closeButtonFont = new Font("Segoe UI", 11f);

                // 5. Title Bar
                int titleHeight = (int)(40 * _scaleFactor);
                pnlTitle = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = titleHeight,
                    BackColor = UIColors.TitleBg,
                };
                pnlTitle.MouseDown += (s, e) => {
                    if (e.Button == MouseButtons.Left)
                    {
                        ReleaseCapture();
                        SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                    }
                };

                pnlTitle.Paint += (s, e) =>
                {
                    e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
                    e.Graphics.CompositingQuality = CompositingQuality.HighQuality;

                    using (var brush = new SolidBrush(UIColors.TitleText))
                    {
                        float y = (pnlTitle.Height - _titleFont.Height) / 2f;
                        e.Graphics.DrawString("Enter Locations", _titleFont, brush, (int)(15 * _scaleFactor), y);
                    }
                };

                btnClose = new Label
                {
                    Text = "✕",
                    Font = _closeButtonFont,
                    ForeColor = UIColors.TitleText,
                    Dock = DockStyle.Right,
                    Width = (int)(40 * _scaleFactor),
                    TextAlign = ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = UIColors.TitleBg
                };
                btnClose.MouseEnter += (s, e) => { btnClose.BackColor = UIColors.Danger; btnClose.ForeColor = SDColor.White; };
                btnClose.MouseLeave += (s, e) => { btnClose.BackColor = UIColors.TitleBg; btnClose.ForeColor = UIColors.TitleText; };
                btnClose.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
                pnlTitle.Controls.Add(btnClose);

                mainPanel.Controls.Add(pnlTitle);

                // 6. Content Area
                instructionLabel = new Label
                {
                    Text = "Continental U.S. locations only. For large entries, paste a ';' delimited list or selection from Excel.",
                    ForeColor = UIColors.TextSecondary,
                    Font = new Font("Segoe UI", 9f)
                };
                mainPanel.Controls.Add(instructionLabel);

                // Instantiate Custom Grid
                grid = new EnterAwareDataGridView
                {
                    AllowUserToAddRows = true,
                    AllowUserToDeleteRows = true,
                    AllowUserToResizeRows = false,
                    RowHeadersWidth = (int)(40 * _scaleFactor),
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    SelectionMode = DataGridViewSelectionMode.CellSelect,
                    EnableHeadersVisualStyles = false,
                    BackgroundColor = UIColors.Surface,
                    GridColor = UIColors.Border,
                    BorderStyle = BorderStyle.FixedSingle,
                    DefaultCellStyle = new DataGridViewCellStyle
                    {
                        BackColor = SDColor.White,
                        ForeColor = UIColors.TextPrimary,
                        SelectionBackColor = UIColors.Accent,
                        SelectionForeColor = SDColor.White,
                        Padding = new Padding(2),
                        Font = new Font("Segoe UI", 9f)
                    },
                    AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
                    {
                        BackColor = UIColors.SurfaceAlt
                    }
                };

                grid.RowTemplate.Height = (int)(28 * _scaleFactor);

                grid.ColumnHeadersDefaultCellStyle.BackColor = UIColors.SurfaceAlt;
                grid.ColumnHeadersDefaultCellStyle.ForeColor = UIColors.TextPrimary;
                grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9f);
                grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                grid.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                grid.ColumnHeadersHeight = (int)(32 * _scaleFactor);

                grid.Columns.Add(new DataGridViewTextBoxColumn
                {
                    Name = "Location",
                    HeaderText = "Location (Address or City/State)",
                    SortMode = DataGridViewColumnSortMode.NotSortable
                });

                grid.RowPostPaint += (s, e) =>
                {
                    string idx = (e.RowIndex + 1).ToString();
                    var fmt = new SDStringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    var rect = new SDRect(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height);
                    using (var brush = new SolidBrush(UIColors.TextSecondary))
                        e.Graphics.DrawString(idx, grid.Font, brush, rect, fmt);
                };

                // Handle Paste
                grid.KeyDown += (s, e) => {
                    if (e.Control && e.KeyCode == Keys.V)
                    {
                        e.Handled = true;
                        Paste();
                    }
                };

                mainPanel.Controls.Add(grid);

                // 7. Bottom Button Panel
                int panelHeight = (int)(50 * _scaleFactor);
                buttonPanel = new Panel
                {
                    Height = panelHeight,
                    Dock = DockStyle.Bottom,
                    BackColor = UIColors.Surface
                };
                buttonPanel.Paint += (s, e) => { using (var pen = new Pen(UIColors.Border)) e.Graphics.DrawLine(pen, 0, 0, buttonPanel.Width, 0); };

                okBtn = new Button
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Size = new SDSize((int)(100 * _scaleFactor), (int)(32 * _scaleFactor)),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = UIColors.Accent,
                    ForeColor = SDColor.White,
                    Cursor = Cursors.Hand,
                    Font = new Font("Segoe UI Semibold", 9.5f)
                };
                okBtn.FlatAppearance.BorderSize = 0;

                // ── CRITICAL FIX: Ensure Data is Committed on OK/Enter ──
                okBtn.Click += (s, e) => {
                    if (grid.IsCurrentCellInEditMode)
                    {
                        grid.EndEdit();
                    }
                };

                cancelBtn = new Button
                {
                    Text = "Cancel",
                    DialogResult = DialogResult.Cancel,
                    Size = new SDSize((int)(100 * _scaleFactor), (int)(32 * _scaleFactor)),
                    FlatStyle = FlatStyle.Flat,
                    BackColor = UIColors.Surface,
                    ForeColor = UIColors.TextSecondary,
                    Cursor = Cursors.Hand,
                    Font = new Font("Segoe UI", 9.5f)
                };
                cancelBtn.FlatAppearance.BorderColor = UIColors.Border;
                cancelBtn.FlatAppearance.MouseOverBackColor = UIColors.SurfaceAlt;

                hintLabel = new Label
                {
                    Text = "Tip: Ctrl+V to paste list from Excel",
                    AutoSize = true,
                    ForeColor = UIColors.TextSecondary,
                    Font = new Font("Segoe UI", 8.5f)
                };

                buttonPanel.Controls.Add(hintLabel);
                buttonPanel.Controls.Add(cancelBtn);
                buttonPanel.Controls.Add(okBtn);
                mainPanel.Controls.Add(buttonPanel);

                AcceptButton = okBtn;
                CancelButton = cancelBtn;

                for (int i = 0; i < 10; i++) grid.Rows.Add("");

                this.Resize += (s, e) => LayoutControls();
                LayoutControls();
            }

            private void LayoutControls()
            {
                int pad = (int)(20 * _scaleFactor);
                int titleH = pnlTitle.Height;
                int bottomH = buttonPanel.Height;

                instructionLabel.Location = new SDPoint(pad, titleH + (int)(15 * _scaleFactor));
                instructionLabel.Size = new SDSize(mainPanel.Width - (2 * pad), (int)(40 * _scaleFactor));

                int gridTop = instructionLabel.Bottom + (int)(10 * _scaleFactor);
                grid.Location = new SDPoint(pad, gridTop);
                grid.Size = new SDSize(mainPanel.Width - (2 * pad), mainPanel.Height - bottomH - gridTop - (int)(10 * _scaleFactor));

                int btnMargin = (int)(10 * _scaleFactor);
                int btnTop = (bottomH - okBtn.Height) / 2;

                okBtn.Location = new SDPoint(buttonPanel.Width - okBtn.Width - pad, btnTop);
                cancelBtn.Location = new SDPoint(okBtn.Left - cancelBtn.Width - btnMargin, btnTop);

                hintLabel.Location = new SDPoint(pad, (bottomH - hintLabel.Height) / 2);
            }

            private void Paste()
            {
                try
                {
                    string clip = Clipboard.GetText();
                    if (string.IsNullOrEmpty(clip)) return;

                    var usage = MapUtility.GetSessionGeocodingUsage();
                    if (usage.remaining <= 0)
                    {
                        MessageBox.Show("The session geocoding limit has been reached.", "Limit Reached", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    List<string> items;
                    if (clip.Contains("\r\n") || clip.Contains("\n"))
                    {
                        items = clip.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                    }
                    else
                    {
                        items = clip.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                                    .Select(s => s.Trim()).Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                    }

                    if (items.Count > usage.remaining)
                    {
                        MessageBox.Show($"You attempted to paste {items.Count:N0} locations but only {usage.remaining:N0} requests remain.\nPlease reduce your selection.", "Limit Exceeded", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    int startRow = grid.CurrentCell != null ? grid.CurrentCell.RowIndex : 0;
                    int row = startRow;

                    foreach (string loc in items)
                    {
                        if (row >= grid.Rows.Count - 1) grid.Rows.Add();
                        grid.Rows[row].Cells[0].Value = loc;
                        row++;
                    }

                    grid.ClearSelection();
                    for (int i = startRow; i < row; i++) grid.Rows[i].Cells[0].Selected = true;
                    if (row - 1 < grid.Rows.Count) grid.FirstDisplayedScrollingRowIndex = Math.Max(0, startRow);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Paste failed: {ex.Message}", "Paste Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            public List<string> GetLocations() =>
                grid.Rows.Cast<DataGridViewRow>()
                    .Select(r => r.Cells[0].Value as string)
                    .Where(s => !string.IsNullOrWhiteSpace(s))
                    .Select(s => s.Trim())
                    .ToList();

            protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
            {
                // Global Catch: If focus is anywhere on the form, Enter triggers OK.
                if (keyData == Keys.Enter)
                {
                    if (grid.IsCurrentCellInEditMode) grid.EndEdit();
                    okBtn.PerformClick();
                    return true;
                }
                if (keyData == (Keys.Control | Keys.V)) { Paste(); return true; }
                return base.ProcessCmdKey(ref msg, keyData);
            }

            protected override void OnFormClosed(FormClosedEventArgs e)
            {
                base.OnFormClosed(e);
                if (_titleFont != null) _titleFont.Dispose();
                if (_closeButtonFont != null) _closeButtonFont.Dispose();
            }
        }
    }

    /* ─────────── MapProgressForm (Responsive) ─────────── */
    public class MapProgressForm : Form
    {
        private int _current;
        private int _total;
        private string _progressText;
        private float _scaleFactor = 1.0f;

        public MapProgressForm()
        {
            // 1. Calculate Scale
            using (Graphics g = CreateGraphics())
            {
                _scaleFactor = g.DpiX / 96f;
            }

            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.Manual;
            this.Size = new System.Drawing.Size((int)(300 * _scaleFactor), (int)(80 * _scaleFactor));
            this.BackColor = SDColor.Black;
            this.Opacity = 0.85;
            this.TopMost = true;
            this.ShowInTaskbar = false;
            this.Name = "MapProgressForm";
            this.ResizeRedraw = true;

            var bounds = Screen.PrimaryScreen.Bounds;
            this.Location = new SDPoint(
                bounds.Width - this.Width - (int)(20 * _scaleFactor),
                bounds.Height - this.Height - (int)(20 * _scaleFactor));

            this.SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            this.Paint += MapProgressForm_Paint;
        }

        private void MapProgressForm_Paint(object sender, PaintEventArgs e)
        {
            if (string.IsNullOrEmpty(_progressText))
                _progressText = $"Processing: {_current} / {_total}";

            e.Graphics.Clear(this.BackColor);

            // Scale Font explicitly
            using (var brush = new SolidBrush(SDColor.White))
            using (var font = new Font("Segoe UI", 10 * _scaleFactor, FontStyle.Bold))
            {
                SizeF textSize = e.Graphics.MeasureString(_progressText, font);
                PointF locationToDraw = new PointF(
                    (this.ClientSize.Width - textSize.Width) / 2,
                    (this.ClientSize.Height - textSize.Height) / 2
                );
                e.Graphics.DrawString(_progressText, font, brush, locationToDraw);
            }
        }

        public void UpdateProgress(int current, int total)
        {
            _current = current;
            _total = total;
            _progressText = $"Processing: {current} / {total}";
            if (this.IsHandleCreated && !this.IsDisposed) this.Invalidate();
        }
    }
}