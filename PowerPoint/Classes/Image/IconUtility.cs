// IconUtility.cs
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SkiaSharp;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Drawing.Text; // Added for TextRenderingHint
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Security.Authentication;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SD = System.Drawing; // Strict alias
using IOFile = System.IO.File;
using IODir = System.IO.Directory;
using ApiKeys = OfficePragmatics_PowerPoint.Settings.ApiKeys;

namespace OfficePragmatics_PowerPoint.Image
{
    /*──────────────────────── DTO ────────────────────────*/
    public class IconResult
    {
        public string Id { get; set; }
        public string ThumbnailUrl { get; set; }
        public string Term { get; set; }
    }

    // NOTE: HttpStatusException and UnsupportedFormatException 
    // are already defined in LogoUtility.cs within this namespace.
    // DO NOT redefine them here.

    /*──────────── MAIN STATIC UTILITY ─────────────*/
    public static class IconUtility
    {
        // Freepik API, called with the user's own key.
        private const string FREEPIK_API_BASE = "https://api.freepik.com/v1/icons";
        private const int HTTP_TIMEOUT_MS = 15000;

        private static readonly HttpClientHandler _handler =
            new HttpClientHandler { SslProtocols = SslProtocols.Tls12, AllowAutoRedirect = true };

        /*──────── persist-able Recents ────────*/
        private const int MAX_RECENT = 24;
        private static readonly List<IconResult> _recent = new List<IconResult>();
        internal static IReadOnlyList<IconResult> Recent => _recent;

        private static readonly string _recentsFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OfficePragmatics", "recent-icons.json");

        static IconUtility() { LoadRecents(); }

        private static void LoadRecents()
        {
            try
            {
                if (IOFile.Exists(_recentsFile))
                {
                    string json = IOFile.ReadAllText(_recentsFile);
                    var list = JsonConvert.DeserializeObject<List<IconResult>>(json) ?? new List<IconResult>();
                    _recent.AddRange(list);
                }
            }
            catch { }
        }
        private static void SaveRecents()
        {
            try
            {
                string dir = Path.GetDirectoryName(_recentsFile);
                if (!IODir.Exists(dir)) IODir.CreateDirectory(dir);
                IOFile.WriteAllText(_recentsFile, JsonConvert.SerializeObject(_recent));
            }
            catch { }
        }
        internal static void AddRecent(IconResult icon)
        {
            if (icon == null) return;
            _recent.RemoveAll(r => r.Id == icon.Id);
            _recent.Insert(0, icon);
            if (_recent.Count > MAX_RECENT) _recent.RemoveRange(MAX_RECENT, _recent.Count - MAX_RECENT);
            SaveRecents();
        }
        public static Task IconLookup(PowerPoint.Application app) => InsertIcon(app);

        /*──────────── PUBLIC ENTRY ────────────*/
        public static async Task InsertIcon(PowerPoint.Application app)
        {
            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

            if (!ApiKeys.EnsureIconSearchKey()) return;

            using (var dlg = new IconSearchForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK || dlg.SelectedIcon == null) return;

                try
                {
                    byte[] svg = await DownloadSvgAsync(dlg.SelectedIcon.Id);
                    InsertIconOnSlide(app.ActiveWindow.View.Slide, svg, dlg.SelectedIcon.Term);
                    AddRecent(dlg.SelectedIcon);
                }
                catch (HttpStatusException hex)
                {
                    MessageBox.Show($"Download failed ({(int)hex.StatusCode}): {hex.ResponseBody}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Download / insert failed:\n{ex.Message}",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        /*────────── SEARCH (Freepik, user's own key) ──────────*/
        private const int MAX_RESULTS = 30;

        private static Dictionary<string, string> FreepikHeaders()
        {
            string key = ApiKeys.FreepikApiKey;
            if (string.IsNullOrEmpty(key))
                throw new InvalidOperationException("Add your Freepik API key in Settings > API Keys to search for icons.");

            return new Dictionary<string, string>
            {
                ["x-freepik-api-key"] = key,
                ["Accept"] = "application/json"
            };
        }

        internal static async Task<List<IconResult>> SearchIconsAsync(string term)
        {
            term = (term ?? string.Empty).Trim();
            var list = new List<IconResult>();
            if (term.Length == 0) return list;

            string url = $"{FREEPIK_API_BASE}?term={Uri.EscapeDataString(term)}" +
                         $"&per_page={MAX_RESULTS}&filters[free_svg]=free";

            var dl = await FreepikRequestAsync(url, "Freepik search failed");
            var arr = JObject.Parse(Encoding.UTF8.GetString(dl.Bytes))["data"] as JArray ?? new JArray();

            foreach (var it in arr)
            {
                string id = it["id"]?.ToString().Trim();
                if (string.IsNullOrEmpty(id)) continue;

                string thumb = (it["thumbnails"] as JArray)?.FirstOrDefault()?["url"]?.ToString();
                if (string.IsNullOrEmpty(thumb)) thumb = it["preview"]?["url"]?.ToString();
                if (string.IsNullOrEmpty(thumb)) thumb = (it["images"] as JArray)?.FirstOrDefault()?["url"]?.ToString();

                list.Add(new IconResult
                {
                    Id = id,
                    ThumbnailUrl = thumb ?? string.Empty,
                    Term = it["name"]?.ToString() ?? term
                });

                if (list.Count >= MAX_RESULTS) break;
            }
            return list;
        }

        /*────────── DOWNLOAD SVG (Freepik, user's own key) ──────────*/
        private static async Task<byte[]> DownloadSvgAsync(string id)
        {
            // 1) Ask Freepik for the SVG URL
            string metaUrl = $"{FREEPIK_API_BASE}/{Uri.EscapeDataString(id)}/download?format=svg";
            var meta = await FreepikRequestAsync(metaUrl, "Freepik download failed");

            string svgUrl = JObject.Parse(Encoding.UTF8.GetString(meta.Bytes))["data"]?["url"]?.ToString();
            if (string.IsNullOrEmpty(svgUrl)) throw new InvalidOperationException("Freepik did not return an SVG URL.");

            // 2) Fetch the actual SVG
            var dl = await DownloadBytesAsync(svgUrl, new Dictionary<string, string>
            {
                ["Accept"] = "image/svg+xml,text/xml,application/xml;q=0.9,*/*;q=0.8"
            });

            // basic sanity: ensure looks like SVG
            if (!LooksLikeSvg(dl.Bytes)) throw new InvalidOperationException("Downloaded file is not a valid SVG.");
            return dl.Bytes;
        }

        /// <summary>Calls the Freepik API and turns HTTP errors into a readable message.</summary>
        private static async Task<DownloadResult> FreepikRequestAsync(string url, string what)
        {
            try
            {
                return await DownloadBytesAsync(url, FreepikHeaders());
            }
            catch (HttpStatusException hex)
            {
                string message = null;
                try { message = JObject.Parse(hex.ResponseBody)["message"]?.ToString(); } catch { }

                string hint;
                switch ((int)hex.StatusCode)
                {
                    case 401:
                    case 403: hint = "Check your Freepik API key in Settings > API Keys."; break;
                    case 429: hint = "Your Freepik API rate limit has been reached. Try again later."; break;
                    default: hint = null; break;
                }

                var sb = new StringBuilder($"{what}: HTTP {(int)hex.StatusCode} ({hex.StatusCode})");
                if (!string.IsNullOrWhiteSpace(message)) sb.Append("\n").Append(message);
                if (hint != null) sb.Append("\n").Append(hint);
                throw new HttpStatusException(hex.StatusCode, hex.ResponseBody, sb.ToString());
            }
        }

        private static bool LooksLikeSvg(byte[] bytes)
        {
            if (bytes == null || bytes.Length < 32) return false;
            string s = Encoding.UTF8.GetString(bytes, 0, Math.Min(bytes.Length, 4096)).ToLowerInvariant().TrimStart('\uFEFF');
            return s.Contains("<svg");
        }

        /*────────── INSERT SVG ──────────*/
        private static void InsertIconOnSlide(PowerPoint.Slide slide, byte[] svg, string alt)
        {
            string tmp = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}.svg");
            try
            {
                IOFile.WriteAllBytes(tmp, svg);
                var pic = slide.Shapes.AddPicture(tmp, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 100, 100);

                const double tgtW = 35.0;
                pic.LockAspectRatio = Office.MsoTriState.msoTrue;
                pic.Width = (float)tgtW;

                var p = FindSpot(slide, pic.Width, pic.Height);
                pic.Left = p.X; pic.Top = p.Y;
                pic.AlternativeText = alt;
            }
            finally
            {
                try { if (IOFile.Exists(tmp)) IOFile.Delete(tmp); } catch { }
            }
        }

        private static SD.PointF FindSpot(PowerPoint.Slide s, float w, float h)
        {
            const float colW = 144f, rowH = 50.4f, margin = 15f, top = 0f;
            float startX = s.Master.Width + margin;
            int maxRows = (int)Math.Floor((s.Master.Height - top) / rowH);
            int col = 0, row = 0;

            for (int i = 0; i < 100; i++)
            {
                float x = startX + col * colW;
                float y = top + row * rowH + (rowH - h) / 2;

                var r = new SD.RectangleF(x, y, w, h);
                if (!Collides(s, r)) return new SD.PointF(x, y);

                if (++row >= maxRows) { row = 0; col++; }
            }
            return new SD.PointF(startX, top);
        }
        private static bool Collides(PowerPoint.Slide s, SD.RectangleF r)
        {
            foreach (PowerPoint.Shape sh in s.Shapes)
                if (r.IntersectsWith(new SD.RectangleF(sh.Left, sh.Top, sh.Width, sh.Height)))
                    return true;
            return false;
        }

        /*────────── THUMBNAILS (direct from Freepik's CDN) ──────────*/
        private const int THUMB_CANVAS = 200;
        private const int THUMB_MARGIN = 16;

        internal static async Task<SD.Bitmap> LoadThumbnailAsync(string thumbUrl, string _)
        {
            if (string.IsNullOrEmpty(thumbUrl)) return null;

            var dl = await DownloadBytesAsync(UnwrapLegacyProxyUrl(thumbUrl), new Dictionary<string, string>
            {
                ["Accept"] = "image/*,*/*;q=0.8"
            });

            string ct = (dl.ContentType ?? string.Empty).ToLowerInvariant();
            if (ct.Length > 0 && !ct.StartsWith("image/") && !ct.StartsWith("application/octet-stream"))
                throw new HttpStatusException((HttpStatusCode)415, string.Empty, "Thumbnail is not an image.");

            using (var ms = new MemoryStream(dl.Bytes))
            using (var src = SD.Image.FromStream(ms))
            {
                var bmp = new SD.Bitmap(THUMB_CANVAS, THUMB_CANVAS);
                using (var g = SD.Graphics.FromImage(bmp))
                {
                    g.Clear(SD.Color.Transparent);
                    g.SmoothingMode = SD.Drawing2D.SmoothingMode.HighQuality;

                    float scale = Math.Min((THUMB_CANVAS - 2f * THUMB_MARGIN) / src.Width,
                                             (THUMB_CANVAS - 2f * THUMB_MARGIN) / src.Height);
                    int w = Math.Max(1, (int)(src.Width * scale));
                    int h = Math.Max(1, (int)(src.Height * scale));
                    int x = (THUMB_CANVAS - w) / 2;
                    int y = (THUMB_CANVAS - h) / 2;
                    g.DrawImage(src, x, y, w, h);
                }
                return bmp;
            }
        }

        /// <summary>
        /// Results saved by older versions point at the hosted Office Pragmatics proxy
        /// (…officepragmatics.com/thumb?url=…). Recover the original URL from them.
        /// </summary>
        private static string UnwrapLegacyProxyUrl(string url)
        {
            try
            {
                var u = new Uri(url);
                if (u.Host.EndsWith(".officepragmatics.com", StringComparison.OrdinalIgnoreCase))
                {
                    string inner = System.Web.HttpUtility.ParseQueryString(u.Query)["url"];
                    if (!string.IsNullOrEmpty(inner)) return inner;
                }
            }
            catch { }
            return url;
        }

        /*────────── Robust downloader ──────────*/
        private const string BrowserUserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)";

        private struct DownloadResult { public byte[] Bytes; public string ContentType; }

        private static async Task<DownloadResult> DownloadBytesAsync(
            string url,
            IDictionary<string, string> headers)
        {
            Exception last = null;
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 | SecurityProtocolType.SystemDefault; } catch { }

            headers = headers ?? new Dictionary<string, string>();
            if (!headers.ContainsKey("Accept")) headers["Accept"] = "*/*";

            // Attempt with HttpClient
            try
            {
                using (var client = new HttpClient(_handler, disposeHandler: false))
                {
                    client.Timeout = TimeSpan.FromMilliseconds(HTTP_TIMEOUT_MS);
                    client.DefaultRequestHeaders.TryAddWithoutValidation("User-Agent", BrowserUserAgent);
                    foreach (var h in headers)
                        client.DefaultRequestHeaders.TryAddWithoutValidation(h.Key, h.Value);

                    using (var resp = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
                    {
                        if (!resp.IsSuccessStatusCode)
                        {
                            string body = string.Empty;
                            try
                            {
                                var buf = await resp.Content.ReadAsByteArrayAsync();
                                body = Encoding.UTF8.GetString(buf, 0, Math.Min(buf.Length, 1024));
                            }
                            catch { }
                            throw new HttpStatusException(resp.StatusCode, body);
                        }

                        var ct = resp.Content.Headers.ContentType?.ToString();
                        var bytes = await resp.Content.ReadAsByteArrayAsync();
                        return new DownloadResult { Bytes = bytes, ContentType = ct ?? "" };
                    }
                }
            }
            catch (HttpStatusException) { throw; }
            catch (Exception ex) { last = ex; }

            // Fallback WebClient
            try
            {
                using (var wc = new WebClient())
                {
                    wc.Headers[HttpRequestHeader.UserAgent] = BrowserUserAgent;
                    foreach (var h in headers)
                        wc.Headers[h.Key] = h.Value;
                    byte[] bytes = await wc.DownloadDataTaskAsync(url);
                    string ct = wc.ResponseHeaders["Content-Type"] ?? "";
                    return new DownloadResult { Bytes = bytes, ContentType = ct };
                }
            }
            catch (WebException wex) when (wex.Response is HttpWebResponse r)
            {
                string body = string.Empty;
                try
                {
                    using (var s = r.GetResponseStream())
                    using (var ms = new MemoryStream())
                    {
                        s.CopyTo(ms);
                        var buf = ms.ToArray();
                        body = Encoding.UTF8.GetString(buf, 0, Math.Min(buf.Length, 1024));
                    }
                }
                catch { }
                throw new HttpStatusException(r.StatusCode, body, wex.Message);
            }
            catch (Exception ex2)
            {
                throw new Exception(ex2.Message, last ?? ex2);
            }
        }

        /*────────────────────────────────────────────────────────────*/
        /*─────────────────── MODERN CUSTOM FORM ───────────────────*/
        /*────────────────────────────────────────────────────────────*/

        private class IconSearchForm : Form
        {
            // ── NATIVE INTEROP ──
            [DllImport("user32.dll")] public static extern bool ReleaseCapture();
            [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            // ── THEME COLORS (Matches LogoUtility) ──
            private struct UIColors
            {
                public static readonly SD.Color Surface = SD.Color.FromArgb(255, 255, 255);
                public static readonly SD.Color SurfaceAlt = SD.Color.FromArgb(248, 248, 248);

                // BRAND COLOR: Deep Navy
                public static readonly SD.Color TitleBg = SD.Color.FromArgb(32, 42, 102);
                public static readonly SD.Color TitleText = SD.Color.White;

                public static readonly SD.Color Border = SD.Color.FromArgb(220, 220, 220);
                public static readonly SD.Color TextPrimary = SD.Color.FromArgb(32, 31, 30);
                public static readonly SD.Color TextSecondary = SD.Color.FromArgb(96, 94, 92);

                // BRAND COLOR: Periwinkle Blue
                public static readonly SD.Color Accent = SD.Color.FromArgb(110, 125, 220);
                public static readonly SD.Color AccentHover = SD.Color.FromArgb(85, 95, 180);

                public static readonly SD.Color Danger = SD.Color.FromArgb(232, 17, 35);
            }

            private readonly int MIN_THUMB;
            private readonly int CANVAS;
            private readonly int MARGIN;

            private readonly TextBox _txtSearch;
            private readonly FlowLayoutPanel _flowResults;
            private readonly Label _lblStatus;
            private bool _inReflow = false;

            public IconResult SelectedIcon { get; private set; }

            public IconSearchForm()
            {
                // 1. Form Settings
                this.FormBorderStyle = FormBorderStyle.None;
                this.DoubleBuffered = true;
                this.StartPosition = FormStartPosition.CenterScreen;
                this.BackColor = UIColors.Border;
                this.Padding = new Padding(1);

                // DPI Scaling
                float dpiX;
                using (var g = CreateGraphics()) { dpiX = g.DpiX; }
                float scale = dpiX / 96f;

                // Dimensions
                int w = (int)(780 * scale);
                int h = (int)(550 * scale);
                this.Size = new SD.Size(w, h);

                MIN_THUMB = (int)(130 * scale);
                CANVAS = (int)(160 * scale);
                MARGIN = (int)(10 * scale);

                // 2. Main Surface
                var mainPanel = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = UIColors.Surface
                };
                this.Controls.Add(mainPanel);

                // ── LAYOUT ORDER ──

                // A. Results Grid (Dock=Fill)
                _flowResults = new FlowLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    AutoScroll = true,
                    Padding = new Padding((int)(20 * scale)),
                    BackColor = UIColors.SurfaceAlt
                };
                typeof(Control).GetProperty("DoubleBuffered", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)
                    .SetValue(_flowResults, true, null);

                _flowResults.Resize += (s, e) => ResizeThumbs();
                _flowResults.Layout += (s, e) =>
                {
                    if (_inReflow) return;
                    _inReflow = true;
                    BeginInvoke((Action)(() => { try { ResizeThumbs(); } finally { _inReflow = false; } }));
                };

                mainPanel.Controls.Add(_flowResults);

                // B. Search Area (Dock=Top)
                var pnlSearchArea = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = (int)(70 * scale),
                    Padding = new Padding((int)(20 * scale), (int)(15 * scale), (int)(20 * scale), (int)(15 * scale)),
                    BackColor = UIColors.Surface
                };

                var pnlInputBorder = new Panel
                {
                    Dock = DockStyle.Fill,
                    Padding = new Padding(1),
                    BackColor = UIColors.Border,
                };
                var pnlInputWhite = new Panel
                {
                    Dock = DockStyle.Fill,
                    BackColor = SD.Color.White,
                    Padding = new Padding(10, 0, 5, 0)
                };

                var tblInput = new TableLayoutPanel
                {
                    Dock = DockStyle.Fill,
                    RowCount = 1,
                    ColumnCount = 1,
                    BackColor = SD.Color.Transparent
                };
                tblInput.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
                tblInput.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

                _txtSearch = new TextBox
                {
                    BorderStyle = BorderStyle.None,
                    Font = new SD.Font("Segoe UI", 11f),
                    ForeColor = UIColors.TextPrimary,
                    Dock = DockStyle.Left,
                    Anchor = AnchorStyles.Left | AnchorStyles.Right,
                    Margin = new Padding(0),
                    Width = 500
                };

                tblInput.Controls.Add(_txtSearch, 0, 0);
                pnlInputWhite.Controls.Add(tblInput);
                pnlInputBorder.Controls.Add(pnlInputWhite);

                var btnSearch = new Button
                {
                    Text = "Search",
                    Dock = DockStyle.Right,
                    Width = (int)(100 * scale),
                    FlatStyle = FlatStyle.Flat,
                    // Match Title Bar (Deep Navy)
                    BackColor = UIColors.TitleBg,
                    ForeColor = UIColors.TitleText,
                    Cursor = Cursors.Hand,
                    Font = new SD.Font("Segoe UI Semibold", 11f)
                };
                btnSearch.FlatAppearance.BorderSize = 0;
                btnSearch.FlatAppearance.MouseOverBackColor = UIColors.Accent; // Hover = Periwinkle
                btnSearch.Click += async (s, e) => await SearchAsync();

                var pnlSpacer = new Panel { Dock = DockStyle.Right, Width = (int)(10 * scale) };

                pnlSearchArea.Controls.Add(pnlInputBorder);
                pnlSearchArea.Controls.Add(pnlSpacer);
                pnlSearchArea.Controls.Add(btnSearch);

                mainPanel.Controls.Add(pnlSearchArea);

                // C. Title Bar (Dock=Top)
                var pnlTitle = new Panel
                {
                    Dock = DockStyle.Top,
                    Height = (int)(60 * scale),
                    BackColor = UIColors.TitleBg,
                };
                pnlTitle.MouseDown += (s, e) =>
                {
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

                    using (var font = new SD.Font("Segoe UI", 14f, SD.FontStyle.Bold))
                    using (var brush = new SD.SolidBrush(UIColors.TitleText))
                    {
                        float y = (pnlTitle.Height - font.Height) / 2f;
                        // Left aligned 30px
                        e.Graphics.DrawString("Insert Icon", font, brush, 30, y);
                    }
                };

                var btnClose = new Label
                {
                    Text = "✕",
                    Font = new SD.Font("Segoe UI", 12f),
                    ForeColor = UIColors.TitleText,
                    Dock = DockStyle.Right,
                    Width = (int)(50 * scale),
                    TextAlign = SD.ContentAlignment.MiddleCenter,
                    Cursor = Cursors.Hand,
                    BackColor = UIColors.TitleBg
                };
                btnClose.MouseEnter += (s, e) => { btnClose.BackColor = UIColors.Danger; btnClose.ForeColor = SD.Color.White; };
                btnClose.MouseLeave += (s, e) => { btnClose.BackColor = UIColors.TitleBg; btnClose.ForeColor = UIColors.TitleText; };
                btnClose.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };

                pnlTitle.Controls.Add(btnClose);
                mainPanel.Controls.Add(pnlTitle);

                // D. Status Bar (Dock=Bottom)
                var pnlStatus = new Panel
                {
                    Dock = DockStyle.Bottom,
                    Height = (int)(30 * scale),
                    // White Background
                    BackColor = UIColors.Surface,
                    Padding = new Padding(10, 0, 0, 0)
                };
                pnlStatus.Paint += (s, e) => {
                    using (var p = new SD.Pen(UIColors.Border)) e.Graphics.DrawLine(p, 0, 0, pnlStatus.Width, 0);
                };

                _lblStatus = new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = SD.ContentAlignment.MiddleLeft,
                    Font = new SD.Font("Segoe UI", 9f),
                    // Gray Text
                    ForeColor = UIColors.TextSecondary
                };
                pnlStatus.Controls.Add(_lblStatus);
                mainPanel.Controls.Add(pnlStatus);

                // Events
                this.Load += (s, e) =>
                {
                    _txtSearch.Select();
                    if (IconUtility.Recent.Count > 0) Populate(IconUtility.Recent);
                };

                _txtSearch.KeyDown += async (s, e) => { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; await SearchAsync(); } };
            }

            private async Task SearchAsync()
            {
                string term = _txtSearch.Text.Trim();
                if (string.IsNullOrEmpty(term)) return;

                _lblStatus.Text = "Searching...";
                Cursor = Cursors.WaitCursor;
                _flowResults.Controls.Clear();
                _flowResults.SuspendLayout();

                try
                {
                    var items = await IconUtility.SearchIconsAsync(term);
                    if (items.Count == 0)
                    {
                        _lblStatus.Text = "No icons found.";
                        MessageBox.Show("No icons found for that keyword.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        _lblStatus.Text = $"Found {items.Count} results.";
                        Populate(items);
                    }
                }
                catch (Exception ex)
                {
                    _lblStatus.Text = "Error";
                    MessageBox.Show($"Search failed:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    _flowResults.ResumeLayout();
                    Cursor = Cursors.Default;
                }
            }

            private void Populate(IEnumerable<IconResult> icons)
            {
                _flowResults.SuspendLayout();
                foreach (var ic in icons)
                {
                    var card = new Panel
                    {
                        Size = new SD.Size(CANVAS, CANVAS),
                        Margin = new Padding(MARGIN),
                        BackColor = UIColors.Surface,
                        Cursor = Cursors.Hand,
                        Tag = ic
                    };

                    var pb = new PictureBox
                    {
                        Dock = DockStyle.Fill,
                        SizeMode = PictureBoxSizeMode.CenterImage,
                        BackColor = SD.Color.White,
                    };

                    card.MouseEnter += (s, e) => { card.BackColor = UIColors.SurfaceAlt; };
                    card.MouseLeave += (s, e) => { card.BackColor = UIColors.Surface; };
                    card.Paint += (s, e) =>
                    {
                        using (var p = new SD.Pen(UIColors.Border))
                            e.Graphics.DrawRectangle(p, 0, 0, card.Width - 1, card.Height - 1);
                    };

                    EventHandler onClick = (s, e) =>
                    {
                        SelectedIcon = ic;
                        DialogResult = DialogResult.OK;
                        Close();
                    };

                    card.Click += onClick;
                    pb.Click += onClick;

                    card.Controls.Add(pb);
                    _flowResults.Controls.Add(card);

                    Task.Run(async () =>
                    {
                        try
                        {
                            var bmp = await IconUtility.LoadThumbnailAsync(ic.ThumbnailUrl, ic.Term);
                            if (pb.IsHandleCreated && !pb.IsDisposed)
                                pb.Invoke((Action)(() => pb.Image = bmp));
                        }
                        catch { }
                    });
                }
                _flowResults.ResumeLayout();
                ResizeThumbs();
            }

            private void ResizeThumbs()
            {
                if (_flowResults.Controls.Count == 0) return;
            }

            protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
            {
                if (keyData == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); return true; }
                return base.ProcessCmdKey(ref msg, keyData);
            }
        }
    }
}