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
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using SD = System.Drawing; // Strict alias for System.Drawing
using IOFile = System.IO.File;
using IODir = System.IO.Directory;
using ApiKeys = OfficePragmatics_PowerPoint.Settings.ApiKeys;

namespace OfficePragmatics_PowerPoint.Image
{
    /*──────────────────────── DTO ────────────────────────*/
    public class LogoResult
    {
        public string ImageUrl { get; set; }
        public string ThumbUrl { get; set; }
        public string Term { get; set; }
        public string Mime { get; set; }
        public string Ext { get; set; }
    }

    /*──────────────────────── Exceptions ────────────────────────*/
    internal sealed class HttpStatusException : Exception
    {
        public HttpStatusCode StatusCode { get; }
        public string ResponseBody { get; }
        public HttpStatusException(HttpStatusCode code, string body, string message = null)
            : base(message ?? $"HTTP {(int)code} ({code})")
        {
            StatusCode = code;
            ResponseBody = body ?? string.Empty;
        }
    }

    internal class UnsupportedFormatException : Exception
    {
        public UnsupportedFormatException(string message) : base(message) { }
    }

    /*──────────── MAIN STATIC UTILITY ─────────────*/
    public static class LogoUtility
    {
        // Brave Search image search, called with the user's own key. One request returns
        // all the results we show.
        private const string BRAVE_IMAGE_SEARCH_ENDPOINT = "https://api.search.brave.com/res/v1/images/search";
        private const int SEARCH_RESULTS = 30;
        private const long MAX_IMAGE_BYTES = 25L * 1024 * 1024;

        private const int MAX_RECENT = 24;
        private const int HTTP_TIMEOUT_MS = 15000;

        private static readonly HttpClientHandler _handler = new HttpClientHandler
        {
            SslProtocols = SslProtocols.None,
            AllowAutoRedirect = true,
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };

        /*──── Recent cache ────*/
        private static readonly List<LogoResult> _recent = new List<LogoResult>();
        internal static IReadOnlyList<LogoResult> Recent => _recent;

        private static readonly string _recentsFile = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "OfficePragmatics", "recent-logos.json");

        static LogoUtility() { LoadRecents(); }

        private static void LoadRecents()
        {
            try
            {
                if (IOFile.Exists(_recentsFile))
                {
                    string json = IOFile.ReadAllText(_recentsFile);
                    var list = JsonConvert.DeserializeObject<List<LogoResult>>(json) ?? new List<LogoResult>();
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

        internal static void AddRecent(LogoResult l)
        {
            if (l == null) return;
            _recent.RemoveAll(x => x.ImageUrl == l.ImageUrl);
            _recent.Insert(0, l);
            if (_recent.Count > MAX_RECENT) _recent.RemoveRange(MAX_RECENT, _recent.Count - MAX_RECENT);
            SaveRecents();
        }

        /*──────────── PUBLIC ENTRY ────────────*/
        public static async Task InsertLogo(PowerPoint.Application app)
        {
            try { ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12 | SecurityProtocolType.SystemDefault; } catch { }

            if (!ApiKeys.EnsureLogoSearchKey()) return;

            using (var dlg = new LogoSearchForm())
            {
                if (dlg.ShowDialog() != DialogResult.OK || dlg.SelectedLogo == null) return;

                try
                {
                    await InsertLogoOnSlide(app.ActiveWindow.View.Slide, dlg.SelectedLogo);
                    AddRecent(dlg.SelectedLogo);
                }
                catch (UnsupportedFormatException uf)
                {
                    MessageBox.Show(uf.Message, "Unsupported Image", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                catch (Exception ex)
                {
                    string msg = ex.Message;

                    if (ex is HttpStatusException hse)
                    {
                        msg = ExplainHttpError(hse);
                    }

                    MessageBox.Show($"Download / insert failed:\n{msg}",
                        "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }
        public static Task LogoLookup(PowerPoint.Application app) => InsertLogo(app);

        /*──────── SEARCH – Brave Search images, user's own key ────────*/
        internal static async Task<List<LogoResult>> SearchLogosAsync(string term)
        {
            term = (term ?? string.Empty).Trim();
            var list = new List<LogoResult>();
            if (term.Length == 0) return list;

            string apiKey = ApiKeys.BraveSearchApiKey;
            if (string.IsNullOrEmpty(apiKey))
                throw new InvalidOperationException(
                    "Add your Brave Search API key in Settings > API Keys to search for logos.");

            string url = BRAVE_IMAGE_SEARCH_ENDPOINT +
                         "?q=" + Uri.EscapeDataString(term + " logo") +
                         "&count=" + SEARCH_RESULTS +
                         "&safesearch=strict";

            using (var client = new HttpClient(_handler, disposeHandler: false))
            {
                client.Timeout = TimeSpan.FromMilliseconds(HTTP_TIMEOUT_MS);
                client.DefaultRequestHeaders.UserAgent.ParseAdd("OfficePragmatics/1.0 (+logo-search)");
                client.DefaultRequestHeaders.Accept.ParseAdd("application/json");
                client.DefaultRequestHeaders.TryAddWithoutValidation("X-Subscription-Token", apiKey);

                using (var resp = await client.GetAsync(url))
                {
                    string body = await resp.Content.ReadAsStringAsync();
                    if (!resp.IsSuccessStatusCode)
                        throw new HttpStatusException(resp.StatusCode, body,
                            "Brave search failed: " + DescribeBraveError(resp.StatusCode, body));

                    var arr = JObject.Parse(body)["results"] as JArray ?? new JArray();
                    var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    foreach (var it in arr)
                    {
                        // properties.url is the full-size image; thumbnail.src is Brave's proxied preview.
                        string imageUrl = it["properties"]?["url"]?.ToString();
                        if (string.IsNullOrEmpty(imageUrl) || !seen.Add(imageUrl)) continue;

                        string thumbUrl = it["thumbnail"]?["src"]?.ToString();
                        list.Add(new LogoResult
                        {
                            ImageUrl = imageUrl,
                            ThumbUrl = string.IsNullOrEmpty(thumbUrl) ? imageUrl : thumbUrl,
                            Term = term,
                            Mime = string.Empty,
                            Ext = SafeExtFromUrl(imageUrl)
                        });

                        if (list.Count >= SEARCH_RESULTS) break;
                    }
                }
            }
            return list;
        }

        private static string DescribeBraveError(HttpStatusCode status, string body)
        {
            // Brave errors look like {"type":"ErrorResponse","error":{"code":"...","detail":"..."}}
            string message = null;
            try
            {
                var err = JObject.Parse(body)["error"];
                message = err?["detail"]?.ToString() ?? err?["code"]?.ToString();
            }
            catch { }

            string hint;
            switch ((int)status)
            {
                case 401:
                case 403:
                case 422: hint = "Check your Brave Search API key in Settings > API Keys."; break;
                case 402: hint = "Your Brave Search plan's credit has run out."; break;
                case 429: hint = "Brave Search rate limit or monthly quota reached. Try again later."; break;
                default: hint = null; break;
            }

            var sb = new StringBuilder($"HTTP {(int)status} ({status})");
            if (!string.IsNullOrWhiteSpace(message)) sb.Append("\n").Append(message);
            if (hint != null) sb.Append("\n").Append(hint);
            return sb.ToString();
        }

        private static readonly Regex ImageExtRegex =
            new Regex(@"\.(svg|webp|heic|heif|avif|png|jpe?g|gif|bmp|tiff?|ico)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static string SafeExtFromUrl(string u)
        {
            try
            {
                var m = ImageExtRegex.Match(new Uri(u).AbsolutePath);
                return m.Success ? m.Value.ToLowerInvariant() : string.Empty;
            }
            catch { return string.Empty; }
        }

        /*──────── INSERT LOGIC ────────*/
        private static async Task InsertLogoOnSlide(PowerPoint.Slide slide, LogoResult lr)
        {
            DownloadResult dl;

            try
            {
                dl = await DownloadBytesAsync(lr.ImageUrl);
            }
            catch (HttpStatusException hex) when (
                   (int)hex.StatusCode == 415 || // not an image / unsupported type
                   (int)hex.StatusCode == 502 || // image host failed (e.g., 403, 404, 5xx)
                   (int)hex.StatusCode == 404 || // explicit not found
                   (int)hex.StatusCode == 403)   // explicit forbidden / hotlink protection
            {
                string thumbSrc = string.IsNullOrEmpty(lr.ThumbUrl) ? lr.ImageUrl : lr.ThumbUrl;
                dl = await DownloadBytesAsync(thumbSrc);
            }
            catch (UnsupportedFormatException)
            {
                string thumbSrc = string.IsNullOrEmpty(lr.ThumbUrl) ? lr.ImageUrl : lr.ThumbUrl;
                dl = await DownloadBytesAsync(thumbSrc);
            }

            byte[] bytes = dl.Bytes;
            string contentType = dl.ContentType ?? "";

            EnsureLooksLikeImage(contentType, lr.ImageUrl, bytes);
            var kind = DetectKind(bytes, lr.ImageUrl, contentType);

            if (kind == ImageKind.Webp)
            {
                bytes = TranscodeToPngWithSkia(bytes, "WebP image");
                kind = ImageKind.Png;
            }

            if (kind == ImageKind.Heic || kind == ImageKind.Avif)
                throw new UnsupportedFormatException("This image format (HEIC/AVIF) is not allowed by policy.");

            string tmp = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + (kind == ImageKind.Svg ? ".svg" : ".png"));

            try
            {
                if (kind == ImageKind.Svg)
                {
                    if (!IsSvgBytes(bytes)) throw new UnsupportedFormatException("Invalid SVG data.");
                    IOFile.WriteAllBytes(tmp, bytes);
                }
                else
                {
                    using (var final = PrepareBitmapForInsert(bytes, kind))
                    {
                        bool hasAlpha = HasTransparencyFast(final);
                        using (var trimmed = TrimWhiteBorder(final, hasAlpha))
                        using (var toSave = HasTransparencyFast(trimmed) ? new SD.Bitmap(trimmed) : RemoveWhiteBackground(trimmed))
                        {
                            toSave.Save(tmp, SD.Imaging.ImageFormat.Png);
                        }
                    }
                }

                var pic = slide.Shapes.AddPicture(tmp, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, 100, 100);
                const double BASE_W = 40.0;
                double ratio = pic.Width / Math.Max(1.0, pic.Height);
                pic.LockAspectRatio = Office.MsoTriState.msoTrue;
                pic.Width = (float)(Math.Pow(ratio, 0.525) * BASE_W);

                var pos = FindSpot(slide, pic.Width, pic.Height);
                pic.Left = pos.X; pic.Top = pos.Y;
                pic.AlternativeText = lr.Term;
            }
            finally { TryDelete(tmp); }
        }

        private static void TryDelete(string path) { try { if (IOFile.Exists(path)) IOFile.Delete(path); } catch { } }

        private struct DownloadResult { public byte[] Bytes; public string ContentType; }

        /// <summary>
        /// Results saved by older versions point at the hosted Office Pragmatics proxy
        /// (…officepragmatics.com/proxy?url=…). Recover the original image URL from them.
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

        private static readonly string[] FetchUserAgents =
        {
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Safari/605.1.15",
            "Googlebot-Image/1.0"
        };

        /// <summary>
        /// Downloads an image straight from its host. Many logo hosts block hotlinking, so each
        /// browser profile is tried with a Google referer, the site's own origin as referer, and
        /// no referer, stopping at the first success or at any error other than 401/403.
        /// Failures are reported with the same status codes the callers already handle:
        /// 502 for an image-host failure, 415 for a non-image, 413 for an oversized file.
        /// </summary>
        private static async Task<DownloadResult> DownloadBytesAsync(string url)
        {
            url = UnwrapLegacyProxyUrl(url);

            if (!Uri.TryCreate(url, UriKind.Absolute, out Uri target) ||
                (target.Scheme != Uri.UriSchemeHttps && target.Scheme != Uri.UriSchemeHttp))
            {
                throw new HttpStatusException(HttpStatusCode.BadRequest,
                    ErrorJson("bad_request", "Invalid image URL.", null), "Invalid image URL.");
            }

            using (var client = new HttpClient(_handler, disposeHandler: false))
            {
                client.Timeout = TimeSpan.FromMilliseconds(HTTP_TIMEOUT_MS);

                var attempt = await SendWithHotlinkRetriesAsync(client, target);
                if (attempt.Response == null)
                {
                    string msg = "Could not connect to image host: " + (attempt.LastError?.Message ?? "Connection failed");
                    throw new HttpStatusException(HttpStatusCode.BadGateway,
                        ErrorJson("upstream_unavailable", msg, target.Host), msg);
                }

                using (var resp = attempt.Response)
                {
                    if (!resp.IsSuccessStatusCode)
                    {
                        int st = (int)resp.StatusCode;
                        string detail = st == 403 ? "Access forbidden (hotlink protection?)" :
                                        st == 404 ? "Image not found" :
                                        st == 401 ? "Authentication required" :
                                        st == 429 ? "Rate limited by origin" :
                                        st >= 500 ? "Origin server error" :
                                        $"HTTP {st}";
                        throw new HttpStatusException(HttpStatusCode.BadGateway,
                            ErrorJson("upstream_fetch_failed", detail, target.Host, st), detail);
                    }

                    long? declared = resp.Content.Headers.ContentLength;
                    if (declared.HasValue && declared.Value > MAX_IMAGE_BYTES)
                        throw TooLarge(declared.Value);

                    byte[] bytes = await ReadAtMostAsync(resp.Content, MAX_IMAGE_BYTES);
                    if (bytes.Length == 0)
                        throw new HttpStatusException(HttpStatusCode.BadGateway,
                            ErrorJson("empty_response", "The image host returned no data.", target.Host), "Empty response.");

                    string actualType = (resp.Content.Headers.ContentType?.ToString() ?? string.Empty).ToLowerInvariant();
                    string detected = DetectFetchedImageType(target, actualType, bytes);
                    if (detected == null)
                        throw new HttpStatusException((HttpStatusCode)415,
                            ErrorJson("unsupported_type", "URL does not point to an image", target.Host),
                            "URL does not point to an image.");

                    return new DownloadResult { Bytes = bytes, ContentType = detected };
                }
            }
        }

        private struct FetchAttempt { public HttpResponseMessage Response; public Exception LastError; }

        private static async Task<FetchAttempt> SendWithHotlinkRetriesAsync(HttpClient client, Uri target)
        {
            string[] referers = { "https://www.google.com/", target.GetLeftPart(UriPartial.Authority) + "/", null };
            var result = new FetchAttempt();

            foreach (string ua in FetchUserAgents)
            {
                for (int i = 0; i < referers.Length; i++)
                {
                    var req = new HttpRequestMessage(HttpMethod.Get, target);
                    req.Headers.TryAddWithoutValidation("User-Agent", ua);
                    req.Headers.TryAddWithoutValidation("Accept", "image/webp,image/svg+xml,image/*,*/*;q=0.8");
                    req.Headers.TryAddWithoutValidation("Accept-Language", "en-US,en;q=0.9");
                    req.Headers.TryAddWithoutValidation("Cache-Control", "no-cache");
                    req.Headers.TryAddWithoutValidation("Pragma", "no-cache");
                    req.Headers.TryAddWithoutValidation("Sec-Fetch-Dest", "image");
                    req.Headers.TryAddWithoutValidation("Sec-Fetch-Mode", "no-cors");
                    req.Headers.TryAddWithoutValidation("Sec-Fetch-Site", i == 1 ? "same-origin" : "cross-site");
                    if (referers[i] != null) req.Headers.Referrer = new Uri(referers[i]);

                    try
                    {
                        var resp = await client.SendAsync(req, HttpCompletionOption.ResponseHeadersRead);
                        result.Response?.Dispose();
                        result.Response = resp;

                        // Success, or an error that retrying with other headers will not fix.
                        if (resp.IsSuccessStatusCode ||
                            (resp.StatusCode != HttpStatusCode.Forbidden && resp.StatusCode != HttpStatusCode.Unauthorized))
                            return result;
                    }
                    catch (Exception ex)
                    {
                        result.LastError = ex;
                    }
                }
            }
            return result;
        }

        private static async Task<byte[]> ReadAtMostAsync(HttpContent content, long maxBytes)
        {
            using (var src = await content.ReadAsStreamAsync())
            using (var ms = new MemoryStream())
            {
                var buffer = new byte[81920];
                int n;
                while ((n = await src.ReadAsync(buffer, 0, buffer.Length)) > 0)
                {
                    if (ms.Length + n > maxBytes) throw TooLarge(ms.Length + n);
                    ms.Write(buffer, 0, n);
                }
                return ms.ToArray();
            }
        }

        private static HttpStatusException TooLarge(long size) =>
            new HttpStatusException((HttpStatusCode)413,
                ErrorJson("entity_too_large", $"Image is larger than {MAX_IMAGE_BYTES / (1024 * 1024)} MB.", null),
                $"Image is too large ({size / (1024 * 1024)} MB).");

        /// <summary>Returns the image content type to use, or null if the bytes are not an image.</summary>
        private static string DetectFetchedImageType(Uri target, string actualType, byte[] bytes)
        {
            string path = target.AbsolutePath.ToLowerInvariant();
            bool hasImageExt = ImageExtRegex.IsMatch(path);

            if (path.EndsWith(".webp") || IsWebPBytes(bytes)) return "image/webp";
            if (path.EndsWith(".svg") || LooksLikeSvgContent(bytes)) return "image/svg+xml";
            if (actualType.StartsWith("image/")) return actualType;

            if (hasImageExt && (actualType.Length == 0 || actualType.StartsWith("application/octet-stream")))
            {
                if (path.EndsWith(".png")) return "image/png";
                if (path.EndsWith(".jpg") || path.EndsWith(".jpeg")) return "image/jpeg";
                if (path.EndsWith(".gif")) return "image/gif";
                if (path.EndsWith(".bmp")) return "image/bmp";
                if (path.EndsWith(".ico")) return "image/x-icon";
                if (path.EndsWith(".tif") || path.EndsWith(".tiff")) return "image/tiff";
                return "image/unknown";
            }
            return null;
        }

        private static bool IsWebPBytes(byte[] b) =>
            b != null && b.Length >= 12 &&
            b[0] == 0x52 && b[1] == 0x49 && b[2] == 0x46 && b[3] == 0x46 &&   // RIFF
            b[8] == 0x57 && b[9] == 0x45 && b[10] == 0x42 && b[11] == 0x50;   // WEBP

        private static bool LooksLikeSvgContent(byte[] b)
        {
            if (b == null || b.Length < 50) return false;
            string s = Encoding.UTF8.GetString(b, 0, Math.Min(b.Length, 64 * 1024)).ToLowerInvariant().TrimStart('﻿');
            return s.Contains("<svg") || s.Contains("<!doctype svg") || (s.Contains("<?xml") && s.Contains("svg"));
        }

        /// <summary>Error body in the shape ExplainHttpError already understands.</summary>
        private static string ErrorJson(string error, string message, string host, int? originStatus = null)
        {
            var j = new JObject { ["error"] = error, ["message"] = message };
            if (originStatus.HasValue) j["status"] = originStatus.Value;
            if (!string.IsNullOrEmpty(host)) j["host"] = host;
            return j.ToString(Formatting.None);
        }



        private static byte[] TranscodeToPngWithSkia(byte[] encoded, string label)
        {
            using (var data = SKData.CreateCopy(encoded))
            using (var codec = SKCodec.Create(data))
            {
                if (codec == null) throw new UnsupportedFormatException($"Invalid {label}.");
                using (var image = SKImage.FromEncodedData(data))
                using (var png = image.Encode(SKEncodedImageFormat.Png, 100))
                    return png.ToArray();
            }
        }

        private enum ImageKind { Unknown, Png, Jpeg, Gif, Bmp, Tiff, Ico, Svg, Webp, Heic, Avif }

        private static ImageKind DetectKind(byte[] bytes, string url, string contentType)
        {
            string ext = (Path.GetExtension(url ?? "") ?? "").ToLowerInvariant();
            string ct = (contentType ?? "").ToLowerInvariant();

            if (!string.IsNullOrEmpty(ct))
            {
                if (ct.Contains("svg")) return ImageKind.Svg;
                if (ct.Contains("webp")) return ImageKind.Webp;
                if (ct.Contains("png")) return ImageKind.Png;
                if (ct.Contains("jpeg") || ct.Contains("jpg")) return ImageKind.Jpeg;
            }
            if (!string.IsNullOrEmpty(ext))
            {
                if (ext == ".svg") return ImageKind.Svg;
                if (ext == ".webp") return ImageKind.Webp;
                if (ext == ".png") return ImageKind.Png;
                if (ext == ".jpg" || ext == ".jpeg") return ImageKind.Jpeg;
            }
            if (bytes != null && bytes.Length >= 4)
            {
                if (bytes[0] == 0x89 && bytes[1] == 0x50) return ImageKind.Png;
                if (bytes[0] == 0xFF && bytes[1] == 0xD8) return ImageKind.Jpeg;
                if (bytes[0] == 0x52 && bytes[1] == 0x49 && bytes[8] == 0x57 && bytes[9] == 0x45) return ImageKind.Webp;
            }
            return ImageKind.Unknown;
        }

        private static SD.Bitmap PrepareBitmapForInsert(byte[] bytes, ImageKind kind)
        {
            if (kind == ImageKind.Ico)
            {
                using (var msIcon = new MemoryStream(bytes))
                using (var icon = new SD.Icon(msIcon))
                    return new SD.Bitmap(icon.ToBitmap());
            }

            using (var data = SKData.CreateCopy(bytes))
            using (var probe = SKCodec.Create(data))
            {
                if (probe == null) throw new UnsupportedFormatException("Downloaded bytes are not a decodable image.");
            }

            using (var ms = new MemoryStream(bytes))
            using (var img = SD.Image.FromStream(ms, useEmbeddedColorManagement: true, validateImageData: true))
            {
                var bmp = new SD.Bitmap(img.Width, img.Height, SD.Imaging.PixelFormat.Format32bppArgb);
                using (var g = SD.Graphics.FromImage(bmp))
                {
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.DrawImage(img, 0, 0, img.Width, img.Height);
                }
                return bmp;
            }
        }

        // ─── Image Processing Helpers ───
        private static bool HasTransparencyFast(SD.Bitmap bmp)
        {
            if ((bmp.PixelFormat & SD.Imaging.PixelFormat.Alpha) == 0 &&
                (bmp.PixelFormat & SD.Imaging.PixelFormat.PAlpha) == 0) return false;

            var rect = new SD.Rectangle(0, 0, bmp.Width, bmp.Height);
            var data = bmp.LockBits(rect, SD.Imaging.ImageLockMode.ReadOnly, SD.Imaging.PixelFormat.Format32bppArgb);
            try
            {
                int bytes = Math.Abs(data.Stride) * bmp.Height;
                var buffer = new byte[bytes];
                Marshal.Copy(data.Scan0, buffer, 0, bytes);
                for (int i = 3; i < bytes; i += 4) if (buffer[i] < 255) return true;
                return false;
            }
            finally { bmp.UnlockBits(data); }
        }

        private static SD.Bitmap TrimWhiteBorder(SD.Bitmap bmp, bool hasAlpha, byte tol = 8)
        {
            bool Border(SD.Color c) => hasAlpha ? (c.A <= tol) : (c.R >= 255 - tol && c.G >= 255 - tol && c.B >= 255 - tol);

            int top = 0; while (top < bmp.Height && RowAll(bmp, top, Border)) top++;
            int bottom = bmp.Height - 1; while (bottom >= top && RowAll(bmp, bottom, Border)) bottom--;
            int left = 0; while (left < bmp.Width && ColAll(bmp, left, top, bottom, Border)) left++;
            int right = bmp.Width - 1; while (right >= left && ColAll(bmp, right, top, bottom, Border)) right--;

            if (right < left || bottom < top) return new SD.Bitmap(bmp);
            return bmp.Clone(new SD.Rectangle(left, top, right - left + 1, bottom - top + 1), bmp.PixelFormat);
        }
        private static bool RowAll(SD.Bitmap b, int y, Func<SD.Color, bool> pred) { for (int x = 0; x < b.Width; x++) if (!pred(b.GetPixel(x, y))) return false; return true; }
        private static bool ColAll(SD.Bitmap b, int x, int t, int bt, Func<SD.Color, bool> pred) { for (int y = t; y <= bt; y++) if (!pred(b.GetPixel(x, y))) return false; return true; }

        private static SD.Bitmap RemoveWhiteBackground(SD.Bitmap src, byte tol = 8)
        {
            var dest = new SD.Bitmap(src.Width, src.Height, SD.Imaging.PixelFormat.Format32bppArgb);
            for (int y = 0; y < src.Height; y++)
                for (int x = 0; x < src.Width; x++)
                {
                    SD.Color c = src.GetPixel(x, y);
                    bool isWhite = c.R >= 255 - tol && c.G >= 255 - tol && c.B >= 255 - tol;
                    dest.SetPixel(x, y, isWhite ? SD.Color.Transparent : c);
                }
            return dest;
        }

        // ─── Placement Helpers ───
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
                if (!Collides(s, new SD.RectangleF(x, y, w, h))) return new SD.PointF(x, y);
                if (++row >= maxRows) { row = 0; col++; }
            }
            return new SD.PointF(startX, top);
        }
        private static bool Collides(PowerPoint.Slide s, SD.RectangleF r)
        {
            foreach (PowerPoint.Shape sh in s.Shapes)
                if (r.IntersectsWith(new SD.RectangleF(sh.Left, sh.Top, sh.Width, sh.Height))) return true;
            return false;
        }

        private static string ExplainHttpError(HttpStatusException hse)
        {
            if (string.IsNullOrWhiteSpace(hse.ResponseBody))
                return $"HTTP {(int)hse.StatusCode} ({hse.StatusCode})";

            try
            {
                var j = JObject.Parse(hse.ResponseBody);
                var error = j["error"]?.ToString();
                var status = j["status"]?.ToString();
                var message = j["message"]?.ToString();
                var host = j["host"]?.ToString();

                var sb = new StringBuilder();
                sb.AppendLine($"HTTP {(int)hse.StatusCode} ({hse.StatusCode})");

                if (!string.IsNullOrWhiteSpace(message))
                    sb.AppendLine(message);
                if (!string.IsNullOrWhiteSpace(error))
                {
                    if (!string.IsNullOrWhiteSpace(status))
                        sb.AppendLine($"Code: {error} (origin status {status})");
                    else
                        sb.AppendLine($"Code: {error}");
                }
                if (!string.IsNullOrWhiteSpace(host))
                    sb.AppendLine($"Host: {host}");

                return sb.ToString().Trim();
            }
            catch
            {
                return $"HTTP {(int)hse.StatusCode} ({hse.StatusCode})";
            }
        }

        private static string GetThumbnailCachePath(string url)
        {
            using (var md5 = MD5.Create())
            {
                byte[] hashBytes = md5.ComputeHash(Encoding.UTF8.GetBytes(url));
                string hash = BitConverter.ToString(hashBytes).Replace("-", "").ToLowerInvariant();
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OfficePragmatics", "thumbnails");
                if (!IODir.Exists(dir)) IODir.CreateDirectory(dir);
                return Path.Combine(dir, hash + ".png");
            }
        }

        private static void EnsureLooksLikeImage(string ct, string url, byte[] b)
        {
            if (b == null || b.Length < 32) throw new UnsupportedFormatException("File too small.");
        }

        private static bool IsSvgBytes(byte[] b)
        {
            if (b == null || b.Length == 0) return false;
            string start = Encoding.ASCII.GetString(b, 0, Math.Min(b.Length, 100)).ToLowerInvariant();
            return start.Contains("<svg") || start.Contains("<?xml");
        }

        // ─── THUMBNAIL GENERATION ───
        internal static async Task<SD.Bitmap> LoadThumbnailBitmapInternal(LogoResult lg, int canvas, int margin)
        {
            string thumbSrc = string.IsNullOrEmpty(lg.ThumbUrl) ? lg.ImageUrl : lg.ThumbUrl;
            string cachePath = GetThumbnailCachePath(thumbSrc);

            if (IOFile.Exists(cachePath)) { try { return new SD.Bitmap(cachePath); } catch { } }

            var dl = await DownloadBytesAsync(thumbSrc);
            byte[] b = dl.Bytes;
            var kind = DetectKind(b, UnwrapLegacyProxyUrl(thumbSrc), dl.ContentType);

            if (kind == ImageKind.Webp) { b = TranscodeToPngWithSkia(b, "thumb"); kind = ImageKind.Png; }
            if (kind == ImageKind.Svg || kind == ImageKind.Unknown) return null;

            using (var src = PrepareBitmapForInsert(b, kind))
            {
                var bmp = new SD.Bitmap(canvas, canvas);
                using (var g = SD.Graphics.FromImage(bmp))
                {
                    g.Clear(SD.Color.Transparent);
                    g.SmoothingMode = SmoothingMode.HighQuality;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;

                    float sc = Math.Min((canvas - 2f * margin) / src.Width, (canvas - 2f * margin) / src.Height);
                    int w = (int)(src.Width * sc);
                    int h = (int)(src.Height * sc);
                    g.DrawImage(src, (canvas - w) / 2, (canvas - h) / 2, w, h);
                }
                try { bmp.Save(cachePath, SD.Imaging.ImageFormat.Png); } catch { }
                return bmp;
            }
        }


        /*────────────────────────────────────────────────────────────*/
        /*─────────────────── MODERN CUSTOM FORM ───────────────────*/
        /*────────────────────────────────────────────────────────────*/

        private class LogoSearchForm : Form
        {
            // ── NATIVE INTEROP ──
            [DllImport("user32.dll")] public static extern bool ReleaseCapture();
            [DllImport("user32.dll")] public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
            private const int WM_NCLBUTTONDOWN = 0xA1;
            private const int HT_CAPTION = 0x2;

            // ── THEME COLORS ──
            private struct UIColors
            {
                public static readonly SD.Color Surface = SD.Color.FromArgb(255, 255, 255);
                public static readonly SD.Color SurfaceAlt = SD.Color.FromArgb(248, 248, 248);

                // BRAND COLOR: Deep Navy from "Office Pragmatics" Logo Text
                public static readonly SD.Color TitleBg = SD.Color.FromArgb(32, 42, 102);
                public static readonly SD.Color TitleText = SD.Color.White;

                public static readonly SD.Color Border = SD.Color.FromArgb(220, 220, 220);
                public static readonly SD.Color TextPrimary = SD.Color.FromArgb(32, 31, 30);
                public static readonly SD.Color TextSecondary = SD.Color.FromArgb(96, 94, 92);

                // BRAND COLOR: Periwinkle Blue from "Microsoft Office" Text / "14 Day Trial" Button
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

            public LogoResult SelectedLogo { get; private set; }

            public LogoSearchForm()
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
                    // Button matches the Title Bar (Deep Navy)
                    BackColor = UIColors.TitleBg,
                    ForeColor = UIColors.TitleText,
                    Cursor = Cursors.Hand,
                    // Slightly larger font for better readability
                    Font = new SD.Font("Segoe UI Semibold", 11f)
                };
                btnSearch.FlatAppearance.BorderSize = 0;
                // Hover uses Periwinkle Accent
                btnSearch.FlatAppearance.MouseOverBackColor = UIColors.Accent;
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

                // Custom Paint logic for High Quality Rendering of Title Text
                pnlTitle.Paint += (s, e) =>
                {
                    e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    e.Graphics.SmoothingMode = SmoothingMode.HighQuality;
                    e.Graphics.CompositingQuality = CompositingQuality.HighQuality;

                    using (var font = new SD.Font("Segoe UI", 14f, SD.FontStyle.Bold))
                    using (var brush = new SD.SolidBrush(UIColors.TitleText)) // Use White text
                    {
                        float y = (pnlTitle.Height - font.Height) / 2f;
                        e.Graphics.DrawString("Insert Logo", font, brush, 30, y);
                    }
                };

                var btnClose = new Label
                {
                    Text = "✕",
                    Font = new SD.Font("Segoe UI", 12f),
                    ForeColor = UIColors.TitleText, // Use White Text
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
                    // White surface
                    BackColor = UIColors.Surface,
                    Padding = new Padding(10, 0, 0, 0)
                };
                // Top Border for separation
                pnlStatus.Paint += (s, e) => {
                    using (var p = new SD.Pen(UIColors.Border)) e.Graphics.DrawLine(p, 0, 0, pnlStatus.Width, 0);
                };

                _lblStatus = new Label
                {
                    Dock = DockStyle.Fill,
                    TextAlign = SD.ContentAlignment.MiddleLeft,
                    Font = new SD.Font("Segoe UI", 9f),
                    // Standard gray text
                    ForeColor = UIColors.TextSecondary
                };
                pnlStatus.Controls.Add(_lblStatus);
                mainPanel.Controls.Add(pnlStatus);

                // Events
                this.Load += (s, e) =>
                {
                    _txtSearch.Select();
                    if (LogoUtility.Recent.Count > 0) Populate(LogoUtility.Recent);
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
                    var items = await LogoUtility.SearchLogosAsync(term);
                    if (items.Count == 0)
                    {
                        _lblStatus.Text = "No logos found.";
                        MessageBox.Show("No logos found.", "Search", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

            private void Populate(IEnumerable<LogoResult> logos)
            {
                _flowResults.SuspendLayout();
                foreach (var lg in logos)
                {
                    var card = new Panel
                    {
                        Size = new SD.Size(CANVAS, CANVAS),
                        Margin = new Padding(MARGIN),
                        BackColor = UIColors.Surface,
                        Cursor = Cursors.Hand,
                        Tag = lg
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
                        SelectedLogo = lg;
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
                            var bmp = await LogoUtility.LoadThumbnailBitmapInternal(lg, CANVAS, MARGIN);
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
                // Flow layout handles wrapping automatically
            }

            protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
            {
                if (keyData == Keys.Escape) { DialogResult = DialogResult.Cancel; Close(); return true; }
                return base.ProcessCmdKey(ref msg, keyData);
            }
        }
    }
}