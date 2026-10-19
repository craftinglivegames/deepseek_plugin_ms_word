using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAIAssistant
{
    /// <summary>
    /// Рисунки: номера и подписи, подготовка SVG, скачивание изображений.
    /// Чистые функции (FormatFigureNumber, NormalizeSvg и т.п.) покрыты logic_test;
    /// COM-кода здесь нет — он в AgentTools.
    /// </summary>
    public static class ImageOps
    {
        private const string RuAlpha = "АБВГДЕЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯ";
        private const string EnAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

        private static readonly Regex RxSvgTag = new Regex(@"<svg\b[^>]*>", RegexOptions.IgnoreCase);

        // ----------------------------------------------------------- номера и подписи

        /// <summary>Номер подписи в выбранном формате: «число» / «буквы ру» / «буквы англ».</summary>
        public static string FormatFigureNumber(int n, string mode)
        {
            if (n < 1) n = 1;
            string alpha = Alpha(mode);
            if (alpha != null && n <= alpha.Length) return alpha[n - 1].ToString();
            return n.ToString(CultureInfo.InvariantCulture);
        }

        /// <summary>Числовое значение номера рисунка («3»→3, «Б»→2), 0 — не распознано.</summary>
        public static int ParseFigureNumber(string num, string mode)
        {
            if (string.IsNullOrWhiteSpace(num)) return 0;
            num = num.Trim().TrimEnd('.', ')');
            int n;
            if (int.TryParse(num, out n)) return n;
            string alpha = Alpha(mode);
            if (alpha == null) alpha = RuAlpha;
            int idx = alpha.IndexOf(num.ToUpperInvariant(), StringComparison.Ordinal);
            return idx >= 0 ? idx + 1 : 0;
        }

        /// <summary>Следующий свободный номер рисунка (максимальный существующий + 1, минимум 1).</summary>
        public static int NextFigureIndex(IEnumerable<string> numbers, string mode)
        {
            int max = 0;
            if (numbers != null)
                foreach (var s in numbers)
                {
                    int v = ParseFigureNumber(s, mode);
                    if (v > max) max = v;
                }
            return max + 1;
        }

        /// <summary>Итоговая подпись: «Рисунок 1. Название» / «Рисунок 1 Название».</summary>
        public static string BuildFigureCaption(string prefix, string number, bool dot, string title)
        {
            string p = (prefix ?? "").Trim();
            if (p.Length == 0) p = "Рисунок";
            string num = (number ?? "").Trim();
            string t = (title ?? "").Trim();
            var sb = new StringBuilder(p);
            if (num.Length > 0) sb.Append(" ").Append(num);
            if (t.Length > 0)
            {
                sb.Append(dot ? ". " : " ").Append(t);
            }
            else if (dot && num.Length > 0)
            {
                sb.Append(".");
            }
            return sb.ToString();
        }

        /// <summary>Превью формата подписи для поля настройки: «Рисунок 1. Название».</summary>
        public static string FigurePreview(string prefix, string mode, bool dot)
        {
            return BuildFigureCaption(prefix, FormatFigureNumber(1, mode), dot, "Название");
        }

        private static string Alpha(string mode)
        {
            if (mode == null) return null;
            string m = mode.Trim().ToLowerInvariant();
            if (m == "буквы англ" || m == "англ" || m == "en") return EnAlpha;
            if (m == "буквы ру" || m == "рус" || m == "ru") return RuAlpha;
            if (m == "число" || m == "цифры" || m.Length == 0) return null;
            return null;
        }

        // ----------------------------------------------------------- SVG

        /// <summary>
        /// Приводит SVG к виду, пригодному для растеризации: вырезает обёртку,
        /// гарантирует width/height в пикселях (масштабирует к targetWidthPx с сохранением пропорций).
        /// null — если это не SVG.
        /// </summary>
        public static string NormalizeSvg(string svg, int targetWidthPx, out int w, out int h)
        {
            w = 0; h = 0;
            if (string.IsNullOrWhiteSpace(svg)) return null;
            if (targetWidthPx < 16) targetWidthPx = 640;

            string s = svg.Trim();
            int fence = s.IndexOf("```", StringComparison.Ordinal);
            if (fence == 0) s = Regex.Replace(s, @"^```[a-zA-Z]*\s*", "");
            s = Regex.Replace(s, @"\s*```\s*$", "");

            int i = s.IndexOf("<svg", StringComparison.OrdinalIgnoreCase);
            if (i < 0) return null;
            if (i > 0) s = s.Substring(i);
            int j = s.LastIndexOf("</svg>", StringComparison.OrdinalIgnoreCase);
            if (j >= 0) s = s.Substring(0, j + "</svg>".Length);
            if (!s.StartsWith("<svg", StringComparison.OrdinalIgnoreCase)) return null;

            var mTag = RxSvgTag.Match(s);
            if (!mTag.Success) return null;
            string tag = mTag.Value;

            double vw = 0, vh = 0, vbW = 0, vbH = 0;
            var mW = Regex.Match(tag, @"\bwidth\s*=\s*['""]([^'""]*)['""]", RegexOptions.IgnoreCase);
            var mH = Regex.Match(tag, @"\bheight\s*=\s*['""]([^'""]*)['""]", RegexOptions.IgnoreCase);
            var mVB = Regex.Match(tag, @"\bviewBox\s*=\s*['""]\s*([-\d.eE]+)[\s,]+([-\d.eE]+)[\s,]+([-\d.eE]+)[\s,]+([-\d.eE]+)",
                                  RegexOptions.IgnoreCase);
            if (mVB.Success)
            {
                vbW = Num(mVB.Groups[3].Value);
                vbH = Num(mVB.Groups[4].Value);
            }
            if (mW.Success) vw = LenPx(mW.Groups[1].Value);
            if (mH.Success) vh = LenPx(mH.Groups[1].Value);
            if (vw <= 0 && vbW > 0) vw = vbW;
            if (vh <= 0 && vbH > 0) vh = vbH;
            if (vw <= 0) vw = targetWidthPx;
            if (vh <= 0) vh = Math.Round(targetWidthPx * 0.6);

            double scale = targetWidthPx / vw;
            if (scale <= 0 || double.IsInfinity(scale) || double.IsNaN(scale)) scale = 1;
            int nw = Math.Max(16, (int)Math.Round(vw * scale));
            int nh = Math.Max(16, (int)Math.Round(vh * scale));
            w = nw; h = nh;

            string newTag = tag;
            newTag = Regex.Replace(newTag, @"(\bwidth\s*=\s*['""])[^'""]*(['""])",
                                   "${1}" + nw + "px${2}", RegexOptions.IgnoreCase);
            newTag = Regex.Replace(newTag, @"(\bheight\s*=\s*['""])[^'""]*(['""])",
                                   "${1}" + nh + "px${2}", RegexOptions.IgnoreCase);
            if (!Regex.IsMatch(newTag, @"\bwidth\s*=", RegexOptions.IgnoreCase))
                newTag = Regex.Replace(newTag, @"^<svg", "<svg width=\"" + nw + "px\"", RegexOptions.IgnoreCase);
            if (!Regex.IsMatch(newTag, @"\bheight\s*=", RegexOptions.IgnoreCase))
                newTag = Regex.Replace(newTag, @"^<svg", "<svg height=\"" + nh + "px\"", RegexOptions.IgnoreCase);

            return s.Substring(0, mTag.Index) + newTag + s.Substring(mTag.Index + tag.Length);
        }

        private static double Num(string s)
        {
            double d;
            return double.TryParse((s ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d) ? d : 0;
        }

        /// <summary>Длина в пикселях: «640», «640px», «480pt». Проценты/эм — 0 (берём viewBox).</summary>
        private static double LenPx(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return 0;
            string v = raw.Trim().ToLowerInvariant();
            if (v.Contains("%") || v.Contains("em") || v.Contains("ex") || v.Contains("vw")) return 0;
            double d = Num(Regex.Replace(v, @"(px|pt|pica|cm|mm|in|pc)$", "").Trim());
            if (d <= 0) return 0;
            if (v.EndsWith("pt")) d = d * 96.0 / 72.0;
            else if (v.EndsWith("cm")) d = d / 2.54 * 96.0;
            else if (v.EndsWith("mm")) d = d / 25.4 * 96.0;
            else if (v.EndsWith("in")) d = d * 96.0;
            return d;
        }

        // ----------------------------------------------------------- файлы

        /// <summary>%TEMP%\WordAIAssistant (создаётся).</summary>
        public static string TempDir()
        {
            string dir = Path.Combine(Path.GetTempPath(), "WordAIAssistant");
            try { Directory.CreateDirectory(dir); } catch { }
            return dir;
        }

        public static string NewPath(string ext)
        {
            return Path.Combine(TempDir(), "fig_" + Guid.NewGuid().ToString("N").Substring(0, 12) + ext);
        }

        /// <summary>Сохраняет текст в temp-файл нужного расширения; null при ошибке.</summary>
        public static string SaveTemp(string content, string ext)
        {
            try
            {
                string p = NewPath(ext);
                File.WriteAllText(p, content ?? "", new UTF8Encoding(false));
                return p;
            }
            catch (Exception ex) { Logger.Log("SaveTemp: " + ex.Message); return null; }
        }

        /// <summary>Скачивает изображение по прямой ссылке. Возвращает путь или null.</summary>
        public static string DownloadImage(string url, out string ext)
        {
            ext = ".png";
            if (string.IsNullOrWhiteSpace(url)) return null;
            try
            {
                var u = new Uri(url);
                if (u.Scheme != Uri.UriSchemeHttp && u.Scheme != Uri.UriSchemeHttps) return null;

                var req = (HttpWebRequest)WebRequest.Create(u);
                req.Method = "GET";
                req.Timeout = 30000;
                req.ReadWriteTimeout = 30000;
                req.Proxy = null;
                req.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 " +
                                "(KHTML, like Gecko) Chrome/124.0 Safari/537.36";
                req.Referer = u.GetLeftPart(UriPartial.Authority) + "/";
                req.AllowAutoRedirect = true;

                using (var resp = (HttpWebResponse)req.GetResponse())
                {
                    string mime = (resp.ContentType ?? "").Split(';')[0].Trim().ToLowerInvariant();
                    ext = ExtFromMime(mime, u.AbsolutePath);
                    string path = NewPath(ext);
                    using (var input = resp.GetResponseStream())
                    using (var output = File.Create(path))
                    {
                        if (input == null) return null;
                        var buf = new byte[81920];
                        int read, total = 0;
                        while ((read = input.Read(buf, 0, buf.Length)) > 0)
                        {
                            total += read;
                            if (total > 40 * 1024 * 1024) break;
                            output.Write(buf, 0, read);
                        }
                        if (total < 64) { try { File.Delete(path); } catch { } return null; }
                    }
                    return path;
                }
            }
            catch (Exception ex) { Logger.Log("DownloadImage: " + ex.Message); return null; }
        }

        /// <summary>Можно ли вставить этот формат в Word как есть.</summary>
        public static bool WordCompatibleExt(string ext)
        {
            string e = (ext ?? "").ToLowerInvariant();
            return e == ".png" || e == ".jpg" || e == ".jpeg" || e == ".gif" || e == ".bmp"
                || e == ".tif" || e == ".tiff" || e == ".emf" || e == ".wmf";
        }

        public static string ExtFromMime(string mime, string path)
        {
            switch (mime ?? "")
            {
                case "image/jpeg": return ".jpg";
                case "image/jpg": return ".jpg";
                case "image/png": return ".png";
                case "image/gif": return ".gif";
                case "image/bmp": case "image/x-ms-bmp": return ".bmp";
                case "image/tiff": return ".tiff";
                case "image/webp": return ".webp";
                case "image/svg+xml": return ".svg";
                case "image/x-icon": case "image/vnd.microsoft.icon": return ".ico";
            }
            try
            {
                string ext = Path.GetExtension(new Uri(path ?? "http://x/").AbsolutePath).ToLowerInvariant();
                if (ext.Length > 1 && ext.Length < 6) return ext;
            }
            catch { }
            return ".png";
        }

        public static string MimeFromExt(string ext)
        {
            switch ((ext ?? "").ToLowerInvariant())
            {
                case ".jpg": case ".jpeg": return "image/jpeg";
                case ".gif": return "image/gif";
                case ".bmp": return "image/bmp";
                case ".webp": return "image/webp";
                case ".svg": return "image/svg+xml";
                case ".ico": return "image/x-icon";
                default: return "image/png";
            }
        }
    }
}
