using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    public class Chapter
    {
        public string Title;
        public int Level;      // 2 = '##', 3 = '###'
        public string Text;    // полный текст главы, включая строку заголовка
    }

    public struct Fmt
    {
        public string Font;
        public double Size;
        public double LineSpacing;
        public double FirstLineIndentCm;
        public bool Justify;
        public double ML, MR, MT, MB;

        // --- заголовки по уровням (задаются шаблоном оформления) ---
        public string HeadingFont;   // пусто = как основной шрифт
        public double H1Size, H2Size, H3Size;  // 0 = авто (Size+2 / Size+1 / Size)
        public bool HeadingBold;
        public string HeadingAlign;  // center | left | right | justify — для уровня 1

        public string HeadFont() { return string.IsNullOrEmpty(HeadingFont) ? Font : HeadingFont; }
        public double Head1() { return H1Size > 0 ? H1Size : Size + 2; }
        public double Head2() { return H2Size > 0 ? H2Size : Size + 1; }
        public double Head3() { return H3Size > 0 ? H3Size : Size; }
    }

    public static class WordOps
    {
        // --- Встроенные стили Word ---
        const int wdStyleNormal = -1;
        const int wdStyleHeading1 = -2;
        const int wdStyleHeading2 = -3;
        const int wdStyleHeading3 = -4;
        const int wdStyleHeading4 = -5;
        const int wdLineSpace1pt5 = 1;
        const int wdLineSpaceExactly = 4;
        const int wdLineSpaceMultiple = 5;
        const int wdCollapseEnd = 0;
        const int wdCollapseStart = 1;
        const int wdStatisticPages = 2;
        const int wdAlignParagraphJustify = 3;
        const int wdAlignParagraphLeft = 0;
        const int wdAlignParagraphCenter = 1;
        const int wdAlignParagraphRight = 2;

        public static Fmt FromSettings(Dictionary<string, object> s)
        {
            return new Fmt
            {
                Font = s.ContainsKey("font") ? (string)s["font"] : "Times New Roman",
                Size = ConvD(s, "size", 14),
                LineSpacing = ConvD(s, "lineSpacing", 1.5),
                FirstLineIndentCm = ConvD(s, "firstLineIndentCm", 1.25),
                Justify = s.ContainsKey("justify") && (bool)s["justify"],
                ML = ConvD(s, "marginLeftCm", 3),
                MR = ConvD(s, "marginRightCm", 1.5),
                MT = ConvD(s, "marginTopCm", 2),
                MB = ConvD(s, "marginBottomCm", 2),
                // шаблон оформления: заголовки по уровням (0 = авто, как раньше)
                HeadingFont = s.ContainsKey("headingFont") ? Convert.ToString(s["headingFont"]) ?? "" : "",
                H1Size = ConvD(s, "h1Size", 0),
                H2Size = ConvD(s, "h2Size", 0),
                H3Size = ConvD(s, "h3Size", 0),
                HeadingBold = !s.ContainsKey("headingBold") ||
                              Convert.ToString(s["headingBold"]).ToLowerInvariant() != "false",
                HeadingAlign = s.ContainsKey("headingAlign")
                    ? Convert.ToString(s["headingAlign"]) ?? "" : ""
            };
        }

        /// <summary>Код wdAlignParagraph* для выравнивания уровня 1 (по умолчанию — по центру).</summary>
        public static int Head1Align(Fmt fmt)
        {
            switch ((fmt.HeadingAlign ?? "").Trim().ToLowerInvariant())
            {
                case "left": return wdAlignParagraphLeft;
                case "right": return wdAlignParagraphRight;
                case "justify": return wdAlignParagraphJustify;
                default: return wdAlignParagraphCenter;
            }
        }

        private static double ConvD(Dictionary<string, object> s, string k, double def)
        {
            if (!s.ContainsKey(k)) return def;
            try { return Convert.ToDouble(s[k], System.Globalization.CultureInfo.InvariantCulture); }
            catch { return def; }
        }

        // ---------------------------------------------------------------
        // Контекст документа (для промптов)
        // ---------------------------------------------------------------
        public static string GetDocumentContext(Word.Document doc, int maxChars)
        {
            var heads = new StringBuilder();
            var body = new StringBuilder();
            try
            {
                int count = 0;
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (++count > 1200) break;
                    string t = p.Range.Text.TrimEnd('\r', '\a', '\x07');
                    t = t.Trim();
                    if (t.Length == 0) continue;
                    int lvl = (int)p.OutlineLevel;
                    if (lvl >= 1 && lvl <= 3 && t.Length < 200)
                        heads.AppendLine(new string('#', lvl) + " " + t);
                    if (body.Length < maxChars)
                    {
                        body.AppendLine(t);
                    }
                }
            }
            catch (Exception ex) { Logger.Log("GetDocumentContext: " + ex.Message); }
            var sb = new StringBuilder();
            if (heads.Length > 0)
            {
                sb.AppendLine("[Структура документа:]");
                sb.AppendLine(heads.ToString());
            }
            sb.AppendLine("[Фрагмент текста документа:]");
            sb.AppendLine(body.Length > maxChars ? body.ToString(0, maxChars) : body.ToString());
            return sb.ToString();
        }

        // ---------------------------------------------------------------
        // Markdown -> Word
        // ---------------------------------------------------------------

        public static void InsertMarkdown(Word.Document doc, Word.Range at, string markdown, Fmt fmt, bool applyFormatting)
        {
            var blocks = ParseBlocks(markdown);
            Word.Range r = at.Duplicate;
            foreach (var b in blocks)
            {
                r.Text = b.PlainText + "\r";
                Word.Paragraph para = r.Paragraphs[1];
                Word.Range pr = para.Range;
                try { pr.set_Style(GetStyleId(b.Kind)); } catch { }
                if (applyFormatting) ApplyBaseFormat(pr, fmt, b.Kind);
                ApplyInline(pr, b);
                object e = wdCollapseEnd;
                r.Collapse(ref e);
            }
        }

        private static object GetStyleId(BlockKind k)
        {
            switch (k)
            {
                case BlockKind.H1: return wdStyleHeading1;
                case BlockKind.H2: return wdStyleHeading2;
                case BlockKind.H3: return wdStyleHeading3;
                case BlockKind.H4: return wdStyleHeading4;
                case BlockKind.Bullet: return wdStyleNormal;
                default: return wdStyleNormal;
            }
        }

        public static void ApplyBaseFormat(Word.Range pr, Fmt fmt, BlockKind kind)
        {
            try
            {
                int hBold = fmt.HeadingBold ? 1 : 0;
                pr.Font.Name = fmt.Font;
                if (kind == BlockKind.H1)
                {
                    pr.Font.Name = fmt.HeadFont();
                    pr.Font.Bold = hBold;
                    pr.Font.Size = (float)fmt.Head1();
                    pr.Font.Color = (Word.WdColor)ToRgb(0, 0, 0);
                    pr.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)Head1Align(fmt);
                    pr.ParagraphFormat.FirstLineIndent = 0;
                }
                else if (kind == BlockKind.H2 || kind == BlockKind.H3 || kind == BlockKind.H4)
                {
                    pr.Font.Name = fmt.HeadFont();
                    pr.Font.Bold = hBold;
                    pr.Font.Size = (float)(kind == BlockKind.H2 ? fmt.Head2() : fmt.Head3());
                    pr.Font.Color = (Word.WdColor)ToRgb(0, 0, 0);
                    pr.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)(fmt.Justify ? wdAlignParagraphJustify : wdAlignParagraphLeft);
                    pr.ParagraphFormat.FirstLineIndent = 0;
                }
                else
                {
                    pr.Font.Bold = 0;
                    pr.Font.Size = (float)fmt.Size;
                    pr.Font.Color = (Word.WdColor)ToRgb(0, 0, 0);
                    if (kind == BlockKind.Bullet) pr.ListFormat.ApplyBulletDefault();
                    pr.ParagraphFormat.Alignment = (Word.WdParagraphAlignment)(fmt.Justify && kind == BlockKind.Para ? wdAlignParagraphJustify : wdAlignParagraphLeft);
                }
                // межстрочный интервал
                var pf = pr.ParagraphFormat;
                if (Math.Abs(fmt.LineSpacing - 1.5) < 0.01) pf.LineSpacingRule = (Word.WdLineSpacing)wdLineSpace1pt5;
                else if (fmt.LineSpacing <= 0.1) pf.LineSpacingRule = (Word.WdLineSpacing)wdLineSpaceExactly;
                else pf.LineSpacingRule = (Word.WdLineSpacing)wdLineSpaceMultiple;
                pf.LineSpacing = (float)(fmt.LineSpacing * 12.0);
                if (kind == BlockKind.Para)
                    pf.FirstLineIndent = (float)doc0(fmt.FirstLineIndentCm);
            }
            catch (Exception ex) { Logger.Log("ApplyBaseFormat: " + ex.Message); }
        }

        // CentimetersToPoints требует Application; дублируем формулу: 1 см = 28.3465 pt
        private static double doc0(double cm) { return cm * 28.3465; }

        private static int ToRgb(int r, int g, int b) { return r + (g * 256) + (b * 65536); }

        private static void ApplyInline(Word.Range pr, Block b)
        {
            try
            {
                foreach (var run in b.Runs)
                {
                    if (run.Start >= run.End || run.End > pr.Text.Length) continue;
                    var rr = pr.Document.Range(pr.Start + run.Start, pr.Start + run.End);
                    if (run.Bold) rr.Font.Bold = 1;
                    if (run.Italic) rr.Font.Italic = 1;
                }
            }
            catch (Exception ex) { Logger.Log("ApplyInline: " + ex.Message); }
        }

        public enum BlockKind { Para, H1, H2, H3, H4, Bullet }

        public class Run { public int Start, End; public bool Bold, Italic; }

        public class Block
        {
            public BlockKind Kind;
            public string PlainText;
            public List<Run> Runs = new List<Run>();
        }

        /// <summary>Разбор markdown на блоки: заголовки, списки, абзацы + инлайн **bold** / *italic*.</summary>
        public static List<Block> ParseBlocks(string md)
        {
            var blocks = new List<Block>();
            if (string.IsNullOrEmpty(md)) return blocks;
            md = md.Replace("\r\n", "\n").Replace("\r", "\n");
            foreach (var rawLine in md.Split('\n'))
            {
                string line = rawLine.TrimEnd();
                if (line.Trim().Length == 0) continue;
                var b = new Block();
                string content;
                if (Regex.IsMatch(line, @"^#{1,6}\s"))
                {
                    int h = line.Length - line.TrimStart('#').Length;
                    content = line.Substring(h).TrimStart();
                    b.Kind = h == 1 ? BlockKind.H1 : h == 2 ? BlockKind.H2 : h == 3 ? BlockKind.H3 : BlockKind.H4;
                }
                else if (Regex.IsMatch(line, @"^[-*•]\s+"))
                {
                    b.Kind = BlockKind.Bullet;
                    content = line.Substring(1).TrimStart();
                }
                else
                {
                    b.Kind = BlockKind.Para;
                    content = line;
                }
                b.PlainText = ParseInline(content, b);
                blocks.Add(b);
            }
            return blocks;
        }

        /// <summary>Убирает **...** и *...*, заполняет список runs (позиции в plain-тексте).</summary>
        private static string ParseInline(string s, Block b)
        {
            var plain = new StringBuilder();
            int i = 0;
            while (i < s.Length)
            {
                if (i + 1 < s.Length && s[i] == '*' && s[i + 1] == '*')
                {
                    int close = s.IndexOf("**", i + 2, StringComparison.Ordinal);
                    if (close > 0)
                    {
                        int start = plain.Length;
                        plain.Append(s, i + 2, close - i - 2);
                        b.Runs.Add(new Run { Start = start, End = plain.Length, Bold = true });
                        i = close + 2;
                        continue;
                    }
                }
                if (s[i] == '*')
                {
                    int close = s.IndexOf('*', i + 1);
                    if (close > 0 && close > i + 1)
                    {
                        int start = plain.Length;
                        plain.Append(s, i + 1, close - i - 1);
                        b.Runs.Add(new Run { Start = start, End = plain.Length, Italic = true });
                        i = close + 1;
                        continue;
                    }
                }
                if (s[i] == '`') { i++; continue; }
                plain.Append(s[i]);
                i++;
            }
            return plain.ToString();
        }

        // ---------------------------------------------------------------
        // Оглавление
        // ---------------------------------------------------------------

        public static void AddToc(Word.Document doc)
        {
            try
            {
                var r = doc.Range(0, 0);
                r.InsertBreak(3 /*wdPageBreak*/);
                var r2 = doc.Range(0, 0);
                r2.Text = "ОГЛАВЛЕНИЕ\r";
                r2.set_Style(wdStyleHeading1);
                doc.TablesOfContents.Add(
                    doc.Range(1, 1),
                    true, 1, 3, true, null, true, true, null, true, true, true);
            }
            catch (Exception ex) { Logger.Log("AddToc: " + ex.Message); }
        }

        public static void UpdateToc(Word.Document doc)
        {
            try
            {
                foreach (Word.TableOfContents toc in doc.TablesOfContents)
                    toc.Update();
                if (doc.TablesOfContents.Count == 0)
                    AddToc(doc);
            }
            catch (Exception ex) { Logger.Log("UpdateToc: " + ex.Message); }
        }

        public static void ApplyPageSetup(Word.Document doc, Fmt fmt)
        {
            try
            {
                foreach (Word.Section s in doc.Sections)
                {
                    s.PageSetup.LeftMargin = (float)doc0(fmt.ML);
                    s.PageSetup.RightMargin = (float)doc0(fmt.MR);
                    s.PageSetup.TopMargin = (float)doc0(fmt.MT);
                    s.PageSetup.BottomMargin = (float)doc0(fmt.MB);
                }
            }
            catch (Exception ex) { Logger.Log("ApplyPageSetup: " + ex.Message); }
        }

        public static void ApplyFormatToDoc(Word.Document doc, Fmt fmt)
        {
            ApplyPageSetup(doc, fmt);
            ApplyBaseFormat(doc.Content, fmt, BlockKind.Para);
            // заголовкам вернуть жирность/размер
            try
            {
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    int lvl = (int)p.OutlineLevel;
                    if (lvl >= 1 && lvl <= 4)
                    {
                        var r = p.Range;
                        r.Font.Bold = fmt.HeadingBold ? 1 : 0;
                        r.Font.Color = (Word.WdColor)ToRgb(0, 0, 0);
                        r.Font.Name = fmt.HeadFont();
                        if (lvl == 1) r.Font.Size = (float)fmt.Head1();
                        else if (lvl == 2) r.Font.Size = (float)fmt.Head2();
                        else r.Font.Size = (float)fmt.Head3();
                        p.Range.ParagraphFormat.FirstLineIndent = 0;
                        if (lvl == 1)
                            p.Range.ParagraphFormat.Alignment =
                                (Word.WdParagraphAlignment)Head1Align(fmt);
                    }
                }
            }
            catch (Exception ex) { Logger.Log("ApplyFormatToDoc headings: " + ex.Message); }
        }

        // ---------------------------------------------------------------
        // Анализ оформления текущего документа → шаблон (задача «Шаблоны»)
        // ---------------------------------------------------------------

        private static void IncD(Dictionary<double, int> d, double v)
        {
            double k = Math.Round(v * 4) / 4;   // округление до 0.25
            if (d.ContainsKey(k)) d[k]++; else d[k] = 1;
        }

        private static void IncS(Dictionary<string, int> d, string v)
        {
            v = (v ?? "").Trim();
            if (v.Length == 0) return;
            if (d.ContainsKey(v)) d[v]++; else d[v] = 1;
        }

        /// <summary>Наиболее частое значение (или def, если пусто).</summary>
        private static double ModeD(Dictionary<double, int> d, double def)
        {
            double best = def; int cnt = -1;
            foreach (var kv in d)
                if (kv.Value > cnt) { cnt = kv.Value; best = kv.Key; }
            return cnt < 0 ? def : best;
        }

        private static string ModeS(Dictionary<string, int> d, string def)
        {
            string best = def; int cnt = -1;
            foreach (var kv in d)
                if (kv.Value > cnt) { cnt = kv.Value; best = kv.Key; }
            return cnt < 0 ? def : best;
        }

        /// <summary>Межстрочный интервал в долях строки (1 / 1.5 / 2 …).</summary>
        private static double LineMult(Word.Paragraph p)
        {
            try
            {
                int rule = (int)p.LineSpacingRule;
                if (rule == 0) return 1;          // одинарный
                if (rule == 1) return 1.5;        // 1.5
                if (rule == 2) return 2;          // двойной
                if (rule == 5)                    // множитель
                {
                    double ls = p.LineSpacing;
                    if (ls > 0 && ls < 400) return Math.Round(ls / 12.0 * 4) / 4;
                }
            }
            catch { }
            return 0;   // AtLeast/Exactly — не показатель интервала
        }

        /// <summary>
        /// Анализирует оформление документа: шрифт/размер/интервал/отступ/выравнивание
        /// основного текста, шрифты и размеры заголовков по уровням, поля страницы.
        /// Возвращает DocTemplate с заполненными параметрами (Name/Description пустые — их даёт ИИ).
        /// </summary>
        public static DocTemplate AnalyzeFormatting(Word.Document doc)
        {
            var t = new DocTemplate();
            var bFont = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var bSize = new Dictionary<double, int>();
            var bSp = new Dictionary<double, int>();
            var bInd = new Dictionary<double, int>();
            var bAl = new Dictionary<double, int>();
            int bodyCount = 0;

            var h1Size = new Dictionary<double, int>();
            var h2Size = new Dictionary<double, int>();
            var h3Size = new Dictionary<double, int>();
            var hFont = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var h1Al = new Dictionary<double, int>();
            int hBold = 0, hPlain = 0, h1Count = 0;

            try
            {
                int n = 0;
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (++n > 3000) break;

                    int lvl;
                    try { lvl = (int)p.OutlineLevel; }
                    catch { continue; }

                    string txt = "";
                    try { txt = (p.Range.Text ?? "").Trim(); } catch { }
                    if (txt.Length == 0) continue;

                    string font = "";
                    double size = 0;
                    bool bold = false;
                    try
                    {
                        var f = p.Range.Font;
                        font = Convert.ToString(f.Name) ?? "";
                        size = f.Size;
                        bold = Convert.ToInt32(f.Bold) != 0;
                    }
                    catch { }
                    if (size <= 0 || size >= 200) size = 0;

                    if (lvl >= 1 && lvl <= 4)
                    {
                        IncS(hFont, font);
                        if (size > 0)
                        {
                            if (lvl == 1) IncD(h1Size, size);
                            else if (lvl == 2) IncD(h2Size, size);
                            else IncD(h3Size, size);
                        }
                        if (bold) hBold++; else hPlain++;
                        if (lvl == 1)
                        {
                            h1Count++;
                            try
                            {
                                int al = (int)p.Alignment;
                                if (al >= 0 && al <= 3) IncD(h1Al, al);
                            }
                            catch { }
                        }
                        continue;
                    }
                    if (lvl < 10) continue;   // 5..9 — промежуточные уровни, не основной текст

                    bodyCount++;
                    IncS(bFont, font);
                    if (size > 0) IncD(bSize, size);
                    double m = LineMult(p);
                    if (m > 0) IncD(bSp, m);
                    try
                    {
                        double fi = p.FirstLineIndent / 28.3465;   // pt → см
                        if (fi > -10 && fi < 10) IncD(bInd, Math.Round(fi * 4) / 4);
                    }
                    catch { }
                    try
                    {
                        int al = (int)p.Alignment;
                        if (al >= 0 && al <= 3) IncD(bAl, al);
                    }
                    catch { }
                }
            }
            catch (Exception ex) { Logger.Log("AnalyzeFormatting: " + ex.Message); }

            t.Font = ModeS(bFont, "Times New Roman");
            t.Size = bodyCount > 0 ? ModeD(bSize, 14) : 14;
            t.LineSpacing = bodyCount > 0 ? ModeD(bSp, 1.5) : 1.5;
            t.FirstLineIndentCm = bodyCount > 0 ? ModeD(bInd, 1.25) : 1.25;
            t.Justify = Math.Abs(ModeD(bAl, 3) - 3) < 0.01;

            t.HeadingFont = ModeS(hFont, "");
            t.H1Size = ModeD(h1Size, 0);
            t.H2Size = ModeD(h2Size, 0);
            t.H3Size = ModeD(h3Size, 0);
            t.HeadingBold = hPlain <= hBold;
            {
                // wdAlignParagraph: 0=left, 1=center, 2=right, 3=justify
                int al = (int)ModeD(h1Al, 1);
                t.HeadingAlign = al == 0 ? "left" : al == 2 ? "right" : al == 3 ? "justify" : "center";
            }

            try
            {
                var ps = doc.PageSetup;
                t.ML = Math.Round(ps.LeftMargin / 28.3465, 2);
                t.MR = Math.Round(ps.RightMargin / 28.3465, 2);
                t.MT = Math.Round(ps.TopMargin / 28.3465, 2);
                t.MB = Math.Round(ps.BottomMargin / 28.3465, 2);
            }
            catch (Exception ex) { Logger.Log("AnalyzeFormatting page: " + ex.Message); }

            return t;
        }

        public static string DocStats(Word.Document doc)
        {
            try
            {
                int pages = doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages);
                int words = doc.ComputeStatistics(Word.WdStatistic.wdStatisticWords);
                return pages + " стр., " + words + " слов";
            }
            catch { return ""; }
        }

        // ---------------------------------------------------------------
        // Замена выделения с сохранением исходного форматирования
        // ---------------------------------------------------------------

        /// <summary>Заменяет выделенный фрагмент текстом, сохраняя шрифт/размер/интервалы выделения.</summary>
        public static void ReplacePreservingFormat(Word.Document doc, string plainText)
        {
            var sel = doc.Application.Selection;
            Word.Range r = sel.Range;
            if (r.End <= r.Start)
                throw new InvalidOperationException("Ничего не выделено — выделите фрагмент в документе.");
            ReplaceRangePreservingFormat(r, plainText);
        }

        /// <summary>Заменяет текст произвольного диапазона, сохраняя формат первого символа и абзаца.</summary>
        public static void ReplaceRangePreservingFormat(Word.Range r, string plainText)
        {
            var doc = r.Document;
            // формат берём с первого символа диапазона и с его первого абзаца
            Word.Range firstChar = doc.Range(r.Start, Math.Min(r.Start + 1, r.End));
            string fontName = firstChar.Font.Name;
            float size = firstChar.Font.Size;
            int bold = firstChar.Font.Bold;
            int italic = firstChar.Font.Italic;
            int color = (int)firstChar.Font.Color;
            var pf = r.Paragraphs.Count > 0 ? r.Paragraphs[1].Format : r.ParagraphFormat;
            int spacingRule = (int)pf.LineSpacingRule;
            float lineSpacing = pf.LineSpacing;
            float firstIndent = pf.FirstLineIndent;
            int align = (int)pf.Alignment;
            int spaceAfter = (int)pf.SpaceAfter;

            r.Text = plainText + "\r"; // диапазон покрывает вставленный текст + абзацную отметку
            Word.Range nr = doc.Range(r.Start, r.Start + plainText.Length + 1);
            nr.Font.Name = fontName;
            nr.Font.Size = size;
            nr.Font.Bold = bold;
            nr.Font.Italic = italic;
            nr.Font.Color = (Word.WdColor)color;
            var npf = nr.ParagraphFormat;
            npf.LineSpacingRule = (Word.WdLineSpacing)spacingRule;
            npf.LineSpacing = lineSpacing;
            npf.FirstLineIndent = firstIndent;
            npf.Alignment = (Word.WdParagraphAlignment)align;
            npf.SpaceAfter = spaceAfter;
        }

        /// <summary>Markdown → чистый текст (без #, **, списковой разметки).</summary>
        public static string MdToPlain(string md)
        {
            var sb = new StringBuilder();
            foreach (var b in ParseBlocks(md ?? ""))
                sb.AppendLine(b.PlainText);
            return sb.ToString().TrimEnd();
        }

        // ---------------------------------------------------------------
        // Абзацы документа: список для проверок и выборочной перезаписи ИИ
        // ---------------------------------------------------------------

        public class DocParagraph
        {
            public int Index;     // 1-based номер абзаца в doc.Paragraphs
            public string Text;   // текст без абзацной отметки
        }

        private const int WdUndefined = 9999999;

        /// <summary>
        /// Непустые «обычные» абзацы документа (заголовки пропускаются),
        /// не более maxCount штук. Index хранит реальный номер абзаца в документе.
        /// </summary>
        public static List<DocParagraph> GetRewritableParagraphs(Word.Document doc, int maxCount)
        {
            var res = new List<DocParagraph>();
            int i = 0;
            try
            {
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (++i > 5000) break;
                    string t;
                    try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); }
                    catch { continue; }
                    if (t.Length < 3) continue;

                    bool heading = false;
                    try
                    {
                        int lvl = (int)p.OutlineLevel;
                        heading = lvl >= 1 && lvl <= 3 && t.Length < 200;
                    }
                    catch { }
                    if (heading) continue;

                    res.Add(new DocParagraph { Index = i, Text = t });
                    if (res.Count >= maxCount) break;
                }
            }
            catch (Exception ex) { Logger.Log("GetRewritableParagraphs: " + ex.Message); }
            return res;
        }

        /// <summary>
        /// Заменяет текст одного абзаца (без абзацной отметки), сохраняя его оформление:
        /// шрифт, размер, начертание, цвет, интервалы и выравнивание.
        /// </summary>
        public static void ReplaceParagraphText(Word.Document doc, int paraIndex, string newText)
        {
            if (string.IsNullOrWhiteSpace(newText))
                throw new InvalidOperationException("Пустой текст — замена не выполнена.");
            newText = CollapseToSingleLine(newText);
            if (newText.Length == 0)
                throw new InvalidOperationException("Пустой текст — замена не выполнена.");

            Word.Paragraph p = doc.Paragraphs[paraIndex];
            Word.Range pr = p.Range;
            int start = pr.Start;
            int end = pr.End;
            string cur = pr.Text ?? "";
            // убираем завершающие знаки (абзацный \r и табличный \a) — они остаются в документе
            int tail = cur.Length;
            while (tail > 0 && (cur[tail - 1] == '\r' || cur[tail - 1] == '\a')) tail--;
            end = start + tail;
            if (end < start) end = start;

            // --- снимаем оформление исходного абзаца ---
            string fontName = "Times New Roman";
            float size = 14f, lineSpacing = 18f, firstIndent = 35.4f, spaceAfter = 0f, spaceBefore = 0f;
            int bold = 0, italic = 0, color = 0, spacingRule = 1, align = 3;
            try
            {
                Word.Range firstChar = doc.Range(start, Math.Min(start + 1, end));
                fontName = firstChar.Font.Name;
                size = firstChar.Font.Size;
                bold = (int)firstChar.Font.Bold;
                italic = (int)firstChar.Font.Italic;
                color = (int)firstChar.Font.Color;
            }
            catch { }
            try
            {
                var pf = pr.ParagraphFormat;
                spacingRule = (int)pf.LineSpacingRule;
                lineSpacing = pf.LineSpacing;
                firstIndent = pf.FirstLineIndent;
                align = (int)pf.Alignment;
                spaceAfter = pf.SpaceAfter;
                spaceBefore = pf.SpaceBefore;
            }
            catch { }

            // значения «не определено» (wdUndefined) — приводим к нормальным
            if (float.IsNaN(size) || size <= 0 || size > 200) size = 14f;
            if (bold == WdUndefined) bold = 0;
            if (italic == WdUndefined) italic = 0;
            if (color == WdUndefined || color < 0) color = 0;
            if (spacingRule == WdUndefined || spacingRule < 0 || spacingRule > 5) { spacingRule = 1; lineSpacing = 18f; }
            if (float.IsNaN(lineSpacing) || lineSpacing <= 0 || lineSpacing > 500) lineSpacing = 18f;
            if (align == WdUndefined || align < 0 || align > 3) align = 3;
            if (float.IsNaN(firstIndent) || Math.Abs(firstIndent) > 1000) firstIndent = 35.4f;
            if (float.IsNaN(spaceAfter) || spaceAfter < 0) spaceAfter = 0f;
            if (float.IsNaN(spaceBefore) || spaceBefore < 0) spaceBefore = 0f;

            // --- замена: диапазон абзаца без его отметки ---
            Word.Range target = doc.Range(start, end);
            target.Text = newText;

            Word.Range nr = doc.Range(start, start + newText.Length);
            nr.Font.Name = fontName;
            nr.Font.Size = size;
            nr.Font.Bold = bold;
            nr.Font.Italic = italic;
            nr.Font.Color = (Word.WdColor)color;
            var npf = nr.ParagraphFormat;
            npf.LineSpacingRule = (Word.WdLineSpacing)spacingRule;
            npf.LineSpacing = lineSpacing;
            npf.FirstLineIndent = firstIndent;
            npf.Alignment = (Word.WdParagraphAlignment)align;
            npf.SpaceAfter = spaceAfter;
            npf.SpaceBefore = spaceBefore;
        }

        /// <summary>Схлопывает переводы строк (в т.ч. markdown-переносы) в один абзац.</summary>
        public static string CollapseToSingleLine(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return Regex.Replace(text.Replace("\r\n", "\n").Replace('\r', '\n').Replace('\v', '\n').Replace('\a', '\n'),
                                 @"\s*\n\s*", " ").Trim();
        }

        /// <summary>Первая непустая строка (строка заголовка главы).</summary>
        public static string FirstLine(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            foreach (var ln in text.Replace("\r\n", "\n").Split('\n'))
                if (ln.Trim().Length > 0) return ln.Trim();
            return "";
        }

        /// <summary>Текст главы без первой строки-заголовка.</summary>
        public static string WithoutFirstLine(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            var lines = text.Replace("\r\n", "\n").Split('\n');
            var sb = new StringBuilder();
            bool skipped = false;
            foreach (var ln in lines)
            {
                if (!skipped && ln.Trim().Length > 0) { skipped = true; continue; }
                sb.AppendLine(ln);
            }
            return sb.ToString().Trim();
        }

        // ---------------------------------------------------------------
        // Разбивка сгенерированного текста на главы
        // ---------------------------------------------------------------

        // эвристика заголовков без markdown: «Глава 1», «1. Название», «1.1. Название», известные разделы
        private static readonly Regex RxChapterWord = new Regex(
            @"^\s*(ГЛАВА|Глава|глава|CHAPTER|Chapter)\s+([0-9]+|[IVXLCivxlc]+)\.?\s*(.*)$");
        private static readonly Regex RxNumbered = new Regex(
            @"^\s*(\d+(?:\.\d+)*)[.)]?\s+([А-ЯЁA-Z0-9«""].{2,120})$");
        private static readonly Regex RxNamed = new Regex(
            @"^\s*(Введение|Заключение|Список литературы|Список использованных источников|Список источников|Выводы|Вывод|Приложения|Приложение|Аннотация|Оглавление|Содержание)\s*$",
            RegexOptions.IgnoreCase);

        /// <summary>Похожа ли строка на заголовок (строки без '#').</summary>
        private static bool TryHeadingLine(string line, out string title, out int level)
        {
            title = null; level = 0;
            if (line == null) return false;
            string t = line.Trim();
            if (t.Length == 0 || t.Length > 130) return false;
            if (t.StartsWith("#")) return false;
            var mWord = RxChapterWord.Match(t);
            if (mWord.Success)
            {
                string rest = mWord.Groups[3].Value.Trim();
                title = rest.Length > 0
                    ? "Глава " + mWord.Groups[2].Value.TrimEnd('.') + ". " + rest
                    : "Глава " + mWord.Groups[2].Value.TrimEnd('.');
                level = 2;
                return true;
            }
            var mNum = RxNumbered.Match(t);
            if (mNum.Success)
            {
                string tail = mNum.Groups[2].Value.Trim();
                // отсекаем обычные предложения: заголовок не заканчивается точкой/запятой/двоеточием
                if (tail.EndsWith(".") || tail.EndsWith(",") || tail.EndsWith(";") || tail.EndsWith(":")) return false;
                string num = mNum.Groups[1].Value.TrimEnd('.');
                title = num + ". " + tail;
                level = num.Contains(".") ? 3 : 2;
                return true;
            }
            if (RxNamed.IsMatch(t)) { title = t; level = 2; return true; }
            return false;
        }

        public static List<Chapter> SplitChapters(string md)
        {
            var res = new List<Chapter>();
            if (string.IsNullOrEmpty(md)) return res;
            var lines = md.Replace("\r\n", "\n").Split('\n');
            string curTitle = null;
            int curLevel = 0;
            var cur = new StringBuilder();
            Action flush = () =>
            {
                if (curTitle != null)
                {
                    var t = cur.ToString().Trim();
                    if (t.Length > 0)
                        res.Add(new Chapter { Title = curTitle, Level = curLevel, Text = t });
                }
            };
            bool sawMarkdownHeading = false;
            foreach (var line in lines)
            {
                var m = Regex.Match(line, @"^(#{1,3})\s+(.+)$");
                if (m.Success)
                {
                    sawMarkdownHeading = true;
                    flush();
                    cur.Clear();
                    curTitle = m.Groups[2].Value.Trim();
                    curLevel = m.Groups[1].Value.Length;
                    cur.AppendLine(line);
                }
                else
                {
                    cur.AppendLine(line);
                }
            }
            flush();

            // Фолбэк: модель прислала главы без '#' — режем по эвристикам
            if (!sawMarkdownHeading)
            {
                res.Clear();
                cur.Clear();
                curTitle = null;
                foreach (var line in lines)
                {
                    if (TryHeadingLine(line, out var hTitle, out var hLevel))
                    {
                        flush();
                        cur.Clear();
                        curTitle = hTitle;
                        curLevel = hLevel;
                        cur.AppendLine(line.Trim());
                    }
                    else
                    {
                        cur.AppendLine(line);
                    }
                }
                flush();
            }

            if (res.Count == 0 && !string.IsNullOrWhiteSpace(md))
                res.Add(new Chapter { Title = "Полный текст", Level = 2, Text = md.Trim() });
            return res;
        }
    }
}
