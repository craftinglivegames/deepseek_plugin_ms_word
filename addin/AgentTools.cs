using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Web.Script.Serialization;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    /// <summary>Описание одного инструмента агента (OpenAI function calling schema).</summary>
    public class AgentTool
    {
        public string Name;
        public string Description;
        public object Parameters;                    // JSON Schema (type/properties/required)
        public Func<string, string> Execute;         // argsJson -> текстовый результат для модели
    }

    /// <summary>
    /// Инструменты ИИ-агента: чтение и правка открытого документа Word.
    /// Схемы отдаются мосту в формате OpenAI tools, выполнение — локально через COM.
    /// </summary>
    public static partial class AgentTools
    {
        // ------------------------------------------------------- аргументы

        private static Dictionary<string, object> ParseArgs(string json)
        {
            if (string.IsNullOrWhiteSpace(json)) return new Dictionary<string, object>();
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            try
            {
                var d = ser.Deserialize<Dictionary<string, object>>(json);
                return d ?? new Dictionary<string, object>();
            }
            catch { return new Dictionary<string, object>(); }
        }

        private static string ArgS(Dictionary<string, object> a, string k, string def = "")
        {
            if (!a.ContainsKey(k) || a[k] == null) return def;
            return Convert.ToString(a[k]) ?? def;
        }

        private static int ArgI(Dictionary<string, object> a, string k, int def)
        {
            if (!a.ContainsKey(k) || a[k] == null) return def;
            try { return Convert.ToInt32(Math.Round(Convert.ToDouble(a[k]))); }
            catch { return def; }
        }

        private static bool ArgB(Dictionary<string, object> a, string k, bool def)
        {
            if (!a.ContainsKey(k) || a[k] == null) return def;
            try { return Convert.ToBoolean(a[k]); }
            catch
            {
                string s = ArgS(a, k).ToLowerInvariant();
                if (s == "true" || s == "1" || s == "да") return true;
                if (s == "false" || s == "0" || s == "нет") return false;
                return def;
            }
        }

        private static double? ArgD(Dictionary<string, object> a, string k)
        {
            if (!a.ContainsKey(k) || a[k] == null) return null;
            try { return Convert.ToDouble(a[k]); } catch { return null; }
        }

        // ------------------------------------------------------- схемы

        private static object S(string type, string desc, object extra = null)
        {
            var d = new Dictionary<string, object> { { "type", type }, { "description", desc } };
            if (extra != null)
                foreach (var kv in (Dictionary<string, object>)extra) d[kv.Key] = kv.Value;
            return d;
        }

        private static object Schema(object properties, params string[] required)
        {
            return new Dictionary<string, object>
            {
                { "type", "object" },
                { "properties", properties },
                { "required", required },
                { "additionalProperties", false }
            };
        }

        private static object Props(params object[] kv)
        {
            var d = new Dictionary<string, object>();
            for (int i = 0; i + 1 < kv.Length; i += 2) d[(string)kv[i]] = kv[i + 1];
            return d;
        }

        // ------------------------------------------------------- инструменты

        /// <summary>Полный список инструментов для передачи мосту (OpenAI tools[]).</summary>
        public static string BuildSchemaJson()
        {
            var tools = new List<AgentTool> { Info(), ReadText(), Selection(), InsertText(),
                                              ReplaceSel(), ReplacePara(), FindReplace(),
                                              FormatSel(), PageSetupRead(), PageSetupWrite(), SaveDoc(),
                                              TplAnalyze(), TplCreate(), TplList(), TplGet(), TplSelect(),
                                              InsertSvg(), FindImages(), InsertImage(),
                                              FindSourcesTool(), ReorderSourcesTool() };
            var arr = new List<object>();
            foreach (var t in tools)
            {
                arr.Add(new Dictionary<string, object>
                {
                    { "type", "function" },
                    { "function", new Dictionary<string, object>
                        { { "name", t.Name }, { "description", t.Description }, { "parameters", t.Parameters } } }
                });
            }
            return new JavaScriptSerializer().Serialize(arr);
        }

        /// <summary>Выполняет инструмент. Ошибки возвращаются строкой (не исключением).</summary>
        public static string Execute(string name, string argsJson, Word.Application app)
        {
            if (app != null) WordApp = app;
            var tools = new Dictionary<string, Func<string, string>>(StringComparer.OrdinalIgnoreCase)
            {
                { "get_document_info", _ => Info().Execute(null) },
                { "get_document_text", a => ReadText().Execute(a) },
                { "get_selection", _ => Selection().Execute(null) },
                { "insert_text", a => InsertText().Execute(a) },
                { "replace_selection", a => ReplaceSel().Execute(a) },
                { "replace_paragraph", a => ReplacePara().Execute(a) },
                { "find_and_replace", a => FindReplace().Execute(a) },
                { "format_selection", a => FormatSel().Execute(a) },
                { "get_page_setup", _ => PageSetupRead().Execute(null) },
                { "set_page_setup", a => PageSetupWrite().Execute(a) },
                { "save_document", _ => SaveDoc().Execute(null) },
                { "analyze_document_format", _ => TplAnalyze().Execute(null) },
                { "create_template", a => TplCreate().Execute(a) },
                { "list_templates", _ => TplList().Execute(null) },
                { "get_template", a => TplGet().Execute(a) },
                { "select_template", a => TplSelect().Execute(a) },
                { "insert_svg_image", a => AllowSvgImages ? InsertSvg().Execute(a) : OffMsg },
                { "find_images", a => AllowWebImages ? FindImages().Execute(a) : OffMsg },
                { "insert_image", a => AllowWebImages ? InsertImage().Execute(a) : OffMsg },
                { "find_sources", a => FindSourcesTool().Execute(a) },
                { "reorder_sources", a => ReorderSourcesTool().Execute(a) },
            };
            Func<string, string> fn;
            if (!tools.TryGetValue(name ?? "", out fn))
                return "ОШИБКА: неизвестный инструмент «" + name + "».";
            try
            {
                return fn(argsJson) ?? "";
            }
            catch (Exception ex)
            {
                Logger.Log("AgentTool " + name + ": " + ex);
                return "ОШИБКА: " + ex.Message;
            }
        }

        // ------------------------------------------------------- реализации

        private static Word.Document Doc(Word.Application app)
        {
            if (app == null)
                throw new InvalidOperationException("Word ещё не запущен — нет открытого документа.");
            var doc = app.ActiveDocument;
            if (doc == null) throw new InvalidOperationException("Нет открытого документа.");
            return doc;
        }

        private static AgentTool Info()
        {
            return new AgentTool
            {
                Name = "get_document_info",
                Description = "Статистика открытого документа Word: имя, число страниц, слов, абзацев " +
                              "и его структура (заголовки). Вызывать первой, чтобы понять документ.",
                Parameters = Schema(Props()),
                Execute = _ =>
                {
                    var doc = Doc(WordApp);
                    var sb = new StringBuilder();
                    sb.AppendLine("Документ: " + doc.Name);
                    try { sb.AppendLine("Страниц: " + doc.ComputeStatistics(Word.WdStatistic.wdStatisticPages)); } catch { }
                    try { sb.AppendLine("Слов: " + doc.ComputeStatistics(Word.WdStatistic.wdStatisticWords)); } catch { }
                    try { sb.AppendLine("Абзацев: " + doc.Paragraphs.Count); } catch { }
                    sb.AppendLine();
                    sb.AppendLine("Структура (заголовки):");
                    int n = 0;
                    foreach (Word.Paragraph p in doc.Paragraphs)
                    {
                        if (++n > 500) break;
                        string t;
                        try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); } catch { continue; }
                        if (t.Length == 0 || t.Length > 200) continue;
                        int lvl;
                        try { lvl = (int)p.OutlineLevel; } catch { continue; }
                        if (lvl >= 1 && lvl <= 3)
                            sb.AppendLine(new string('#', lvl) + " " + t);
                    }
                    return sb.ToString();
                }
            };
        }

        private static AgentTool ReadText()
        {
            return new AgentTool
            {
                Name = "get_document_text",
                Description = "Текст документа с номерами абзацев. Номера совпадают с инструментом " +
                              "replace_paragraph. Параметр numbered=false вернёт чистый текст без номеров.",
                Parameters = Schema(Props(
                    "numbered", S("boolean", "Нумеровать абзацы [N] (по умолчанию true)"),
                    "max_chars", S("integer", "Максимум символов в ответе (по умолчанию 20000, максимум 60000)"),
                    "from_paragraph", S("integer", "Начать с абзаца № N (по умолчанию с первого)"),
                    "to_paragraph", S("integer", "Остановиться после абзаца № N (не включительно — по умолчанию до конца)"))),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);
                    bool numbered = ArgB(d, "numbered", true);
                    int maxChars = Math.Min(Math.Max(ArgI(d, "max_chars", 20000), 500), 60000);
                    int from = Math.Max(ArgI(d, "from_paragraph", 1), 1);
                    int to = ArgI(d, "to_paragraph", int.MaxValue);
                    var sb = new StringBuilder();
                    int i = 0, shown = 0;
                    foreach (Word.Paragraph p in doc.Paragraphs)
                    {
                        if (++i > 5000 || i < from || i >= to) { if (i >= to) break; continue; }
                        string t;
                        try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); } catch { continue; }
                        if (t.Length == 0) continue;
                        string line = numbered ? "[" + i + "] " + t : t;
                        if (sb.Length + line.Length + 2 > maxChars)
                        {
                            sb.AppendLine();
                            sb.AppendLine("…(достигнут лимит " + maxChars + " символов, продолжите с абзаца " + (i + 1) + ")");
                            break;
                        }
                        sb.AppendLine(line);
                        shown++;
                    }
                    if (shown == 0) return "Документ пуст (или указанный диапазон абзацев пуст).";
                    return sb.ToString();
                }
            };
        }

        private static AgentTool Selection()
        {
            return new AgentTool
            {
                Name = "get_selection",
                Description = "Возвращает выделенный в документе фрагмент текста. Если ничего не выделено — сообщает об этом.",
                Parameters = Schema(Props()),
                Execute = _ =>
                {
                    var doc = Doc(WordApp);
                    string text = null;
                    try { text = doc.Application.Selection.Text; } catch { }
                    if (string.IsNullOrEmpty(text))
                        return "Ничего не выделено в документе.";
                    text = text.TrimEnd('\r', '\v', '\a', '\f');
                    if (text.Length > 20000) text = text.Substring(0, 20000) + "\n…(обрезано)";
                    return text;
                }
            };
        }

        private static AgentTool InsertText()
        {
            return new AgentTool
            {
                Name = "insert_text",
                Description = "Вставляет текст в текущую позицию курсора (или вместо выделения). " +
                              "Обычный текст вставляется как есть; markdown=true преобразует заголовки #, " +
                              "жирный/курсив и списки в оформление Word.",
                Parameters = Schema(Props(
                    "text", S("string", "Текст для вставки"),
                    "markdown", S("boolean", "Преобразовать markdown-разметку в оформление Word (по умолчанию false)")),
                    "text"),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);
                    string text = ArgS(d, "text");
                    if (string.IsNullOrEmpty(text)) return "ОШИБКА: пустой текст.";
                    if (ArgB(d, "markdown", false))
                    {
                        var sel = doc.Application.Selection;
                        WordOps.InsertMarkdown(doc, sel.Range, text, WordOps.FromSettings(Store.LoadSettings()), true);
                        return "Вставлено (markdown): " + text.Length + " символов.";
                    }
                    doc.Application.Selection.TypeText(text);
                    return "Вставлено: " + text.Length + " символов.";
                }
            };
        }

        private static AgentTool ReplaceSel()
        {
            return new AgentTool
            {
                Name = "replace_selection",
                Description = "Заменяет выделенный фрагмент указанным текстом, сохраняя исходное оформление " +
                              "(шрифт, размер, начертание, интервалы). Ничего не выделено — вернёт ошибку.",
                Parameters = Schema(Props(
                    "text", S("string", "Новый текст взамен выделенного")),
                    "text"),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    string text = ArgS(ParseArgs(a), "text");
                    if (string.IsNullOrEmpty(text)) return "ОШИБКА: пустой текст.";
                    WordOps.ReplacePreservingFormat(doc, text);
                    return "Выделение заменено (" + text.Length + " символов), оформление сохранено.";
                }
            };
        }

        private static AgentTool ReplacePara()
        {
            return new AgentTool
            {
                Name = "replace_paragraph",
                Description = "Заменяет текст одного абзаца (номер — из get_document_text), сохраняя оформление " +
                              "этого абзаца. Пустой текст запрещён.",
                Parameters = Schema(Props(
                    "index", S("integer", "Номер абзаца в документе (как в get_document_text)"),
                    "text", S("string", "Новый текст абзаца")),
                    "index", "text"),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);
                    int index = ArgI(d, "index", -1);
                    string text = ArgS(d, "text");
                    if (index < 1) return "ОШИБКА: index должен быть >= 1.";
                    if (string.IsNullOrWhiteSpace(text)) return "ОШИБКА: пустой текст.";
                    WordOps.ReplaceParagraphText(doc, index, text);
                    return "Абзац " + index + " заменён (" + text.Length + " символов).";
                }
            };
        }

        private static AgentTool FindReplace()
        {
            return new AgentTool
            {
                Name = "find_and_replace",
                Description = "Находит все вхождения фразы в документе и заменяет их (с учётом регистра, если указано).",
                Parameters = Schema(Props(
                    "find", S("string", "Что искать"),
                    "replace", S("string", "Чем заменить"),
                    "match_case", S("boolean", "Учитывать регистр (по умолчанию false)")),
                    "find", "replace"),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);
                    string find = ArgS(d, "find");
                    string repl = ArgS(d, "replace");
                    if (string.IsNullOrEmpty(find)) return "ОШИБКА: пустой поисковый запрос.";
                    bool matchCase = ArgB(d, "match_case", false);
                    int count = 0;
                    var r = doc.Content;
                    r.Find.ClearFormatting();
                    r.Find.Forward = true;
                    r.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    r.Find.Format = false;
                    r.Find.MatchCase = matchCase;
                    while (count < 500)
                    {
                        r.Find.Text = find;
                        if (!r.Find.Execute()) break;
                        int s = r.Start;
                        r.Text = repl;
                        r.SetRange(s + repl.Length, doc.Content.End);
                        r.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        r.Find.MatchCase = matchCase;
                        count++;
                    }
                    if (count == 0) return "Вхождений «" + find + "» не найдено.";
                    return "Заменено вхождений: " + count + ".";
                }
            };
        }

        private static AgentTool FormatSel()
        {
            return new AgentTool
            {
                Name = "format_selection",
                Description = "Форматирует текст: шрифт, размер, жирный/курсив, выравнивание. " +
                              "Цель задаётся параметром paragraph (номер абзаца) или paragraphs (несколько номеров) — " +
                              "это основной способ, когда в документе ничего не выделено. Номера абзацев бери из " +
                              "get_document_text с numbered=true. Если номера не указаны — форматируется текущее выделение.",
                Parameters = Schema(Props(
                    "paragraph", S("number", "Номер абзаца (с 1) — применить форматирование ко всему абзацу"),
                    "paragraphs", S("array", "Несколько номеров абзацев — применить ко всем сразу",
                        new Dictionary<string, object> { { "items", new Dictionary<string, object> { { "type", "number" } } } }),
                    "font_name", S("string", "Имя шрифта, например «Times New Roman»"),
                    "font_size", S("number", "Размер шрифта в пунктах, например 14"),
                    "bold", S("boolean", "Жирный: true/false"),
                    "italic", S("boolean", "Курсив: true/false"),
                    "align", S("string", "Выравнивание абзаца: left | center | right | justify"))),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);

                    // номера абзацев (приоритет над выделением)
                    var nums = new List<int>();
                    if (d.ContainsKey("paragraphs"))
                    {
                        var arr = d["paragraphs"] as System.Collections.IEnumerable;
                        if (arr != null)
                            foreach (object o in arr)
                            {
                                try { nums.Add(Convert.ToInt32(Math.Round(Convert.ToDouble(o)))); }
                                catch { }
                            }
                    }
                    int one = ArgI(d, "paragraph", 0);
                    if (one > 0) nums.Add(one);
                    nums.Sort();

                    var ranges = new List<Word.Range>();
                    string what;
                    if (nums.Count > 0)
                    {
                        var uniq = new List<int>();
                        foreach (int n in nums) if (!uniq.Contains(n)) uniq.Add(n);
                        foreach (int n in uniq)
                        {
                            if (n < 1 || n > doc.Paragraphs.Count)
                                return "ОШИБКА: абзаца №" + n + " нет — в документе " + doc.Paragraphs.Count + " абзацев.";
                            ranges.Add(doc.Paragraphs[n].Range);
                        }
                        what = "к абзацам " + string.Join(", ", Array.ConvertAll(uniq.ToArray(), x => x.ToString()));
                    }
                    else
                    {
                        var sel = doc.Application.Selection;
                        if (sel.Start == sel.End)
                            return "ОШИБКА: ничего не выделено. Укажите параметр paragraph (номер абзаца, " +
                                   "номера бывают в get_document_text с numbered=true) либо выделите фрагмент в документе.";
                        ranges.Add(sel.Range);
                        what = "к выделению";
                    }

                    foreach (var r in ranges)
                    {
                        string font = ArgS(d, "font_name");
                        if (font.Length > 0) r.Font.Name = font;
                        double? size = ArgD(d, "font_size");
                        if (size.HasValue) r.Font.Size = (float)size.Value;
                        if (d.ContainsKey("bold")) r.Font.Bold = ArgB(d, "bold", false) ? -1 : 0;
                        if (d.ContainsKey("italic")) r.Font.Italic = ArgB(d, "italic", false) ? -1 : 0;
                        string align = ArgS(d, "align").ToLowerInvariant();
                        switch (align)
                        {
                            case "left": r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft; break;
                            case "center": r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter; break;
                            case "right": r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight; break;
                            case "justify": r.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify; break;
                        }
                    }
                    return "Форматирование применено " + what + ".";
                }
            };
        }

        private static double PtToCm(double pt) { return Math.Round(pt / 28.346456693, 2); }

        private static string Fmt(double v) { return v.ToString("0.##", System.Globalization.CultureInfo.InvariantCulture); }

        /// <summary>Человекочитаемое описание PageSetup (ориентация, размер, поля).</summary>
        private static string Describe(Word.PageSetup ps)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Ориентация: " + (ps.Orientation == Word.WdOrientation.wdOrientLandscape
                          ? "альбомная (landscape)" : "портретная (portrait)"));
            string paper = PaperName(ps);
            if (paper != null) sb.AppendLine("Формат бумаги: " + paper);
            sb.AppendLine("Размер страницы: " + Fmt(PtToCm(ps.PageWidth)) + " × " + Fmt(PtToCm(ps.PageHeight)) + " см");
            sb.AppendLine("Поля, см: верх " + Fmt(PtToCm(ps.TopMargin)) + ", низ " + Fmt(PtToCm(ps.BottomMargin)) +
                          ", слева " + Fmt(PtToCm(ps.LeftMargin)) + ", справа " + Fmt(PtToCm(ps.RightMargin)));
            try { sb.AppendLine("Колонок: " + ps.TextColumns.Count); } catch { }
            return sb.ToString();
        }

        private static string PaperName(Word.PageSetup ps)
        {
            try
            {
                var s = ps.PaperSize;
                if (s == Word.WdPaperSize.wdPaperA4) return "A4";
                if (s == Word.WdPaperSize.wdPaperA3) return "A3";
                if (s == Word.WdPaperSize.wdPaperA5) return "A5";
                if (s == Word.WdPaperSize.wdPaperB4) return "B4";
                if (s == Word.WdPaperSize.wdPaperB5) return "B5";
                if (s == Word.WdPaperSize.wdPaperLetter) return "Letter";
                if (s == Word.WdPaperSize.wdPaperLegal) return "Legal";
                if (s == Word.WdPaperSize.wdPaperExecutive) return "Executive";
            }
            catch { }
            return null; // пользовательский размер
        }

        private static AgentTool PageSetupRead()
        {
            return new AgentTool
            {
                Name = "get_page_setup",
                Description = "Параметры страницы документа: ориентация (портретная/альбомная), размер бумаги, " +
                              "поля в сантиметрах, колонки, число разделов. Вызывать перед изменением параметров страницы.",
                Parameters = Schema(Props()),
                Execute = _ =>
                {
                    var doc = Doc(WordApp);
                    var sb = new StringBuilder(Describe(doc.PageSetup));
                    sb.AppendLine("Разделов в документе: " + doc.Sections.Count);
                    int cur = 1;
                    try { cur = doc.Application.Selection.Sections[1].Index; } catch { }
                    sb.AppendLine("Текущий раздел (курсор): " + cur);
                    return sb.ToString();
                }
            };
        }

        private static AgentTool PageSetupWrite()
        {
            return new AgentTool
            {
                Name = "set_page_setup",
                Description = "Меняет параметры страницы Word: ориентацию (портретная/альбомная), формат бумаги " +
                              "(A4, A3, A5, Letter…), поля в сантиметрах. По умолчанию применяется ко всему документу, " +
                              "с apply_to=\"section\" — только к одному разделу. Можно передать только нужные параметры.",
                Parameters = Schema(Props(
                    "orientation", S("string", "Ориентация: portrait (портретная) | landscape (альбомная/горизонтальная)"),
                    "paper_size", S("string", "Формат бумаги: A4 | A3 | A5 | B4 | B5 | Letter | Legal | Executive"),
                    "width_cm", S("number", "Ширина страницы в сантиметрах (свой размер; указывать вместе с height_cm)"),
                    "height_cm", S("number", "Высота страницы в сантиметрах (свой размер; указывать вместе с width_cm)"),
                    "margin_top_cm", S("number", "Верхнее поле, см"),
                    "margin_bottom_cm", S("number", "Нижнее поле, см"),
                    "margin_left_cm", S("number", "Левое поле, см (в Word левое поле включает подшивку)"),
                    "margin_right_cm", S("number", "Правое поле, см"),
                    "apply_to", S("string", "Где применить: document (весь документ, по умолчанию) | section (один раздел)"),
                    "section", S("integer", "Номер раздела при apply_to=section (по умолчанию — текущий раздел курсора)"))),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);

                    string orientation = ArgS(d, "orientation").Trim().ToLowerInvariant();
                    string paper = ArgS(d, "paper_size").Trim();
                    double? w = ArgD(d, "width_cm"), h = ArgD(d, "height_cm");
                    double? mt = ArgD(d, "margin_top_cm"), mb = ArgD(d, "margin_bottom_cm");
                    double? ml = ArgD(d, "margin_left_cm"), mr = ArgD(d, "margin_right_cm");

                    if (orientation.Length == 0 && paper.Length == 0 && !w.HasValue && !h.HasValue &&
                        !mt.HasValue && !mb.HasValue && !ml.HasValue && !mr.HasValue)
                        return "ОШИБКА: не указано ни одного параметра страницы (ориентация, формат бумаги, поля).";

                    // область применения
                    string applyTo = ArgS(d, "apply_to", "document").Trim().ToLowerInvariant();
                    Word.PageSetup ps;
                    string scope;
                    if (applyTo == "section")
                    {
                        int sec = ArgI(d, "section", 0);
                        if (sec == 0)
                        {
                            try { sec = doc.Application.Selection.Sections[1].Index; } catch { sec = 1; }
                        }
                        int total = doc.Sections.Count;
                        if (sec < 1 || sec > total)
                            return "ОШИБКА: раздел " + sec + " не найден (в документе разделов: " + total + ").";
                        ps = doc.Sections[sec].PageSetup;
                        scope = "раздел " + sec;
                    }
                    else
                    {
                        ps = doc.PageSetup;
                        scope = "весь документ";
                    }

                    // формат бумаги / свой размер
                    if (w.HasValue != h.HasValue)
                        return "ОШИБКА: width_cm и height_cm указываются вместе.";
                    if (paper.Length > 0)
                    {
                        string p = paper.Replace(" ", "").ToUpperInvariant();
                        if (p == "A4") ps.PaperSize = Word.WdPaperSize.wdPaperA4;
                        else if (p == "A3") ps.PaperSize = Word.WdPaperSize.wdPaperA3;
                        else if (p == "A5") ps.PaperSize = Word.WdPaperSize.wdPaperA5;
                        else if (p == "B4") ps.PaperSize = Word.WdPaperSize.wdPaperB4;
                        else if (p == "B5") ps.PaperSize = Word.WdPaperSize.wdPaperB5;
                        else if (p == "LETTER") ps.PaperSize = Word.WdPaperSize.wdPaperLetter;
                        else if (p == "LEGAL") ps.PaperSize = Word.WdPaperSize.wdPaperLegal;
                        else if (p == "EXECUTIVE") ps.PaperSize = Word.WdPaperSize.wdPaperExecutive;
                        else return "ОШИБКА: неизвестный формат бумаги «" + paper +
                                    "». Доступны: A4, A3, A5, B4, B5, Letter, Legal, Executive.";
                    }
                    if (w.HasValue && h.HasValue)
                    {
                        if (w.Value < 5 || h.Value < 5 || w.Value > 200 || h.Value > 200)
                            return "ОШИБКА: недопустимый размер страницы (допустимо 5…200 см).";
                        ps.PageWidth = (float)(w.Value * 28.346456693);
                        ps.PageHeight = (float)(h.Value * 28.346456693);
                    }

                    // ориентация
                    if (orientation.Length > 0)
                    {
                        if (orientation == "landscape" || orientation == "альбомная" || orientation == "горизонтальная")
                            ps.Orientation = Word.WdOrientation.wdOrientLandscape;
                        else if (orientation == "portrait" || orientation == "портретная" || orientation == "вертикальная")
                            ps.Orientation = Word.WdOrientation.wdOrientPortrait;
                        else return "ОШИБКА: orientation должен быть portrait или landscape (получено: «" + orientation + "»).";
                    }

                    // поля
                    string[] keys = { "margin_top_cm", "margin_bottom_cm", "margin_left_cm", "margin_right_cm" };
                    double?[] vals = { mt, mb, ml, mr };
                    for (int i = 0; i < keys.Length; i++)
                    {
                        if (!vals[i].HasValue) continue;
                        double v = vals[i].Value;
                        if (v < 0 || v > 30)
                            return "ОШИБКА: " + keys[i] + " вне диапазона 0…30 см.";
                        float pt = (float)(v * 28.346456693);
                        if (i == 0) ps.TopMargin = pt;
                        else if (i == 1) ps.BottomMargin = pt;
                        else if (i == 2) ps.LeftMargin = pt;
                        else ps.RightMargin = pt;
                    }

                    Logger.Log("AgentTool set_page_setup: " + scope +
                               ", orientation=" + orientation + ", paper=" + paper);
                    return "Параметры страницы применены (" + scope + ").\n" + Describe(ps);
                }
            };
        }

        private static AgentTool SaveDoc()
        {
            return new AgentTool
            {
                Name = "save_document",
                Description = "Сохраняет документ на диск (Ctrl+S). Если документ ещё не сохранялся — просит сохранить вручную.",
                Parameters = Schema(Props()),
                Execute = _ =>
                {
                    var doc = Doc(WordApp);
                    if (string.IsNullOrEmpty(doc.FullName))
                        return "Документ ещё не сохранён на диске — сохраните его вручную (Ctrl+S), затем повторите вызов.";
                    doc.Save();
                    return "Сохранено: " + doc.FullName;
                }
            };
        }

        // ------------------------------------- шаблоны оформления документа

        /// <summary>Ключи параметров оформления, общие для настроек и шаблона.</summary>
        private static string[] TplKeys()
        {
            return new[]
            {
                "font", "size", "lineSpacing", "firstLineIndentCm", "justify",
                "marginLeftCm", "marginRightCm", "marginTopCm", "marginBottomCm",
                "headingFont", "h1Size", "h2Size", "h3Size", "headingBold", "headingAlign"
            };
        }

        /// <summary>Текущие настройки оформления как параметры шаблона (заполнение по умолчанию).</summary>
        private static Dictionary<string, object> TplDefaults()
        {
            var d = new Dictionary<string, object>();
            try
            {
                var s = Store.LoadSettings();
                foreach (var k in TplKeys())
                    if (s.ContainsKey(k) && s[k] != null) d[k] = s[k];
            }
            catch (Exception ex) { Logger.Log("TplDefaults: " + ex.Message); }
            return d;
        }

        private static string N0(double v) { return v.ToString("0.##", CultureInfo.InvariantCulture); }

        private static void SetNum(Dictionary<string, object> a, string from, Dictionary<string, object> to, string toKey)
        {
            if (!a.ContainsKey(from)) return;
            double? v = ArgD(a, from);
            if (v.HasValue) to[toKey] = v.Value;
        }

        private static string TplBrief(DocTemplate t, bool active)
        {
            if (t == null) return "";
            var sb = new StringBuilder();
            sb.Append(active ? "[x] " : "[  ] ");
            sb.Append(t.Name).Append(" (id ").Append(t.Id).Append(") — ")
              .Append(t.Font).Append(' ').Append(N0(t.Size)).Append(" пт, интервал ").Append(N0(t.LineSpacing))
              .Append(", отступ ").Append(N0(t.FirstLineIndentCm)).Append(" см, поля ")
              .Append(N0(t.ML)).Append('/').Append(N0(t.MR)).Append('/')
              .Append(N0(t.MT)).Append('/').Append(N0(t.MB)).Append(" см, ")
              .Append(Templates.SourceLabel(t.Source));
            if (!string.IsNullOrEmpty(t.Description)) sb.Append(". ").Append(t.Description);
            return sb.ToString();
        }

        /// <summary>Анализ оформления открытого документа (задача «шаблоны»).</summary>
        private static AgentTool TplAnalyze()
        {
            return new AgentTool
            {
                Name = "analyze_document_format",
                Description = "Анализирует оформление открытого документа: шрифт, кегль, межстрочный интервал " +
                              "и абзацный отступ основного текста, шрифты и кегли заголовков по уровням, поля страницы. " +
                              "Результат напрямую подставляется в параметры create_template.",
                Parameters = Schema(Props()),
                Execute = _ =>
                {
                    var doc = Doc(WordApp);
                    var t = WordOps.AnalyzeFormatting(doc);
                    var sb = new StringBuilder();
                    sb.AppendLine("АНАЛИЗ ОФОРМЛЕНИЯ: " + doc.Name);
                    sb.AppendLine();
                    sb.Append(Templates.FormatSummary(t));
                    sb.AppendLine("Параметры для create_template:");
                    sb.AppendLine("  font=\"" + t.Font + "\", size=" + N0(t.Size) +
                                  ", line_spacing=" + N0(t.LineSpacing) +
                                  ", first_line_indent_cm=" + N0(t.FirstLineIndentCm) +
                                  ", justify=" + (t.Justify ? "true" : "false"));
                    sb.AppendLine("  margin_left_cm=" + N0(t.ML) + ", margin_right_cm=" + N0(t.MR) +
                                  ", margin_top_cm=" + N0(t.MT) + ", margin_bottom_cm=" + N0(t.MB));
                    sb.AppendLine("  heading_font=\"" + t.HeadingFont + "\", h1_size=" + N0(t.H1Size) +
                                  ", h2_size=" + N0(t.H2Size) + ", h3_size=" + N0(t.H3Size) +
                                  ", heading_bold=" + (t.HeadingBold ? "true" : "false") +
                                  ", heading_align=" + t.HeadingAlign);
                    return sb.ToString();
                }
            };
        }

        /// <summary>Создание/обновление шаблона оформления.</summary>
        private static AgentTool TplCreate()
        {
            return new AgentTool
            {
                Name = "create_template",
                Description = "Создаёт (или обновляет по совпадающему id) шаблон оформления документа. " +
                              "Параметры можно взять из analyze_document_format либо задать по описанию пользователя. " +
                              "Не указанные параметры берутся из текущих настроек оформления.",
                Parameters = Schema(Props(
                    "name", S("string", "Название шаблона, например «Курсовая по ГОСТ Р 7.0.5»"),
                    "description", S("string", "Подробное описание: для чего шаблон и какие требования к оформлению"),
                    "id", S("string", "Существующий id — чтобы обновить ранее созданный шаблон"),
                    "activate", S("boolean", "Сразу сделать шаблон активным (галочка + применить параметры к настройкам)"),
                    "font", S("string", "Шрифт основного текста"),
                    "size", S("number", "Кегль основного текста, пт"),
                    "line_spacing", S("number", "Межстрочный интервал (1, 1.5, 2)"),
                    "first_line_indent_cm", S("number", "Абзацный отступ, см"),
                    "justify", S("boolean", "Выравнивание основного текста по ширине"),
                    "margin_left_cm", S("number", "Левое поле, см"),
                    "margin_right_cm", S("number", "Правое поле, см"),
                    "margin_top_cm", S("number", "Верхнее поле, см"),
                    "margin_bottom_cm", S("number", "Нижнее поле, см"),
                    "heading_font", S("string", "Шрифт заголовков (пусто — как основной)"),
                    "h1_size", S("number", "Кегль заголовка 1 уровня, пт (0 = авто = основной + 2)"),
                    "h2_size", S("number", "Кегль заголовка 2 уровня, пт (0 = авто = основной + 1)"),
                    "h3_size", S("number", "Кегль заголовка 3 уровня, пт (0 = авто = основной)"),
                    "heading_bold", S("boolean", "Заголовки жирным"),
                    "heading_align", S("string", "Выравнивание заголовка 1 уровня: center | left | right | justify"),
                    "source", S("string", "Откуда шаблон: document | description | chat")
                ), "name"),
                Execute = a =>
                {
                    var d = ParseArgs(a);
                    string name = ArgS(d, "name").Trim();
                    if (name.Length == 0) return "ОШИБКА: не указано обязательное поле name (название шаблона).";

                    var dict = TplDefaults();
                    dict["name"] = name;
                    if (d.ContainsKey("description")) dict["description"] = ArgS(d, "description");
                    if (d.ContainsKey("id")) dict["id"] = ArgS(d, "id").Trim();
                    if (d.ContainsKey("font") && ArgS(d, "font").Trim().Length > 0)
                        dict["font"] = ArgS(d, "font").Trim();
                    if (d.ContainsKey("heading_font")) dict["headingFont"] = ArgS(d, "heading_font").Trim();
                    if (d.ContainsKey("justify")) dict["justify"] = ArgB(d, "justify", true);
                    if (d.ContainsKey("heading_bold")) dict["headingBold"] = ArgB(d, "heading_bold", true);
                    if (d.ContainsKey("heading_align")) dict["headingAlign"] = ArgS(d, "heading_align").Trim();
                    SetNum(d, "size", dict, "size");
                    SetNum(d, "line_spacing", dict, "lineSpacing");
                    SetNum(d, "first_line_indent_cm", dict, "firstLineIndentCm");
                    SetNum(d, "margin_left_cm", dict, "marginLeftCm");
                    SetNum(d, "margin_right_cm", dict, "marginRightCm");
                    SetNum(d, "margin_top_cm", dict, "marginTopCm");
                    SetNum(d, "margin_bottom_cm", dict, "marginBottomCm");
                    SetNum(d, "h1_size", dict, "h1Size");
                    SetNum(d, "h2_size", dict, "h2Size");
                    SetNum(d, "h3_size", dict, "h3Size");

                    string src = ArgS(d, "source").Trim().ToLowerInvariant();
                    if (src != "document" && src != "description" && src != "chat") src = "chat";
                    dict["source"] = src;

                    var tpl = Templates.FromDict(dict, src);
                    if (tpl == null) return "ОШИБКА: не удалось разобрать параметры шаблона.";

                    string active;
                    var items = Templates.Load(out active);
                    bool replaced = false;
                    for (int i = 0; i < items.Count; i++)
                        if (items[i].Id == tpl.Id) { items[i] = tpl; replaced = true; }
                    if (!replaced) items.Add(tpl);

                    bool activate = ArgB(d, "activate", false);
                    if (activate)
                    {
                        active = tpl.Id;
                        Templates.ApplyToSettings(tpl, Store.LoadSettings());
                    }
                    Templates.Save(items, active);

                    Logger.Log("AgentTool create_template: " + tpl.Name + " (id " + tpl.Id +
                               ", source=" + src + ", activate=" + activate + ")");
                    var sb = new StringBuilder();
                    sb.AppendLine((replaced ? "Шаблон обновлён: " : "Шаблон сохранён: ") +
                                  tpl.Name + " (id " + tpl.Id + ").");
                    if (activate) sb.AppendLine("Шаблон сделан активным — параметры применены к настройкам оформления.");
                    sb.AppendLine();
                    sb.Append(Templates.FormatSummary(tpl));
                    return sb.ToString();
                }
            };
        }

        /// <summary>Список сохранённых шаблонов.</summary>
        private static AgentTool TplList()
        {
            return new AgentTool
            {
                Name = "list_templates",
                Description = "Список сохранённых шаблонов оформления: id, название, описание и параметры. " +
                              "Активный шаблон отмечен [x]. Для выбора активного шаблона — select_template, " +
                              "для подробностей — get_template.",
                Parameters = Schema(Props()),
                Execute = _ =>
                {
                    string active;
                    var items = Templates.Load(out active);
                    if (items.Count == 0)
                        return "Сохранённых шаблонов оформления пока нет. Создай их через create_template " +
                               "(параметры можно получить через analyze_document_format).";
                    var sb = new StringBuilder();
                    sb.AppendLine("ШАБЛОНЫ ОФОРМЛЕНИЯ (всего " + items.Count + ", активный отмечен [x]):");
                    foreach (var t in items)
                        sb.AppendLine(TplBrief(t, t.Id == active));
                    sb.AppendLine();
                    sb.AppendLine("Подробности: get_template(query=id или название). Выбор активного: select_template(query).");
                    return sb.ToString();
                }
            };
        }

        /// <summary>Подробное описание одного шаблона.</summary>
        private static AgentTool TplGet()
        {
            return new AgentTool
            {
                Name = "get_template",
                Description = "Подробное описание шаблона оформления по id, названию или фрагменту названия/описания. " +
                              "Возвращает все параметры и текст описания.",
                Parameters = Schema(Props(
                    "query", S("string", "Id шаблона, название или часть названия/описания")), "query"),
                Execute = a =>
                {
                    var d = ParseArgs(a);
                    string q = ArgS(d, "query").Trim();
                    if (q.Length == 0) return "ОШИБКА: не указан query (id или название шаблона).";
                    string active;
                    var items = Templates.Load(out active);
                    if (items.Count == 0) return "Сохранённых шаблонов оформления пока нет.";
                    var t = Templates.Find(items, q);
                    if (t == null)
                        return "Шаблон по запросу «" + q + "» не найден. Вызови list_templates — там все названия и id.";
                    var sb = new StringBuilder();
                    sb.Append(Templates.Describe(t));
                    if (t.Id == active)
                        sb.AppendLine("Состояние: АКТИВНЫЙ (галочка в списке шаблонов).");
                    else
                        sb.AppendLine("Состояние: не активный.");
                    return sb.ToString();
                }
            };
        }

        /// <summary>Выбор активного шаблона — авто-подбор по названию/описанию.</summary>
        private static AgentTool TplSelect()
        {
            return new AgentTool
            {
                Name = "select_template",
                Description = "Выбирает активный шаблон оформления: ищет лучшее совпадение по названию и описанию " +
                              "(авто-подбор, без учёта падежей), ставит галочку и применяет параметры шаблона " +
                              "к настройкам оформления документа.",
                Parameters = Schema(Props(
                    "query", S("string", "Название или описание шаблона, например «курсовая по ГОСТ»"),
                    "apply_to_document", S("boolean", "Сразу применить оформление шаблона ко всему открытому документу (по умолчанию false — только настройки)")
                ), "query"),
                Execute = a =>
                {
                    var d = ParseArgs(a);
                    string q = ArgS(d, "query").Trim();
                    if (q.Length == 0) return "ОШИБКА: не указан query (название или описание шаблона).";
                    string active;
                    var items = Templates.Load(out active);
                    if (items.Count == 0)
                        return "Сохранённых шаблонов оформления пока нет — сначала create_template.";
                    var t = Templates.Find(items, q);
                    if (t == null)
                    {
                        var sb0 = new StringBuilder();
                        sb0.AppendLine("Подходящего шаблона по запросу «" + q + "» не найдено. Есть:");
                        foreach (var x in items) sb0.AppendLine("  - " + x.Name + " (id " + x.Id + "): " + x.Description);
                        return sb0.ToString();
                    }
                    Templates.Save(items, t.Id);
                    Templates.ApplyToSettings(t, Store.LoadSettings());
                    Logger.Log("AgentTool select_template: " + t.Name + " (id " + t.Id + ", query=" + q + ")");

                    bool applyDoc = ArgB(d, "apply_to_document", false);
                    string applied = "";
                    if (applyDoc)
                    {
                        try
                        {
                            var doc = Doc(WordApp);
                            WordOps.ApplyFormatToDoc(doc, WordOps.FromSettings(Store.LoadSettings()));
                            applied = "Оформление шаблона применено ко всему документу «" + doc.Name + "».";
                        }
                        catch (Exception ex) { applied = "Не удалось применить к документу: " + ex.Message; }
                    }

                    var sb = new StringBuilder();
                    sb.AppendLine("Активный шаблон: " + t.Name + " (id " + t.Id + ").");
                    sb.AppendLine("Параметры применены к настройкам оформления.");
                    if (applied.Length > 0) sb.AppendLine(applied);
                    sb.AppendLine();
                    sb.Append(Templates.FormatSummary(t));
                    return sb.ToString();
                }
            };
        }

        /// <summary>Word.Application подставляется извне (см. AgentPanel).</summary>
        public static Word.Application WordApp;
    }
}
