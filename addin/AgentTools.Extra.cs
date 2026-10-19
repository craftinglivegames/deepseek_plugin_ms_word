using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    /// <summary>
    /// Дополнительные инструменты ИИ-агента (partial-часть AgentTools):
    /// изображения (SVG и из интернета) и работа с источниками.
    /// </summary>
    public static partial class AgentTools
    {
        /// <summary>Клиент моста — нужен для SVG→PNG и поиска картинок. Задаётся перед циклом агента.</summary>
        public static BridgeClient Bridge;

        /// <summary>Галочки вкладки «ИИ-агент»: разрешают агенту работать с изображениями.</summary>
        public static bool AllowSvgImages;
        public static bool AllowWebImages;

        private const string OffMsg =
            "ОШИБКА: функция отключена — включите соответствующую галочку во вкладке «ИИ-агент».";

        private class CitePara
        {
            public int Index;    // реальный номер абзаца в doc.Paragraphs
            public string Text;  // текст без абзацной отметки
        }

        // ============================================================ схемы инструментов

        internal static AgentTool InsertSvg()
        {
            return new AgentTool
            {
                Name = "insert_svg_image",
                Description = "Строит рисунок/схему по SVG-коду, конвертирует его в картинку и вставляет в документ " +
                              "вместе с подписью в формате «Рисунок N. Название». SVG должен быть самодостаточным " +
                              "(без внешних CSS/шрифтов), с viewBox и width/height в пикселях. " +
                              "after_paragraph — номер абзаца, ПОСЛЕ которого вставить (по умолчанию 0 — в конец документа). " +
                              "Требуется галочка «SVG-изображения».",
                Parameters = Schema(Props(
                    "svg", S("string", "Полный SVG-код: <svg …>…</svg>"),
                    "title", S("string", "Название рисунка для подписи (обязательно)"),
                    "width_px", S("integer", "Ширина рисунка в пикселях, по умолчанию 640"),
                    "after_paragraph", S("integer", "Вставить после абзаца с этим номером (0 = в конец документа)")),
                    "svg", "title"),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);
                    string svg = ArgS(d, "svg");
                    string title = ArgS(d, "title").Trim();
                    if (title.Length == 0) return "ОШИБКА: не задан title (название рисунка).";

                    int target = ArgI(d, "width_px", 640);
                    if (target < 100) target = 100;
                    if (target > 2000) target = 2000;

                    int w, h;
                    string norm = ImageOps.NormalizeSvg(svg, target, out w, out h);
                    if (norm == null) return "ОШИБКА: svg не распознан (нет тега <svg>).";

                    string svgPath = ImageOps.SaveTemp(norm, ".svg");
                    if (svgPath == null) return "ОШИБКА: не удалось сохранить SVG во временный файл.";

                    bool raster = false;
                    string imgPath = null;
                    try { imgPath = RenderPng(norm, null, w, h); }
                    catch (Exception ex) { Logger.Log("svg render: " + ex.Message); }
                    if (imgPath != null) { raster = true; }
                    else imgPath = svgPath;

                    string caption = MakeCaption(doc, title);
                    InsertFigure(doc, imgPath, caption, ArgI(d, "after_paragraph", 0));
                    return "Рисунок вставлен (" + (raster ? "SVG → PNG" : "как SVG") + "): " + caption;
                }
            };
        }

        internal static AgentTool FindImages()
        {
            return new AgentTool
            {
                Name = "find_images",
                Description = "Ищет картинки в интернете по запросу и возвращает список прямых ссылок с названиями. " +
                              "Ссылку можно передать в insert_image. Требуется галочка «Изображения из интернет».",
                Parameters = Schema(Props(
                    "query", S("string", "Что искать, например «схема процесса»"),
                    "count", S("integer", "Сколько результатов вернуть (1–10, по умолчанию 6)")),
                    "query"),
                Execute = a =>
                {
                    if (Bridge == null) return "ОШИБКА: локальный мост не запущен.";
                    var d = ParseArgs(a);
                    string q = ArgS(d, "query").Trim();
                    if (q.Length == 0) return "ОШИБКА: не задан query.";
                    int count = ArgI(d, "count", 6);
                    if (count < 1) count = 1;
                    if (count > 10) count = 10;

                    var resp = Bridge.ImageSearch(q, count);
                    if (resp == null) return "ОШИБКА: мост не ответил.";
                    string err;
                    if (!Ok(resp, out err)) return "ОШИБКА: поиск изображений не удался (" + err + ").";
                    var items = resp.ContainsKey("items") ? resp["items"] as List<object> : null;
                    if (items == null || items.Count == 0)
                        return "Ничего не найдено по запросу «" + q + "».";

                    var sb = new StringBuilder();
                    sb.AppendLine("Найдено " + items.Count + " по запросу «" + q + "»:");
                    for (int i = 0; i < items.Count; i++)
                    {
                        var it = items[i] as Dictionary<string, object>;
                        if (it == null) continue;
                        string url = AsStr(it, "url");
                        if (url.Length == 0) continue;
                        sb.AppendLine((i + 1) + ") " + url + " — " + AsStr(it, "title"));
                    }
                    return sb.ToString().TrimEnd();
                }
            };
        }

        internal static AgentTool InsertImage()
        {
            return new AgentTool
            {
                Name = "insert_image",
                Description = "Скачивает картинку по прямой ссылке url и вставляет её в документ с подписью " +
                              "в формате «Рисунок N. Название». Ссылку берите из find_images. " +
                              "after_paragraph — номер абзаца, ПОСЛЕ которого вставить (по умолчанию 0 — в конец документа). " +
                              "Требуется галочка «Изображения из интернет».",
                Parameters = Schema(Props(
                    "url", S("string", "Прямая ссылка на изображение (http/https)"),
                    "title", S("string", "Название рисунка для подписи (обязательно)"),
                    "after_paragraph", S("integer", "Вставить после абзаца с этим номером (0 = в конец документа)")),
                    "url", "title"),
                Execute = a =>
                {
                    var doc = Doc(WordApp);
                    var d = ParseArgs(a);
                    string url = ArgS(d, "url").Trim();
                    string title = ArgS(d, "title").Trim();
                    if (url.Length == 0) return "ОШИБКА: не задан url.";
                    if (title.Length == 0) return "ОШИБКА: не задан title (название рисунка).";

                    string ext;
                    string path = ImageOps.DownloadImage(url, out ext);
                    if (path == null) return "ОШИБКА: не удалось скачать изображение по ссылке.";

                    if (!ImageOps.WordCompatibleExt(ext))
                    {
                        string png;
                        try { png = RenderPng(null, path, 0, 0); }
                        catch (Exception ex) { png = null; Logger.Log("img convert: " + ex.Message); }
                        if (png == null)
                            return "ОШИБКА: формат " + ext + " Word не поддерживает и конвертировать не удалось.";
                        path = png;
                    }

                    string caption = MakeCaption(doc, title);
                    InsertFigure(doc, path, caption, ArgI(d, "after_paragraph", 0));
                    return "Изображение вставлено: " + caption;
                }
            };
        }

        internal static AgentTool FindSourcesTool()
        {
            return new AgentTool
            {
                Name = "find_sources",
                Description = "Подбирает реальные источники для абзацев документа, расставляет ссылки [1], [2] в тексте " +
                              "и создаёт/обновляет раздел «Список литературы» в конце работы. Документ обрабатывается " +
                              "целиком, включая разделы 2, 3 и заключение. " +
                              "Параметр query — уточнение темы (можно пустой).",
                Parameters = Schema(Props(
                    "query", S("string", "Уточнение запроса: на какие источники ориентироваться"),
                    "max_sources", S("integer", "Максимальное число источников в итоговом списке (0 = без ограничения). " +
                                    "Если пользователь назвал нужное количество — передай его сюда")),
                    "query"),
                Execute = a =>
                {
                    var d = ParseArgs(a);
                    return RunFindSources(WordApp, Bridge, ArgS(d, "query"), ArgI(d, "max_sources", 0));
                }
            };
        }

        internal static AgentTool ReorderSourcesTool()
        {
            return new AgentTool
            {
                Name = "reorder_sources",
                Description = "Перенумеровывает список литературы и ссылки [n] в тексте. " +
                              "mode=mention — источники идут в порядке первого упоминания в тексте (текст получает " +
                              "сплошные [1], [2], [3]…); mode=alphabet — список по алфавиту, а в тексте номера " +
                              "переставляются соответственно.",
                Parameters = Schema(Props(
                    "mode", S("string", "mention | alphabet (по умолчанию mention)"))),
                Execute = a =>
                {
                    var d = ParseArgs(a);
                    string mode = ArgS(d, "mode", "mention").Trim().ToLowerInvariant();
                    bool byMention = !(mode == "alphabet" || mode == "alpha" || mode == "abc" || mode == "алфавит");
                    return RunReorderSources(WordApp, byMention);
                }
            };
        }

        // ============================================================ рабочие процессы

        /// <summary>
        /// Ищет источники для абзацев, ставит [n] и обновляет список литературы.
        /// Документ отправляется модели чанками (по ~20k символов), иначе длинные работы
        /// обрезались и разделы 2–3 оставались без ссылок. Уже подобранные источники
        /// передаются в следующие чанки, чтобы модель переиспользовала их, а не плодила дубли.
        /// </summary>
        public static string RunFindSources(Word.Application app, BridgeClient bridge, string query, int maxSources = 0)
        {
            if (app == null) return "ОШИБКА: нет открытого документа.";
            if (bridge == null) return "ОШИБКА: локальный мост не запущен.";
            Word.Document doc = Doc(app);

            int bibIdx;
            var body = CollectBody(doc, out bibIdx);
            if (body.Count == 0) return "ОШИБКА: в документе нет текстовых абзацев.";

            var settings = Store.LoadSettings();
            int timeout = SettingsInt(settings, "chatTimeout", 900);

            // --- существующий раздел «Список литературы»: канонический список с дедупликацией ---
            var entries = new List<string>();
            if (bibIdx > 0)
            {
                int pi = 0;
                foreach (Word.Paragraph p in doc.Paragraphs)
                {
                    if (++pi > 8000) break;
                    if (pi <= bibIdx) continue;
                    string t;
                    try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); }
                    catch { continue; }
                    if (t.Length == 0) continue;
                    entries.Add(Citations.StripEntryNumber(t));
                }
            }
            var dd = Citations.Deduplicate(entries);
            var refs = dd.Refs;
            var valid = new HashSet<int>();
            foreach (var cp in body) valid.Add(cp.Index);

            // --- чанки абзацев: один запрос на ~20k символов ---
            const int ChunkChars = 20000;
            var chunks = new List<List<CitePara>>();
            var cur = new List<CitePara>();
            int used = 0;
            foreach (var cp in body)
            {
                int lineLen = cp.Text.Length + 12;
                if (used + lineLen > ChunkChars && cur.Count > 0)
                {
                    chunks.Add(cur);
                    cur = new List<CitePara>();
                    used = 0;
                }
                cur.Add(cp);
                used += lineLen;
            }
            if (cur.Count > 0) chunks.Add(cur);

            // --- запросы по чанкам ---
            var paraOld = new Dictionary<int, int>();   // абзац -> номер источника (1-based, канонический)
            int added = 0;
            int chunkErrors = 0;
            for (int c = 0; c < chunks.Count; c++)
            {
                var sb = new StringBuilder();
                foreach (var cp in chunks[c])
                    sb.AppendLine("P" + cp.Index + ": " + cp.Text);
                if (sb.Length == 0) continue;

                // передаём уже выбранные источники — модель переиспользует их для смежных чанков
                string q = query;
                if (refs.Count > 0)
                {
                    var kn = new StringBuilder();
                    kn.AppendLine();
                    kn.AppendLine("Уже выбранные источники работы (если подходят абзацу — используй ИМЕННО эту запись, не создавай вариант с другим оформлением):");
                    for (int k = 0; k < refs.Count; k++)
                        kn.Append(k + 1).Append(") ").AppendLine(refs[k]);
                    q = string.IsNullOrWhiteSpace(q) ? kn.ToString().Trim() : q.Trim() + "\n" + kn;
                }
                if (maxSources > 0)
                    q = (string.IsNullOrWhiteSpace(q) ? "" : q.Trim() + "\n") +
                        "В итоговом списке литературы должно быть НЕ БОЛЕЕ " + maxSources +
                        " источников. Новые добавляй только если среди выбранных нет подходящего.";

                string prompt = Prompts.BuildFindSourcesPrompt(sb.ToString(), q);
                var res = bridge.Chat(prompt, null, false, true, true, timeout);
                if (res == null || !res.ContainsKey("ok") || !Convert.ToBoolean(res["ok"]))
                {
                    string err = res != null && res.ContainsKey("error") ? Convert.ToString(res["error"]) : "нет ответа";
                    Logger.Log("find_sources chunk " + (c + 1) + "/" + chunks.Count + ": " + err);
                    chunkErrors++;
                    continue;
                }
                string answer = Convert.ToString(res.ContainsKey("text") ? res["text"] : "") ?? "";

                foreach (var kv in Prompts.ParseFindSources(answer))
                {
                    if (!valid.Contains(kv.Key)) continue;
                    string r = (kv.Value ?? "").Trim();
                    if (r.Length < 5) continue;
                    if (paraOld.ContainsKey(kv.Key)) continue;
                    int n = Citations.FindCanonical(r, refs);
                    if (n == 0)
                    {
                        if (maxSources > 0 && refs.Count >= maxSources) continue;
                        refs.Add(r);
                        n = refs.Count;
                        added++;
                    }
                    paraOld[kv.Key] = n;
                }
            }

            if (refs.Count == 0) return "ОШИБКА: источники подобрать не удалось — модель не прислала записи " +
                (chunkErrors > 0 ? "(ошибок запросов: " + chunkErrors + ")" : "(обычно это значит, что в тексте нет темы для цитирования)") +
                ". Уточните параметр query или попросите список литературы для осмысленного содержания работы.";

            // абзацы, для которых ИИ источника не прислал, — оставляем их прежние ссылки
            // (номера в документе пересчитываются через карту дедупликации)
            var paraNums = new List<List<int>>();
            for (int b = 0; b < body.Count; b++)
            {
                int n;
                if (paraOld.TryGetValue(body[b].Index, out n)) { paraNums.Add(new List<int> { n }); continue; }
                var keep = new List<int>();
                foreach (var o in Citations.Extract(body[b].Text))
                {
                    int m;
                    if (dd.Map.TryGetValue(o, out m)) { if (!keep.Contains(m)) keep.Add(m); }
                    else if (o >= 1 && o <= refs.Count && !keep.Contains(o)) keep.Add(o);
                }
                paraNums.Add(keep);
            }

            var rr = Citations.Reorder(refs, paraNums, true);

            int updated = 0;
            for (int b = 0; b < body.Count; b++)
            {
                int newN = 0;
                if (paraNums[b].Count > 0 && rr.Map.ContainsKey(paraNums[b][0])) newN = rr.Map[paraNums[b][0]];
                bool had = Citations.Extract(body[b].Text).Count > 0;
                if (newN == 0 && !had) continue;
                string newText = Citations.Append(body[b].Text, newN);
                if (newText.Length == 0 || newText == body[b].Text) continue;
                try { WordOps.ReplaceParagraphText(doc, body[b].Index, newText); updated++; }
                catch (Exception ex) { Logger.Log("find_sources p" + body[b].Index + ": " + ex.Message); }
            }

            try { WriteBibliography(doc, rr.Refs); }
            catch (Exception ex) { Logger.Log("find_sources bib: " + ex.Message); return "ОШИБКА: не удалось записать список литературы: " + ex.Message; }

            return "Готово. Источников в списке: " + rr.Refs.Count + " (новых: " + added +
                   (dd.Removed > 0 ? ", дублей убрано: " + dd.Removed : "") +
                   "), обновлено абзацев: " + updated + ". Раздел «Список литературы» " +
                   (bibIdx > 0 ? "обновлён." : "создан.") +
                   (chunkErrors > 0 ? " ВНИМАНИЕ: не удалось обработать чанков: " + chunkErrors + " из " + chunks.Count + " — часть разделов могла остаться без ссылок." : "");
        }

        /// <summary>Перенумеровывает список литературы и ссылки [n] в тексте.</summary>
        public static string RunReorderSources(Word.Application app, bool byMention)
        {
            if (app == null) return "ОШИБКА: нет открытого документа.";
            Word.Document doc = Doc(app);

            int bibIdx;
            var body = CollectBody(doc, out bibIdx);
            if (bibIdx < 0) return "ОШИБКА: в документе нет раздела со списком литературы («Список литературы» и т.п.).";
            if (body.Count == 0) return "ОШИБКА: в тексте нет абзацев.";

            var entries = new List<string>();
            int i = 0;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                if (++i > 8000) break;
                if (i <= bibIdx) continue;
                string t;
                try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); } catch { continue; }
                if (t.Length > 0) entries.Add(t);
            }
            if (entries.Count == 0) return "ОШИБКА: список литературы пуст.";

            var refs0 = new List<string>();
            foreach (var e in entries) refs0.Add(Citations.StripEntryNumber(e));

            // почти-дубликаты (одна работа в разных вариантах оформления) схлопываем,
            // номера в тексте пересчитываются через карту
            var dd = Citations.Deduplicate(refs0);

            var paraNums = new List<List<int>>();
            int outOfRange = 0;
            foreach (var cp in body)
            {
                var nums = new List<int>();
                foreach (var n in Citations.Extract(cp.Text))
                {
                    int m;
                    if (n > entries.Count) { outOfRange++; continue; }
                    if (dd.Map.TryGetValue(n, out m)) { if (!nums.Contains(m)) nums.Add(m); }
                }
                paraNums.Add(nums);
            }

            var rr = Citations.Reorder(dd.Refs, paraNums, byMention);

            // сквозная карта для текста: номер в текущем списке -> канонический -> новый
            var remapAll = new Dictionary<int, int>();
            for (int n = 1; n <= entries.Count; n++)
            {
                int nw;
                if (rr.Map.TryGetValue(dd.Map[n], out nw)) remapAll[n] = nw;
            }

            int changed = 0;
            for (int b = 0; b < body.Count; b++)
            {
                string nt = Citations.Remap(body[b].Text, remapAll);
                if (nt == body[b].Text) continue;
                try { WordOps.ReplaceParagraphText(doc, body[b].Index, nt); changed++; }
                catch (Exception ex) { Logger.Log("reorder p" + body[b].Index + ": " + ex.Message); }
            }

            try { WriteBibliography(doc, rr.Refs); }
            catch (Exception ex) { Logger.Log("reorder bib: " + ex.Message); return "ОШИБКА: не удалось переписать список литературы: " + ex.Message; }

            string mode = byMention ? "по порядку упоминания в тексте" : "по алфавиту";
            string extra = (dd.Removed > 0 ? ", удалено дублей: " + dd.Removed : "") +
                           (outOfRange > 0 ? ", цитат вне списка: " + outOfRange : "");
            return "Готово: источник " + rr.Refs.Count + ", порядок — " + mode +
                   ", обновлено абзацев: " + changed + extra + ".";
        }

        // ============================================================ помощники

        /// <summary>Абзацы основной части (до раздела «Список литературы»), без заголовков.</summary>
        private static List<CitePara> CollectBody(Word.Document doc, out int bibIdx)
        {
            bibIdx = -1;
            var res = new List<CitePara>();
            int i = 0;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                if (++i > 8000) break;
                string t;
                try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); }
                catch { continue; }
                if (bibIdx < 0 && Citations.IsBibliographyHeading(t)) { bibIdx = i; continue; }
                if (bibIdx > 0) break;
                if (t.Length < 3) continue;
                bool heading = false;
                try
                {
                    int lvl = (int)p.OutlineLevel;
                    heading = lvl >= 1 && lvl <= 3 && t.Length < 200;
                }
                catch { }
                if (heading) continue;
                res.Add(new CitePara { Index = i, Text = t });
            }
            return res;
        }

        /// <summary>Дописывает абзацы в конец документа, не «приклеивая» их к последнему абзацу.</summary>
        private static void AppendParagraphs(Word.Document doc, string text)
        {
            bool lastIsEmpty = false;
            try
            {
                string lt = doc.Paragraphs[doc.Paragraphs.Count].Range.Text;
                lastIsEmpty = lt.TrimEnd('\r', '\v', '\a', '\f').Trim().Length == 0;
            }
            catch { }
            Word.Range at = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
            at.InsertAfter((lastIsEmpty ? "" : "\r") + (text ?? ""));
        }

        /// <summary>Записывает раздел «Список литературы» (существующий — переписывает, иначе создаёт).
        /// Записям явно задаётся оформление основного текста — иначе они наследуют стиль Normal
        /// документа (случается Courier New 10 без отступов), что выглядит как «испорченный список».</summary>
        private static void WriteBibliography(Word.Document doc, List<string> refs)
        {
            int bibIdx = -1;
            int i = 0;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                if (++i > 8000) break;
                string t;
                try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f'); } catch { continue; }
                if (Citations.IsBibliographyHeading(t)) { bibIdx = i; break; }
            }

            bool created = bibIdx <= 0;
            if (!created)
            {
                try
                {
                    int start = doc.Paragraphs[bibIdx].Range.End;
                    int end = doc.Content.End - 1;
                    if (end > start)
                    {
                        Word.Range del = doc.Range(start, end);
                        object unit = Type.Missing, cnt = Type.Missing;
                        del.Delete(ref unit, ref cnt);
                    }
                }
                catch (Exception ex) { Logger.Log("bib delete: " + ex.Message); }
            }
            else
            {
                AppendParagraphs(doc, "Список литературы\r");
            }

            if (refs == null || refs.Count == 0) return;
            var sb = new StringBuilder();
            for (int k = 0; k < refs.Count; k++)
                sb.Append(Citations.RenumberEntry(refs[k], k + 1)).Append("\r");
            AppendParagraphs(doc, sb.ToString());

            // заголовок после вставок мог появиться только что — ищем заново
            int headIdx = -1;
            i = 0;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                if (++i > 8000) break;
                string t;
                try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f'); } catch { continue; }
                if (Citations.IsBibliographyHeading(t)) { headIdx = i; break; }
            }
            if (headIdx <= 0) return;

            string fontName; float size, firstIndent, lineSpacing, spaceAfter, spaceBefore;
            int spacingRule, align;
            SampleBodyFormat(doc, headIdx, out fontName, out size, out firstIndent,
                             out lineSpacing, out spacingRule, out spaceAfter, out spaceBefore, out align);

            if (created)
            {
                // свежесозданный заголовок — как заголовки разделов: шрифт текста, по центру, полужирный
                try
                {
                    Word.Range hr = doc.Paragraphs[headIdx].Range;
                    hr.Font.Name = fontName;
                    hr.Font.Size = size;
                    hr.Font.Bold = 1;
                    hr.Font.Italic = 0;
                    hr.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    hr.ParagraphFormat.FirstLineIndent = 0;
                }
                catch (Exception ex) { Logger.Log("bib head fmt: " + ex.Message); }
            }

            try
            {
                Word.Range rng = doc.Range(doc.Paragraphs[headIdx].Range.End, doc.Content.End - 1);
                rng.Font.Name = fontName;
                rng.Font.Size = size;
                rng.Font.Bold = 0;
                rng.Font.Italic = 0;
                var pf = rng.ParagraphFormat;
                pf.LineSpacingRule = (Word.WdLineSpacing)spacingRule;
                pf.LineSpacing = lineSpacing;
                pf.FirstLineIndent = firstIndent;
                pf.LeftIndent = 0;
                pf.Alignment = (Word.WdParagraphAlignment)align;
                pf.SpaceAfter = spaceAfter;
                pf.SpaceBefore = spaceBefore;
            }
            catch (Exception ex) { Logger.Log("bib fmt: " + ex.Message); }
        }

        /// <summary>Образец оформления основного текста: ближайший содержательный абзац перед списком литературы.</summary>
        private static void SampleBodyFormat(Word.Document doc, int beforeIdx,
            out string fontName, out float size, out float firstIndent,
            out float lineSpacing, out int spacingRule, out float spaceAfter, out float spaceBefore, out int align)
        {
            // разумные значения по умолчанию (типовая курсовая)
            fontName = "Times New Roman"; size = 14f; firstIndent = 35.4f;
            lineSpacing = 18f; spacingRule = 1; spaceAfter = 0f; spaceBefore = 0f; align = 3;

            for (int idx = beforeIdx - 1; idx >= 1; idx--)
            {
                Word.Paragraph p;
                try { p = doc.Paragraphs[idx]; } catch { break; }
                string t;
                try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f').Trim(); } catch { continue; }
                if (t.Length < 150) continue;   // пропускаем заголовки и короткие строки
                Word.Range r = p.Range;
                try
                {
                    Word.Range fc = doc.Range(r.Start, r.Start + 1);
                    fontName = Convert.ToString(fc.Font.Name);
                    size = fc.Font.Size;
                }
                catch { }
                try
                {
                    var pf = r.ParagraphFormat;
                    spacingRule = (int)pf.LineSpacingRule;
                    lineSpacing = pf.LineSpacing;
                    firstIndent = pf.FirstLineIndent;
                    align = (int)pf.Alignment;
                    spaceAfter = pf.SpaceAfter;
                    spaceBefore = pf.SpaceBefore;
                }
                catch { }
                break;
            }

            // значения «не определено» приводим к нормальным
            if (string.IsNullOrWhiteSpace(fontName) || fontName == "\u0001") fontName = "Times New Roman";
            if (float.IsNaN(size) || size <= 0 || size > 200) size = 14f;
            if (spacingRule < 0 || spacingRule > 5) { spacingRule = 1; lineSpacing = 18f; }
            if (float.IsNaN(lineSpacing) || lineSpacing <= 0 || lineSpacing > 500) lineSpacing = 18f;
            if (float.IsNaN(firstIndent) || Math.Abs(firstIndent) > 1000) firstIndent = 35.4f;
            if (align < 0 || align > 3) align = 3;
            if (float.IsNaN(spaceAfter) || spaceAfter < 0) spaceAfter = 0f;
            if (float.IsNaN(spaceBefore) || spaceBefore < 0) spaceBefore = 0f;
        }

        /// <summary>Подпись для нового рисунка с учётом настроек и уже существующих номеров.</summary>
        private static string MakeCaption(Word.Document doc, string title)
        {
            var s = Store.LoadSettings();
            string prefix = SettingsStr(s, "figPrefix", "Рисунок");
            string mode = SettingsStr(s, "figNumber", "число");
            bool dot = SettingsBool(s, "figDot", true);
            int next = ImageOps.NextFigureIndex(ExistingFigureNumbers(doc, prefix), mode);
            return ImageOps.BuildFigureCaption(prefix, ImageOps.FormatFigureNumber(next, mode), dot, title);
        }

        /// <summary>Номера уже существующих подписей «Рисунок N …».</summary>
        public static List<string> ExistingFigureNumbers(Word.Document doc, string prefix)
        {
            var res = new List<string>();
            if (string.IsNullOrWhiteSpace(prefix)) return res;
            var rx = new System.Text.RegularExpressions.Regex(
                @"(?i)^\s*" + System.Text.RegularExpressions.Regex.Escape(prefix.Trim()) +
                @"\s+(?-i:([0-9]{1,4}|[A-ZА-ЯЁ]))");
            int i = 0;
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                if (++i > 8000) break;
                string t;
                try { t = p.Range.Text.TrimEnd('\r', '\v', '\a', '\f'); } catch { continue; }
                var m = rx.Match(t);
                if (m.Success && m.Groups[1].Success) res.Add(m.Groups[1].Value);
            }
            return res;
        }

        /// <summary>Вставка картинки + подпись в следующем абзаце (по центру).
        /// afterPara &lt;= 0 — в конец документа; &gt; 0 — после абзаца с этим номером.</summary>
        private static void InsertFigure(Word.Document doc, string imagePath, string caption, int afterPara = 0)
        {
            int pos = doc.Content.End - 1;
            if (afterPara > 0)
            {
                try
                {
                    int n = doc.Paragraphs.Count;
                    if (afterPara > n) afterPara = n;
                    // конец (без параграфной отметки) — картинка встанет сразу после абзаца
                    pos = doc.Paragraphs[afterPara].Range.End;
                }
                catch (Exception ex) { Logger.Log("figure pos: " + ex.Message); pos = doc.Content.End - 1; }
            }
            Word.Range r = doc.Range(pos, pos);
            r.InsertAfter(caption + "\r");
            r = doc.Range(pos, pos);
            r.InsertAfter("\r");

            object link = false, save = true;
            object range = doc.Range(pos, pos);
            Word.InlineShape sh = doc.InlineShapes.AddPicture(imagePath, ref link, ref save, ref range);

            try
            {
                float usable = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;
                float ow = sh.Width, oh = sh.Height;
                float k = 1f;
                if (usable > 50f && ow > usable) k = usable / ow;
                if (oh > 560f && oh * k > 560f) k = 560f / oh;
                if (k < 0.999f && ow > 0 && oh > 0)
                {
                    sh.Width = ow * k;
                    sh.Height = oh * k;
                }
            }
            catch (Exception ex) { Logger.Log("figure size: " + ex.Message); }

            try
            {
                Word.Range picP = sh.Range.Paragraphs[1].Range;
                picP.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                Word.Range capP = doc.Range(picP.End, picP.End);
                capP.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
            }
            catch (Exception ex) { Logger.Log("figure align: " + ex.Message); }
        }

        /// <summary>Растеризация через мост: svg-текст ИЛИ файл с картинкой → temp-файл PNG (null — нельзя).</summary>
        private static string RenderPng(string svg, string filePath, int w, int h)
        {
            if (Bridge == null) return null;
            string dataUrl = null;
            if (filePath != null)
            {
                try
                {
                    byte[] bytes = File.ReadAllBytes(filePath);
                    string ext = Path.GetExtension(filePath);
                    dataUrl = "data:" + ImageOps.MimeFromExt(ext) + ";base64," + Convert.ToBase64String(bytes);
                }
                catch (Exception ex) { Logger.Log("render data: " + ex.Message); return null; }
            }
            if (svg == null && dataUrl == null) return null;

            var resp = Bridge.RenderImage(svg, dataUrl, w, h);
            if (resp == null) return null;
            string err;
            if (!Ok(resp, out err)) { Logger.Log("render_image: " + err); return null; }
            string b64 = resp.ContainsKey("png") ? Convert.ToString(resp["png"]) : null;
            if (string.IsNullOrEmpty(b64)) return null;
            try
            {
                string p = ImageOps.NewPath(".png");
                File.WriteAllBytes(p, Convert.FromBase64String(b64));
                return p;
            }
            catch (Exception ex) { Logger.Log("render save: " + ex.Message); return null; }
        }

        // ============================================================ мелочи

        private static bool Ok(Dictionary<string, object> resp, out string err)
        {
            err = "";
            if (resp == null) { err = "нет ответа"; return false; }
            if (resp.ContainsKey("error")) { err = Convert.ToString(resp["error"]); return false; }
            if (resp.ContainsKey("ok"))
            {
                bool ok;
                try { ok = Convert.ToBoolean(resp["ok"]); }
                catch { ok = true; }
                if (!ok) { err = "ok=false"; return false; }
            }
            return true;
        }

        private static string AsStr(Dictionary<string, object> d, string k)
        {
            if (d == null || !d.ContainsKey(k) || d[k] == null) return "";
            return Convert.ToString(d[k]) ?? "";
        }

        private static int SettingsInt(Dictionary<string, object> s, string k, int def)
        {
            if (s == null || !s.ContainsKey(k) || s[k] == null) return def;
            try { return Convert.ToInt32(Math.Round(Convert.ToDouble(s[k]))); }
            catch { return def; }
        }

        private static string SettingsStr(Dictionary<string, object> s, string k, string def)
        {
            if (s == null || !s.ContainsKey(k) || s[k] == null) return def;
            string v = Convert.ToString(s[k]);
            return string.IsNullOrWhiteSpace(v) ? def : v;
        }

        private static bool SettingsBool(Dictionary<string, object> s, string k, bool def)
        {
            if (s == null || !s.ContainsKey(k) || s[k] == null) return def;
            try { return Convert.ToBoolean(s[k]); }
            catch { return def; }
        }
    }
}
