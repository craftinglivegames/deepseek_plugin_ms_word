using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordAIAssistant
{
    /// <summary>
    /// Сборка промптов и очистка ответов модели от «мусора»
    /// («Конечно, вот ваш текст…», лишние ```-блоки и т.п.).
    /// </summary>
    public static class Prompts
    {
        public const string SystemRulesRu =
            "Ты — профессиональный академический ассистент, помогаешь писать рефераты, курсовые, дипломные работы. " +
            "СТРОГИЕ ПРАВИЛА ВЫВОДА:\n" +
            "1. Выводи ТОЛЬКО готовый текст работы. Никаких вступлений («Конечно, вот ваш реферат…», «Ниже представлен…»), " +
            "никаких заключительных фраз («Если нужно, могу…»), никаких вопросов и пояснений от себя.\n" +
            "2. Не выводи код, блоки ``` и не упоминай форматирование — просто текст работы.\n" +
            "3. Используй markdown-разметку структуры: '# ' — название работы, '## ' — главы/разделы, '### ' — подразделы, " +
            "'#### ' — пункты. Нумерацию глав веди сквозную (1., 1.1., 1.2. и т.д.).\n" +
            "4. В конце работы обязательно добавь раздел '## Список литературы' с реальными, проверяемыми источниками " +
            "(оформи по ГОСТ Р 7.0.100-2018, если требования не говорят иное).\n" +
            "5. Пиши академическим стилем, развёрнуто, без воды и повторов. Объём соблюдай по запросу.\n" +
            "6. Не выдумывай факты, цифры и цитаты. Если данных не хватает — опирайся на общеизвестные проверяемые сведения.";

        public static string BuildMainPrompt(string workType, string requirements, string query,
                                             string docContext, string[] imagesNote)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(SystemRulesRu);
            sb.AppendLine();
            if (!string.IsNullOrWhiteSpace(docContext))
            {
                sb.AppendLine("--- КОНТЕКСТ ТЕКУЩЕГО ДОКУМЕНТА (уже сделанная работа) ---");
                sb.AppendLine(docContext);
                sb.AppendLine("--- КОНЕЦ КОНТЕКСТА ---");
                sb.AppendLine("Продолжай работу в том же стиле и той же терминологии, не переписывай уже написанное без просьбы.");
                sb.AppendLine();
            }
            if (!string.IsNullOrWhiteSpace(workType))
            {
                sb.AppendLine("Тип работы: " + workType.Trim());
            }
            if (!string.IsNullOrWhiteSpace(requirements))
            {
                sb.AppendLine("Требования к оформлению и содержанию:");
                sb.AppendLine(requirements.Trim());
            }
            if (!string.IsNullOrWhiteSpace(query))
            {
                sb.AppendLine();
                sb.AppendLine("ЗАПРОС (что именно нужно сделать):");
                sb.AppendLine(query.Trim());
            }
            return sb.ToString();
        }

        public static string BuildChapterPrompt(string chapterTitle, string chapterText, string userRequest, string docContext)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(SystemRulesRu);
            sb.AppendLine("Сейчас идёт КОРРЕКТИРОВКА одной главы работы. Верни ТОЛЬКО полный новый текст этой главы " +
                          "(включая её заголовок в формате '## ...'), без каких-либо комментариев.");
            if (!string.IsNullOrWhiteSpace(docContext))
            {
                sb.AppendLine("--- КОНТЕКСТ РАБОТЫ ---");
                sb.AppendLine(docContext);
                sb.AppendLine("--- КОНЕЦ КОНТЕКСТА ---");
            }
            sb.AppendLine("ТЕКУЩИЙ ТЕКСТ ГЛАВЫ («" + chapterTitle + "»):");
            sb.AppendLine(chapterText);
            sb.AppendLine();
            sb.AppendLine("ЗАПРОС К ГЛАВЕ: " + (string.IsNullOrWhiteSpace(userRequest) ? "улучшить и расширить главу" : userRequest.Trim()));
            return sb.ToString();
        }

        public static string BuildAiDetectPrompt(string text)
        {
            return "Ты — эксперт по определению текстов, сгенерированных нейросетями. Проанализируй текст работы ниже и составь отчёт на русском языке в формате:\n" +
                   "ВЕРОЯТНОСТЬ ИИ-ГЕНЕРАЦИИ: N% (обоснование в 2–3 предложениях)\n" +
                   "ПРИЗНАКИ ИИ: маркированный список конкретных признаков (шаблонность, клише, гладкость стиля, повторов и т.д.)\n" +
                   "СЛАБЫЕ МЕСТА: список\n" +
                   "РЕКОМЕНДАЦИИ: конкретные шаги по «очеловечиванию» текста (какие фразы переписать, что добавить)\n" +
                   "Отвечай только отчётом, без вступлений.\n\nТЕКСТ:\n" + text;
        }

        public static string BuildFactCheckPrompt(string text)
        {
            return "Ты — научный рецензент. Проверь работу ниже на достоверность: факты, цифры, термины, источники. " +
                   "Если доступен поиск в интернете — используй его для проверки ключевых утверждений. " +
                   "Составь отчёт на русском языке в формате:\n" +
                   "ИТОГ: насколько работе можно доверять (кратко)\n" +
                   "ОШИБКИ И СОМНИТЕЛЬНЫЕ УТВЕРЖДЕНИЯ: нумерованный список — утверждение, почему сомнительно, как правильно\n" +
                   "ИСТОЧНИКИ: надежные источники по теме, которые стоит добавить в работу (с выходными данными)\n" +
                   "Отвечай только отчётом, без вступлений.\n\nТЕКСТ:\n" + text;
        }

        /// <summary>Аудит абзацев: какие выглядят ИИ-шными / содержат сомнительные факты.</summary>
        public static string BuildParagraphAuditPrompt(string numberedParagraphs, string mode)
        {
            string what = mode == "ai"
                ? "определи, какие абзацы выглядят написанными нейросетью (шаблонность, клише, гладкость, «водянистость», повтор конструкций)"
                : "определи, какие абзацы содержат сомнительные, непроверенные или потенциально ложные факты, цифры и утверждения";

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Ты — эксперт по академическим текстам. Ниже даны абзацы работы, пронумерованные как 'N. текст'.");
            sb.AppendLine("Задача: " + what + ".");
            sb.AppendLine("Ответь СТРОГО в таком формате, без вступлений и без markdown-разметки:");
            sb.AppendLine("ПОМЕЧЕННЫЕ: 2,5,7");
            sb.AppendLine("2 | краткая причина пометки");
            sb.AppendLine("5 | краткая причина пометки");
            sb.AppendLine("ОТЧЁТ:");
            sb.AppendLine("краткий общий отчёт по работе (3-6 предложений).");
            sb.AppendLine("Номера бери ТОЛЬКО из списка абзацев. Если помеченных абзацев нет — оставь строку 'ПОМЕЧЕННЫЕ: ' без чисел, но отчёт напиши обязательно.");
            sb.AppendLine();
            sb.AppendLine("АБЗАЦЫ:");
            sb.Append(numberedParagraphs ?? "");
            return sb.ToString();
        }

        /// <summary>
        /// Разбор ответа аудита абзацев: номера помеченных абзацев и причины пометок.
        /// Номера — позиции в списке (1..N), отправленном в BuildParagraphAuditPrompt.
        /// </summary>
        public static void ParseParagraphAudit(string response,
                                               out System.Collections.Generic.HashSet<int> marked,
                                               out System.Collections.Generic.Dictionary<int, string> reasons)
        {
            marked = new System.Collections.Generic.HashSet<int>();
            reasons = new System.Collections.Generic.Dictionary<int, string>();
            if (string.IsNullOrEmpty(response)) return;

            var m = Regex.Match(response, @"ПОМЕЧЕННЫЕ\s*:\s*([0-9,\s;]*[0-9])", RegexOptions.IgnoreCase);
            if (m.Success)
            {
                foreach (var part in m.Groups[1].Value.Split(new[] { ',', ';', ' ', '\t', '\r', '\n' },
                                                              StringSplitOptions.RemoveEmptyEntries))
                {
                    int n;
                    if (int.TryParse(part.Trim(), out n)) marked.Add(n);
                }
            }

            foreach (var raw in response.Replace("\r\n", "\n").Split('\n'))
            {
                var line = raw.Trim();
                var mm = Regex.Match(line, @"^(\d{1,4})\s*[|:]\s+(\S.*)$");
                if (!mm.Success) continue;
                int num;
                if (!int.TryParse(mm.Groups[1].Value, out num)) continue;
                string reason = mm.Groups[2].Value.Trim();
                if (reason.Length > 0) { reasons[num] = reason; marked.Add(num); }
            }
        }

        /// <summary>Переписывание одного абзаца (анти-ИИ или факт-фикс).</summary>
        public static string BuildParagraphRewritePrompt(string mode, string paragraphText, string reason, string docContext)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine(SystemRulesRu);
            if (mode == "ai")
            {
                sb.AppendLine("Задача: перепиши ОДИН абзац так, чтобы он звучал естественно и по-человечески " +
                              "(живой ритм предложений, конкретика, никакого шаблонного ИИ-стиля и клише), " +
                              "сохранив смысл и академический уровень. Причина, по которой абзац помечен: " + (reason ?? "шаблонность") + ".");
            }
            else
            {
                sb.AppendLine("Задача: перепиши ОДИН абзац, исправив сомнительные факты и утверждения " +
                              "(при необходимости опирайся на общеизвестные проверяемые сведения). " +
                              "Причина пометки: " + (reason ?? "сомнительные факты") + ".");
            }
            if (!string.IsNullOrWhiteSpace(docContext))
            {
                sb.AppendLine("--- КОНТЕКСТ РАБОТЫ ---");
                sb.AppendLine(docContext);
                sb.AppendLine("--- КОНЕЦ КОНТЕКСТА ---");
            }
            sb.AppendLine("АБЗАЦ:");
            sb.AppendLine(paragraphText);
            sb.AppendLine("Верни ТОЛЬКО переписанный абзац, без заголовков, кавычек и комментариев.");
            return sb.ToString();
        }

        public static string BuildFormatExtractPrompt(string requirements)
        {
            return "Извлеки из требований к оформлению параметры и верни СТРОГО валидный JSON без пояснений, по схеме:\n" +
                   "{\"font\": \"название шрифта\", \"size\": число, \"lineSpacing\": число, \"firstLineIndentCm\": число, " +
                   "\"justify\": true/false, \"marginLeftCm\": число, \"marginRightCm\": число, \"marginTopCm\": число, \"marginBottomCm\": число}\n" +
                   "Если параметр в требованиях не указан, используй: font=Times New Roman, size=14, lineSpacing=1.5, " +
                   "firstLineIndentCm=1.25, justify=true, marginLeftCm=3, marginRightCm=1.5, marginTopCm=2, marginBottomCm=2.\n\n" +
                   "ТРЕБОВАНИЯ:\n" + requirements;
        }

        /// <summary>Общая схема JSON шаблона оформления (для обоих промптов ниже).</summary>
        public const string TemplateJsonSchema =
            "{\"name\": \"короткое понятное название шаблона\", " +
            "\"description\": \"подробное описание: для каких работ и какие требования к оформлению\", " +
            "\"font\": \"шрифт основного текста\", \"size\": 14, " +
            "\"lineSpacing\": 1.5, \"firstLineIndentCm\": 1.25, \"justify\": true, " +
            "\"marginLeftCm\": 3, \"marginRightCm\": 1.5, \"marginTopCm\": 2, \"marginBottomCm\": 2, " +
            "\"headingFont\": \"шрифт заголовков, пустая строка = как основной\", " +
            "\"h1Size\": 0, \"h2Size\": 0, \"h3Size\": 0, " +
            "\"headingBold\": true, \"headingAlign\": \"center\"}";

        /// <summary>Пояснение к схеме шаблона (типы и допустимые значения полей).</summary>
        public const string TemplateJsonHint =
            "Где: size — кегль основного текста в птах (14), lineSpacing — межстрочный интервал (1, 1.5, 2), " +
            "firstLineIndentCm — абзацный отступ в см, marginLeftCm/Right/Top/Bottom — поля страницы в см, " +
            "h1Size/h2Size/h3Size — кегли заголовков 1–3 уровней в птах, 0 = «авто» (кегль основного +2/+1/+0), " +
            "headingAlign — выравнивание заголовка 1 уровня: center, left, right или justify.";

        /// <summary>ИИ создаёт шаблон по анализу текущего документа.</summary>
        public static string BuildTemplateFromDocPrompt(string docTitle, string analysis)
        {
            return "Нужно создать шаблон оформления документа Word на основе анализа его текущего форматирования.\n" +
                   "Верни СТРОГО валидный JSON без пояснений, без обёртки в код и без текста вокруг, по схеме:\n" +
                   TemplateJsonSchema + "\n" +
                   "name придумай коротко и ёмко (например «Курсовая по ГОСТ Р 7.0.5»), " +
                   "description — 1–3 предложения, для каких работ подойдёт этот шаблон. " +
                   "Остальные поля перенеси из анализа как есть, не «улучшай» их.\n" +
                   TemplateJsonHint + "\n\n" +
                   "НАЗВАНИЕ ДОКУМЕНТА:\n" + (docTitle ?? "—") + "\n\n" +
                   "АНАЛИЗ ФОРМАТИРОВАНИЯ:\n" + (analysis ?? "—");
        }

        /// <summary>ИИ создаёт шаблон по текстовому описанию пользователя.</summary>
        public static string BuildTemplateFromDescriptionPrompt(string description)
        {
            return "Нужно создать шаблон оформления документа Word по текстовому описанию требований.\n" +
                   "Верни СТРОГО валидный JSON без пояснений, без обёртки в код и без текста вокруг, по схеме:\n" +
                   TemplateJsonSchema + "\n" +
                   "Если параметр в описании не указан, используй: font=Times New Roman, size=14, lineSpacing=1.5, " +
                   "firstLineIndentCm=1.25, justify=true, marginLeftCm=3, marginRightCm=1.5, marginTopCm=2, " +
                   "marginBottomCm=2, headingFont=\"\", h1Size=0, h2Size=0, h3Size=0, headingBold=true, " +
                   "headingAlign=center.\n" +
                   TemplateJsonHint + "\n\n" +
                   "ОПИСАНИЕ ТРЕБОВАНИЙ:\n" + description;
        }

        // ------------------------------------------------------ источники и изображения

        /// <summary>Подбор источников для каждого абзаца (JSON-ответ модели).</summary>
        public static string BuildFindSourcesPrompt(string numberedParagraphs, string query)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("Ты — библиограф. Перед тобой абзацы работы, каждый помечен кодом вида 'P42:' " +
                          "(42 — номер абзаца в документе).");
            sb.AppendLine("Подбери для КАЖДОГО смыслового абзаца реальный существующий источник (книгу, статью, " +
                          "государственный стандарт, официальный сайт, статью в журнале), на который этот абзац можно " +
                          "сослаться, и приведи готовое оформление записи списка литературы по ГОСТ Р 7.0.5-2008 / " +
                          "ГОСТ Р 7.0.100-2018.");
            sb.AppendLine("Требования к источнику: реальный и проверяемый (известный автор, издательство, журнал, " +
                          "организация или стандарт), по теме абзаца, год не позднее текущего. НЕ выдумывай несуществующие " +
                          "издания, журналы и ISBN. Если абзац — служебный (заголовок, «Введение», номера страниц, " +
                          "колонтитул), источник ему не нужен: просто не включай его в ответ.");
            if (!string.IsNullOrWhiteSpace(query))
                sb.AppendLine("Уточнение пользователя: " + query.Trim());
            sb.AppendLine();
            sb.AppendLine("Ответь СТРОГО валидным JSON, без markdown-обёртки и без пояснений вокруг:");
            sb.AppendLine("{\"sources\": [{\"p\": 42, \"ref\": \"Иванов И. И. Название книги. — М.: Издательство, 2020. — 320 с.\"}]}");
            sb.AppendLine("где p — код абзаца из списка (буква P не нужна), ref — одна оформленная запись источника. " +
                          "Один и тот же источник можно указать для нескольких абзацев — они получат один номер.");
            sb.AppendLine("Если абзацы не содержат содержательной темы (заглушки, служебный текст, случайные строки) — " +
                          "не выдумывай источники, ответь точно {\"sources\": []}.");
            sb.AppendLine();
            sb.AppendLine("АБЗАЦЫ:");
            sb.Append(numberedParagraphs ?? "");
            return sb.ToString();
        }

        /// <summary>
        /// Разбор ответа модели с подбором источников: пары (номер абзаца, оформление источника).
        /// Понимает и JSON-массив, и объект {"sources": [...]}, и запасной текстовый формат «42 | источник».
        /// </summary>
        public static System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<int, string>>
            ParseFindSources(string response)
        {
            var res = new System.Collections.Generic.List<System.Collections.Generic.KeyValuePair<int, string>>();
            if (string.IsNullOrWhiteSpace(response)) return res;

            string s = response.Replace("\r\n", "\n").Trim();
            s = Regex.Replace(s, @"^\s*```[a-zA-Z]*\s*\n?", "");
            s = Regex.Replace(s, @"\n?```\s*$", "");

            int iArr = s.IndexOf('['), iObj = s.IndexOf('{');
            int start = iArr < 0 ? iObj : iObj < 0 ? iArr : Math.Min(iArr, iObj);
            if (start >= 0)
            {
                int endArr = s.LastIndexOf(']'), endObj = s.LastIndexOf('}');
                int end = Math.Max(endArr, endObj);
                if (end > start)
                {
                    try
                    {
                        var ser = new System.Web.Script.Serialization.JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                        object parsed = ser.DeserializeObject(s.Substring(start, end - start + 1));
                        var items = ItemsOf(parsed);
                        foreach (object it in items)
                        {
                            var d = it as System.Collections.Generic.Dictionary<string, object>;
                            if (d == null) continue;
                            string refText = FirstStr(d, "ref", "source", "citation", "bib", "text", "title");
                            if (string.IsNullOrWhiteSpace(refText)) continue;
                            int p = FirstInt(d, -1, "p", "paragraph", "para", "n", "index", "num");
                            res.Add(new System.Collections.Generic.KeyValuePair<int, string>(p, refText.Trim()));
                        }
                    }
                    catch { res.Clear(); }
                }
            }

            if (res.Count > 0) return res;

            // запасной формат: «42 | источник»
            foreach (var raw in s.Split('\n'))
            {
                var line = raw.Trim();
                var m = Regex.Match(line, @"^P?(\d{1,4})\s*[|:]\s+(\S.*)$");
                if (!m.Success) continue;
                int p;
                if (!int.TryParse(m.Groups[1].Value, out p)) continue;
                res.Add(new System.Collections.Generic.KeyValuePair<int, string>(p, m.Groups[2].Value.Trim()));
            }
            return res;
        }

        private static System.Collections.IList ItemsOf(object parsed)
        {
            var list = parsed as System.Collections.IList;
            if (list != null) return list;
            var dict = parsed as System.Collections.Generic.Dictionary<string, object>;
            if (dict != null)
                foreach (var k in new[] { "sources", "items", "refs", "references", "results", "data" })
                    if (dict.ContainsKey(k) && dict[k] is System.Collections.IList)
                        return (System.Collections.IList)dict[k];
            return new System.Collections.ArrayList();
        }

        private static string FirstStr(System.Collections.Generic.Dictionary<string, object> d, params string[] keys)
        {
            foreach (var k in keys)
                if (d.ContainsKey(k) && d[k] != null)
                {
                    string v = Convert.ToString(d[k]);
                    if (!string.IsNullOrWhiteSpace(v)) return v;
                }
            return null;
        }

        private static int FirstInt(System.Collections.Generic.Dictionary<string, object> d, int def, params string[] keys)
        {
            foreach (var k in keys)
                if (d.ContainsKey(k) && d[k] != null)
                    try { return Convert.ToInt32(Math.Round(Convert.ToDouble(d[k]))); }
                    catch { }
            return def;
        }

        /// <summary>Правила про изображения для промптов (галочки вкладки «ИИ-агент»).</summary>
        public static string ImageRules(bool svg, bool web, string captionSample)
        {
            string cap = string.IsNullOrWhiteSpace(captionSample) ? "Рисунок 1. Название" : captionSample.Trim();
            var sb = new System.Text.StringBuilder();
            if (svg)
            {
                sb.AppendLine("РИСУНКИ И СХЕМЫ (галочка «SVG-изображения» включена): если материалу нужен рисунок, " +
                              "схема, диаграмма или блок-схема — сгенерируй его как SVG (чистый XML в блоке ```svg, без внешних " +
                              "CSS и ссылок, обязательно с viewBox и width/height в пикселях) и вставь в документ " +
                              "инструментом insert_svg_image, указав краткое осмысленное title. Подпись будет оформлена так: «" +
                              cap + "». Если пользователь просил вставить рисунок после конкретного абзаца — передай номер " +
                              "этого абзаца в after_paragraph (0 — в конец документа).");
            }
            else
            {
                sb.AppendLine("Рисунки и схемы пока не вставляй — галочка «SVG-изображения» выключена.");
            }
            if (web)
            {
                sb.AppendLine("ИЗОБРАЖЕНИЯ ИЗ ИНТЕРНЕТА (галочка «Изображения из интернет» включена): если нужна реальная " +
                              "картинка, найди её инструментом find_images и вставь ссылку из результата через insert_image " +
                              "с обязательным title. Подпись будет оформлена так: «" + cap + "». После какого абзаца вставить " +
                              "(пользователь просил) — указывай в after_paragraph (0 — в конец документа).");
            }
            else
            {
                sb.AppendLine("Картинки из интернета не используй — галочка «Изображения из интернет» выключена.");
            }
            return sb.ToString();
        }

        /// <summary>Очистка ответа модели от «служебного мусора».</summary>
        public static string Clean(string text)
        {
            if (string.IsNullOrEmpty(text)) return text;

            // 1) типовые вступления (только в начале)
            var leadPatterns = new[]
            {
                @"^\s*(Конечно|Хорошо|Отлично|Понял[а]?|Ясно)[^\n]{0,200}?[:.]\s*\n+",
                @"^\s*(Вот|Ниже)\s+(ваш|твоя|твой|ваша|представлен|представляю|готов)[^\n]{0,200}[!:.]?\s*\n+",
                @"^\s*Готов[оы]?[^\n]{0,150}[!:.]?\s*\n+",
                @"^\s*Ваш(а|) (готовый |полный |)(реферат|курсовая|дипломн\w+|работа|текст)[^\n]{0,150}[!:.]?\s*\n+",
                @"^\s*Извините[^\n]{0,200}[:.]\s*\n+"
            };
            bool changed = true;
            while (changed)
            {
                changed = false;
                foreach (var p in leadPatterns)
                {
                    var m = Regex.Match(text, p, RegexOptions.IgnoreCase);
                    if (m.Success && m.Index == 0)
                    {
                        text = text.Substring(m.Length);
                        changed = true;
                    }
                }
            }

            // 2) типовые концовки
            var tailPatterns = new[]
            {
                @"\n+\s*(Если|При необходимости|Нужно|Хотите|Могу)[^\n]{0,300}$",
                @"\n+\s*(Удачи|Успехов)[^\n]{0,100}[.!]?\s*$",
                @"\n+\s*—{3,}\s*$"
            };
            foreach (var p in tailPatterns)
            {
                var m = Regex.Match(text, p + @"\s*$", RegexOptions.IgnoreCase);
                if (m.Success) text = text.Substring(0, m.Index);
            }

            // 3) обёртки ```markdown / ``` — убрать только ограждающие, содержимое оставить
            text = Regex.Replace(text, @"^\s*```[a-zA-Z]*\s*\n", "");
            text = Regex.Replace(text, @"\n?```\s*$", "");

            // 4) HTML-теги и <think>...</think>, если просочились
            text = Regex.Replace(text, @"<think>.*?</think>", "", RegexOptions.Singleline);
            text = Regex.Replace(text, @"</?[a-zA-Z][a-zA-Z0-9]*(\s[^>]*)?/?>", "");

            return text.Trim();
        }
    }
}
