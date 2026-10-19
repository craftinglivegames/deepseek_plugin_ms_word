using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;

namespace WordAIAssistant
{
    /// <summary>
    /// Шаблон оформления документа: параметры основного текста и заголовков по уровням.
    /// Хранится в %APPDATA%\WordAIAssistant\templates.json.
    /// </summary>
    public class DocTemplate
    {
        public string Id = "";
        public string Name = "";
        public string Description = "";
        public string Source = "description";      // document | description | chat
        public string CreatedAt = "";              // yyyy-MM-dd HH:mm

        // --- основной текст ---
        public string Font = "Times New Roman";
        public double Size = 14;
        public double LineSpacing = 1.5;
        public double FirstLineIndentCm = 1.25;
        public bool Justify = true;
        public double ML = 3, MR = 1.5, MT = 2, MB = 2;

        // --- заголовки по уровням (0 = авто: Size+2 / Size+1 / Size) ---
        public string HeadingFont = "";
        public double H1Size, H2Size, H3Size;
        public bool HeadingBold = true;
        public string HeadingAlign = "center";     // выравнивание уровня 1

        public string EffectiveHeadingFont { get { return string.IsNullOrEmpty(HeadingFont) ? Font : HeadingFont; } }
        public double EffectiveH1 { get { return H1Size > 0 ? H1Size : Size + 2; } }
        public double EffectiveH2 { get { return H2Size > 0 ? H2Size : Size + 1; } }
        public double EffectiveH3 { get { return H3Size > 0 ? H3Size : Size; } }
    }

    /// <summary>
    /// Шаблоны оформления: загрузка/сохранение JSON и чистые функции
    /// (разбор ответа ИИ, авто-подбор по названию/описанию, подробное описание).
    /// Никакого UI и COM — покрывается unit-тестами (logic_test).
    /// </summary>
    public static class Templates
    {
        public const string FileName = "templates.json";

        public static string FilePath()
        {
            return Path.Combine(Store.AppDataDir(), FileName);
        }

        // ------------------------------------------------- JSON (чистые)

        public static string SerializeList(List<DocTemplate> items, string activeId)
        {
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var arr = new List<object>();
            if (items != null)
                foreach (var t in items) arr.Add(ToDict(t));
            var root = new Dictionary<string, object>
            {
                { "active", activeId ?? "" },
                { "items", arr }
            };
            return ser.Serialize(root);
        }

        public static void ParseList(string json, out List<DocTemplate> items, out string activeId)
        {
            items = new List<DocTemplate>();
            activeId = "";
            if (string.IsNullOrWhiteSpace(json)) return;
            try
            {
                var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                var root = ser.Deserialize<Dictionary<string, object>>(json);
                if (root == null) return;
                activeId = Str(root, "active", "");
                var arr = GetList(root, "items");
                if (arr == null) return;
                foreach (var o in arr)
                {
                    var d = o as Dictionary<string, object>;
                    if (d == null) continue;
                    var t = FromDict(d);
                    if (t != null) items.Add(t);
                }
            }
            catch (Exception ex) { Logger.Log("templates.json: " + ex.Message); }
        }

        /// <summary>Разбирает JSON-ответ ИИ {name, description, font, size, …} → DocTemplate (или null).</summary>
        public static DocTemplate FromJson(string raw, string source)
        {
            if (string.IsNullOrWhiteSpace(raw)) return null;
            int a = raw.IndexOf('{'), b = raw.LastIndexOf('}');
            if (a < 0 || b <= a) return null;
            try
            {
                var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
                var d = ser.Deserialize<Dictionary<string, object>>(raw.Substring(a, b - a + 1));
                if (d == null) return null;
                return FromDict(d, source);
            }
            catch (Exception ex) { Logger.Log("Templates.FromJson: " + ex.Message); return null; }
        }

        /// <summary>Словарь → DocTemplate. Пустое name/отсутствующее оформление → null.</summary>
        public static DocTemplate FromDict(Dictionary<string, object> d, string source = null)
        {
            if (d == null) return null;
            string name = Str(d, "name", "").Trim();
            string font = Str(d, "font", "").Trim();
            string hFont = Str(d, "headingFont", "").Trim();
            double size = Num(d, "size", 0);
            if (name.Length == 0 && font.Length == 0 && size <= 0) return null;
            if (name.Length == 0) name = "Без названия";

            string src = Str(d, "source", "").Trim();
            if (src.Length == 0) src = source ?? "description";

            return new DocTemplate
            {
                Id = Str(d, "id", "").Trim().Length > 0 ? Str(d, "id", "").Trim() : NewId(),
                Name = name,
                Description = Str(d, "description", "").Trim(),
                Source = src,
                CreatedAt = Str(d, "createdAt", "").Trim().Length > 0
                    ? Str(d, "createdAt", "").Trim()
                    : DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                Font = font.Length > 0 ? font : "Times New Roman",
                Size = size > 0 ? size : 14,
                LineSpacing = Num(d, "lineSpacing", 1.5),
                FirstLineIndentCm = Num(d, "firstLineIndentCm", 1.25),
                Justify = Bool(d, "justify", true),
                ML = Num(d, "marginLeftCm", 3),
                MR = Num(d, "marginRightCm", 1.5),
                MT = Num(d, "marginTopCm", 2),
                MB = Num(d, "marginBottomCm", 2),
                HeadingFont = hFont,
                H1Size = Num(d, "h1Size", 0),
                H2Size = Num(d, "h2Size", 0),
                H3Size = Num(d, "h3Size", 0),
                HeadingBold = Bool(d, "headingBold", true),
                HeadingAlign = NormalizeAlign(Str(d, "headingAlign", "center"))
            };
        }

        public static Dictionary<string, object> ToDict(DocTemplate t)
        {
            return new Dictionary<string, object>
            {
                { "id", t.Id }, { "name", t.Name }, { "description", t.Description },
                { "source", t.Source }, { "createdAt", t.CreatedAt },
                { "font", t.Font }, { "size", t.Size }, { "lineSpacing", t.LineSpacing },
                { "firstLineIndentCm", t.FirstLineIndentCm }, { "justify", t.Justify },
                { "marginLeftCm", t.ML }, { "marginRightCm", t.MR },
                { "marginTopCm", t.MT }, { "marginBottomCm", t.MB },
                { "headingFont", t.HeadingFont }, { "h1Size", t.H1Size },
                { "h2Size", t.H2Size }, { "h3Size", t.H3Size },
                { "headingBold", t.HeadingBold }, { "headingAlign", t.HeadingAlign }
            };
        }

        // ------------------------------------------------- хранилище

        public static List<DocTemplate> Load(out string activeId)
        {
            activeId = "";
            try
            {
                var p = FilePath();
                if (File.Exists(p))
                {
                    List<DocTemplate> items;
                    string a;
                    ParseList(File.ReadAllText(p), out items, out a);
                    activeId = a;
                    return items;
                }
            }
            catch (Exception ex) { Logger.Log("Templates.Load: " + ex.Message); }
            return new List<DocTemplate>();
        }

        public static void Save(List<DocTemplate> items, string activeId)
        {
            try { File.WriteAllText(FilePath(), SerializeList(items, activeId)); }
            catch (Exception ex) { Logger.Log("Templates.Save: " + ex.Message); }
        }

        /// <summary>Активный шаблон (или null).</summary>
        public static DocTemplate GetActive()
        {
            string id;
            var items = Load(out id);
            if (string.IsNullOrEmpty(id)) return null;
            return items.FirstOrDefault(t => t.Id == id);
        }

        /// <summary>Сохраняет список и делает шаблон с Id активным (галочка в списке).</summary>
        public static void SaveActive(List<DocTemplate> items, DocTemplate active)
        {
            Save(items, active == null ? "" : active.Id);
        }

        public static string NewId()
        {
            return Guid.NewGuid().ToString("N").Substring(0, 8);
        }

        public static string SourceLabel(string src)
        {
            switch ((src ?? "").ToLowerInvariant())
            {
                case "document": return "из документа";
                case "chat": return "из чата";
                default: return "по описанию";
            }
        }

        // ------------------------------------------------- авто-подбор

        /// <summary>Нормализация строки для сравнения: нижний регистр, ё→е, без лишних пробелов.</summary>
        public static string Norm(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new StringBuilder(s.Length);
            bool prevSpace = false;
            foreach (char c in s.ToLowerInvariant().Replace('ё', 'е').Trim())
            {
                if (char.IsWhiteSpace(c) || c == ',' || c == ';' || c == '.')
                {
                    if (!prevSpace && sb.Length > 0) sb.Append(' ');
                    prevSpace = true;
                }
                else { sb.Append(c); prevSpace = false; }
            }
            return sb.ToString().Trim();
        }

        /// <summary>Очки совпадения запроса с названием/описанием шаблона (0 = не подходит).</summary>
        public static int Score(DocTemplate t, string query)
        {
            if (t == null) return 0;
            string q = Norm(query);
            if (q.Length == 0) return 0;
            string name = Norm(t.Name), desc = Norm(t.Description);
            int score = 0;
            if (name.Length > 0 && name == q) score += 100;
            else if (name.Length > 0 && name.Contains(q)) score += 60;
            else if (q.Length > 3 && name.Length > 0 && q.Contains(name)) score += 40;
            if (desc.Length > 0 && desc.Contains(q)) score += 25;

            var nameKeys = WordKeys(name);
            var descKeys = WordKeys(desc);
            foreach (var w in SplitWords(q))
            {
                if (w.Length < 2) continue;
                string k = WordKey(w);
                if (nameKeys.Contains(k)) score += 15;      // совпадение слова в названии
                if (descKeys.Contains(k)) score += 6;       // совпадение слова в описании
            }
            return score;
        }

        /// <summary>Ключ слова для сравнения без учёта падежей: первые 4 символа.</summary>
        public static string WordKey(string w)
        {
            return w != null && w.Length > 4 ? w.Substring(0, 4) : w;
        }

        private static HashSet<string> WordKeys(string s)
        {
            var set = new HashSet<string>();
            foreach (var w in SplitWords(s))
                if (w.Length >= 2) set.Add(WordKey(w));
            return set;
        }

        /// <summary>Лучший шаблон по названию/описанию (авто-подбор в чате); null — если ничего не подошло.</summary>
        public static DocTemplate Match(List<DocTemplate> items, string query)
        {
            if (items == null || items.Count == 0) return null;
            DocTemplate best = null;
            int bestScore = 0;
            foreach (var t in items)
            {
                int s = Score(t, query);
                if (s > bestScore) { bestScore = s; best = t; }
            }
            return bestScore >= 10 ? best : null;
        }

        /// <summary>Точный поиск по id, затем по названию, затем авто-подбор.</summary>
        public static DocTemplate Find(List<DocTemplate> items, string idOrQuery)
        {
            if (items == null || items.Count == 0) return null;
            if (string.IsNullOrWhiteSpace(idOrQuery)) return null;
            string key = idOrQuery.Trim();
            var byId = items.FirstOrDefault(t => t.Id == key);
            if (byId != null) return byId;
            var byName = items.FirstOrDefault(t =>
                string.Equals(Norm(t.Name), Norm(key), StringComparison.Ordinal));
            if (byName != null) return byName;
            return Match(items, key);
        }

        private static IEnumerable<string> SplitWords(string s)
        {
            return (s ?? "").Split(new[] { ' ', '\t', '-', '/', '(', ')' }, StringSplitOptions.RemoveEmptyEntries);
        }

        // ------------------------------------------------- описание

        /// <summary>Краткая строка для списка.</summary>
        public static string SummaryLine(DocTemplate t)
        {
            if (t == null) return "";
            return t.Name + "  [" + t.Font + " " + N0(t.Size) + " пт, интервал " + N0(t.LineSpacing) + "]";
        }

        /// <summary>Только параметры оформления (для промпта/лога).</summary>
        public static string FormatSummary(DocTemplate t)
        {
            if (t == null) return "";
            var sb = new StringBuilder();
            sb.AppendLine("Основной текст:");
            sb.AppendLine("  шрифт: " + t.Font + ", размер: " + N0(t.Size) + " пт");
            sb.AppendLine("  межстрочный интервал: " + N0(t.LineSpacing));
            sb.AppendLine("  абзацный отступ: " + N0(t.FirstLineIndentCm) + " см");
            sb.AppendLine("  выравнивание: " + (t.Justify ? "по ширине" : "по левому краю"));
            sb.AppendLine("  поля, см: левое " + N0(t.ML) + ", правое " + N0(t.MR) +
                          ", верхнее " + N0(t.MT) + ", нижнее " + N0(t.MB));
            sb.AppendLine("Заголовки:");
            sb.AppendLine("  шрифт: " + t.EffectiveHeadingFont + (string.IsNullOrEmpty(t.HeadingFont) ? " (как основной)" : ""));
            sb.AppendLine("  уровень 1: " + N0(t.EffectiveH1) + " пт, " +
                          (t.HeadingBold ? "жирный" : "обычный") + ", " + AlignLabel(t.HeadingAlign));
            sb.AppendLine("  уровень 2: " + N0(t.EffectiveH2) + " пт, " + (t.HeadingBold ? "жирный" : "обычный"));
            sb.AppendLine("  уровень 3: " + N0(t.EffectiveH3) + " пт, " + (t.HeadingBold ? "жирный" : "обычный"));
            return sb.ToString();
        }

        /// <summary>Полное описание шаблона — для отдельного окна подробностей.</summary>
        public static string Describe(DocTemplate t)
        {
            if (t == null) return "Шаблон не выбран.";
            var sb = new StringBuilder();
            sb.AppendLine("Шаблон: " + t.Name);
            sb.AppendLine("Идентификатор: " + t.Id);
            sb.AppendLine("Создан: " + (string.IsNullOrEmpty(t.CreatedAt) ? "—" : t.CreatedAt));
            sb.AppendLine("Источник: " + SourceLabel(t.Source));
            sb.AppendLine();
            sb.AppendLine("ОПИСАНИЕ");
            sb.AppendLine(string.IsNullOrEmpty(t.Description) ? "(описание не заполнено)" : t.Description);
            sb.AppendLine();
            sb.AppendLine("ПАРАМЕТРЫ ОФОРМЛЕНИЯ");
            sb.Append(FormatSummary(t));
            sb.AppendLine();
            sb.AppendLine("ПРИМЕНЕНИЕ");
            sb.AppendLine("  Параметры основного текста используются при вставке глав и в окне");
            sb.AppendLine("  «Оформление документа». Поля страницы применяются к разделу документа.");
            sb.AppendLine("  Шаблон также доступен ИИ-агенту инструментами list_templates /");
            sb.AppendLine("  get_template / select_template.");
            return sb.ToString();
        }

        /// <summary>Применяет параметры шаблона к словарю настроек (и сохраняет их).</summary>
        public static void ApplyToSettings(DocTemplate t, Dictionary<string, object> settings)
        {
            if (t == null || settings == null) return;
            settings["font"] = t.Font;
            settings["size"] = t.Size;
            settings["lineSpacing"] = t.LineSpacing;
            settings["firstLineIndentCm"] = t.FirstLineIndentCm;
            settings["justify"] = t.Justify;
            settings["marginLeftCm"] = t.ML;
            settings["marginRightCm"] = t.MR;
            settings["marginTopCm"] = t.MT;
            settings["marginBottomCm"] = t.MB;
            settings["headingFont"] = t.HeadingFont;
            settings["h1Size"] = t.H1Size;
            settings["h2Size"] = t.H2Size;
            settings["h3Size"] = t.H3Size;
            settings["headingBold"] = t.HeadingBold;
            settings["headingAlign"] = t.HeadingAlign;
            Store.SaveSettings(settings);
        }

        // ------------------------------------------------- служебное

        private static string NormalizeAlign(string a)
        {
            a = (a ?? "").Trim().ToLowerInvariant();
            if (a == "left" || a == "влево" || a == "по левому краю") return "left";
            if (a == "right" || a == "вправо" || a == "по правому краю") return "right";
            if (a == "justify" || a == "по ширине") return "justify";
            return "center";
        }

        public static string AlignLabel(string a)
        {
            switch ((a ?? "").ToLowerInvariant())
            {
                case "left": return "по левому краю";
                case "right": return "по правому краю";
                case "justify": return "по ширине";
                default: return "по центру";
            }
        }

        private static string N0(double v)
        {
            return v.ToString("0.##", CultureInfo.InvariantCulture);
        }

        private static string Str(Dictionary<string, object> d, string k, string def)
        {
            object v;
            if (d == null || !d.TryGetValue(k, out v) || v == null) return def;
            return Convert.ToString(v, CultureInfo.InvariantCulture) ?? def;
        }

        private static double Num(Dictionary<string, object> d, string k, double def)
        {
            object v;
            if (d == null || !d.TryGetValue(k, out v) || v == null) return def;
            try
            {
                if (v is string) return Convert.ToDouble(((string)v).Replace(',', '.'), CultureInfo.InvariantCulture);
                return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            }
            catch { return def; }
        }

        private static bool Bool(Dictionary<string, object> d, string k, bool def)
        {
            object v;
            if (d == null || !d.TryGetValue(k, out v) || v == null) return def;
            if (v is bool) return (bool)v;
            var s = Str(d, k, "").Trim().ToLowerInvariant();
            if (s == "true" || s == "1" || s == "да") return true;
            if (s == "false" || s == "0" || s == "нет") return false;
            return def;
        }

        private static List<object> GetList(Dictionary<string, object> d, string k)
        {
            object v;
            if (d == null || !d.TryGetValue(k, out v) || v == null) return null;
            return v as List<object> ?? (v as System.Collections.IList)?.Cast<object>().ToList();
        }
    }
}
