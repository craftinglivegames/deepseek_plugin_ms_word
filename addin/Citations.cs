using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace WordAIAssistant
{
    /// <summary>Результат перенумерации источников: новый порядок списка и карта старых номеров в новые.</summary>
    public class RefReorder
    {
        public List<string> Refs = new List<string>();                 // ссылки в новом порядке (без номеров)
        public Dictionary<int, int> Map = new Dictionary<int, int>();  // старый номер -> новый номер
    }

    /// <summary>Результат дедупликации: канонические записи и карта старых позиций (1-based) в новые.</summary>
    public class RefDedup
    {
        public List<string> Refs = new List<string>();
        public Dictionary<int, int> Map = new Dictionary<int, int>();  // старая позиция -> новая (1-based)
        public int Removed;
    }

    /// <summary>
    /// Ссылки [1] в тексте и список литературы: разбор, очистка, перенумерация.
    /// Здесь только чистые функции (без Word COM) — они покрыты logic_test.
    /// </summary>
    public static class Citations
    {
        // « [1]» с пробелом/табом перед скобками — для Strip (поглощает лишний пробел)
        private static readonly Regex RxCite = new Regex(@"\s*\[(\d{1,3})\]");
        // только скобки — для Remap (пробелы сохраняем)
        private static readonly Regex RxCiteTight = new Regex(@"\[(\d{1,3})\]");
        private static readonly Regex RxEntryNum = new Regex(
            @"^\s*(\d{1,3})(?:[.)]|\s+)\s*(.*)$", RegexOptions.Singleline);
        private static readonly Regex RxMultiSpace = new Regex(@"[ \t]{2,}");

        /// <summary>Заголовки разделов со списком литературы.</summary>
        private static readonly Regex RxBibHeading = new Regex(
            @"^\s*(Список литературы|Список использованных источников|Список источников|Список использованной литературы|Библиографический список|Библиография|Литература)\s*([(:].*)?$",
            RegexOptions.IgnoreCase);

        // ---------------------------------------------------------- цитаты

        /// <summary>Номера [n] в порядке появления, без повторов.</summary>
        public static List<int> Extract(string text)
        {
            var res = new List<int>();
            if (string.IsNullOrEmpty(text)) return res;
            foreach (Match m in RxCite.Matches(text))
            {
                int n;
                if (int.TryParse(m.Groups[1].Value, out n) && n > 0 && !res.Contains(n)) res.Add(n);
            }
            return res;
        }

        /// <summary>Убирает все цитаты [n] и подчищает двойные пробелы.</summary>
        public static string Strip(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            string s = RxCite.Replace(text, "");
            s = RxMultiSpace.Replace(s, " ");
            return s.TrimEnd(' ', '\t');
        }

        /// <summary>Чистый текст + цитата [n] в конце (n=0 — просто чистый текст).</summary>
        public static string Append(string text, int n)
        {
            string s = Strip(text);
            if (n <= 0) return s;
            return (s.Length > 0 ? s + " " : "") + "[" + n + "]";
        }

        /// <summary>Заменяет номера цитат по карте (старый -> новый). Неизвестные номера не трогает.</summary>
        public static string Remap(string text, Dictionary<int, int> map)
        {
            if (string.IsNullOrEmpty(text) || map == null || map.Count == 0) return text;
            return RxCiteTight.Replace(text, m =>
            {
                int old;
                if (!int.TryParse(m.Groups[1].Value, out old)) return m.Value;
                int nw;
                if (!map.TryGetValue(old, out nw)) return m.Value;
                return "[" + nw + "]";
            });
        }

        // ---------------------------------------------------------- список литературы

        /// <summary>Похожа ли строка на заголовок раздела со списком литературы.</summary>
        public static bool IsBibliographyHeading(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return false;
            string t = text.Trim();
            if (t.Length == 0 || t.Length > 140) return false;
            return RxBibHeading.IsMatch(t);
        }

        /// <summary>Если запись начинается с номера («1.», «2)», «3 ») — true и номер.</summary>
        public static bool EntryNumber(string entry, out int n)
        {
            n = 0;
            if (string.IsNullOrEmpty(entry)) return false;
            var m = RxEntryNum.Match(entry);
            if (!m.Success) return false;
            return int.TryParse(m.Groups[1].Value, out n) && n > 0;
        }

        /// <summary>Текст записи без ведущего номера.</summary>
        public static string StripEntryNumber(string entry)
        {
            if (string.IsNullOrEmpty(entry)) return "";
            var m = RxEntryNum.Match(entry);
            return (m.Success ? m.Groups[2].Value : entry).Trim();
        }

        /// <summary>Запись с новым номером: «3. Текст».</summary>
        public static string RenumberEntry(string entry, int n)
        {
            return n + ". " + StripEntryNumber(entry);
        }

        /// <summary>Ключ сортировки записи (без номера, в нижнем регистре, без лишних пробелов).</summary>
        public static string SortKey(string entry)
        {
            string s = StripEntryNumber(entry ?? "");
            s = Regex.Replace(s, @"[\s\u00A0]+", " ").Trim().ToLowerInvariant();
            return s;
        }

        /// <summary>
        /// Перенумеровывает список источников.
        /// refs — тексты записей в порядке их текущих номеров (refs[0] = источник [1]).
        /// paraNums — для каждого абзаца списка его цитаты в порядке появления.
        /// byMention=true  — порядок по первому упоминанию в тексте (неназванные — в конец);
        /// byMention=false — по алфавиту.
        /// </summary>
        public static RefReorder Reorder(List<string> refs, List<List<int>> paraNums, bool byMention)
        {
            var res = new RefReorder();
            if (refs == null || refs.Count == 0) return res;
            int n = refs.Count;
            var order = new List<int>();

            if (byMention)
            {
                var seen = new HashSet<int>();
                if (paraNums != null)
                    foreach (var list in paraNums)
                        if (list != null)
                            foreach (var num in list)
                                if (num >= 1 && num <= n && !seen.Contains(num))
                                {
                                    seen.Add(num);
                                    order.Add(num);
                                }
                for (int i = 1; i <= n; i++)
                    if (!seen.Contains(i)) order.Add(i);
            }
            else
            {
                for (int i = 1; i <= n; i++) order.Add(i);
                var keys = new Dictionary<int, string>();
                foreach (int i in order) keys[i] = SortKey(refs[i - 1]);
                order.Sort((a, b) =>
                {
                    int c = string.Compare(keys[a], keys[b], StringComparison.CurrentCulture);
                    return c != 0 ? c : a.CompareTo(b);
                });
            }

            for (int pos = 0; pos < order.Count; pos++)
            {
                int oldNum = order[pos];
                res.Refs.Add(refs[oldNum - 1]);
                res.Map[oldNum] = pos + 1;
            }
            return res;
        }

        /// <summary>Перенумерованный список записей для вставки в документ.</summary>
        public static List<string> RenumberAll(List<string> refs)
        {
            var res = new List<string>();
            if (refs == null) return res;
            for (int i = 0; i < refs.Count; i++) res.Add(RenumberEntry(refs[i], i + 1));
            return res;
        }

        // ---------------------------------------------------------- дедупликация

        /// <summary>Голова записи: автор и название (часть до « // » — разделителя источника по ГОСТ).</summary>
        public static string EntryHead(string entry)
        {
            string s = StripEntryNumber(entry ?? "");
            int cut = s.IndexOf("//");
            return cut > 10 ? s.Substring(0, cut).Trim() : s;
        }

        /// <summary>Нормализация для сравнения: буквы и цифры в нижнем регистре, без пунктуации и «ё».</summary>
        private static string FuzzyNorm(string s)
        {
            if (string.IsNullOrEmpty(s)) return "";
            var sb = new System.Text.StringBuilder(s.Length);
            foreach (char ch in s.ToLowerInvariant())
            {
                char c = ch == 'ё' ? 'е' : ch;
                if (char.IsLetterOrDigit(c)) sb.Append(c);
            }
            return sb.ToString();
        }

        /// <summary>Похожесть двух нормализованных строк: коэффициент Дайса по биграммам (0..1).</summary>
        private static double DiceBigrams(string a, string b)
        {
            if (a.Length < 2 || b.Length < 2)
                return string.Equals(a, b, StringComparison.Ordinal) ? 1.0 : 0.0;
            var ga = new Dictionary<string, int>();
            for (int i = 0; i < a.Length - 1; i++)
            {
                string bg = a.Substring(i, 2);
                int v; ga.TryGetValue(bg, out v); ga[bg] = v + 1;
            }
            var gb = new Dictionary<string, int>();
            for (int i = 0; i < b.Length - 1; i++)
            {
                string bg = b.Substring(i, 2);
                int v; gb.TryGetValue(bg, out v); gb[bg] = v + 1;
            }
            int inter = 0;
            foreach (var kv in ga)
            {
                int v; if (gb.TryGetValue(kv.Key, out v)) inter += Math.Min(kv.Value, v);
            }
            return 2.0 * inter / ((a.Length - 1) + (b.Length - 1));
        }

        /// <summary>
        /// Почти-дубликаты: одна и та же работа в разных вариантах оформления
        /// (например, один и тот же тест, опубликованный в разных сборниках).
        /// Сравниваем голову записи (автор + название) — выходные данные («// журнал…») могут отличаться.
        /// </summary>
        public static bool AreNearDuplicates(string a, string b)
        {
            string na = FuzzyNorm(EntryHead(a)), nb = FuzzyNorm(EntryHead(b));
            if (na.Length < 10 || nb.Length < 10) return false;
            if (na.StartsWith(nb, StringComparison.Ordinal) || nb.StartsWith(na, StringComparison.Ordinal))
                return true;
            return DiceBigrams(na, nb) >= 0.85;
        }

        /// <summary>Ищет среди refs точный или почти-дубликат записи. Возвращает номер (1-based) или 0.</summary>
        public static int FindCanonical(string entry, List<string> refs)
        {
            if (refs == null || refs.Count == 0 || string.IsNullOrWhiteSpace(entry)) return 0;
            string key = SortKey(entry);
            for (int i = 0; i < refs.Count; i++)
                if (SortKey(refs[i]) == key) return i + 1;
            for (int i = 0; i < refs.Count; i++)
                if (AreNearDuplicates(refs[i], entry)) return i + 1;
            return 0;
        }

        /// <summary>
        /// Убирает почти-дубликаты из списка записей. Map покрывает каждую исходную
        /// позицию (1-based), так что все ссылки в тексте остаются валидными.
        /// </summary>
        public static RefDedup Deduplicate(List<string> entries)
        {
            var res = new RefDedup();
            if (entries == null) return res;
            var seen = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < entries.Count; i++)
            {
                string e = (entries[i] ?? "").Trim();
                if (e.Length == 0) { res.Map[i + 1] = Math.Max(1, res.Refs.Count); continue; }
                int canon = FindCanonical(e, res.Refs);
                if (canon == 0)
                {
                    res.Refs.Add(e);
                    canon = res.Refs.Count;
                }
                else res.Removed++;
                res.Map[i + 1] = canon;
            }
            return res;
        }
    }
}
