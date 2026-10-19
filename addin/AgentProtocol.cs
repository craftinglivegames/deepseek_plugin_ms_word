using System;
using System.Collections.Generic;

namespace WordAIAssistant
{
    /// <summary>
    /// Разбор pseudo-OpenAI ответа моста (/v1/chat/completions) —
    /// чистые функции без UI и COM, покрыты юнит-тестами (agent_test).
    /// </summary>
    public static class AgentProtocol
    {
        /// <summary>choices[0] из ответа OpenAI-формата (массив — object[] или ArrayList).</summary>
        public static Dictionary<string, object> FirstChoice(Dictionary<string, object> resp)
        {
            object ch;
            if (resp == null || !resp.TryGetValue("choices", out ch) || ch == null) return null;
            var list = ch as System.Collections.IList;
            if (list == null || list.Count == 0) return null;
            return list[0] as Dictionary<string, object>;
        }

        /// <summary>Текст из поля message.content (или null).</summary>
        public static string GetMessageContent(Dictionary<string, object> choice)
        {
            object m;
            if (choice == null || !choice.TryGetValue("message", out m) || !(m is Dictionary<string, object>))
                return null;
            return GetString((Dictionary<string, object>)m, "content");
        }

        /// <summary>
        /// tool_calls из message -> упорядоченный список пар (имя инструмента, arguments JSON-строка).
        /// null, если вызовов нет. Разбирает и нормализованный формат моста,
        /// и «сырой» формат OpenAI (arguments строка или объект).
        /// </summary>
        public static List<KeyValuePair<string, string>> GetToolCalls(Dictionary<string, object> message)
        {
            object tc;
            if (message == null || !message.TryGetValue("tool_calls", out tc) || tc == null) return null;
            var arr = tc as System.Collections.IList;
            if (arr == null) return null;

            var res = new List<KeyValuePair<string, string>>();
            foreach (var item in arr)
            {
                var c = item as Dictionary<string, object>;
                if (c == null) continue;

                object fnObj;
                var fn = (c.TryGetValue("function", out fnObj) && fnObj is Dictionary<string, object>)
                    ? (Dictionary<string, object>)fnObj : c;

                string name = GetString(fn, "name");
                if (string.IsNullOrEmpty(name)) continue;

                object argsObj;
                string args;
                if (fn.TryGetValue("arguments", out argsObj) && argsObj != null)
                {
                    args = argsObj is string
                        ? (string)argsObj
                        : new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(argsObj);
                }
                else args = "{}";
                res.Add(new KeyValuePair<string, string>(name, args));
            }
            return res.Count > 0 ? res : null;
        }

        /// <summary>finish_reason == "tool_calls".</summary>
        public static bool IsToolCallFinish(Dictionary<string, object> choice)
        {
            object fr;
            return choice != null && choice.TryGetValue("finish_reason", out fr)
                   && string.Equals(Convert.ToString(fr), "tool_calls", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>Сообщение assistant с tool_calls в формате, принимаемом мостом.</summary>
        public static Dictionary<string, object> BuildAssistantToolCallMessage(
            string content, List<KeyValuePair<string, string>> calls)
        {
            var list = new List<object>();
            int n = 0;
            foreach (var kv in calls)
                list.Add(new Dictionary<string, object>
                {
                    { "id", "call_" + (++n) },
                    { "type", "function" },
                    { "function", new Dictionary<string, object> { { "name", kv.Key }, { "arguments", kv.Value } } }
                });
            return new Dictionary<string, object>
            {
                { "role", "assistant" },
                { "content", content },
                { "tool_calls", list }
            };
        }

        /// <summary>Сообщение role=tool с результатом инструмента.</summary>
        public static Dictionary<string, object> BuildToolResultMessage(string name, string result)
        {
            return new Dictionary<string, object>
            {
                { "role", "tool" },
                { "name", name },
                { "content", result ?? "" }
            };
        }

        /// <summary>Текст ошибки из тела ответа {"error": {...}} / {"error": "str"}.</summary>
        public static string ExtractError(Dictionary<string, object> resp)
        {
            object errObj;
            if (resp == null || !resp.TryGetValue("error", out errObj) || errObj == null)
                return null;
            var d = errObj as Dictionary<string, object>;
            if (d != null)
            {
                object m;
                if (d.TryGetValue("message", out m) && m != null) return Convert.ToString(m);
            }
            return Convert.ToString(errObj);
        }

        private static string GetString(Dictionary<string, object> d, string key)
        {
            object v;
            return (d != null && d.TryGetValue(key, out v) && v != null) ? Convert.ToString(v) : null;
        }
    }
}
