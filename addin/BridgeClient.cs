using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Web.Script.Serialization;

namespace WordAIAssistant
{
    /// <summary>
    /// Клиент локального моста DeepSeek (Python, http://127.0.0.1:8765).
    /// </summary>
    public class BridgeClient
    {
        public const string Dir = "bridge";

        public bool Ping()
        {
            try
            {
                string resp = Http("GET", "/health", null, 1500);
                return resp != null && resp.Contains("ok");
            }
            catch { return false; }
        }

        /// <summary>Если мост не запущен — стартует Python-процесс и ждёт /health.</summary>
        public bool EnsureRunning()
        {
            if (Ping()) return true;
            StartBridgeProcess();
            for (int i = 0; i < 60; i++)
            {
                System.Threading.Thread.Sleep(500);
                if (Ping()) return true;
            }
            return false;
        }

        public void StartBridgeProcess()
        {
            try
            {
                string bridgeDir = FindBridgeDir();
                if (bridgeDir == null) throw new DirectoryNotFoundException("Не найдена папка bridge");
                string py = Path.Combine(bridgeDir, "venv", "Scripts", "python.exe");
                string psi = File.Exists(py) ? py : "python";
                var startInfo = new ProcessStartInfo
                {
                    FileName = psi,
                    Arguments = File.Exists(py) ? "deepseek_bridge.py" : "deepseek_bridge.py",
                    WorkingDirectory = bridgeDir,
                    UseShellExecute = true,
                    WindowStyle = ProcessWindowStyle.Minimized
                };
                Process.Start(startInfo);
            }
            catch (Exception ex)
            {
                Logger.Log("Не удалось запустить мост: " + ex.Message);
            }
        }

        public static string FindBridgeDir()
        {
            // рядом с DLL, рядом с Word.exe, стандартное место установки
            string dllDir = AppDomain.CurrentDomain.BaseDirectory;
            string install = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "WordAIAssistant");
            foreach (string root in new[] { dllDir, Path.GetFullPath(Path.Combine(dllDir, "..")), install })
            {
                if (root == null) continue;
                try
                {
                    string cand = Path.Combine(root, Dir);
                    if (File.Exists(Path.Combine(cand, "deepseek_bridge.py"))) return cand;
                    if (File.Exists(Path.Combine(root, "deepseek_bridge.py"))) return root;
                }
                catch { }
            }
            return null;
        }

        public Dictionary<string, object> Status()
        {
            return Parse(Http("GET", "/status", null, 4000));
        }

        public Dictionary<string, object> OpenLogin()
        {
            return Parse(Http("POST", "/open_login", "{}", 60000));
        }

        public Dictionary<string, object> ConfirmLogin()
        {
            return Parse(Http("POST", "/confirm_login", "{}", 120000));
        }

        public void Shutdown()
        {
            try { Http("POST", "/shutdown", "{}", 3000); } catch { }
        }

        /// <summary>
        /// Pseudo-OpenAI Chat Completions: история сообщений + инструменты (OpenAI tools[]).
        /// Возвращает словарь ответа OpenAI-формата:
        /// choices[0].message.{content, tool_calls}, choices[0].finish_reason.
        /// При ошибке HTTP — {"error": {...}} или null.
        /// </summary>
        public Dictionary<string, object> ChatCompletion(List<Dictionary<string, object>> messages,
                                                         string toolsJson, bool thinking, bool search,
                                                         int timeoutSec)
        {
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var payload = new Dictionary<string, object>
            {
                { "model", "deepseek-chat" },
                { "messages", messages },
                { "stream", false },
                { "thinking", thinking },
                { "search", search },
                { "new_chat", true },
                { "timeout", timeoutSec }
            };
            if (!string.IsNullOrEmpty(toolsJson))
                payload["tools"] = ser.Deserialize<object>(toolsJson);
            string body = ser.Serialize(payload);
            string resp = Http("POST", "/v1/chat/completions", body, (timeoutSec + 60) * 1000);
            return Parse(resp);
        }

        /// <summary>Отправка запроса в чат DeepSeek (синхронно; вызывать из фонового потока).</summary>
        public Dictionary<string, object> Chat(string prompt, string[] images, bool thinking, bool search,
                                               bool newChat, int timeoutSec)
        {
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var payload = new Dictionary<string, object>
            {
                { "prompt", prompt },
                { "images", images ?? new string[0] },
                { "thinking", thinking },
                { "search", search },
                { "new_chat", newChat },
                { "timeout", timeoutSec }
            };
            string body = ser.Serialize(payload);
            string resp = Http("POST", "/chat", body, (timeoutSec + 60) * 1000);
            return Parse(resp);
        }

        /// <summary>Поиск картинок в интернете: {"ok":true,"items":[{"url","title","page"}]}.</summary>
        public Dictionary<string, object> ImageSearch(string query, int count)
        {
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var payload = new Dictionary<string, object> { { "query", query ?? "" }, { "count", count } };
            string resp = Http("POST", "/image_search", ser.Serialize(payload), 75000);
            return Parse(resp);
        }

        /// <summary>Растеризация: svg-текст ИЛИ data-URL картинки → {"ok":true,"png":"&lt;base64&gt;"}.</summary>
        public Dictionary<string, object> RenderImage(string svg, string dataUrl, int width, int height)
        {
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            var payload = new Dictionary<string, object>
            {
                { "svg", svg ?? "" },
                { "data_url", dataUrl ?? "" },
                { "width", width },
                { "height", height }
            };
            string resp = Http("POST", "/render_image", ser.Serialize(payload), 60000);
            return Parse(resp);
        }

        private static Dictionary<string, object> Parse(string json)
        {
            if (string.IsNullOrEmpty(json)) return null;
            var ser = new JavaScriptSerializer { MaxJsonLength = int.MaxValue };
            try { return ser.Deserialize<Dictionary<string, object>>(json); }
            catch { return null; }
        }

        private static string Http(string method, string path, string body, int timeoutMs)
        {
            var req = (HttpWebRequest)WebRequest.Create(Constants.BridgeBase + path);
            req.Method = method;
            req.Timeout = timeoutMs;
            req.ReadWriteTimeout = timeoutMs;
            req.ContentType = "application/json";
            req.Proxy = null;
            if (body != null)
            {
                byte[] data = Encoding.UTF8.GetBytes(body);
                req.ContentLength = data.Length;
                using (var s = req.GetRequestStream()) s.Write(data, 0, data.Length);
            }
            try
            {
                using (var resp = (HttpWebResponse)req.GetResponse())
                using (var sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    return sr.ReadToEnd();
            }
            catch (WebException wex)
            {
                var hr = wex.Response as HttpWebResponse;
                if (hr != null)
                {
                    using (var sr = new StreamReader(hr.GetResponseStream(), Encoding.UTF8))
                        return sr.ReadToEnd();
                }
                throw;
            }
        }
    }

    public static class Logger
    {
        private static readonly object gate = new object();
        public static string LogPath
        {
            get { return Path.Combine(Store.AppDataDir(), "addin.log"); }
        }
        public static void Log(string message)
        {
            try
            {
                lock (gate)
                    File.AppendAllText(LogPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + message + Environment.NewLine);
            }
            catch { }
        }
    }
}
