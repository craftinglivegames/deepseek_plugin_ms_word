using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web.Script.Serialization;

namespace WordAIAssistant
{
    /// <summary>Избранное (значения полей) и настройки, хранятся в %APPDATA%\WordAIAssistant.</summary>
    public static class Store
    {
        private static string _favPath, _setPath;

        public static string AppDataDir()
        {
            string dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                Constants.AppDataDirName);
            Directory.CreateDirectory(dir);
            return dir;
        }

        static Store()
        {
            string dir = AppDataDir();
            _favPath = Path.Combine(dir, "favorites.json");
            _setPath = Path.Combine(dir, "settings.json");
        }

        public static Dictionary<string, List<string>> LoadFavorites()
        {
            var defaults = new Dictionary<string, List<string>>
            {
                { "type", new List<string>() },
                { "req", new List<string>() },
                { "query", new List<string>() }
            };
            try
            {
                if (File.Exists(_favPath))
                {
                    var ser = new JavaScriptSerializer();
                    var data = ser.Deserialize<Dictionary<string, List<string>>>(File.ReadAllText(_favPath));
                    if (data != null)
                    {
                        foreach (var k in defaults.Keys.ToList())
                            if (data.ContainsKey(k) && data[k] != null) defaults[k] = data[k];
                    }
                }
            }
            catch (Exception ex) { Logger.Log("favorites.json: " + ex.Message); }
            return defaults;
        }

        public static void SaveFavorites(Dictionary<string, List<string>> fav)
        {
            try
            {
                var ser = new JavaScriptSerializer();
                File.WriteAllText(_favPath, ser.Serialize(fav));
            }
            catch (Exception ex) { Logger.Log("save favorites: " + ex.Message); }
        }

        public static Dictionary<string, object> LoadSettings()
        {
            var def = new Dictionary<string, object>
            {
                { "font", "Times New Roman" },
                { "size", 14.0 },
                { "lineSpacing", 1.5 },
                { "firstLineIndentCm", 1.25 },
                { "justify", true },
                { "marginLeftCm", 3.0 },
                { "marginRightCm", 1.5 },
                { "marginTopCm", 2.0 },
                { "marginBottomCm", 2.0 },
                { "contextChars", 6000 },
                { "autoClean", true },
                { "chatTimeout", 1500 },
                { "svgImages", false },
                { "webImages", false },
                { "figPrefix", "Рисунок" },
                { "figNumber", "число" },
                { "figDot", true },
                { "humanMode", true }
            };
            try
            {
                if (File.Exists(_setPath))
                {
                    var ser = new JavaScriptSerializer();
                    var data = ser.Deserialize<Dictionary<string, object>>(File.ReadAllText(_setPath));
                    if (data != null)
                        foreach (var k in data.Keys.ToList()) def[k] = data[k];
                }
            }
            catch (Exception ex) { Logger.Log("settings.json: " + ex.Message); }
            return def;
        }

        public static void SaveSettings(Dictionary<string, object> s)
        {
            try
            {
                var ser = new JavaScriptSerializer();
                File.WriteAllText(_setPath, ser.Serialize(s));
            }
            catch (Exception ex) { Logger.Log("save settings: " + ex.Message); }
        }

        // ====================================================== живой режим запросов

        /// <summary>Срабатывает при переключении галочки «Живые запросы» — чтобы
        /// все места с этой галочкой (окно запроса, вкладка генератора) оставались в согласии.</summary>
        public static event Action<bool> HumanModeChanged;

        private static bool ReadHumanMode(Dictionary<string, object> s)
        {
            try
            {
                if (s != null && s.ContainsKey("humanMode") && s["humanMode"] != null)
                    return Convert.ToBoolean(s["humanMode"]);
            }
            catch (Exception ex) { Logger.Log("humanMode: " + ex.Message); }
            return true;   // по умолчанию живой режим включён — бережём аккаунт
        }

        /// <summary>Живой режим: имитация человеческого темпа ввода при отправке
        /// запросов в чат DeepSeek (паузы + набор текста). Меняется пользователем
        /// отдельно от остальных настроек; читается мостом прямо из settings.json.</summary>
        public static bool HumanMode
        {
            get { return ReadHumanMode(LoadSettings()); }
            set
            {
                try
                {
                    var s = LoadSettings();
                    s["humanMode"] = value;
                    SaveSettings(s);
                }
                catch (Exception ex) { Logger.Log("save humanMode: " + ex.Message); }
                var h = HumanModeChanged;
                if (h != null) h(value);
            }
        }
    }
}
