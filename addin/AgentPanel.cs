using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    /// <summary>
    /// Вкладка «ИИ-агент»: чат с DeepSeek в pseudo-OpenAI совместимости.
    /// Агент сам вызывает инструменты работы с документом (tool calling)
    /// и выполняет их локально, пока не получит финальный ответ.
    /// </summary>
    public class AgentPanel : UserControl
    {
        private readonly Word.Application app;
        private readonly BridgeClient bridge;
        private const int MaxSteps = 8;           // защита от бесконечного цикла вызовов

        private RichTextBox log;
        private TextBox input;
        private Button btnSend, btnNew;
        private CheckBox chkThinking, chkSearch, chkSvg, chkWeb;
        private Button btnAiCheck, btnFactCheck, btnSources, btnOrdMention, btnOrdAlpha;
        private FlowLayoutPanel row, row2, capRow;
        private TextBox txtCapWord, txtCapPreview;
        private ComboBox cboCapNum, cboCapDot;
        private Panel bottomPanel;
        private Label status;
        private readonly Action onAiCheck, onFactCheck;

        private readonly List<Dictionary<string, object>> messages =
            new List<Dictionary<string, object>>();
        private bool busy;

        public AgentPanel(Word.Application app, BridgeClient bridge,
                          Action aiCheck = null, Action factCheck = null)
        {
            this.app = app;
            this.bridge = bridge;
            this.onAiCheck = aiCheck;
            this.onFactCheck = factCheck;
            InitUi();
        }

        // =========================================================== UI

        private void InitUi()
        {
            Dock = DockStyle.Fill;
            Padding = new Padding(6);

            log = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Font = new Font("Segoe UI", 9f),
                WordWrap = true
            };

            var bottom = new Panel { Dock = DockStyle.Bottom, Height = 150 };
            bottomPanel = bottom;

            input = new TextBox
            {
                Multiline = true,
                Dock = DockStyle.Top,
                Height = 74,
                ScrollBars = ScrollBars.Vertical,
                Font = new Font("Segoe UI", 9f)
            };

            // строка управления: при нехватке ширины кнопки переносятся на вторую линию,
            // высота блока и нижней панели пересчитываются автоматически — ничего не обрезается
            row = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, WrapContents = true, Padding = new Padding(0, 2, 0, 0) };
            btnSend = new Button { Text = "Отправить", AutoSize = true, Padding = new Padding(14, 0, 14, 0), Height = 27 };
            btnNew = new Button { Text = "Новый чат", AutoSize = true, Padding = new Padding(14, 0, 14, 0), Height = 27 };
            chkThinking = new CheckBox { Text = "◉ Размышление", AutoSize = true, Margin = new Padding(8, 5, 0, 0) };
            chkSearch = new CheckBox { Text = "⌕ Поиск в интернете", AutoSize = true, Margin = new Padding(8, 5, 0, 0) };
            chkSvg = new CheckBox { Text = "SVG-изображения", AutoSize = true, Margin = new Padding(8, 5, 0, 0) };
            chkWeb = new CheckBox { Text = "Изображения из интернет", AutoSize = true, Margin = new Padding(8, 5, 0, 0) };
            btnSend.Click += (s, e) => Send();
            btnNew.Click += (s, e) => NewChat();
            row.Controls.AddRange(new Control[] { btnSend, btnNew, chkThinking, chkSearch, chkSvg, chkWeb });

            // строка проверок (дублирует вкладку «ИИ-генератор») и работы с источниками
            row2 = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, WrapContents = true, Padding = new Padding(0, 0, 0, 0) };
            btnAiCheck = new Button { Text = "✦ Проверка на ИИ", AutoSize = true, Padding = new Padding(10, 0, 10, 0), Height = 27, Margin = new Padding(0, 2, 6, 0) };
            btnFactCheck = new Button { Text = "☰ Факты и источники", AutoSize = true, Padding = new Padding(10, 0, 10, 0), Height = 27, Margin = new Padding(0, 2, 6, 0) };
            btnSources = new Button { Text = "Найти источники", AutoSize = true, Padding = new Padding(10, 0, 10, 0), Height = 27, Margin = new Padding(0, 2, 6, 0) };
            btnOrdMention = new Button { Text = "Порядок: по упоминанию", AutoSize = true, Padding = new Padding(10, 0, 10, 0), Height = 27, Margin = new Padding(0, 2, 6, 0) };
            btnOrdAlpha = new Button { Text = "Порядок: по алфавиту", AutoSize = true, Padding = new Padding(10, 0, 10, 0), Height = 27, Margin = new Padding(0, 2, 6, 0) };
            btnAiCheck.Visible = onAiCheck != null;
            btnFactCheck.Visible = onFactCheck != null;
            btnAiCheck.Click += (s, e) => { if (onAiCheck != null) onAiCheck(); };
            btnFactCheck.Click += (s, e) => { if (onFactCheck != null) onFactCheck(); };
            btnSources.Click += (s, e) => FindSources();
            btnOrdMention.Click += (s, e) => ReorderSources(true);
            btnOrdAlpha.Click += (s, e) => ReorderSources(false);
            row2.Controls.AddRange(new Control[] { btnAiCheck, btnFactCheck, btnSources, btnOrdMention, btnOrdAlpha });

            // строка формата подписи рисунков: слово + нумерация + превью + точка
            capRow = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 32, WrapContents = true, Padding = new Padding(0, 0, 0, 0) };
            var lblCap = new Label { Text = "Формат подписи рисунков:", AutoSize = true, Margin = new Padding(0, 6, 0, 0) };
            txtCapWord = new TextBox { Width = 74, Text = "Рисунок", Margin = new Padding(6, 3, 0, 0) };
            cboCapNum = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 96, Margin = new Padding(6, 3, 0, 0) };
            cboCapNum.Items.AddRange(new object[] { "число", "буквы ру", "буквы англ" });
            cboCapDot = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 78, Margin = new Padding(6, 3, 0, 0) };
            cboCapDot.Items.AddRange(new object[] { "с точкой", "без точки" });
            txtCapPreview = new TextBox
            {
                Width = 172,
                ReadOnly = true,
                TabStop = false,
                BackColor = SystemColors.Window,
                Cursor = Cursors.Default,
                Margin = new Padding(6, 3, 0, 0)
            };
            capRow.Controls.AddRange(new Control[] { lblCap, txtCapWord, cboCapNum, txtCapPreview, cboCapDot });

            status = new Label
            {
                Dock = DockStyle.Bottom,
                Height = 22,
                Text = "Готов",
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };

            // порядок в Controls задаёт порядок размещения (WinForms идёт с конца):
            // строка -> строка проверок -> подпись -> поле ввода -> статус
            bottom.Controls.Add(input);
            bottom.Controls.Add(capRow);
            bottom.Controls.Add(row2);
            bottom.Controls.Add(row);
            bottom.Controls.Add(status);

            Controls.Add(log);
            Controls.Add(bottom);

            // пересчёт высоты строк/нижней панели при изменении ширины:
            // если кнопки не влезают — строка занимает две линии, панель растёт
            this.Resize += (s, e) => FitRow();
            this.HandleCreated += (s, e) => FitRow();
            this.VisibleChanged += (s, e) => FitRow();
            row.Resize += (s, e) => FitRow();
            row2.Resize += (s, e) => FitRow();
            capRow.Resize += (s, e) => FitRow();

            EventHandler optsChanged = (s, e) => { UpdateCaptionPreview(); SaveOpts(); };
            txtCapWord.TextChanged += optsChanged;
            cboCapNum.TextChanged += optsChanged;
            cboCapNum.SelectedIndexChanged += optsChanged;
            cboCapDot.TextChanged += optsChanged;
            cboCapDot.SelectedIndexChanged += optsChanged;
            chkSvg.CheckedChanged += optsChanged;
            chkWeb.CheckedChanged += optsChanged;

            LoadOpts();

            input.KeyDown += (s, e) =>
            {
                if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Send(); }
            };

            AppendInfo("Чат с ИИ-агентом. Агент сам читает и правит открытый документ с помощью инструментов: " +
                       "текст, выделение, замена абзацев, поиск-замена, форматирование, параметры страницы, " +
                       "сохранение, шаблоны оформления, изображения (SVG и из интернета), подбор и упорядочивание " +
                       "источников. Проверки и работа с источниками — кнопками ниже, формат подписи рисунков — " +
                       "в строке «Формат подписи рисунков».");
        }

        private bool fitGuard;

        /// <summary>Высота одного ряда, если он переносится на две линии.</summary>
        private static int MeasureRow(FlowLayoutPanel p)
        {
            if (p == null) return 0;
            int avail = p.ClientSize.Width;
            if (avail <= 40) return p.Height;
            int x = 0, total = 0, lineH = 0;
            foreach (Control c in p.Controls)
            {
                if (!c.Visible) continue;
                int cw = c.AutoSize ? c.PreferredSize.Width : c.Width;
                int ch = c.AutoSize ? c.PreferredSize.Height : c.Height;
                cw += c.Margin.Horizontal; ch += c.Margin.Vertical;
                if (x > 0 && x + cw > avail) { total += lineH; x = cw; lineH = ch; }
                else { x += cw; if (ch > lineH) lineH = ch; }
            }
            total += lineH + p.Padding.Vertical;
            return Math.Max(28, total);
        }

        private static int RowH(Control c) { return (c == null || !c.Visible) ? 0 : c.Height; }

        /// <summary>Пересчитывает высоты строк и нижней панели под текущую ширину:
        /// если содержимое не влезает в одну линию, строка занимает две (WrapContents),
        /// а панель растёт — текст кнопок никогда не обрезается.</summary>
        private void FitRow()
        {
            if (fitGuard || bottomPanel == null) return;
            fitGuard = true;
            try
            {
                foreach (var p in new[] { row, row2, capRow })
                {
                    if (p == null || !p.Visible) continue;
                    int need = MeasureRow(p);
                    if (p.Height != need) p.Height = need;
                }
                int botH = 74 + RowH(row) + RowH(row2) + RowH(capRow) + 22 + 8;
                if (bottomPanel.Height != botH) bottomPanel.Height = botH;
            }
            finally { fitGuard = false; }
        }

        // =========================================================== вывод

        private void Append(string who, string text, Color color)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke((Action)(() => Append(who, text, color))); return; }
            log.SelectionStart = log.TextLength;
            log.SelectionLength = 0;
            log.SelectionColor = Color.DimGray;
            log.AppendText("[" + DateTime.Now.ToString("HH:mm:ss") + "] ");
            log.SelectionColor = color;
            log.AppendText(who + text + "\n");
            log.SelectionColor = log.ForeColor;
            log.ScrollToCaret();
        }

        private void AppendInfo(string text) { Append("Агент: ", text, Color.DimGray); }
        private void AppendUser(string text) { Append("Вы: ", text, Color.DarkBlue); }
        private void AppendBot(string text) { Append("DeepSeek: ", text, Color.Black); }
        private void AppendTool(string text) { Append("⚙ ", text, Color.DarkOrange); }
        private void AppendErr(string text) { Append("Ошибка: ", text, Color.Firebrick); }

        private void SetStatus(string text, bool error = false)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke((Action)(() => SetStatus(text, error))); return; }
            status.Text = text;
            status.ForeColor = error ? Color.Firebrick : Color.DimGray;
            BusyForm.UpdateStatus(text);
        }

        private void SetBusy(bool b)
        {
            if (IsDisposed) return;
            if (InvokeRequired) { BeginInvoke((Action)(() => SetBusy(b))); return; }
            busy = b;
            btnSend.Enabled = !b;
            btnNew.Enabled = !b;
            btnSources.Enabled = !b;
            btnOrdMention.Enabled = !b;
            btnOrdAlpha.Enabled = !b;
            if (btnAiCheck.Visible) btnAiCheck.Enabled = !b;
            if (btnFactCheck.Visible) btnFactCheck.Enabled = !b;
            UseWaitCursor = b;
            if (b) BusyForm.Open(FindForm(), status.Text);
            else BusyForm.HideWindow();
        }

        // =========================================================== настройки вкладки

        private bool loadingOpts;

        private static bool SBool(Dictionary<string, object> s, string k, bool def)
        {
            if (s == null || !s.ContainsKey(k) || s[k] == null) return def;
            try { return Convert.ToBoolean(s[k]); } catch { return def; }
        }

        private static string SStr(Dictionary<string, object> s, string k, string def)
        {
            if (s == null || !s.ContainsKey(k) || s[k] == null) return def;
            string v = Convert.ToString(s[k]);
            return string.IsNullOrWhiteSpace(v) ? def : v;
        }

        private void LoadOpts()
        {
            loadingOpts = true;
            try
            {
                var s = Store.LoadSettings();
                chkSvg.Checked = SBool(s, "svgImages", false);
                chkWeb.Checked = SBool(s, "webImages", false);
                txtCapWord.Text = SStr(s, "figPrefix", "Рисунок");
                string num = SStr(s, "figNumber", "число");
                if (!cboCapNum.Items.Contains(num)) cboCapNum.SelectedIndex = 0;
                else cboCapNum.SelectedItem = num;
                cboCapDot.SelectedItem = SBool(s, "figDot", true) ? "с точкой" : "без точки";
                if (cboCapDot.SelectedItem == null) cboCapDot.SelectedIndex = 0;
                UpdateCaptionPreview();
            }
            finally { loadingOpts = false; }
        }

        private void SaveOpts()
        {
            if (loadingOpts || IsDisposed) return;
            try
            {
                var s = Store.LoadSettings();
                s["svgImages"] = chkSvg.Checked;
                s["webImages"] = chkWeb.Checked;
                s["figPrefix"] = (txtCapWord.Text ?? "").Trim();
                s["figNumber"] = cboCapNum.Text.Length > 0 ? cboCapNum.Text : "число";
                s["figDot"] = cboCapDot.Text != "без точки";
                Store.SaveSettings(s);
            }
            catch (Exception ex) { Logger.Log("SaveOpts: " + ex.Message); }
            AgentTools.AllowSvgImages = chkSvg.Checked;
            AgentTools.AllowWebImages = chkWeb.Checked;
        }

        private void UpdateCaptionPreview()
        {
            if (txtCapPreview == null) return;
            txtCapPreview.Text = ImageOps.FigurePreview(
                (txtCapWord.Text ?? "").Trim(),
                cboCapNum.Text.Length > 0 ? cboCapNum.Text : "число",
                cboCapDot.Text != "без точки");
        }

        /// <summary>Формат подписи одной строкой — для промптов.</summary>
        public string CaptionSample()
        {
            return ImageOps.FigurePreview(
                (txtCapWord.Text ?? "").Trim(),
                cboCapNum.Text.Length > 0 ? cboCapNum.Text : "число",
                cboCapDot.Text != "без точки");
        }

        // =========================================================== кнопки источников

        private void FindSources()
        {
            if (busy) return;
            string q = (input.Text ?? "").Trim();
            SetBusy(true);
            Task.Run(() =>
            {
                try
                {
                    AgentTools.WordApp = app;
                    AgentTools.Bridge = bridge;
                    if (!bridge.EnsureRunning()) { AppendErr("не удалось запустить локальный мост DeepSeek."); return; }
                    SetStatus("Ищу источники для абзацев…");
                    AppendTool("find_sources(" + (q.Length > 0 ? q : "по содержанию документа") + ")");
                    string res = AgentTools.RunFindSources(app, bridge, q);
                    if (res.StartsWith("ОШИБКА")) AppendErr(res); else AppendInfo(res);
                    SetStatus("Готово");
                }
                catch (Exception ex) { Logger.Log("FindSources: " + ex); AppendErr(ex.Message); SetStatus("Ошибка", true); }
                finally { SetBusy(false); }
            });
        }

        private void ReorderSources(bool byMention)
        {
            if (busy) return;
            SetBusy(true);
            Task.Run(() =>
            {
                try
                {
                    AgentTools.WordApp = app;
                    SetStatus(byMention ? "Переставляю источники по порядку упоминания…" : "Сортирую источники по алфавиту…");
                    string res = AgentTools.RunReorderSources(app, byMention);
                    if (res.StartsWith("ОШИБКА")) AppendErr(res); else AppendInfo(res);
                    SetStatus("Готово");
                }
                catch (Exception ex) { Logger.Log("ReorderSources: " + ex); AppendErr(ex.Message); SetStatus("Ошибка", true); }
                finally { SetBusy(false); }
            });
        }

        // =========================================================== системный промпт агента

        /// <summary>Системное сообщение агента: правила + текущие галочки (обновляется перед каждым запросом).</summary>
        private void EnsureSystemMessage()
        {
            var s = Store.LoadSettings();
            bool svg = SBool(s, "svgImages", false);
            bool web = SBool(s, "webImages", false);
            AgentTools.AllowSvgImages = svg;
            AgentTools.AllowWebImages = web;

            var sb = new StringBuilder();
            sb.AppendLine("Ты — ИИ-агент внутри Microsoft Word. Ты работаешь с открытым документом через инструменты " +
                          "(function calling): чтение текста, правка абзацев и выделения, поиск-замена, форматирование, " +
                          "параметры страницы, сохранение, шаблоны оформления, изображения и источники.");
            sb.AppendLine("ПРАВИЛА:");
            sb.AppendLine("1. Отвечай пользователю по-русски, коротко и по делу.");
            sb.AppendLine("2. Не сообщай о правках, которых ты не сделал — сначала вызови инструмент, затем подведи итог.");
            sb.AppendLine("3. Перед большими правками можно прочитать документ (get_document_info / get_document_text).");
            sb.AppendLine("4. Текст верни без markdown-обёрток и пояснений от себя.");
            sb.AppendLine();
            sb.AppendLine("ИСТОЧНИКИ: find_sources подбирает реальные источники, ставит ссылки [1], [2] в тексте и " +
                          "создаёт/обновляет раздел «Список литературы»; reorder_sources перенумеровывает список и " +
                          "ссылки в тексте (mode=mention — по порядку упоминания, mode=alphabet — по алфавиту).");
            sb.AppendLine(Prompts.ImageRules(svg, web, CaptionSample()));
            string text = sb.ToString();

            if (messages.Count > 0 && messages[0].ContainsKey("role") && (string)messages[0]["role"] == "system")
            {
                messages[0]["content"] = text;
            }
            else
            {
                messages.Insert(0, new Dictionary<string, object> { { "role", "system" }, { "content", text } });
            }
        }

        // =========================================================== чат

        private void NewChat()
        {
            if (busy) return;
            messages.Clear();
            log.Clear();
            AppendInfo("Новый чат начат.");
            SetStatus("Готов");
        }

        private void Send()
        {
            if (busy) return;
            string text = (input.Text ?? "").Trim();
            if (text.Length == 0) { SetStatus("Введите сообщение", true); return; }

            input.Clear();
            AppendUser(text);
            EnsureSystemMessage();
            messages.Add(new Dictionary<string, object> { { "role", "user" }, { "content", text } });

            bool thinking = chkThinking.Checked, search = chkSearch.Checked;
            SetBusy(true);
            Task.Run(() => RunAgentLoop(thinking, search));
        }

        /// <summary>
        /// Цикл агента: запрос -> если модель просит вызвать инструмент, выполняем его
        /// локально и возвращаем результат модели, повторяем до финального ответа.
        /// </summary>
        private void RunAgentLoop(bool thinking, bool search)
        {
            try
            {
                AgentTools.WordApp = app;
                AgentTools.Bridge = bridge;
                if (!bridge.EnsureRunning())
                {
                    AppendErr("не удалось запустить локальный мост DeepSeek.");
                    return;
                }

                string tools = AgentTools.BuildSchemaJson();
                int timeout = 900;

                for (int step = 1; step <= MaxSteps; step++)
                {
                    SetStatus("Запрос к DeepSeek… (шаг " + step + " из " + MaxSteps + ")");
                    Logger.Log("Agent: step " + step + ", messages=" + messages.Count);

                    var resp = bridge.ChatCompletion(messages, tools, thinking, search, timeout);
                    if (resp == null) { AppendErr("мост не ответил."); return; }

                    string err = AgentProtocol.ExtractError(resp);
                    if (err != null) { AppendErr(err); return; }

                    var choice = AgentProtocol.FirstChoice(resp);
                    if (choice == null) { AppendErr("неожиданный формат ответа моста."); return; }

                    object msgObj;
                    var msg = (choice.TryGetValue("message", out msgObj) && msgObj is Dictionary<string, object>)
                        ? (Dictionary<string, object>)msgObj : new Dictionary<string, object>();

                    string content = AgentProtocol.GetMessageContent(choice);

                    // --- tool_calls? ---
                    var calls = AgentProtocol.GetToolCalls(msg);
                    if (calls != null && calls.Count > 0)
                    {
                        // сохраняем сообщение модели в истории (OpenAI-формат)
                        messages.Add(AgentProtocol.BuildAssistantToolCallMessage(content, calls));

                        foreach (var call in calls)
                        {
                            string name = call.Key;
                            string args = call.Value;
                            AppendTool(name + "(" + Truncate(args, 300) + ")");
                            SetStatus("Выполняю инструмент: " + name + "…");

                            string result = AgentTools.Execute(name, args, app);
                            AppendTool("→ " + Truncate((result ?? "").Replace("\r", " ").Replace("\n", " "), 300));

                            messages.Add(AgentProtocol.BuildToolResultMessage(name, result));
                        }
                        continue;   // возвращаемся к модели с результатами
                    }

                    // --- финальный текстовый ответ ---
                    if (string.IsNullOrWhiteSpace(content)) content = "(пустой ответ)";
                    AppendBot(content.Trim());
                    messages.Add(new Dictionary<string, object> { { "role", "assistant" }, { "content", content } });
                    SetStatus("Готово (шагов агента: " + step + ")");
                    return;
                }
                AppendErr("достигнут лимит вызовов инструментов (" + MaxSteps + ").");
                SetStatus("Лимит шагов агента", true);
            }
            catch (Exception ex)
            {
                Logger.Log("Agent: " + ex);
                AppendErr(ex.Message);
                SetStatus("Ошибка", true);
            }
            finally
            {
                SetBusy(false);
            }
        }

        // =========================================================== разбор ответа

        private static string Truncate(string s, int max)
        {
            if (string.IsNullOrEmpty(s)) return "";
            return s.Length <= max ? s : s.Substring(0, max) + "…";
        }
    }
}
