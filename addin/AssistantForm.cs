using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    /// <summary>
    /// Главная панель ассистента (модальное окно рядом с Word).
    /// </summary>
    public class AssistantForm : Form
    {
        private readonly Word.Application app;
        private readonly BridgeClient bridge = new BridgeClient();
        private Dictionary<string, object> settings = Store.LoadSettings();
        private Dictionary<string, List<string>> favorites = Store.LoadFavorites();
        private List<Chapter> chapters = new List<Chapter>();

        // --- контролы ---
        private Label lblStatus;
        private ComboBox cboType;
        private TextBox txtReq, txtQuery;
        private ComboBox cboFavReq, cboFavQuery, cboFavType;
        private CheckBox chkThinking, chkSearch, chkNewChat, chkToc, chkUpdateToc, chkApplyFmt, chkAutoFmtFromReq;
        private CheckBox chkHuman;                 // «Живые запросы» — живой режим ввода
        private bool loadingHuman;                 // не перезаписывать настройку при синхронизации галочек
        private Button btnAttach, btnClearImages, btnGenerate;
        private Label lblImages;
        private ListBox lstChapters;
        private TextBox txtChapterRequest;
        private Button btnInsChapter, btnReplChapter, btnReplFmt, btnRewriteChapter;
        // карта строк списка глав: rowChapter[i] — индекс главы, rowPart[i]: 0=всё, 1=заголовок, 2=текст
        private readonly List<int> rowChapter = new List<int>();
        private readonly List<int> rowPart = new List<int>();
        private Button btnInsertAll, btnReplaceAll;
        private Button btnAiDetect, btnFactCheck, btnLogin, btnFormat, btnTemplates, btnOpenBridge;
        private List<string> images = new List<string>();
        private System.Windows.Forms.Timer posTimer;
        private bool busy;

        // адаптивная раскладка: строки переносятся на вторую линию, а не обрезаются
        private readonly List<FlowLayoutPanel> rowList = new List<FlowLayoutPanel>();
        private readonly List<Control> stretchList = new List<Control>();
        private readonly Dictionary<Control, int> stretchReserve = new Dictionary<Control, int>();
        private FlowLayoutPanel root;

        public AssistantForm(Word.Application app)
        {
            this.app = app;
            InitUi();
            LoadFavsToUi();
            StartPosTimer();
        }

        // =========================================================== UI

        private static Label L(string text) { return new Label { Text = text, AutoSize = true, Padding = new Padding(0, 6, 0, 2) }; }

        /// <summary>Строка-контейнер, которая при нехватке ширины переносит содержимое на вторую линию.</summary>
        private FlowLayoutPanel Row(params Control[] ctrls)
        {
            var p = new FlowLayoutPanel { WrapContents = true, Margin = new Padding(0), Padding = new Padding(0) };
            p.Controls.AddRange(ctrls);
            rowList.Add(p);
            return p;
        }

        /// <summary>Контрол, которому подгоняем ширину под доступное пространство.</summary>
        private T Stretch<T>(T c, int reserve = 0) where T : Control
        {
            stretchList.Add(c);
            if (reserve > 0) stretchReserve[c] = reserve;
            return c;
        }

        /// <summary>Подгоняет все строки/растягиваемые контролы под текущую ширину панели.
        /// Кнопки с автоподбором ширины никогда не обрезают текст: при нехватке места
        /// строка просто переносится на следующую линию.</summary>
        private void FitAll()
        {
            if (root == null || !root.IsHandleCreated) return;
            int avail = root.ClientSize.Width - root.Padding.Horizontal;
            if (avail < 120) return;

            foreach (var c in stretchList)
            {
                int reserve;
                stretchReserve.TryGetValue(c, out reserve);
                c.Width = Math.Max(120, avail - reserve);
            }

            foreach (var r in rowList)
            {
                r.Width = avail;
                // пересчёт высоты под фактическое число линий после переноса
                int x = 0, total = 0, lineH = 0;
                foreach (Control c in r.Controls)
                {
                    int cw, ch;
                    if (c.AutoSize) { cw = c.PreferredSize.Width; ch = c.PreferredSize.Height; }
                    else { cw = c.Width; ch = c.Height; }
                    cw += c.Margin.Horizontal; ch += c.Margin.Vertical;
                    if (x > 0 && x + cw > avail) { total += lineH; x = cw; lineH = ch; }
                    else { x += cw; if (ch > lineH) lineH = ch; }
                }
                total += lineH;
                if (total > 0 && r.Height != total) r.Height = total;
            }
        }

        private void InitUi()
        {
            Text = "AI Ассистент (DeepSeek)";
            Width = 470;
            Height = Math.Min(940, Screen.PrimaryScreen.WorkingArea.Height - 20);
            FormBorderStyle = FormBorderStyle.SizableToolWindow;
            StartPosition = FormStartPosition.Manual;
            ShowInTaskbar = true;
            Font = new Font("Segoe UI", 9f);

            var tabs = new TabControl { Dock = DockStyle.Fill };
            Controls.Add(tabs);
            this.Shown += (s2, e2) => { CheckBridgeStatus(); FitAll(); };
            this.Resize += (s2, e2) => FitAll();

            var tabGen = new TabPage("ИИ-генератор");
            var tabAgent = new TabPage("ИИ-агент");
            tabs.TabPages.Add(tabGen);
            tabs.TabPages.Add(tabAgent);

            root = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                // Слева/справа 8px воздуха; сверху/снизу 0: WinForms вычитает
                // вертикальный отступ из диапазона прокрутки дважды, из-за чего
                // нижняя строка кнопок не долистывалась до конца и обрезалась.
                // Визуально разницы нет — низ и так был бы пустым.
                Padding = new Padding(8, 0, 8, 0)
            };
            tabGen.Controls.Add(root);

            // вкладка «ИИ-агент»: чат с инструментами документа
            tabAgent.Controls.Add(new AgentPanel(app, bridge, RunAiDetect, RunFactCheck));

            // --- статус ---
            lblStatus = new Label { Text = "Готов", AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(280, 0), TextAlign = ContentAlignment.MiddleLeft };
            btnOpenBridge = new Button { Text = "Мост: ?", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnOpenBridge.Click += (s, e) => Task.Run(() => { bridge.EnsureRunning(); });
            var stRow = Row(lblStatus, btnOpenBridge);
            root.Controls.Add(stRow);
            this.FormClosed += (s, e) =>
            {
                Store.HumanModeChanged -= OnHumanModeChanged;
                BusyForm.Shutdown();
            };

            // --- тип работы ---
            root.Controls.Add(L("Тип работы:"));
            cboType = Stretch(new ComboBox { Width = 250, DropDownStyle = ComboBoxStyle.DropDown }, 210);
            var btnFavT = new Button { Text = "★", AutoSize = true, Padding = new Padding(4, 0, 4, 0) };
            var btnDelT = new Button { Text = "✕", AutoSize = true, Padding = new Padding(4, 0, 4, 0) };
            cboFavType = new ComboBox { Width = 96, DropDownStyle = ComboBoxStyle.DropDownList };
            btnFavT.Click += (s, e) => AddFav("type", cboType.Text);
            btnDelT.Click += (s, e) => DelFav("type", cboFavType);
            cboFavType.SelectedIndexChanged += (s, e) => { if (cboFavType.SelectedIndex >= 0) cboType.Text = cboFavType.SelectedItem.ToString(); };
            root.Controls.Add(Row(cboType, btnFavT, btnDelT, cboFavType));

            // --- требования ---
            root.Controls.Add(L("Требования к оформлению работы:"));
            txtReq = Stretch(new TextBox { Multiline = true, Height = 70, ScrollBars = ScrollBars.Vertical });
            var btnFavR = new Button { Text = "★ Добавить в избранное", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            cboFavReq = new ComboBox { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList };
            var btnDelR = new Button { Text = "✕ Удалить", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnFavR.Click += (s, e) => AddFav("req", txtReq.Text);
            btnDelR.Click += (s, e) => DelFav("req", cboFavReq);
            cboFavReq.SelectedIndexChanged += (s, e) => { if (cboFavReq.SelectedIndex >= 0) txtReq.Text = cboFavReq.SelectedItem.ToString(); };
            root.Controls.Add(txtReq);
            root.Controls.Add(Row(btnFavR, cboFavReq, btnDelR));

            // --- запрос ---
            root.Controls.Add(L("Запрос:"));
            txtQuery = Stretch(new TextBox { Multiline = true, Height = 90, ScrollBars = ScrollBars.Vertical });
            var btnFavQ = new Button { Text = "★ Добавить в избранное", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            cboFavQuery = new ComboBox { Width = 190, DropDownStyle = ComboBoxStyle.DropDownList };
            var btnDelQ = new Button { Text = "✕ Удалить", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnFavQ.Click += (s, e) => AddFav("query", txtQuery.Text);
            btnDelQ.Click += (s, e) => DelFav("query", cboFavQuery);
            cboFavQuery.SelectedIndexChanged += (s, e) => { if (cboFavQuery.SelectedIndex >= 0) txtQuery.Text = cboFavQuery.SelectedItem.ToString(); };
            root.Controls.Add(txtQuery);
            root.Controls.Add(Row(btnFavQ, cboFavQuery, btnDelQ));

            // --- режимы ---
            chkThinking = new CheckBox { Text = "◉ Размышление", AutoSize = true, Checked = false };
            chkSearch = new CheckBox { Text = "⌕ Поиск в интернете", AutoSize = true, Checked = false };
            chkNewChat = new CheckBox { Text = "Новый чат", AutoSize = true, Checked = true };
            // живой режим: темп ввода запросов в чате. То же переключение показано в окне
            // «Выполнение запроса...» — чтобы успеть выключить его прямо во время ожидания.
            chkHuman = new CheckBox { Text = "Живые запросы", AutoSize = true, Checked = Store.HumanMode };
            chkHuman.CheckedChanged += (s, e) => { if (!loadingHuman) Store.HumanMode = chkHuman.Checked; };
            Store.HumanModeChanged += OnHumanModeChanged;
            root.Controls.Add(Row(chkThinking, chkSearch, chkNewChat, chkHuman));

            // --- картинки ---
            btnAttach = new Button { Text = "⛓ Прикрепить фото", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            lblImages = new Label { Text = "нет файлов", AutoSize = true, ForeColor = Color.DimGray, MaximumSize = new Size(180, 0) };
            btnClearImages = new Button { Text = "Очистить", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnAttach.Click += OnAttach;
            btnClearImages.Click += (s, e) => { images.Clear(); UpdateImagesLabel(); };
            root.Controls.Add(Row(btnAttach, lblImages, btnClearImages));

            // --- генерация ---
            btnGenerate = new Button { Text = "⚡ Сгенерировать работу", Height = 36, BackColor = Color.FromArgb(210, 235, 255) };
            Stretch(btnGenerate);
            btnGenerate.Click += (s, e) => Generate(false);
            root.Controls.Add(btnGenerate);

            root.Controls.Add(new Label
            {
                Text = "— Главы (по каждой можно просить правки) —",
                AutoSize = true,
                ForeColor = Color.DimGray,
                Padding = new Padding(0, 8, 0, 2)
            });

            lstChapters = Stretch(new ListBox { Height = 130 });
            root.Controls.Add(lstChapters);
            txtChapterRequest = Stretch(new TextBox());
            root.Controls.Add(txtChapterRequest);

            btnInsChapter = new Button { Text = "Вставить в курсор", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnReplChapter = new Button { Text = "Заменить выделение", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnReplFmt = new Button { Text = "Заменить (сохр. формат)", AutoSize = true, Padding = new Padding(8, 0, 8, 0) };
            btnInsChapter.Click += (s, e) => InsertChapter(false, false);
            btnReplChapter.Click += (s, e) => InsertChapter(true, false);
            btnReplFmt.Click += (s, e) => InsertChapter(true, true);
            root.Controls.Add(Row(btnInsChapter, btnReplChapter, btnReplFmt));
            btnRewriteChapter = new Button { Text = "♻ Переписать главу по запросу", Height = 30 };
            Stretch(btnRewriteChapter);
            btnRewriteChapter.Click += (s, e) => RewriteChapter();
            root.Controls.Add(btnRewriteChapter);

            root.Controls.Add(new Label { Text = "— Вся работа в документ —", AutoSize = true, ForeColor = Color.DimGray, Padding = new Padding(0, 8, 0, 2) });

            chkToc = new CheckBox { Text = "Добавить оглавление (авто)", AutoSize = true };
            chkUpdateToc = new CheckBox { Text = "Обновить существующее оглавление", AutoSize = true, Checked = true };
            chkApplyFmt = new CheckBox { Text = "Применить оформление к вставленному", AutoSize = true, Checked = true };
            chkAutoFmtFromReq = new CheckBox { Text = "Оформление брать из поля «Требования» (через ИИ)", AutoSize = true };
            root.Controls.Add(Row(chkToc, chkUpdateToc));
            root.Controls.Add(Row(chkApplyFmt, chkAutoFmtFromReq));

            btnInsertAll = new Button { Text = "Вставить всё в курсор", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
            btnReplaceAll = new Button { Text = "Заменить весь документ", AutoSize = true, Padding = new Padding(10, 0, 10, 0), BackColor = Color.FromArgb(255, 224, 224) };
            btnInsertAll.Click += (s, e) => InsertAll(false);
            btnReplaceAll.Click += (s, e) => InsertAll(true);
            root.Controls.Add(Row(btnInsertAll, btnReplaceAll));

            root.Controls.Add(new Label { Text = "— Проверки и оформление —", AutoSize = true, ForeColor = Color.DimGray, Padding = new Padding(0, 8, 0, 2) });

            btnAiDetect = new Button { Text = "✦ Проверка на ИИ", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
            btnFactCheck = new Button { Text = "☰ Факты и источники", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
            btnAiDetect.Click += (s, e) => RunAiDetect();
            btnFactCheck.Click += (s, e) => RunFactCheck();
            root.Controls.Add(Row(btnAiDetect, btnFactCheck));

            btnFormat = new Button { Text = "Оформление документа…", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
            btnTemplates = new Button { Text = "Шаблоны оформления…", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
            btnLogin = new Button { Text = "Вход в DeepSeek", AutoSize = true, Padding = new Padding(10, 0, 10, 0) };
            btnFormat.Click += (s, e) => ShowFormatDialog();
            btnTemplates.Click += (s, e) => ShowTemplates();
            btnLogin.Click += (s, e) => RunLoginFlow();
            root.Controls.Add(Row(btnFormat, btnTemplates, btnLogin));

            // первичная подгонка (до показа — по уже известной ширине)
            FitAll();
        }

        // =========================================================== Избранное

        private void LoadFavsToUi()
        {
            cboType.Items.Clear(); cboFavType.Items.Clear();
            foreach (var v in favorites["type"]) { cboType.Items.Add(v); cboFavType.Items.Add(v); }
            cboFavReq.Items.Clear();
            foreach (var v in favorites["req"]) cboFavReq.Items.Add(v);
            cboFavQuery.Items.Clear();
            foreach (var v in favorites["query"]) cboFavQuery.Items.Add(v);
        }

        private void AddFav(string key, string value)
        {
            value = (value ?? "").Trim();
            if (value.Length == 0) { SetStatus("Нечего добавлять в избранное", true); return; }
            if (!favorites[key].Contains(value)) favorites[key].Insert(0, value);
            Store.SaveFavorites(favorites);
            LoadFavsToUi();
            SetStatus("Добавлено в избранное", false);
        }

        private void DelFav(string key, ComboBox which)
        {
            if (which.SelectedIndex < 0) { SetStatus("Выберите элемент в списке избранного", true); return; }
            favorites[key].Remove(which.SelectedItem.ToString());
            Store.SaveFavorites(favorites);
            LoadFavsToUi();
            SetStatus("Удалено из избранного", false);
        }

        // =========================================================== Служебное

        private void SetStatus(string text) { SetStatus(text, false); }

        /// <summary>Держит галочку «Живые запросы» в согласии с настройкой
        /// (её могли переключить в окне «Выполнение запроса...»).</summary>
        private void OnHumanModeChanged(bool v)
        {
            if (IsDisposed || chkHuman == null) return;
            if (InvokeRequired)
            {
                try { BeginInvoke((Action)(() => OnHumanModeChanged(v))); } catch { }
                return;
            }
            loadingHuman = true;
            try { if (chkHuman.Checked != v) chkHuman.Checked = v; }
            finally { loadingHuman = false; }
        }

        private void SetStatus(string text, bool error)
        {
            if (InvokeRequired) { BeginInvoke((Action)(() => SetStatus(text, error))); return; }
            lblStatus.Text = text;
            lblStatus.ForeColor = error ? Color.Firebrick : Color.DimGray;
            BusyForm.UpdateStatus(text);
        }

        private void SetBusy(bool b, string text)
        {
            if (InvokeRequired) { BeginInvoke((Action)(() => SetBusy(b, text))); return; }
            busy = b;
            UseWaitCursor = b;
            foreach (Control c in new Control[] { btnGenerate, btnInsertAll, btnReplaceAll, btnAiDetect, btnFactCheck, btnInsChapter, btnReplChapter, btnReplFmt, btnRewriteChapter })
                c.Enabled = !b;
            if (b) BusyForm.Open(this, text ?? lblStatus.Text);
            else BusyForm.HideWindow();
            if (!b) SetStatus(text ?? "Готов", false);
        }

        private void CheckBridgeStatus()
        {
            bool ok = bridge.Ping();
            btnOpenBridge.Text = ok ? "Мост: работает" : "Мост: не запущен";
            btnOpenBridge.BackColor = ok ? Color.FromArgb(215, 245, 215) : Color.FromArgb(250, 220, 220);
            if (ok)
            {
                Task.Run(() =>
                {
                    var st = bridge.Status();
                    if (st != null)
                    {
                        bool logged = st.ContainsKey("logged_in") && (bool)st["logged_in"];
                        BeginInvoke((Action)(() => SetStatus(logged ? "DeepSeek: вход выполнен" : "DeepSeek: НЕ выполнен вход", !logged)));
                    }
                });
            }
        }

        public void ShowPanelWindow()
        {
            PositionNextToWord();
            Show();
            if (WindowState == FormWindowState.Minimized) WindowState = FormWindowState.Normal;
            Activate();
        }

        public void CloseFormSafe()
        {
            try { posTimer?.Stop(); Close(); } catch { }
        }

        private void StartPosTimer()
        {
            posTimer = new System.Windows.Forms.Timer { Interval = 700 };
            posTimer.Tick += (s, e) => PositionNextToWord();
            posTimer.Start();
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtrW")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll", EntryPoint = "SetWindowLongW")]
        private static extern IntPtr SetWindowLong32(IntPtr hWnd, int nIndex, IntPtr dwNewLong);
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowLongPtrW(IntPtr hWnd, int nIndex);

        private const int GWL_HWNDPARENT = -8;
        private IntPtr _lastOwner = IntPtr.Zero;

        /// <summary>Делает панель «принадлежащей» окну Word — всегда поверх него.</summary>
        private void MakeOwnedByWord(IntPtr wordHwnd)
        {
            if (wordHwnd == IntPtr.Zero || !IsHandleCreated) return;
            IntPtr cur = GetWindowLongPtrW(Handle, GWL_HWNDPARENT);
            if (cur == wordHwnd) { _lastOwner = wordHwnd; return; }
            try
            {
                if (IntPtr.Size == 8) SetWindowLongPtr64(Handle, GWL_HWNDPARENT, wordHwnd);
                else SetWindowLong32(Handle, GWL_HWNDPARENT, wordHwnd);
                _lastOwner = wordHwnd;
            }
            catch (Exception ex) { Logger.Log("MakeOwnedByWord: " + ex.Message); }
        }
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        private void PositionNextToWord()
        {
            try
            {
                if (!Visible) return;
                IntPtr hwnd = IntPtr.Zero;
                try
                {
                    var w = app.ActiveWindow;
                    if (w != null) hwnd = (IntPtr)w.Hwnd;
                }
                catch { }
                MakeOwnedByWord(hwnd);
                if (hwnd == IntPtr.Zero)
                {
                    foreach (var p in System.Diagnostics.Process.GetProcessesByName("WINWORD"))
                        if (p.MainWindowHandle != IntPtr.Zero) { hwnd = p.MainWindowHandle; break; }
                }
                var wa = Screen.FromHandle(Handle).WorkingArea;
                RECT r = default(RECT);
                bool ok = hwnd != IntPtr.Zero && GetWindowRect(hwnd, out r);
                if (!ok || r.Top < -20000 || r.Left < -20000 || r.Right - r.Left <= 0)
                {
                    // окно Word скрыто/свёрнуто — ставим панель у правого края рабочего стола
                    if (Math.Abs(Left - (wa.Right - Width)) > 2 || Top < wa.Top)
                    {
                        Left = wa.Right - Width;
                        Top = wa.Top + 10;
                    }
                    return;
                }
                int x = Math.Max(wa.Left, Math.Min(r.Right, wa.Right - Width));
                int y = Math.Max(wa.Top, Math.Min(r.Top, wa.Bottom - Height));
                if (Math.Abs(Left - x) > 2 || Math.Abs(Top - y) > 2)
                {
                    Left = x; Top = y;
                    Height = Math.Min(Math.Max(r.Bottom - r.Top, 400), wa.Height - 10);
                }
            }
            catch { }
        }

        private void OnAttach(object sender, EventArgs e)
        {
            using (var dlg = new OpenFileDialog { Multiselect = true, Filter = "Изображения|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp|Все файлы|*.*" })
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    images.AddRange(dlg.FileNames);
                    UpdateImagesLabel();
                }
            }
        }

        private void UpdateImagesLabel()
        {
            lblImages.Text = images.Count == 0 ? "нет файлов" : images.Count + " файл(ов)";
        }

        // =========================================================== Генерация

        private Word.Document SafeDoc()
        {
            try { return app.ActiveDocument; }
            catch { return null; }
        }

        private void Generate(bool silent)
        {
            if (busy) { MessageBox.Show("Уже выполняется запрос, подождите."); return; }
            string query = txtQuery.Text.Trim();
            if (query.Length == 0 && images.Count == 0)
            {
                MessageBox.Show("Введите запрос (или прикрепите фото с заданием).");
                return;
            }
            SetBusy(true, "Запуск моста…");
            lstChapters.Items.Clear(); rowChapter.Clear(); rowPart.Clear();
            chapters.Clear();

            string type = cboType.Text.Trim();
            string req = txtReq.Text.Trim();
            bool thinking = chkThinking.Checked, search = chkSearch.Checked, newChat = chkNewChat.Checked;
            int ctxChars = (int)Conv(settings, "contextChars", 6000);
            int timeout = (int)Conv(settings, "chatTimeout", 900);
            string[] imgs = images.ToArray();
            bool autoFmt = chkAutoFmtFromReq.Checked;

            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning())
                    {
                        BeginInvoke((Action)(() => { SetBusy(false, null); SetStatus("Мост не запустился — см. bridge.log", true); }));
                        return;
                    }
                    BeginInvoke((Action)(() => SetStatus("Отправка запроса в DeepSeek…")));

                    string docContext = "";
                    var doc = SafeDoc();
                    if (doc != null)
                    {
                        try { docContext = WordOps.GetDocumentContext(doc, ctxChars); }
                        catch { }
                    }
                    string prompt = Prompts.BuildMainPrompt(type, req, query, docContext, null);
                    var res = bridge.Chat(prompt, imgs, thinking, search, newChat, timeout);
                    if (res == null)
                    {
                        BeginInvoke((Action)(() => { SetBusy(false, null); SetStatus("Нет ответа от моста", true); }));
                        return;
                    }
                    if (res.ContainsKey("error") && (string)res["error"] == "not_logged_in")
                    {
                        BeginInvoke((Action)(() =>
                        {
                            SetBusy(false, null);
                            SetStatus("Требуется вход в DeepSeek", true);
                            if (MessageBox.Show("Выполнить вход в DeepSeek сейчас? Откроется браузер.",
                                "Вход", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                                RunLoginFlow();
                        }));
                        return;
                    }
                    bool ok = res.ContainsKey("ok") && (bool)res["ok"];
                    string text = res.ContainsKey("text") ? (res["text"] as string) : null;
                    if (!ok || string.IsNullOrWhiteSpace(text))
                    {
                        string err = res.ContainsKey("error") ? (res["error"] as string) : "unknown";
                        BeginInvoke((Action)(() => { SetBusy(false, null); SetStatus("Ошибка: " + err, true); }));
                        return;
                    }
                    if ((bool)settings["autoClean"]) text = Prompts.Clean(text);

                    var parsed = WordOps.SplitChapters(text);
                    BeginInvoke((Action)(() =>
                    {
                        lastResult = text;
                        chapters = parsed;
                        chapters = parsed;
                        RepopulateChapters();
                        SetBusy(false, "Готово: " + parsed.Count + " глав. Вставляйте в документ.");
                    }));

                    // опционально сразу извлечь параметры оформления из требований
                    if (autoFmt && !string.IsNullOrWhiteSpace(req))
                    {
                        try
                        {
                            var fr = bridge.Chat(Prompts.BuildFormatExtractPrompt(req), null, false, false, true, 300);
                            if (fr != null && fr.ContainsKey("ok") && (bool)fr["ok"])
                            {
                                var fmtSettings = ParseFmtJson((string)fr["text"]);
                                if (fmtSettings != null)
                                {
                                    foreach (var kv in fmtSettings) settings[kv.Key] = kv.Value;
                                    Store.SaveSettings(settings);
                                }
                            }
                        }
                        catch { }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log("Generate: " + ex);
                    BeginInvoke((Action)(() => { SetBusy(false, null); SetStatus("Ошибка: " + ex.Message, true); }));
                }
            });
        }

        private string lastResult = "";

        private static double Conv(Dictionary<string, object> s, string k, double def)
        {
            if (!s.ContainsKey(k)) return def;
            try { return Convert.ToDouble(s[k], System.Globalization.CultureInfo.InvariantCulture); }
            catch { return def; }
        }

        private Dictionary<string, object> ParseFmtJson(string raw)
        {
            try
            {
                int a = raw.IndexOf('{'), b = raw.LastIndexOf('}');
                if (a < 0 || b <= a) return null;
                raw = raw.Substring(a, b - a + 1);
                var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                return ser.Deserialize<Dictionary<string, object>>(raw);
            }
            catch { return null; }
        }

        // =========================================================== Вставка в документ

        private Fmt CurrentFmt()
        {
            // перечитываем настройки: шаблон мог быть выбран и из окна, и через чат (select_template)
            try { settings = Store.LoadSettings(); }
            catch (Exception ex) { Logger.Log("CurrentFmt reload: " + ex.Message); }
            return WordOps.FromSettings(settings);
        }

        /// <summary>Перезаполняет список глав: на каждую главу 3 строки — всё / заголовок / текст.</summary>
        private void RepopulateChapters()
        {
            lstChapters.Items.Clear(); rowChapter.Clear(); rowPart.Clear();
            rowChapter.Clear();
            rowPart.Clear();
            for (int i = 0; i < chapters.Count; i++)
            {
                lstChapters.Items.Add("Всё       | " + chapters[i].Title);
                rowChapter.Add(i); rowPart.Add(0);
                lstChapters.Items.Add("Заголовок | " + chapters[i].Title);
                rowChapter.Add(i); rowPart.Add(1);
                lstChapters.Items.Add("Текст     | " + chapters[i].Title);
                rowChapter.Add(i); rowPart.Add(2);
            }
        }

        private void InsertChapter(bool replaceSelection, bool preserveFormat)
        {
            if (lstChapters.SelectedIndex < 0) { MessageBox.Show("Выберите строку в списке глав."); return; }
            var ch = chapters[rowChapter[lstChapters.SelectedIndex]];
            int part = rowPart[lstChapters.SelectedIndex];
            string content = part == 1 ? WordOps.FirstLine(ch.Text)
                           : part == 2 ? WordOps.WithoutFirstLine(ch.Text)
                           : ch.Text;
            var doc = SafeDoc();
            if (doc == null) { MessageBox.Show("Нет открытого документа."); return; }
            try
            {
                Word.Range r = app.Selection.Range;
                bool hasSelection = replaceSelection && app.Selection.Type != Word.WdSelectionType.wdSelectionIP;
                if (preserveFormat)
                {
                    if (!hasSelection)
                    {
                        MessageBox.Show("Выделите фрагмент в документе — он будет заменён текстом с сохранением его форматирования.");
                        return;
                    }
                    WordOps.ReplacePreservingFormat(doc, WordOps.MdToPlain(content));
                    SetStatus("Заменено с сохранением форматирования", false);
                    return;
                }
                if (!hasSelection)
                {
                    r = doc.Range(r.Start, r.Start); // вставка в курсор
                }
                bool useFmt = chkApplyFmt.Checked;
                WordOps.InsertMarkdown(doc, r, content, CurrentFmt(), useFmt);
                SetStatus("Вставлено", false);
            }
            catch (Exception ex) { Logger.Log("InsertChapter: " + ex); MessageBox.Show("Ошибка вставки: " + ex.Message); }
        }

        private void InsertAll(bool replaceAll)
        {
            if (string.IsNullOrWhiteSpace(lastResult)) { MessageBox.Show("Сначала сгенерируйте работу."); return; }
            var doc = SafeDoc();
            if (doc == null) { MessageBox.Show("Нет открытого документа."); return; }
            if (replaceAll)
            {
                if (MessageBox.Show("Заменить ВСЁ содержимое документа сгенерированной работой?",
                    "Подтверждение", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            }
            try
            {
                bool useFmt = chkApplyFmt.Checked;
                Word.Range at;
                if (replaceAll)
                {
                    doc.Content.Delete();
                    at = doc.Range(0, 0);
                }
                else
                {
                    var sel = app.Selection.Range;
                    at = doc.Range(sel.Start, sel.Start);
                }
                WordOps.InsertMarkdown(doc, at, lastResult, CurrentFmt(), useFmt);
                if (chkToc.Checked)
                    WordOps.AddToc(doc);
                else if (chkUpdateToc.Checked)
                    WordOps.UpdateToc(doc);
                SetStatus("Работа вставлена (" + WordOps.DocStats(doc) + ")", false);
            }
            catch (Exception ex) { Logger.Log("InsertAll: " + ex); MessageBox.Show("Ошибка вставки: " + ex.Message); }
        }

        private void RewriteChapter()
        {
            if (lstChapters.SelectedIndex < 0) { MessageBox.Show("Выберите главу в списке."); return; }
            var idx = rowChapter[lstChapters.SelectedIndex];
            var ch = chapters[idx];
            string userReq = txtChapterRequest.Text.Trim();
            SetBusy(true, "Переписываю главу…");
            bool thinking = chkThinking.Checked, search = chkSearch.Checked;
            int timeout = (int)Conv(settings, "chatTimeout", 900);
            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning()) { BeginInvoke((Action)(() => SetBusy(false, "Мост не запущен"))); return; }
                    string docContext = "";
                    var doc = SafeDoc();
                    if (doc != null)
                    {
                        try { docContext = WordOps.GetDocumentContext(doc, (int)Conv(settings, "contextChars", 6000)); } catch { }
                    }
                    string prompt = Prompts.BuildChapterPrompt(ch.Title, ch.Text, userReq, docContext);
                    var res = bridge.Chat(prompt, null, thinking, search, true, timeout);
                    if (res == null || !res.ContainsKey("ok") || !(bool)res["ok"] || string.IsNullOrEmpty(res.ContainsKey("text") ? res["text"] as string : null))
                    {
                        BeginInvoke((Action)(() => SetBusy(false, "Не удалось переписать главу")));
                        return;
                    }
                    string text = (string)res["text"];
                    if ((bool)settings["autoClean"]) text = Prompts.Clean(text);
                    BeginInvoke((Action)(() =>
                    {
                        ch.Text = text;
                        var t = WordOps.SplitChapters(text);
                        if (t.Count > 0) ch.Title = t[0].Title;
                        RepopulateChapters();
                        for (int i = 0; i < rowChapter.Count; i++)
                            if (rowChapter[i] == idx) { lstChapters.SelectedIndex = i; break; }
                        SetBusy(false, "Глава обновлена — вставьте её в документ");
                    }));
                }
                catch (Exception ex)
                {
                    Logger.Log("RewriteChapter: " + ex);
                    BeginInvoke((Action)(() => { SetBusy(false, null); SetStatus("Ошибка: " + ex.Message, true); }));
                }
            });
        }

        // =========================================================== Проверки

        private void RunCheck(string kind)
        {
            var doc = SafeDoc();
            if (doc == null) { MessageBox.Show("Нет открытого документа."); return; }
            string text;
            List<WordOps.DocParagraph> paras;
            try
            {
                text = doc.Content.Text;
                if (string.IsNullOrWhiteSpace(text)) { MessageBox.Show("Документ пустой."); return; }
                if (text.Length > 30000) text = text.Substring(0, 30000) + "\n…(текст обрезан)";

                // абзацы для списка «переписать через ИИ» (красные подсказки — отдельным запросом)
                paras = WordOps.GetRewritableParagraphs(doc, 400);
                paras = LimitParagraphs(paras, 24000);
            }
            catch { return; }
            SetBusy(true, kind == "ai" ? "Анализирую на признаки ИИ…" : "Проверяю факты и источники…");
            bool search = kind == "fact";
            int timeout = (int)Conv(settings, "chatTimeout", 900);
            string payload = text;
            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning()) { BeginInvoke((Action)(() => SetBusy(false, "Мост не запущен"))); return; }
                    string prompt = kind == "ai" ? Prompts.BuildAiDetectPrompt(payload) : Prompts.BuildFactCheckPrompt(payload);
                    var res = bridge.Chat(prompt, null, false, search, true, timeout);
                    string report = res != null && res.ContainsKey("ok") && (bool)res["ok"] ? (string)res["text"] : null;
                    if (report != null && (bool)settings["autoClean"]) report = Prompts.Clean(report);
                    BeginInvoke((Action)(() =>
                    {
                        SetBusy(false, null);
                        if (report == null) { SetStatus("Проверка не удалась", true); return; }
                        var form = new CheckReportForm(
                            report,
                            kind == "ai" ? "Отчёт: проверка на ИИ" : "Отчёт: факты и источники",
                            kind, app, bridge, settings, paras);
                        form.Show();
                        form.BringToFront();
                        form.Activate();
                    }));
                }
                catch (Exception ex)
                {
                    Logger.Log("RunCheck: " + ex);
                    BeginInvoke((Action)(() => { SetBusy(false, null); SetStatus("Ошибка: " + ex.Message, true); }));
                }
            });
        }

        /// <summary>Ограничивает список абзацев так, чтобы пронумерованный текст влезал в запрос к ИИ.</summary>
        private static List<WordOps.DocParagraph> LimitParagraphs(List<WordOps.DocParagraph> src, int maxChars)
        {
            var res = new List<WordOps.DocParagraph>();
            int len = 0;
            for (int i = 0; i < src.Count; i++)
            {
                int add = src[i].Text.Length + 12;
                if (len + add > maxChars) break;
                len += add;
                res.Add(src[i]);
            }
            return res;
        }

        public void RunAiDetect() { RunCheck("ai"); }
        public void RunFactCheck() { RunCheck("fact"); }

        // =========================================================== Вход / оформление

        public void RunLoginFlow()
        {
            SetStatus("Открываю браузер для входа…");
            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning()) { BeginInvoke((Action)(() => SetStatus("Мост не запустился", true))); return; }
                    bridge.OpenLogin();
                    BeginInvoke((Action)(() => SetStatus("Войдите в аккаунт в открывшемся браузере, затем нажмите «Проверить вход».")));
                    // диалог подтверждения
                    BeginInvoke((Action)(() =>
                    {
                        if (MessageBox.Show("После входа в аккаунт DeepSeek в браузере нажмите OK, чтобы проверить вход.",
                            "Вход в DeepSeek", MessageBoxButtons.OKCancel, MessageBoxIcon.Information) == DialogResult.OK)
                        {
                            Task.Run(() =>
                            {
                                var res = bridge.ConfirmLogin();
                                bool ok = res != null && res.ContainsKey("ok") && (bool)res["ok"] && res.ContainsKey("logged_in") && (bool)res["logged_in"];
                                BeginInvoke((Action)(() =>
                                {
                                    SetStatus(ok ? "Вход выполнен ✔" : "Вход НЕ выполнен", !ok);
                                    CheckBridgeStatus();
                                }));
                            });
                        }
                    }));
                }
                catch (Exception ex) { Logger.Log("RunLoginFlow: " + ex); }
            });
        }

        /// <summary>Окно шаблонов оформления (создание через ИИ, выбор активного галочкой).</summary>
        private void ShowTemplates()
        {
            try
            {
                var doc = app != null ? app.ActiveDocument : null;
                if (doc == null)
                {
                    SetStatus("Откройте документ — шаблоны работают с ним", true);
                    return;
                }
                using (var f = new TemplatesForm(app, bridge))
                    f.ShowDialog(this);
                // шаблон мог изменить настройки оформления — перечитываем
                settings = Store.LoadSettings();
                SetStatus("Шаблоны: активный — " + ActiveTemplateLabel(), false);
            }
            catch (Exception ex)
            {
                Logger.Log("ShowTemplates: " + ex.Message);
                SetStatus("Ошибка шаблонов: " + ex.Message, true);
            }
        }

        private string ActiveTemplateLabel()
        {
            var t = Templates.GetActive();
            return t == null ? "не выбран" : "«" + t.Name + "»";
        }

        private void ShowFormatDialog()
        {
            using (var dlg = new FormatDialog(settings, txtReq.Text, bridge))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    settings = dlg.Settings;
                    Store.SaveSettings(settings);
                    if (dlg.ApplyToDocument)
                    {
                        var doc = SafeDoc();
                        if (doc != null)
                        {
                            WordOps.ApplyFormatToDoc(doc, WordOps.FromSettings(settings));
                            if (dlg.UpdateToc) WordOps.UpdateToc(doc);
                            SetStatus("Оформление применено", false);
                        }
                    }
                }
            }
        }
    }
}
