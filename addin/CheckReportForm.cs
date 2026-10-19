using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    /// <summary>
    /// Окно отчёта проверки (на ИИ / факты и источники).
    /// Слева — текст отчёта, справа — список абзацев документа с чекбоксами:
    /// красным ИИ отмечает абзацы, которые стоит переписать; поставив галочки,
    /// пользователь запускает перезапись выбранных (или одного — двойным кликом).
    /// </summary>
    public class CheckReportForm : Form
    {
        private readonly Word.Application app;
        private readonly BridgeClient bridge;
        private readonly Dictionary<string, object> settings;
        private readonly string mode;                     // "ai" | "fact"
        private readonly List<WordOps.DocParagraph> paras;

        private readonly Dictionary<int, string> reasons = new Dictionary<int, string>();
        private readonly HashSet<int> flagged = new HashSet<int>();   // номера в списке (1..N)

        private TextBox txtReport;
        private ListView lst;
        private Label lblHint, lblReason, lblStatus;
        private Button btnRewrite, btnPickRed, btnUncheck, btnAll, btnInsert, btnClose;
        private CheckBox chkThinking;

        private string docContext;
        private readonly List<int> queue = new List<int>();
        private int totalQueue, doneCount;
        private bool busy, auditRunning;

        public CheckReportForm(string report, string title, string mode, Word.Application app,
                               BridgeClient bridge, Dictionary<string, object> settings,
                               List<WordOps.DocParagraph> paragraphs)
        {
            this.app = app;
            this.bridge = bridge;
            this.settings = settings;
            this.mode = mode;
            this.paras = paragraphs ?? new List<WordOps.DocParagraph>();

            Text = title;
            Width = 1160;
            Height = 700;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9f);
            MinimumSize = new Size(860, 460);

            BuildUi(report);

            Shown += (s, e) => StartAudit();
        }

        // =========================================================== UI

        private void BuildUi(string report)
        {
            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Orientation = Orientation.Vertical
            };
            // размер/минимумы выставляем в строгом порядке: сначала ширина контейнера,
            // затем ограничения панелей, и только потом SplitterDistance (иначе InvalidOperationException)
            split.Size = new Size(1148, 640);
            split.SplitterWidth = 6;
            split.Panel1MinSize = 260;
            split.Panel2MinSize = 340;
            split.SplitterDistance = 500;
            split.FixedPanel = FixedPanel.None;

            // --- слева: отчёт ---
            txtReport = new TextBox
            {
                Multiline = true, ReadOnly = true, Dock = DockStyle.Fill,
                ScrollBars = ScrollBars.Vertical, WordWrap = true,
                Text = report ?? "", Font = new Font("Consolas", 9.5f),
                BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle
            };
            split.Panel1.Controls.Add(txtReport);

            // --- справа: абзацы с чекбоксами ---
            var right = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 6, Padding = new Padding(6, 0, 0, 0)
            };
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 48));    // 0: подсказка
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 1: кнопки выбора
            right.RowStyles.Add(new RowStyle(SizeType.Percent, 100));    // 2: список
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 32));    // 3: причина выбранного
            right.RowStyles.Add(new RowStyle(SizeType.AutoSize));        // 4: «переписать»
            right.RowStyles.Add(new RowStyle(SizeType.Absolute, 26));    // 5: статус

            lblHint = new Label
            {
                Dock = DockStyle.Fill,
                Text = "ИИ определяет абзацы, которые стоит переписать…",
                ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            right.Controls.Add(lblHint, 0, 0);

            var topRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            btnPickRed = new Button { Text = "● Выбрать отмеченные", Width = 165, Height = 27 };
            btnUncheck = new Button { Text = "Снять все галочки", Width = 135, Height = 27 };
            btnAll = new Button { Text = "Выбрать все", Width = 105, Height = 27 };
            btnPickRed.Click += (s, e) => PickFlagged();
            btnUncheck.Click += (s, e) => { foreach (ListViewItem i in lst.Items) i.Checked = false; };
            btnAll.Click += (s, e) => { foreach (ListViewItem i in lst.Items) i.Checked = true; };
            topRow.Controls.AddRange(new Control[] { btnPickRed, btnUncheck, btnAll });
            right.Controls.Add(topRow, 0, 1);

            lst = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                HideSelection = false,
                ShowItemToolTips = true,
                GridLines = false,
                HeaderStyle = ColumnHeaderStyle.Nonclickable
            };
            lst.Columns.Add("№", 42);
            lst.Columns.Add("Абзац", 560);
            lst.Resize += (s, e) =>
            {
                if (lst.Columns.Count > 1)
                    lst.Columns[1].Width = Math.Max(120, lst.ClientSize.Width - 48);
            };
            lst.DoubleClick += (s, e) => RewriteOneSelected();
            lst.SelectedIndexChanged += (s, e) => ShowSelectedReason();
            right.Controls.Add(lst, 0, 2);

            lblReason = new Label
            {
                AutoSize = false, Dock = DockStyle.Fill, Height = 30,
                Text = "Красным — подсказка ИИ. Галочка = переписать. Двойной клик — переписать один абзац.",
                ForeColor = Color.DimGray, TextAlign = ContentAlignment.MiddleLeft
            };
            right.Controls.Add(lblReason, 0, 3);

            var midRow = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Padding = new Padding(0, 4, 0, 0) };
            btnRewrite = new Button
            {
                Text = "♻ Переписать выбранные абзацы",
                Width = 215, Height = 32,
                BackColor = Color.FromArgb(210, 235, 255)
            };
            chkThinking = new CheckBox { Text = "◉ Размышление", AutoSize = true, Padding = new Padding(6, 8, 0, 0) };
            btnRewrite.Click += (s, e) => StartRewrite(GetCheckedNumbers());
            midRow.Controls.AddRange(new Control[] { btnRewrite, chkThinking });
            right.Controls.Add(midRow, 0, 4);

            lblStatus = new Label
            {
                AutoSize = false, Dock = DockStyle.Fill, Height = 24,
                Text = "Готово", ForeColor = Color.DimGray,
                TextAlign = ContentAlignment.MiddleLeft
            };
            right.Controls.Add(lblStatus, 0, 5);

            split.Panel2.Controls.Add(right);
            Controls.Add(split);

            // --- нижняя панель кнопок ---
            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom, Height = 42, FlowDirection = FlowDirection.RightToLeft,
                Padding = new Padding(6, 4, 6, 4)
            };
            btnClose = new Button { Text = "Закрыть", Width = 110, Height = 30 };
            btnInsert = new Button { Text = "Вставить отчёт в документ", Width = 210, Height = 30 };
            btnClose.Click += (s, e) => Close();
            btnInsert.Click += (s, e) => InsertReport();
            bottom.Controls.Add(btnClose);
            bottom.Controls.Add(btnInsert);
            Controls.Add(bottom);

            for (int i = 0; i < paras.Count; i++)
            {
                var it = new ListViewItem((i + 1).ToString()) { Tag = i + 1 };   // Tag = номер в списке (1..N)
                it.SubItems.Add(Snippet(paras[i].Text));
                lst.Items.Add(it);
            }
            if (paras.Count == 0)
                lblHint.Text = "Список абзацев пуст — нечего переписывать.";
        }

        private static string Snippet(string text)
        {
            if (string.IsNullOrEmpty(text)) return "";
            return text.Length > 160 ? text.Substring(0, 160) + "…" : text;
        }

        private void SetStatus(string text, bool error)
        {
            if (IsDisposed) return;
            lblStatus.Text = text;
            lblStatus.ForeColor = error ? Color.Firebrick : Color.DimGray;
        }

        private void ShowSelectedReason()
        {
            if (lst.SelectedItems.Count == 0) return;
            int n = (int)lst.SelectedItems[0].Tag;
            string r;
            if (flagged.Contains(n) && reasons.TryGetValue(n, out r))
                lblReason.Text = "⚠ ИИ: " + r;
            else if (flagged.Contains(n))
                lblReason.Text = "⚠ ИИ рекомендует переписать этот абзац.";
            else
                lblReason.Text = "Галочка = переписать этот абзац. Двойной клик — переписать один.";
        }

        private ListViewItem FindItem(int number)
        {
            foreach (ListViewItem it in lst.Items)
                if (it.Tag is int t && t == number) return it;
            return null;
        }

        private void PickFlagged()
        {
            if (flagged.Count == 0)
            {
                SetStatus(auditRunning ? "ИИ ещё анализирует абзацы…" : "ИИ не отметил ни одного абзаца", true);
                return;
            }
            foreach (ListViewItem it in lst.Items)
                it.Checked = flagged.Contains((int)it.Tag);
            SetStatus("Отмечено галочками: " + flagged.Count + " абзац(ов) — нажмите «Переписать выбранные»", false);
        }

        private List<int> GetCheckedNumbers()
        {
            var res = new List<int>();
            foreach (ListViewItem it in lst.Items)
                if (it.Checked) res.Add((int)it.Tag);
            res.Sort();
            return res;
        }

        // =========================================================== Аудит абзацев (красная подсказка ИИ)

        private void StartAudit()
        {
            if (auditRunning || paras.Count == 0) return;
            auditRunning = true;
            string numbered = BuildNumberedText();
            int timeout = SettingInt("chatTimeout", 1500);
            lblHint.Text = "ИИ анализирует абзацы (" + paras.Count + " шт.)…";
            SetStatus("Анализ абзацев…", false);

            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning())
                    {
                        SafeBeginInvoke(() => AuditDone(null, "Мост не запущен — список останется без подсказок"));
                        return;
                    }
                    var res = bridge.Chat(Prompts.BuildParagraphAuditPrompt(numbered, mode), null,
                                          false, mode == "fact", true, timeout);
                    string txt = res != null && res.ContainsKey("ok") && (bool)res["ok"]
                        ? res["text"] as string : null;
                    if (txt != null && SettingBool("autoClean", true)) txt = Prompts.Clean(txt);
                    if (string.IsNullOrWhiteSpace(txt))
                    {
                        SafeBeginInvoke(() => AuditDone(null, "Не удалось получить подсказку ИИ по абзацам"));
                        return;
                    }
                    HashSet<int> marked;
                    Dictionary<int, string> why;
                    Prompts.ParseParagraphAudit(txt, out marked, out why);
                    SafeBeginInvoke(() => AuditDone(marked, why, null));
                }
                catch (Exception ex)
                {
                    Logger.Log("CheckReportForm.audit: " + ex);
                    SafeBeginInvoke(() => AuditDone(null, "Ошибка анализа: " + ex.Message));
                }
            });
        }

        private void AuditDone(HashSet<int> marked, string error)
        {
            AuditDone(marked, new Dictionary<int, string>(), error);
        }

        private void AuditDone(HashSet<int> marked, Dictionary<int, string> why, string error)
        {
            if (IsDisposed) return;
            auditRunning = false;
            if (marked != null)
            {
                foreach (int n in marked)
                    if (n >= 1 && n <= paras.Count) flagged.Add(n);
                if (why != null)
                    foreach (var kv in why)
                        if (kv.Key >= 1 && kv.Key <= paras.Count) reasons[kv.Key] = kv.Value;

                foreach (ListViewItem it in lst.Items)
                {
                    int n = (int)it.Tag;
                    if (!flagged.Contains(n)) continue;
                    it.ForeColor = Color.Firebrick;
                    it.BackColor = Color.FromArgb(255, 244, 244);
                    string r;
                    it.ToolTipText = reasons.TryGetValue(n, out r)
                        ? "⚠ ИИ: " + r
                        : "⚠ ИИ рекомендует переписать этот абзац.";
                }
                if (flagged.Count > 0)
                    lblHint.Text = "● Красным ИИ отметило " + flagged.Count +
                                   " абзац(ов) — поставьте на них галочки и нажмите «Переписать выбранные».";
                else
                    lblHint.Text = "✅ ИИ не нашёл абзацев, требующих переписывания. Можно отметить нужные вручную.";
            }
            else
            {
                lblHint.Text = "Слева — отчёт проверки. Справа — абзацы документа для ручного выбора.";
            }
            SetStatus(error ?? ("Готово: " + paras.Count + " абзацев в списке"), error != null);
        }

        private string BuildNumberedText()
        {
            var sb = new StringBuilder();
            for (int i = 0; i < paras.Count; i++)
                sb.AppendLine((i + 1) + ". " + paras[i].Text);
            return sb.ToString();
        }

        // =========================================================== Переписывание абзацев

        private void RewriteOneSelected()
        {
            if (lst.SelectedItems.Count == 0)
            {
                SetStatus("Выберите абзац в списке", true);
                return;
            }
            StartRewrite(new List<int> { (int)lst.SelectedItems[0].Tag });
        }

        private void StartRewrite(List<int> numbers)
        {
            if (busy) { SetStatus("Дождитесь окончания текущей задачи", true); return; }
            if (numbers == null || numbers.Count == 0)
            {
                SetStatus("Отметьте галочками абзацы в списке (или двойной клик по одному)", true);
                return;
            }
            if (GetDoc() == null) { SetStatus("Нет открытого документа", true); return; }

            queue.Clear();
            queue.AddRange(numbers.Where(n => n >= 1 && n <= paras.Count).Distinct());
            if (queue.Count == 0) { SetStatus("Некорректный выбор абзацев", true); return; }

            busy = true;
            totalQueue = queue.Count;
            doneCount = 0;
            SetUiEnabled(false);
            NextRewrite();
        }

        private void NextRewrite()
        {
            if (queue.Count == 0)
            {
                busy = false;
                SetUiEnabled(true);
                SetStatus("Готово: переписано абзац(ов) — " + doneCount + ". Проверьте текст в документе.", false);
                return;
            }
            int n = queue[0];
            string text = paras[n - 1].Text;
            string reason;
            reasons.TryGetValue(n, out reason);
            SetStatus("Переписываю абзац " + n + "… (" + (doneCount + 1) + "/" + totalQueue + ")", false);

            if (docContext == null)
            {
                docContext = "";
                try
                {
                    var d = GetDoc();
                    if (d != null) docContext = WordOps.GetDocumentContext(d, SettingInt("contextChars", 6000));
                }
                catch { }
            }
            string ctx = docContext;
            bool thinking = chkThinking.Checked;
            int timeout = SettingInt("chatTimeout", 1500);
            bool clean = SettingBool("autoClean", true);

            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning())
                    {
                        Fail("Мост не запущен — перезапустите ПК или нажмите «Мост: ?» на панели");
                        return;
                    }
                    var res = bridge.Chat(Prompts.BuildParagraphRewritePrompt(mode, text, reason, ctx),
                                          null, thinking, false, true, timeout);
                    if (res == null || !res.ContainsKey("ok") || !(bool)res["ok"])
                    {
                        Fail("DeepSeek не ответил (абзац " + n + ")");
                        return;
                    }
                    string raw = res.ContainsKey("text") ? res["text"] as string : null;
                    if (string.IsNullOrWhiteSpace(raw))
                    {
                        Fail("Пустой ответ по абзацу " + n);
                        return;
                    }
                    if (clean) raw = Prompts.Clean(raw);
                    string plain = WordOps.CollapseToSingleLine(WordOps.MdToPlain(raw));
                    if (plain.Length == 0)
                    {
                        Fail("Пустой результат по абзацу " + n);
                        return;
                    }
                    SafeBeginInvoke(() => ApplyRewrite(n, plain));
                }
                catch (Exception ex)
                {
                    Logger.Log("CheckReportForm.rewrite: " + ex);
                    Fail("Ошибка: " + ex.Message);
                }
            });
        }

        private void ApplyRewrite(int n, string plain)
        {
            try
            {
                var doc = GetDoc();
                if (doc == null) { Fail("Документ закрыт — замена не выполнена"); return; }
                WordOps.ReplaceParagraphText(doc, paras[n - 1].Index, plain);

                paras[n - 1].Text = plain;
                var it = FindItem(n);
                if (it != null)
                {
                    it.SubItems[1].Text = Snippet(plain);
                    it.Checked = false;
                    it.ToolTipText = "";
                }
                if (flagged.Remove(n))
                {
                    if (it != null)
                    {
                        it.ForeColor = SystemColors.ControlText;
                        it.BackColor = SystemColors.Window;
                    }
                    reasons.Remove(n);
                }
                doneCount++;
                queue.RemoveAt(0);
                NextRewrite();
            }
            catch (Exception ex)
            {
                Logger.Log("CheckReportForm.apply: " + ex);
                Fail("Не удалось заменить абзац " + n + ": " + ex.Message);
            }
        }

        private void Fail(string message)
        {
            SafeBeginInvoke(() =>
            {
                busy = false;
                queue.Clear();
                SetUiEnabled(true);
                SetStatus(message, true);
            });
        }

        private void SetUiEnabled(bool enabled)
        {
            btnRewrite.Enabled = enabled;
            btnPickRed.Enabled = enabled;
            btnUncheck.Enabled = enabled;
            btnAll.Enabled = enabled;
            btnInsert.Enabled = enabled;
            chkThinking.Enabled = enabled;
            lst.Enabled = enabled;
            UseWaitCursor = !enabled;
        }

        // =========================================================== Служебное

        private Word.Document GetDoc()
        {
            try { return app.ActiveDocument; }
            catch { return null; }
        }

        private void SafeBeginInvoke(Action action)
        {
            try
            {
                if (IsDisposed || !IsHandleCreated) return;
                BeginInvoke(action);
            }
            catch (ObjectDisposedException) { }
            catch (InvalidOperationException) { }
        }

        private void InsertReport()
        {
            try
            {
                var doc = GetDoc();
                if (doc == null) { MessageBox.Show("Нет открытого документа."); return; }
                var r = doc.Range(doc.Content.End - 1, doc.Content.End - 1);
                r.InsertBreak(3 /*wdPageBreak*/);
                WordOps.InsertMarkdown(doc, doc.Range(doc.Content.End - 1, doc.Content.End - 1),
                                       txtReport.Text, WordOps.FromSettings(Store.LoadSettings()), true);
                Close();
            }
            catch (Exception ex) { MessageBox.Show("Ошибка: " + ex.Message); }
        }

        private int SettingInt(string key, int def)
        {
            if (settings == null || !settings.ContainsKey(key)) return def;
            try { return Convert.ToInt32(settings[key]); }
            catch { return def; }
        }

        private bool SettingBool(string key, bool def)
        {
            if (settings == null || !settings.ContainsKey(key)) return def;
            try { return Convert.ToBoolean(settings[key]); }
            catch { return def; }
        }
    }
}
