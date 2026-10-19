using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAIAssistant
{
    /// <summary>
    /// Диалог «Редактировать с ИИ» (ПКМ по выделенному тексту или кнопка на ленте):
    /// поля «Текущий фрагмент» и «Запрос», результат заменяет фрагмент.
    /// </summary>
    public class EditSelectionForm : Form
    {
        private readonly Word.Application app;
        private readonly BridgeClient bridge;
        private readonly Dictionary<string, object> settings;
        private TextBox txtFragment, txtRequest, txtResult;
        private CheckBox chkThinking, chkSearch;
        private Button btnRun, btnReplace, btnInsertBelow, btnClose;
        private FlowLayoutPanel btnRow2;
        private Label lblState;
        private bool busy;

        public EditSelectionForm(Word.Application app, BridgeClient bridge, Dictionary<string, object> settings)
        {
            this.app = app;
            this.bridge = bridge;
            this.settings = settings;
            Text = "Редактировать с ИИ";
            Width = 580;
            Height = 720;
            MinimumSize = new Size(460, 480);
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9f);

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 9, Padding = new Padding(8)
            };
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 0: подпись фрагмента
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 26));                    // 1: фрагмент
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 2: подпись запроса
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 3: запрос
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 4: режимы
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 5: подпись результата
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 34));                    // 6: результат
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 7: кнопка «Выполнить»
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));                       // 8: остальные кнопки

            root.Controls.Add(new Label { Text = "Текущий фрагмент:", AutoSize = true, Padding = new Padding(0, 4, 0, 2) });
            txtFragment = new TextBox
            {
                Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle
            };
            root.Controls.Add(txtFragment, 0, 1);

            root.Controls.Add(new Label { Text = "Запрос (что сделать с фрагментом):", AutoSize = true, Padding = new Padding(0, 6, 0, 2) });
            txtRequest = new TextBox
            {
                Multiline = true, Height = 70, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, BorderStyle = BorderStyle.FixedSingle
            };
            root.Controls.Add(txtRequest, 0, 3);

            var modeRow = new FlowLayoutPanel { AutoSize = true, Padding = new Padding(0, 4, 0, 2) };
            chkThinking = new CheckBox { Text = "◉ Размышление", AutoSize = true };
            chkSearch = new CheckBox { Text = "⌕ Поиск", AutoSize = true };
            modeRow.Controls.Add(chkThinking);
            modeRow.Controls.Add(chkSearch);
            root.Controls.Add(modeRow, 0, 4);

            root.Controls.Add(new Label { Text = "Результат (перед применением можно поправить):", AutoSize = true, Padding = new Padding(0, 6, 0, 2) });
            txtResult = new TextBox
            {
                Multiline = true, ReadOnly = false, ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle
            };
            root.Controls.Add(txtResult, 0, 6);

            var btnRow1 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Padding = new Padding(0, 6, 0, 0) };
            btnRun = new Button { Text = "⚡ Выполнить", Width = 160, Height = 32, BackColor = Color.FromArgb(210, 235, 255) };
            btnRun.Click += (s, e) => Run();
            lblState = new Label
            {
                Text = "Введите запрос и нажмите «Выполнить»",
                AutoSize = true, ForeColor = Color.DimGray, Padding = new Padding(10, 8, 0, 0)
            };
            btnRow1.Controls.Add(btnRun);
            btnRow1.Controls.Add(lblState);
            root.Controls.Add(btnRow1, 0, 7);

            btnRow2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Padding = new Padding(0, 6, 0, 0) };
            btnReplace = new Button { Text = "Заменить фрагмент результатом", Width = 220, Height = 32, Enabled = false };
            var btnReplaceFmt = new Button { Text = "Заменить (сохр. формат)", Width = 180, Height = 32, Enabled = false };
            btnInsertBelow = new Button { Text = "Вставить ниже", Width = 130, Height = 32, Enabled = false };
            btnClose = new Button { Text = "Закрыть", Width = 90, Height = 32 };
            btnReplace.Click += (s, e) => Apply(true);
            btnReplaceFmt.Click += (s, e) => ApplyFmt();
            btnInsertBelow.Click += (s, e) => Apply(false);
            btnClose.Click += (s, e) => Close();
            btnRow2.Controls.AddRange(new Control[] { btnReplace, btnReplaceFmt, btnInsertBelow, btnClose });
            root.Controls.Add(btnRow2, 0, 8);

            Controls.Add(root);
            Load += (s, e) => LoadFragment();
        }

        private void LoadFragment()
        {
            try
            {
                var sel = app.Selection;
                string text = sel != null ? sel.Text : null;
                if (string.IsNullOrWhiteSpace(text))
                {
                    txtFragment.Text = "(ничего не выделено — выделите текст в документе и повторите, либо введите запрос для вставки ниже)";
                }
                else
                {
                    txtFragment.Text = text.TrimEnd('\r', '\n');
                }
            }
            catch { }
        }

        private void Run()
        {
            string request = txtRequest.Text.Trim();
            if (request.Length == 0) { MessageBox.Show("Введите запрос."); return; }
            if (busy) return;
            busy = true;
            UseWaitCursor = true;
            btnRun.Enabled = false;
            lblState.Text = "DeepSeek выполняет запрос…";
            lblState.ForeColor = Color.DimGray;
            SetTexts();
        }

        private void SetTexts()
        {
            string fragment = txtFragment.Text;
            string request = txtRequest.Text.Trim();
            bool thinking = chkThinking.Checked, search = chkSearch.Checked;
            int timeout = settings.ContainsKey("chatTimeout") ? Convert.ToInt32(settings["chatTimeout"]) : 900;
            string ctx = "";
            try
            {
                var doc = app.ActiveDocument;
                if (doc != null) ctx = WordOps.GetDocumentContext(doc, settings.ContainsKey("contextChars") ? Convert.ToInt32(settings["contextChars"]) : 6000);
            }
            catch { }

            Task.Run(() =>
            {
                try
                {
                    if (!bridge.EnsureRunning())
                    {
                        Done("Мост не запущен. Запустите bridge/start_bridge.bat");
                        return;
                    }
                    var sb = new System.Text.StringBuilder();
                    sb.AppendLine(Prompts.SystemRulesRu);
                    sb.AppendLine("Сейчас — редактирование фрагмента документа. Верни ТОЛЬКО новый вариант текста (без заголовков-обёрток, без комментариев), готовый к замене исходного фрагмента.");
                    if (!string.IsNullOrWhiteSpace(ctx))
                    {
                        sb.AppendLine("--- КОНТЕКСТ ДОКУМЕНТА ---");
                        sb.AppendLine(ctx);
                        sb.AppendLine("--- КОНЕЦ КОНТЕКСТА ---");
                    }
                    sb.AppendLine("ТЕКУЩИЙ ФРАГМЕНТ:");
                    sb.AppendLine(string.IsNullOrWhiteSpace(fragment) || fragment.StartsWith("(ничего") ? "(фрагмент не выделен)" : fragment);
                    sb.AppendLine();
                    sb.AppendLine("ЗАПРОС: " + request);

                    var res = bridge.Chat(sb.ToString(), null, thinking, search, true, timeout);
                    if (res != null && res.ContainsKey("ok") && (bool)res["ok"])
                    {
                        string text = (string)res["text"];
                        if ((bool)settings["autoClean"]) text = Prompts.Clean(text);
                        Done(text);
                    }
                    else
                    {
                        Done(null);
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log("EditSelection: " + ex);
                    Done("Ошибка: " + ex.Message);
                }
            });

            void Done(string result)
            {
                bool isError = result == null ||
                               result.StartsWith("Ошибка") ||
                               result.StartsWith("Мост не");
                BeginInvoke((Action)(() =>
                {
                    busy = false;
                    UseWaitCursor = false;
                    btnRun.Enabled = true;
                    if (!isError)
                    {
                        txtResult.Text = result;
                        btnReplace.Enabled = btnInsertBelow.Enabled = true;
                        foreach (Control c in btnRow2.Controls)
                            if (c is Button b && b.Text.StartsWith("Заменить (сохр")) b.Enabled = true;
                        lblState.Text = "Готово — при желании поправьте текст и замените фрагмент";
                        lblState.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        string msg = result ?? "Не удалось получить ответ от DeepSeek.";
                        lblState.Text = msg;
                        lblState.ForeColor = Color.Firebrick;
                        MessageBox.Show(msg, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }));
            }
        }

        /// <summary>Замена выделения с сохранением его форматирования (plain-текст результата).</summary>
        private void ApplyFmt()
        {
            string result = txtResult.Text.Trim();
            if (result.Length == 0) { MessageBox.Show("Нет результата."); return; }
            try
            {
                var doc = app.ActiveDocument;
                WordOps.ReplacePreservingFormat(doc, WordOps.MdToPlain(result));
                Close();
            }
            catch (InvalidOperationException ex)
            {
                MessageBox.Show(ex.Message);
            }
            catch (Exception ex)
            {
                Logger.Log("EditSelection.ApplyFmt: " + ex);
                MessageBox.Show("Ошибка вставки: " + ex.Message);
            }
        }

        private void Apply(bool replace)
        {
            string result = txtResult.Text.Trim();
            if (result.Length == 0) { MessageBox.Show("Нет результата."); return; }
            try
            {
                var doc = app.ActiveDocument;
                var sel = app.Selection;
                if (replace && sel != null && sel.Type != Word.WdSelectionType.wdSelectionIP)
                {
                    int start = sel.Range.Start;
                    sel.Range.Delete();
                    WordOps.InsertMarkdown(doc, doc.Range(start, start), result, WordOps.FromSettings(settings), true);
                }
                else
                {
                    var r = doc.Range(sel.Range.Start, sel.Range.Start);
                    WordOps.InsertMarkdown(doc, r, result + "\r", WordOps.FromSettings(settings), true);
                }
                Close();
            }
            catch (Exception ex)
            {
                Logger.Log("EditSelection.Apply: " + ex);
                MessageBox.Show("Ошибка вставки: " + ex.Message);
            }
        }
    }
}
