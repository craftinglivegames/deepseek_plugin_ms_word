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
    /// Окно подробного описания одного шаблона оформления (отдельное окно, как в задании).
    /// </summary>
    public class TemplateDetailsForm : Form
    {
        public TemplateDetailsForm(DocTemplate t)
        {
            Text = "Шаблон: " + (t == null ? "" : t.Name);
            Width = 560;
            Height = 520;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9f);

            var txt = new TextBox
            {
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                Font = new Font("Consolas", 10f),
                BackColor = Color.White
            };
            txt.Text = Templates.Describe(t);
            txt.SelectionStart = 0;
            txt.SelectionLength = 0;

            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 46,
                Padding = new Padding(6)
            };
            var btnClose = new Button { Text = "Закрыть", AutoSize = true, DialogResult = DialogResult.Cancel };
            bottom.Controls.Add(btnClose);

            var pad = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10) };
            pad.Controls.Add(txt);

            Controls.Add(pad);
            Controls.Add(bottom);
            CancelButton = btnClose;
        }
    }

    /// <summary>Простой диалог ввода многострочного текста (описание шаблона).</summary>
    public class TextPromptForm : Form
    {
        public string Value { get { return txt.Text; } }
        private readonly TextBox txt;

        public TextPromptForm(string title, string label, string initial)
        {
            Text = title;
            Width = 560;
            Height = 340;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9f);

            var lbl = new Label
            {
                Text = label,
                Dock = DockStyle.Top,
                Height = 42,
                Padding = new Padding(10, 8, 10, 0)
            };

            txt = new TextBox
            {
                Multiline = true,
                ScrollBars = ScrollBars.Vertical,
                Dock = DockStyle.Fill,
                Text = initial ?? "",
                Font = new Font("Segoe UI", 10f)
            };

            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.RightToLeft,
                Height = 46,
                Padding = new Padding(6)
            };
            var btnOk = new Button { Text = "Создать через ИИ", AutoSize = true, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", AutoSize = true, DialogResult = DialogResult.Cancel };
            bottom.Controls.Add(btnCancel);
            bottom.Controls.Add(btnOk);

            var pad = new Panel { Dock = DockStyle.Fill, Padding = new Padding(10, 0, 10, 0) };
            pad.Controls.Add(txt);
            Controls.Add(pad);
            Controls.Add(lbl);
            Controls.Add(bottom);

            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }
    }

    /// <summary>
    /// Окно шаблонов оформления: список сохранённых шаблонов с галочкой активного,
    /// создание через ИИ из документа или по описанию, подробности в отдельном окне.
    /// </summary>
    public class TemplatesForm : Form
    {
        private readonly Word.Application app;
        private readonly BridgeClient bridge;

        private ListView lst;
        private Label lblStatus;
        private Button btnFromDoc, btnFromText, btnDetails, btnApply, btnDelete, btnRefresh;
        private List<DocTemplate> items = new List<DocTemplate>();
        private string activeId = "";
        private bool guard;   // защита от рекурсии ItemCheck

        public TemplatesForm(Word.Application app, BridgeClient bridge)
        {
            this.app = app;
            this.bridge = bridge;
            InitUi();
            Reload();
        }

        private void InitUi()
        {
            Text = "Шаблоны оформления документа";
            Width = 880;
            Height = 540;
            StartPosition = FormStartPosition.CenterParent;
            MinimizeBox = false;
            Font = new Font("Segoe UI", 9f);

            lst = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                CheckBoxes = true,
                FullRowSelect = true,
                GridLines = true,
                HideSelection = false,
                MultiSelect = false
            };
            lst.Columns.Add("Название", 210);
            lst.Columns.Add("Описание", 300);
            lst.Columns.Add("Шрифт / кегль", 150);
            lst.Columns.Add("Источник", 95);
            lst.Columns.Add("Создан", 110);
            lst.ItemCheck += Lst_ItemCheck;
            lst.DoubleClick += (s, e) => ShowDetails();

            var top = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(8, 8, 8, 2),
                WrapContents = true
            };
            btnFromDoc = new Button { Text = "Из документа (ИИ)…", AutoSize = true, Padding = new Padding(10, 2, 10, 2) };
            btnFromText = new Button { Text = "По описанию (ИИ)…", AutoSize = true, Padding = new Padding(10, 2, 10, 2) };
            btnDetails = new Button { Text = "Подробности", AutoSize = true, Padding = new Padding(10, 2, 10, 2) };
            btnRefresh = new Button { Text = "Обновить", AutoSize = true, Padding = new Padding(10, 2, 10, 2) };
            top.Controls.Add(btnFromDoc);
            top.Controls.Add(btnFromText);
            top.Controls.Add(btnDetails);
            top.Controls.Add(btnRefresh);

            var bottom = new FlowLayoutPanel
            {
                Dock = DockStyle.Bottom,
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Padding = new Padding(8, 4, 8, 8),
                WrapContents = true
            };
            btnApply = new Button { Text = "Применить к документу", AutoSize = true, Padding = new Padding(10, 2, 10, 2) };
            btnDelete = new Button { Text = "Удалить", AutoSize = true, Padding = new Padding(10, 2, 10, 2) };
            var btnClose = new Button { Text = "Закрыть", AutoSize = true, DialogResult = DialogResult.Cancel };
            bottom.Controls.Add(btnClose);
            bottom.Controls.Add(btnDelete);
            bottom.Controls.Add(btnApply);

            lblStatus = new Label
            {
                Dock = DockStyle.Bottom,
                AutoSize = false,
                Height = 40,
                Padding = new Padding(10, 6, 10, 4),
                Text = "",
                MaximumSize = new Size(0, 40)
            };

            Controls.Add(lst);
            Controls.Add(top);
            Controls.Add(bottom);
            Controls.Add(lblStatus);

            btnFromDoc.Click += (s, e) => CreateFromDocument();
            btnFromText.Click += (s, e) => CreateFromDescription();
            btnDetails.Click += (s, e) => ShowDetails();
            btnRefresh.Click += (s, e) => Reload();
            btnApply.Click += (s, e) => ApplyToDocument();
            btnDelete.Click += (s, e) => DeleteSelected();
            CancelButton = btnClose;
        }

        // ------------------------------------------------------------ список

        private void Reload()
        {
            activeId = "";
            items = Templates.Load(out activeId);
            guard = true;
            lst.Items.Clear();
            foreach (var t in items)
            {
                var it = new ListViewItem(t.Name)
                {
                    Tag = t,
                    Checked = t.Id == activeId,
                    ToolTipText = t.Description
                };
                it.SubItems.Add(string.IsNullOrEmpty(t.Description) ? "—" : t.Description);
                it.SubItems.Add(t.Font + ", " + N0(t.Size) + " пт, интервал " + N0(t.LineSpacing));
                it.SubItems.Add(Templates.SourceLabel(t.Source));
                it.SubItems.Add(t.CreatedAt);
                lst.Items.Add(it);
            }
            guard = false;
            lblStatus.Text = items.Count == 0
                ? "Сохранённых шаблонов пока нет. Создайте шаблон из документа или по описанию."
                : "Шаблонов: " + items.Count + ". Галочкой отмечен активный шаблон" +
                  (activeId.Length > 0 ? ": " + ActiveName() : " (нет)") + ".";
            btnDetails.Enabled = btnApply.Enabled = btnDelete.Enabled = items.Count > 0;
        }

        private string ActiveName()
        {
            foreach (var t in items)
                if (t.Id == activeId) return "«" + t.Name + "»";
            return "";
        }

        private DocTemplate Selected()
        {
            if (lst.SelectedItems.Count == 0) return null;
            return lst.SelectedItems[0].Tag as DocTemplate;
        }

        private static string N0(double v) { return v.ToString("0.##"); }

        /// <summary>Галочка в списке = выбор активного шаблона (только одна).</summary>
        private void Lst_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (guard) return;
            guard = true;
            try
            {
                var t = lst.Items[e.Index].Tag as DocTemplate;
                if (t == null) return;

                if (e.NewValue == CheckState.Checked)
                {
                    // снять все остальные галочки
                    foreach (ListViewItem i in lst.Items)
                        if (i.Index != e.Index && i.Checked) i.Checked = false;
                    activeId = t.Id;
                    Templates.Save(items, t.Id);
                    Templates.ApplyToSettings(t, Store.LoadSettings());
                    lblStatus.Text = "Активный шаблон: «" + t.Name + "». Параметры применены к настройкам оформления.";
                }
                else
                {
                    activeId = "";
                    Templates.Save(items, "");
                    lblStatus.Text = "Активный шаблон не выбран.";
                }
            }
            catch (Exception ex)
            {
                Logger.Log("TemplatesForm check: " + ex.Message);
                lblStatus.Text = "Ошибка: " + ex.Message;
            }
            finally { guard = false; }
        }

        // ------------------------------------------------------------ создание через ИИ

        private void Busy(bool on, string text)
        {
            btnFromDoc.Enabled = !on;
            btnFromText.Enabled = !on;
            btnDetails.Enabled = !on && items.Count > 0;
            btnApply.Enabled = !on && items.Count > 0;
            btnDelete.Enabled = !on && items.Count > 0;
            btnRefresh.Enabled = !on;
            if (on) lblStatus.Text = text;
        }

        /// <summary>Шаблон из текущего документа: локальный анализ форматирования + ИИ (название/описание).</summary>
        private void CreateFromDocument()
        {
            Word.Document doc;
            try
            {
                doc = app.ActiveDocument;
                if (doc == null) throw new Exception("Нет открытого документа.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Не удалось открыть документ: " + ex.Message,
                                "Шаблоны", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string title = "";
            try { title = doc.Name; } catch { }
            Busy(true, "Анализирую оформление документа…");
            Task.Run(() =>
            {
                string err = null;
                DocTemplate local = null;
                DocTemplate ai = null;
                try
                {
                    local = WordOps.AnalyzeFormatting(doc);
                    if (!bridge.EnsureRunning()) err = "Мост не запущен (DeepSeek недоступен)";
                    else
                    {
                        var res = bridge.Chat(Prompts.BuildTemplateFromDocPrompt(
                            title, Templates.FormatSummary(local)), null, false, false, true, 300);
                        if (res != null && res.ContainsKey("ok") && (bool)res["ok"])
                            ai = Templates.FromJson((string)res["text"], "document");
                        else err = "DeepSeek не ответил";
                    }
                }
                catch (Exception ex) { err = ex.Message; }

                BeginInvoke((Action)(() =>
                {
                    Busy(false, "");
                    DocTemplate created;
                    if (ai != null)
                    {
                        // ИИ дал название/описание — параметры берём из локального анализа
                        created = Merge(ai, local, "document");
                    }
                    else if (local != null)
                    {
                        created = local;
                        created.Name = "Шаблон от " + DateTime.Now.ToString("dd.MM.yyyy HH:mm");
                        created.Description = "Создан по оформлению документа «" + title + "»." +
                                              (err != null ? " (ИИ недоступен: " + err + ")" : "");
                        created.Source = "document";
                    }
                    else
                    {
                        lblStatus.Text = "Не удалось создать шаблон: " + (err ?? "анализ не дал результата");
                        return;
                    }
                    SaveNew(created);
                    if (ai == null)
                        lblStatus.Text = "Шаблон создан по оформлению документа. ИИ недоступен: " + err;
                }));
            });
        }

        /// <summary>Шаблон по текстовому описанию: ИИ возвращает JSON целиком.</summary>
        private void CreateFromDescription()
        {
            using (var dlg = new TextPromptForm(
                "Создать шаблон по описанию",
                "Опишите требования к оформлению (шрифт, кегль, интервал, отступы, поля, заголовки):", ""))
            {
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                string desc = dlg.Value;
                if (string.IsNullOrWhiteSpace(desc))
                {
                    lblStatus.Text = "Описание пустое — шаблон не создан.";
                    return;
                }

                Busy(true, "Создаю шаблон через DeepSeek…");
                Task.Run(() =>
                {
                    string err = null;
                    DocTemplate tpl = null;
                    try
                    {
                        if (!bridge.EnsureRunning()) err = "Мост не запущен (DeepSeek недоступен)";
                        else
                        {
                            var res = bridge.Chat(Prompts.BuildTemplateFromDescriptionPrompt(desc),
                                                  null, false, false, true, 300);
                            if (res != null && res.ContainsKey("ok") && (bool)res["ok"])
                                tpl = Templates.FromJson((string)res["text"], "description");
                            else err = "DeepSeek не ответил";
                        }
                    }
                    catch (Exception ex) { err = ex.Message; }

                    BeginInvoke((Action)(() =>
                    {
                        Busy(false, "");
                        if (tpl == null)
                        {
                            lblStatus.Text = "Не удалось создать шаблон: " + (err ?? "неверный ответ ИИ");
                            MessageBox.Show(this,
                                "ИИ не смог разобрать описание." + (err != null ? "\n" + err : ""),
                                "Шаблоны", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        SaveNew(tpl);
                    }));
                });
            }
        }

        /// <summary>Название/описание от ИИ + параметры из локального анализа документа.</summary>
        private static DocTemplate Merge(DocTemplate ai, DocTemplate local, string source)
        {
            if (local == null) { ai.Source = source; return ai; }
            var t = local;
            t.Name = string.IsNullOrEmpty(ai.Name) ? t.Name : ai.Name;
            t.Description = ai.Description;
            t.Source = source;
            // ИИ мог уточнить кегли/выравнивание заголовков — переносим только явно заданные
            if (ai.H1Size > 0) t.H1Size = ai.H1Size;
            if (ai.H2Size > 0) t.H2Size = ai.H2Size;
            if (ai.H3Size > 0) t.H3Size = ai.H3Size;
            if (!string.IsNullOrEmpty(ai.HeadingAlign)) t.HeadingAlign = ai.HeadingAlign;
            if (!string.IsNullOrEmpty(ai.HeadingFont)) t.HeadingFont = ai.HeadingFont;
            return t;
        }

        private void SaveNew(DocTemplate t)
        {
            if (t == null) return;
            if (string.IsNullOrEmpty(t.Id)) t.Id = Templates.NewId();
            if (string.IsNullOrEmpty(t.CreatedAt)) t.CreatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
            if (string.IsNullOrEmpty(t.Name)) t.Name = "Без названия";
            items.Add(t);
            activeId = t.Id;
            Templates.Save(items, activeId);
            Templates.ApplyToSettings(t, Store.LoadSettings());
            Reload();
            SelectId(t.Id);
            lblStatus.Text = "Шаблон «" + t.Name + "» создан и выбран активным.";
            Logger.Log("TemplatesForm: создан шаблон " + t.Name + " (id " + t.Id + ", source=" + t.Source + ")");
        }

        private void SelectId(string id)
        {
            foreach (ListViewItem i in lst.Items)
            {
                var t = i.Tag as DocTemplate;
                if (t != null && t.Id == id) { i.Selected = true; i.EnsureVisible(); break; }
            }
        }

        // ------------------------------------------------------------ прочие действия

        private void ShowDetails()
        {
            var t = Selected();
            if (t == null)
            {
                lblStatus.Text = "Выберите шаблон в списке, чтобы посмотреть подробности.";
                return;
            }
            using (var f = new TemplateDetailsForm(t)) f.ShowDialog(this);
        }

        private void ApplyToDocument()
        {
            var t = Selected();
            if (t == null)
            {
                lblStatus.Text = "Выберите шаблон в списке.";
                return;
            }
            try
            {
                var doc = app.ActiveDocument;
                if (doc == null) throw new Exception("Нет открытого документа.");
                var settings = Store.LoadSettings();
                Templates.ApplyToSettings(t, settings);
                WordOps.ApplyFormatToDoc(doc, WordOps.FromSettings(settings));
                activeId = t.Id;
                Templates.Save(items, activeId);
                Reload();
                SelectId(t.Id);
                lblStatus.Text = "Шаблон «" + t.Name + "» применён к документу «" + doc.Name + "».";
            }
            catch (Exception ex)
            {
                Logger.Log("TemplatesForm apply: " + ex.Message);
                lblStatus.Text = "Не удалось применить: " + ex.Message;
            }
        }

        private void DeleteSelected()
        {
            var t = Selected();
            if (t == null)
            {
                lblStatus.Text = "Выберите шаблон в списке.";
                return;
            }
            if (MessageBox.Show(this, "Удалить шаблон «" + t.Name + "»?", "Шаблоны",
                                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            items.Remove(t);
            if (activeId == t.Id) activeId = "";
            Templates.Save(items, activeId);
            Reload();
            lblStatus.Text = "Шаблон «" + t.Name + "» удалён.";
        }
    }
}
