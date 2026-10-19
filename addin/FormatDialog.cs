using System;
using System.Collections.Generic;
using System.Drawing;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAIAssistant
{
    /// <summary>
    /// Диалог оформления: параметры форматирования (можно извлечь ИИ из поля «Требования»).
    /// </summary>
    public class FormatDialog : Form
    {
        public Dictionary<string, object> Settings { get; private set; }
        public bool ApplyToDocument { get; private set; }
        public bool UpdateToc { get; private set; }

        private readonly string requirements;
        private readonly BridgeClient bridge;

        private TextBox txtFont;
        private NumericUpDown numSize, numSpacing, numIndent;
        private NumericUpDown numML, numMR, numMT, numMB;
        private CheckBox chkJustify;
        private CheckBox chkApply, chkToc;
        private Button btnFromReq;

        public FormatDialog(Dictionary<string, object> settings, string requirements, BridgeClient bridge)
        {
            Settings = new Dictionary<string, object>(settings);
            this.requirements = requirements;
            this.bridge = bridge;

            Text = "Оформление документа";
            Width = 440;
            Height = 470;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterParent;
            Font = new Font("Segoe UI", 9f);

            var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Padding = new Padding(10), AutoSize = true };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));

            void AddL(string t) { root.Controls.Add(new Label { Text = t, AutoSize = true, Anchor = AnchorStyles.Left }); }

            AddL("Шрифт:");
            txtFont = new TextBox { Text = S("font", "Times New Roman") };
            root.Controls.Add(txtFont);

            AddL("Размер шрифта (пт):");
            numSize = new NumericUpDown { Value = D("size", 14), Minimum = 8, Maximum = 28, DecimalPlaces = 1 };
            root.Controls.Add(numSize);

            AddL("Межстрочный интервал:");
            numSpacing = new NumericUpDown { Value = D("lineSpacing", 1.5), Minimum = 1, Maximum = 3, DecimalPlaces = 2, Increment = 0.25m };
            root.Controls.Add(numSpacing);

            AddL("Абзацный отступ (см):");
            numIndent = new NumericUpDown { Value = D("firstLineIndentCm", 1.25m), Minimum = 0, Maximum = 3, DecimalPlaces = 2, Increment = 0.05m };
            root.Controls.Add(numIndent);

            AddL("Выравнивание по ширине:");
            chkJustify = new CheckBox { Checked = B("justify", true), Text = "" };
            root.Controls.Add(chkJustify);

            AddL("Поля: левое / правое (см):");
            var mlr = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
            numML = new NumericUpDown { Value = D("marginLeftCm", 3m), Minimum = 0.5m, Maximum = 6, DecimalPlaces = 2, Width = 70 };
            numMR = new NumericUpDown { Value = D("marginRightCm", 1.5m), Minimum = 0.5m, Maximum = 6, DecimalPlaces = 2, Width = 70 };
            mlr.Controls.Add(numML); mlr.Controls.Add(numMR);
            root.Controls.Add(mlr);

            AddL("Поля: верхнее / нижнее (см):");
            var mtb = new TableLayoutPanel { ColumnCount = 2, AutoSize = true };
            numMT = new NumericUpDown { Value = D("marginTopCm", 2m), Minimum = 0.5m, Maximum = 6, DecimalPlaces = 2, Width = 70 };
            numMB = new NumericUpDown { Value = D("marginBottomCm", 2m), Minimum = 0.5m, Maximum = 6, DecimalPlaces = 2, Width = 70 };
            mtb.Controls.Add(numMT); mtb.Controls.Add(numMB);
            root.Controls.Add(mtb);

            btnFromReq = new Button { Text = "➤ Взять из «Требований» через ИИ", Width = 300, Height = 30 };
            btnFromReq.Click += (s, e) => ExtractFromRequirements();
            root.SetColumnSpan(btnFromReq, 2);
            root.Controls.Add(btnFromReq);

            chkApply = new CheckBox { Text = "Применить к текущему документу", Checked = true, AutoSize = true };
            root.SetColumnSpan(chkApply, 2);
            root.Controls.Add(chkApply);
            chkToc = new CheckBox { Text = "Обновить оглавление после применения", Checked = true, AutoSize = true };
            root.SetColumnSpan(chkToc, 2);
            root.Controls.Add(chkToc);

            var btnRow = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Fill };
            var btnOk = new Button { Text = "OK", Width = 100, DialogResult = DialogResult.OK };
            var btnCancel = new Button { Text = "Отмена", Width = 100, DialogResult = DialogResult.Cancel };
            btnOk.Click += (s, e) => SaveBack();
            btnRow.Controls.Add(btnOk);
            btnRow.Controls.Add(btnCancel);
            root.SetColumnSpan(btnRow, 2);
            root.Controls.Add(btnRow);

            Controls.Add(root);
            AcceptButton = btnOk;
            CancelButton = btnCancel;
        }

        private string S(string k, string def) { return Settings.ContainsKey(k) && Settings[k] is string ? (string)Settings[k] : def; }
        private decimal D(string k, object def)
        {
            try { return Convert.ToDecimal(Settings.ContainsKey(k) ? Settings[k] : def, System.Globalization.CultureInfo.InvariantCulture); }
            catch { try { return Convert.ToDecimal(def); } catch { return 0; } }
        }
        private bool B(string k, bool def) { return Settings.ContainsKey(k) && Settings[k] is bool ? (bool)Settings[k] : def; }

        private void SaveBack()
        {
            Settings["font"] = txtFont.Text.Trim();
            Settings["size"] = numSize.Value;
            Settings["lineSpacing"] = numSpacing.Value;
            Settings["firstLineIndentCm"] = numIndent.Value;
            Settings["justify"] = chkJustify.Checked;
            Settings["marginLeftCm"] = numML.Value;
            Settings["marginRightCm"] = numMR.Value;
            Settings["marginTopCm"] = numMT.Value;
            Settings["marginBottomCm"] = numMB.Value;
            ApplyToDocument = chkApply.Checked;
            UpdateToc = chkToc.Checked;
        }

        private void ExtractFromRequirements()
        {
            if (string.IsNullOrWhiteSpace(requirements))
            {
                MessageBox.Show("Поле «Требования к оформлению» на главной панели пустое.");
                return;
            }
            btnFromReq.Enabled = false;
            btnFromReq.Text = "Запрашиваю у DeepSeek…";
            var req = requirements;
            Task.Run(() =>
            {
                string err = null;
                Dictionary<string, object> parsed = null;
                try
                {
                    if (!bridge.EnsureRunning()) err = "Мост не запущен";
                    else
                    {
                        var res = bridge.Chat(Prompts.BuildFormatExtractPrompt(req), null, false, false, true, 300);
                        if (res != null && res.ContainsKey("ok") && (bool)res["ok"])
                        {
                            var raw = (string)res["text"];
                            int a = raw.IndexOf('{'), b = raw.LastIndexOf('}');
                            if (a >= 0 && b > a)
                            {
                                var ser = new System.Web.Script.Serialization.JavaScriptSerializer();
                                parsed = ser.Deserialize<Dictionary<string, object>>(raw.Substring(a, b - a + 1));
                            }
                        }
                        else err = "Не удалось получить ответ";
                    }
                }
                catch (Exception ex) { err = ex.Message; }
                BeginInvoke((Action)(() =>
                {
                    btnFromReq.Enabled = true;
                    btnFromReq.Text = "➤ Взять из «Требований» через ИИ";
                    if (parsed != null)
                    {
                        foreach (var kv in parsed) Settings[kv.Key] = kv.Value;
                        txtFont.Text = S("font", txtFont.Text);
                        numSize.Value = D("size", numSize.Value);
                        numSpacing.Value = D("lineSpacing", numSpacing.Value);
                        numIndent.Value = D("firstLineIndentCm", numIndent.Value);
                        chkJustify.Checked = B("justify", chkJustify.Checked);
                        numML.Value = D("marginLeftCm", numML.Value);
                        numMR.Value = D("marginRightCm", numMR.Value);
                        numMT.Value = D("marginTopCm", numMT.Value);
                        numMB.Value = D("marginBottomCm", numMB.Value);
                        Store.SaveSettings(Settings);
                        MessageBox.Show("Параметры оформления извлечены из требований.");
                    }
                    else MessageBox.Show("Ошибка извлечения: " + (err ?? "неизвестная"));
                }));
            });
        }
    }
}
