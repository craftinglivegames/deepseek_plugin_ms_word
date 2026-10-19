using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordAIAssistant
{
    /// <summary>
    /// Окно «Выполнение запроса...» — показывается, пока идёт запрос к ИИ (обе вкладки).
    /// Поясняет, почему работа идёт медленнее (живой режим = защита аккаунта от
    /// блокировки за подозрительную активность), показывает текущий статус и таймер
    /// и даёт живой режим включить/выключить прямо во время ожидания.
    ///
    /// Чтобы не ломать E2E-тесты (перед каждым кликом проверяется WindowFromPoint):
    /// окно НЕ перехватывает фокус (ShowWithoutActivation) и размещается СЛЕВА от панели,
    /// т.е. никогда не перекрывает её прямоугольник.
    /// </summary>
    public class BusyForm : Form
    {
        private const string Note =
            "Чтобы чат не блокировал аккаунт за подозрительную активность, надстройка симулирует " +
            "живость запросов — из-за чего скорость работы медленнее. Ожидайте терпеливо!";

        private static BusyForm inst;

        private readonly Label lblTitle, lblStatus, lblNote;
        private readonly CheckBox chkHuman;
        private readonly Button btnHide;
        private readonly Timer tmr;
        private DateTime started;
        private int dots;
        private bool loading;
        private Control ownerCtrl;

        public BusyForm()
        {
            Text = "Выполнение запроса...";
            FormBorderStyle = FormBorderStyle.FixedToolWindow;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            Font = new Font("Segoe UI", 9f);
            ClientSize = new Size(440, 200);

            lblTitle = new Label
            {
                AutoSize = true,
                Font = new Font(Font, FontStyle.Bold),
                ForeColor = Color.DarkBlue,
                Location = new Point(14, 9),
                Text = "Выполнение запроса..."
            };

            lblStatus = new Label
            {
                AutoSize = false,
                Size = new Size(412, 32),
                Location = new Point(14, 33),
                ForeColor = Color.DimGray,
                AutoEllipsis = true,
                Text = ""
            };

            lblNote = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(412, 0),
                Location = new Point(14, 69),
                ForeColor = Color.FromArgb(70, 70, 70),
                Text = Note
            };

            chkHuman = new CheckBox
            {
                AutoSize = true,
                Location = new Point(14, 134),
                Text = "Живые запросы (темп ввода и паузы)"
            };
            chkHuman.CheckedChanged += (s, e) =>
            {
                if (loading) return;
                Store.HumanMode = chkHuman.Checked;
            };

            btnHide = new Button
            {
                Size = new Size(94, 27),
                Location = new Point(332, 164),
                Text = "Скрыть"
            };
            btnHide.Click += (s, e) => HideWindow();

            Controls.Add(lblTitle);
            Controls.Add(lblStatus);
            Controls.Add(lblNote);
            Controls.Add(chkHuman);
            Controls.Add(btnHide);

            // окно живёт столько же, сколько и надстройка — держим галочку в согласии
            Store.HumanModeChanged += OnHumanModeChanged;

            tmr = new Timer { Interval = 500 };
            tmr.Tick += (s, e) => UpdateClock();
        }

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                tmr.Stop();
                tmr.Dispose();
                Store.HumanModeChanged -= OnHumanModeChanged;
            }
            base.Dispose(disposing);
        }

        private void OnHumanModeChanged(bool v)
        {
            if (IsDisposed) return;
            if (InvokeRequired)
            {
                try { BeginInvoke((Action)(() => OnHumanModeChanged(v))); } catch { }
                return;
            }
            loading = true;
            try { if (chkHuman.Checked != v) chkHuman.Checked = v; }
            finally { loading = false; }
        }

        // ====================================================== публичный API

        /// <summary>Показывает (или лишь обновляет) окно для текущего запроса.</summary>
        public static void Open(Control owner, string status)
        {
            try
            {
                if (inst == null || inst.IsDisposed) inst = new BusyForm();
                inst.Loaded(owner, status);
                if (!inst.Visible) inst.Show();
                inst.BringToFront();
            }
            catch (Exception ex) { Logger.Log("BusyForm.Open: " + ex.Message); }
        }

        /// <summary>Обновляет строку статуса (шаг, что делает мост и т.п.).</summary>
        public static void UpdateStatus(string text)
        {
            try
            {
                if (inst == null || !inst.Visible || inst.IsDisposed) return;
                inst.SetStatusText(text);
            }
            catch (Exception ex) { Logger.Log("BusyForm.Status: " + ex.Message); }
        }

        /// <summary>Прячет окно — запрос завершён.</summary>
        public static void HideWindow()
        {
            try
            {
                if (inst == null || !inst.Visible || inst.IsDisposed) return;
                inst.tmr.Stop();
                inst.Hide();
            }
            catch (Exception ex) { Logger.Log("BusyForm.Hide: " + ex.Message); }
        }

        /// <summary>Закрывает окно при выгрузке надстройки.</summary>
        public static void Shutdown()
        {
            try
            {
                if (inst == null) return;
                inst.tmr.Stop();
                inst.Close();
                inst.Dispose();
            }
            catch (Exception ex) { Logger.Log("BusyForm.Shutdown: " + ex.Message); }
            finally { inst = null; }
        }

        // ====================================================== внутреннее

        private void Loaded(Control owner, string status)
        {
            ownerCtrl = owner;
            loading = true;
            try
            {
                chkHuman.Checked = Store.HumanMode;
                if (string.IsNullOrEmpty(status) || status.Trim() == "Готов" || status.Trim() == "Готово")
                    lblStatus.Text = "Отправляю запрос...";
                else
                    lblStatus.Text = status;
            }
            finally { loading = false; }

            PlaceNextTo(owner);
            dots = 0;
            started = DateTime.Now;
            UpdateClock();
            tmr.Start();
        }

        private void SetStatusText(string text)
        {
            if (string.IsNullOrEmpty(text)) return;
            lblStatus.Text = text;
        }

        private void UpdateClock()
        {
            int sec = (int)(DateTime.Now - started).TotalSeconds;
            string d = new string('.', 1 + dots % 3);
            dots++;
            lblTitle.Text = "Выполнение запроса" + d + "  " +
                            (sec / 60).ToString("00") + ":" + (sec % 60).ToString("00");
            KeepClear();
        }

        /// <summary>Панель умеет сама ездить за окном Word — если она наехала на наше
        /// окно, возвращаем его на место (иначе E2E-клик по панели промахнётся).</summary>
        private void KeepClear()
        {
            if (ownerCtrl == null || ownerCtrl.IsDisposed) return;
            Form f = ownerCtrl as Form;
            if (f == null) f = ownerCtrl.FindForm();
            if (f == null || f.IsDisposed || !f.IsHandleCreated) return;
            if (Bounds.IntersectsWith(f.Bounds)) PlaceNextTo(ownerCtrl);
        }

        /// <summary>
        /// Ставим окно слева от панели (панель занимает правый край экрана) — так оно
        /// гарантированно не перекрывает контролы, по которым кликают тесты.
        /// </summary>
        private void PlaceNextTo(Control owner)
        {
            Rectangle wa = Screen.PrimaryScreen.WorkingArea;
            Rectangle or = Rectangle.Empty;
            Form f = owner as Form;
            if (f == null && owner != null) f = owner.FindForm();
            if (f != null && f.IsHandleCreated && !f.IsDisposed) or = f.Bounds;

            int x, y;
            if (or.Width > 0)
            {
                x = or.Left - Width - 8;                 // слева от панели
                if (x < wa.Left + 4)                      // не влезло — справа
                {
                    x = or.Right + 8;
                    if (x + Width > wa.Right - 4) x = wa.Left + 4;
                }
                y = or.Top + Math.Max(24, (or.Height - Height) / 2);
            }
            else
            {
                x = wa.Left + 4;
                y = wa.Bottom - Height - 4;
            }

            // последняя страховка: окно не должно заходить за границы экрана
            if (x + Width > wa.Right - 4) x = wa.Right - Width - 4;
            if (x < wa.Left + 4) x = wa.Left + 4;
            if (y + Height > wa.Bottom - 4) y = wa.Bottom - Height - 4;
            if (y < wa.Top + 4) y = wa.Top + 4;

            // ...и не должно перекрывать саму панель (иначе E2E-клик по ней промахнётся)
            var r = new Rectangle(x, y, Width, Height);
            if (or.Width > 0 && r.IntersectsWith(or))
            {
                x = or.Left - Width - 8;
                if (x < wa.Left + 4) { x = wa.Left + 4; y = wa.Bottom - Height - 4; }
                r = new Rectangle(x, y, Width, Height);
                if (r.IntersectsWith(or)) y = Math.Max(wa.Top + 4, or.Bottom + 4);
            }

            if (x < wa.Left + 4) x = wa.Left + 4;
            if (y + Height > wa.Bottom - 4) y = wa.Bottom - Height - 4;
            if (y < wa.Top + 4) y = wa.Top + 4;

            Location = new Point(x, y);
        }
    }
}
