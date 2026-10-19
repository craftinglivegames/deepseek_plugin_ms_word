using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace WordAIAssistant
{
    /// <summary>
    /// Точка входа COM-надстройки Word: панель ассистента (Ctrl+Alt+D), контекстное меню.
    /// ПРИМЕЧАНИЕ: кастомная лента отключена — на этой системе (Windows 11 24H2 + .NET Framework)
    /// вызов IDispatch managed-надстройки приводит к падению Word (fail-fast).
    /// Поэтому панель открывается горячей клавишей и контекстным меню.
    /// </summary>
    [ComVisible(true)]
    [Guid("F3B7A2E1-6C4D-4B8A-9E2F-51D0A7C8B3F6")]
    [ProgId("WordAIAssistant.Connect")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IDTExtensibility2))]
    public class Connect : IDTExtensibility2, IRibbonExtensibility, ICustomQueryInterface
    {
        private Word.Application app;
        private AssistantForm panel;
        private Office.CommandBarButton ctxButton;
        private HotkeyWindow hotkey;

        public CustomQueryInterfaceResult GetInterface(ref Guid iid, out IntPtr ppv)
        {
            ppv = IntPtr.Zero;
            // Word 16.0.14228 + .NET Framework: любой Invoke в managed-IDispatch роняет Word.
            // Отклоняем IDispatch — Word работает с нами через vtable, как с нативной надстройкой.
            if (iid == new Guid("00020400-0000-0000-C000-000000000046"))
                return CustomQueryInterfaceResult.Failed;
            return CustomQueryInterfaceResult.NotHandled;
        }

        public Connect() { }

        // -------------------------------------------------------- IDTExtensibility2

        public void OnConnection(object Application, int ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                app = Application as Word.Application;
                if (app == null) return;
                Logger.Log("OnConnection OK");
            }
            catch (Exception ex) { Logger.Log("OnConnection: " + ex); }
        }

        public void OnStartupComplete(ref Array custom)
        {
            try
            {
                EnsureContextMenu();
                ShowHotkey();
                ShowPanel();
                Logger.Log("Startup complete: панель показана, ПКМ и Ctrl+Alt+D активны");
            }
            catch (Exception ex) { Logger.Log("OnStartupComplete: " + ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnDisconnection(int RemoveMode, ref Array custom)
        {
            try
            {
                hotkey?.Dispose();
                if (panel != null) { panel.CloseFormSafe(); panel = null; }
                if (ctxButton != null) { ctxButton.Delete(false); ctxButton = null; }
            }
            catch { }
        }

        public void OnBeginShutdown(ref Array custom)
        {
            try
            {
                hotkey?.Dispose();
                if (panel != null) { panel.CloseFormSafe(); panel = null; }
            }
            catch { }
        }

        // -------------------------------------------------------- Лента (не используется)

        public string GetCustomUI(string RibbonID) { return ""; }

        // -------------------------------------------------------- Контекстное меню

        private void EnsureContextMenu()
        {
            if (app == null) return;
            const string tag = "WordAIAssistantEdit";
            const string caption = "Редактировать с ИИ";
            try
            {
                Office.CommandBar textBar = app.CommandBars["Text"];

                // Собираем ВСЕ варианты нашего пункта (по Tag и по подписи): кнопка могла
                // остаться от прошлой загрузки надстройки — тогда у неё нет обработчика Click
                // и пункт меню «не нажимается». Оставляем одну, остальные удаляем.
                Office.CommandBarControl found = null;
                var stale = new System.Collections.Generic.List<Office.CommandBarControl>();
                foreach (Office.CommandBarControl c in textBar.Controls)
                {
                    bool ours;
                    try { ours = c.Tag == tag || c.Caption == caption; }
                    catch { continue; }
                    if (!ours) continue;
                    if (found == null && c is Office.CommandBarButton) found = c;
                    else stale.Add(c);   // чужой тип элемента или дубликат — удаляем
                }
                foreach (var c in stale)
                    try { c.Delete(true); } catch { }

                if (found == null)
                {
                    found = textBar.Controls.Add(
                        Office.MsoControlType.msoControlButton, Type.Missing, Type.Missing, 1 /*Position*/, true /*Temporary*/);
                }

                ctxButton = (Office.CommandBarButton)found;
                ctxButton.Tag = tag;
                ctxButton.Caption = caption;
                ctxButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                ctxButton.TooltipText = "Переписать выделенный текст с помощью ИИ (также Ctrl+Alt+E)";
                ctxButton.Click -= OnCtxEditClick;
                ctxButton.Click += OnCtxEditClick;
                Logger.Log("ПКМ: пункт «" + caption + "» привязан (удалено дубликатов: " + stale.Count + ")");
            }
            catch (Exception ex) { Logger.Log("EnsureContextMenu: " + ex); }
        }

        private void OnCtxEditClick(Office.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            Logger.Log("CtxEditClick: открытие диалога");
            CancelDefault = false;
            ShowEditDialog();
        }

        // -------------------------------------------------------- Горячая клавиша и окна

        private void ShowHotkey()
        {
            if (hotkey == null)
            {
                hotkey = new HotkeyWindow();
                hotkey.HotkeyPressed += (key) =>
                {
                    if (key == "D") ShowPanel();
                    else if (key == "E") ShowEditDialog();
                };
            }
            hotkey.Register();
        }

        internal void ShowPanel()
        {
            try
            {
                if (app == null) return;
                EnsureContextMenu();   // «самоисцеление»: пункт мог не создаться при старте
                if (panel == null || panel.IsDisposed)
                    panel = new AssistantForm(app);
                panel.ShowPanelWindow();
            }
            catch (Exception ex) { Logger.Log("ShowPanel: " + ex); }
        }

        internal void ShowEditDialog()
        {
            try
            {
                if (app == null) { Logger.Log("ShowEditDialog: app == null"); return; }
                Word.Document doc = null;
                try { doc = app.ActiveDocument; } catch { }
                if (doc == null)
                {
                    Logger.Log("ShowEditDialog: нет открытого документа");
                    MessageBox.Show("Сначала откройте документ в Word.", "Редактировать с ИИ",
                                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                IntPtr hwnd = IntPtr.Zero;
                try { if (app.ActiveWindow != null) hwnd = (IntPtr)app.ActiveWindow.Hwnd; }
                catch (Exception ex) { Logger.Log("ShowEditDialog: ActiveWindow — " + ex.Message); }
                if (hwnd == IntPtr.Zero) hwnd = GetForegroundWindow();

                using (var dlg = new EditSelectionForm(app, new BridgeClient(), Store.LoadSettings()))
                {
                    if (hwnd != IntPtr.Zero)
                        dlg.ShowDialog(new Win32Window(hwnd));
                    else
                        dlg.ShowDialog();
                }
                Logger.Log("ShowEditDialog: диалог закрыт");
            }
            catch (Exception ex) { Logger.Log("ShowEditDialog: " + ex); }
        }

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();
    }

    /// <summary>Скрытое окно для глобальных горячих клавиш Ctrl+Alt+D (панель) и Ctrl+Alt+E (редактирование).</summary>
    public class HotkeyWindow : IDisposable
    {
        public event Action<string> HotkeyPressed;
        private NativeWindowEx wnd;
        private const int WM_HOTKEY = 0x0312;
        private const int ID_PANEL = 0xB00B;
        private const int ID_EDIT = 0xB00C;
        private bool registered;

        private class NativeWindowEx : NativeWindow
        {
            public Action<string> OnHotkey;
            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WM_HOTKEY)
                {
                    if (m.WParam.ToInt32() == ID_PANEL) OnHotkey?.Invoke("D");
                    else if (m.WParam.ToInt32() == ID_EDIT) OnHotkey?.Invoke("E");
                }
                base.WndProc(ref m);
            }
        }

        public void Register()
        {
            if (registered) return;
            wnd = new NativeWindowEx();
            var cp = new System.Windows.Forms.CreateParams();
            cp.Parent = new IntPtr(-3); // HWND_MESSAGE
            cp.Caption = "WordAIAssistantHotkey";
            wnd.CreateHandle(cp);
            wnd.OnHotkey = (k) => HotkeyPressed?.Invoke(k);
            // MOD_ALT=1, MOD_CONTROL=2; 'D'=0x44, 'E'=0x45 → Ctrl+Alt+D / Ctrl+Alt+E
            bool r1 = RegisterHotKey(wnd.Handle, ID_PANEL, 2 | 1, 0x44);
            bool r2 = RegisterHotKey(wnd.Handle, ID_EDIT, 2 | 1, 0x45);
            registered = r1 && r2;
            Logger.Log("Hotkeys Ctrl+Alt+D=" + r1 + ", Ctrl+Alt+E=" + r2);
        }

        public void Dispose()
        {
            try
            {
                if (registered && wnd != null && wnd.Handle != IntPtr.Zero)
                {
                    UnregisterHotKey(wnd.Handle, ID_PANEL);
                    UnregisterHotKey(wnd.Handle, ID_EDIT);
                }
            }
            catch { }
            try { wnd?.DestroyHandle(); } catch { }
            registered = false;
        }

        [DllImport("user32.dll")]
        private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint modifiers, uint vk);
        [DllImport("user32.dll")]
        private static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    }

    /// <summary>IWin32Window-обёртка для owner-окон.</summary>
    public class Win32Window : System.Windows.Forms.IWin32Window
    {
        private readonly IntPtr _h;
        public Win32Window(IntPtr h) { _h = h; }
        public IntPtr Handle { get { return _h; } }
    }
}
