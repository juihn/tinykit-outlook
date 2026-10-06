using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// Keyboard shortcuts for add-in commands (ribbon buttons cannot have Ctrl/Alt shortcuts of their own): a keyboard
    /// hook on Outlook's UI thread only, so it sees keys only while an Outlook window has the focus. A shortcut's handler
    /// returns false to let Outlook have the key (e.g. in a window where it does not apply). Handlers run just after the
    /// hook returns, not inside it.
    /// </summary>
    internal sealed class KeyboardShortcuts : IDisposable
    {
        private const int WH_KEYBOARD = 2;
        private const int HC_ACTION = 0;

        private delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll")]
        private static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        private readonly HookProc _proc; // kept alive: the hook calls it
        private readonly Dictionary<Keys, Func<bool>> _shortcuts = new Dictionary<Keys, Func<bool>>();
        private IntPtr _hook;

        public KeyboardShortcuts()
        {
            _proc = OnKey;
            _hook = SetWindowsHookEx(WH_KEYBOARD, _proc, IntPtr.Zero, GetCurrentThreadId());
            if (_hook == IntPtr.Zero)
                Log.Info("Keyboard shortcuts unavailable: SetWindowsHookEx failed (" + Marshal.GetLastWin32Error() + ")");
        }

        /// <summary>
        /// Adds a shortcut, e.g. Keys.Control | Keys.Alt | Keys.D2. <paramref name="applies"/> says whether it applies now
        /// (checked in the hook, so keep it cheap); if it does, the key is taken and <paramref name="run"/> runs.
        /// </summary>
        public void Add(Keys keys, Func<bool> applies, Action run)
        {
            _shortcuts[keys] = () =>
            {
                if (!applies())
                    return false;
                var timer = new Timer { Interval = 1 };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    timer.Dispose();
                    try
                    {
                        run();
                    }
                    catch (Exception ex)
                    {
                        Log.Error("Shortcut " + keys, ex);
                    }
                };
                timer.Start();
                return true;
            };
        }

        /// <summary>
        /// Runs <paramref name="run"/> a moment after <paramref name="key"/> (with or without Shift, not with Ctrl or Alt)
        /// has gone through to Outlook, e.g. to look where it moved the focus; it gets whether Shift was down at the key
        /// press (it may be up by then). The key itself is not taken.
        /// </summary>
        public void After(Keys key, Action<bool> run)
        {
            _after[key] = run;
        }

        private readonly Dictionary<Keys, Action<bool>> _after = new Dictionary<Keys, Action<bool>>();

        private void RunSoon(Action run, string what)
        {
            var timer = new Timer { Interval = 50 };
            timer.Tick += (s, e) =>
            {
                timer.Stop();
                timer.Dispose();
                try
                {
                    run();
                }
                catch (Exception ex)
                {
                    Log.Error(what, ex);
                }
            };
            timer.Start();
        }

        private IntPtr OnKey(int code, IntPtr wParam, IntPtr lParam)
        {
            try
            {
                long flags = lParam.ToInt64();
                bool keyDown = (flags & 0x80000000) == 0;
                bool repeat = (flags & 0x40000000) != 0;
                Action<bool> after;
                if (code == HC_ACTION && keyDown && (Control.ModifierKeys & (Keys.Control | Keys.Alt)) == Keys.None
                    && _after.TryGetValue((Keys)wParam.ToInt32(), out after))
                {
                    var shift = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
                    RunSoon(() => after(shift), "After " + (Keys)wParam.ToInt32());
                }
                if (code == HC_ACTION && keyDown)
                {
                    var keys = (Keys)wParam.ToInt32() | Control.ModifierKeys;
                    Func<bool> handler;
                    if (_shortcuts.TryGetValue(keys, out handler))
                    {
                        if (repeat)
                            return new IntPtr(1); // held down: taken, but run once
                        if (handler())
                            return new IntPtr(1);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Keyboard hook", ex);
            }
            return CallNextHookEx(_hook, code, wParam, lParam);
        }

        public void Dispose()
        {
            if (_hook != IntPtr.Zero)
            {
                UnhookWindowsHookEx(_hook);
                _hook = IntPtr.Zero;
            }
        }
    }
}
