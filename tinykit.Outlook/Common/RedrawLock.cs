using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// Stops an Office window (and its children) from painting until disposed, then repaints it once: Outlook has no
    /// ScreenUpdating switch, so this uses WM_SETREDRAW. Released after <see cref="MaxHold"/> at the latest, so a
    /// failure between suspend and dispose cannot leave the window frozen.
    /// </summary>
    internal sealed class RedrawLock : IDisposable
    {
        private const int WM_SETREDRAW = 0x000B;
        private const uint RDW_INVALIDATE = 0x0001, RDW_ERASE = 0x0004, RDW_ALLCHILDREN = 0x0080, RDW_UPDATENOW = 0x0100, RDW_FRAME = 0x0400;
        private const int MaxHold = 3000; // ms

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool RedrawWindow(IntPtr hWnd, IntPtr rect, IntPtr region, uint flags);

        private IntPtr _hwnd;
        private readonly Timer _safety;

        private RedrawLock(IntPtr hwnd)
        {
            _hwnd = hwnd;
            SendMessage(_hwnd, WM_SETREDRAW, IntPtr.Zero, IntPtr.Zero);
            _safety = new Timer { Interval = MaxHold };
            _safety.Tick += (s, e) => Dispose();
            _safety.Start();
        }

        /// <summary>Suspends painting of the given Explorer or Inspector; null when it has no window.</summary>
        public static RedrawLock Suspend(object officeWindow)
        {
            var owner = WindowOwner.From(officeWindow);
            return owner == null ? null : new RedrawLock(owner.Handle);
        }

        public void Dispose()
        {
            if (_hwnd == IntPtr.Zero)
                return;
            _safety.Stop();
            _safety.Dispose();
            SendMessage(_hwnd, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero);
            RedrawWindow(_hwnd, IntPtr.Zero, IntPtr.Zero, RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_UPDATENOW | RDW_FRAME);
            _hwnd = IntPtr.Zero;
        }
    }
}
