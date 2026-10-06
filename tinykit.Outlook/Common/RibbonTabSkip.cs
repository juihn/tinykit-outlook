using System;
using System.Runtime.InteropServices;
using System.Text;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// Tab and Shift+Tab in the ribbon skip the Quick Filter's recent-value lists, so Tab goes from the box to the filter
    /// buttons. Ribbon XML has no tab-stop setting: a focus event hook (as screen readers use; asking the ribbon for its
    /// focused control only answers "Ribbon") notes which ribbon control gets the focus, and a moment after a Tab, when
    /// that is such a list, another Tab is sent. Such a list reports itself through MSAA as its group (a toolbar named
    /// "Quick Filter"), while the buttons report themselves; a collapsed group is a drop-down button, so it is not skipped.
    /// </summary>
    internal static class RibbonTabSkip
    {
        /// <summary>The label of the Quick Filter galleries (OutlookRibbon.AppendQuick), as UI Automation names them.</summary>
        public const string GalleryName = "Recent values";

        /// <summary>The label of the Quick Filter group (OutlookRibbon.BuildXml), the name a gallery has in MSAA focus events.</summary>
        public const string GroupName = "Quick Filter";

        private const int ROLE_SYSTEM_TOOLBAR = 0x16;

        private const uint EVENT_OBJECT_FOCUS = 0x8005;
        private const uint WINEVENT_OUTOFCONTEXT = 0;
        private const int MaxSkips = 4; // consecutive lists skipped by one key press, in case every stop were one

        private static int _skips;
        private static string _focusedName; // the ribbon control that last got the focus
        private static int _focusedRole;
        private static IntPtr _focusedWindow;
        private static WinEventProc _proc; // kept alive: the hook calls it
        private static IntPtr _hook;

        private delegate void WinEventProc(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time);

        [ComImport, Guid("618736e0-3c3d-11cf-810c-00aa00389b71"), InterfaceType(ComInterfaceType.InterfaceIsDual)]
        private interface IAccessible
        {
            [return: MarshalAs(UnmanagedType.IDispatch)] object get_accParent();
            int get_accChildCount();
            [return: MarshalAs(UnmanagedType.IDispatch)] object get_accChild(object varChild);
            string get_accName(object varChild);
            string get_accValue(object varChild);
            string get_accDescription(object varChild);
            object get_accRole(object varChild);
        }

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromEvent(IntPtr hwnd, int idObject, int idChild, out IAccessible accessible,
            [MarshalAs(UnmanagedType.Struct)] out object child);

        [DllImport("user32.dll")]
        private static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr module, WinEventProc proc, uint process,
            uint thread, uint flags);

        [DllImport("user32.dll")]
        private static extern bool UnhookWinEvent(IntPtr hook);

        [DllImport("user32.dll")]
        private static extern IntPtr GetFocus();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int count);

        [DllImport("user32.dll")]
        private static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);

        [DllImport("user32.dll")]
        private static extern short GetKeyState(int vk);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentProcessId();

        /// <summary>Starts following the focus of this (UI) thread's windows.</summary>
        public static void Start()
        {
            if (_hook != IntPtr.Zero)
                return;
            _proc = OnFocus;
            _hook = SetWinEventHook(EVENT_OBJECT_FOCUS, EVENT_OBJECT_FOCUS, IntPtr.Zero, _proc, GetCurrentProcessId(),
                GetCurrentThreadId(), WINEVENT_OUTOFCONTEXT);
            if (_hook == IntPtr.Zero)
                Log.Info("Ribbon Tab: focus events unavailable (SetWinEventHook failed)");
        }

        public static void Stop()
        {
            if (_hook != IntPtr.Zero)
            {
                UnhookWinEvent(_hook);
                _hook = IntPtr.Zero;
            }
        }

        private static void OnFocus(IntPtr hook, uint eventType, IntPtr hwnd, int idObject, int idChild, uint thread, uint time)
        {
            try
            {
                if (!IsRibbon(hwnd))
                {
                    _focusedWindow = hwnd;
                    _focusedName = null;
                    return;
                }
                IAccessible acc;
                object child;
                if (AccessibleObjectFromEvent(hwnd, idObject, idChild, out acc, out child) != 0 || acc == null)
                    return;
                // Each move also raises focus events without a name (UI Automation's own); they say nothing new.
                var name = acc.get_accName(child ?? 0);
                if (string.IsNullOrEmpty(name))
                    return;
                _focusedWindow = hwnd;
                _focusedName = name;
                var role = acc.get_accRole(child ?? 0);
                _focusedRole = role is int r ? r : 0;
            }
            catch (Exception ex)
            {
                _focusedName = null;
                Log.Error("Ribbon Tab focus", ex);
            }
        }

        /// <summary>
        /// Called a moment after a Tab key went through (<paramref name="shift"/>: Shift+Tab): skips the list the focus
        /// landed on, if it is one, in the same direction.
        /// </summary>
        public static void AfterTab(bool shift)
        {
            try
            {
                if (_skips < MaxSkips && GetFocus() == _focusedWindow && IsRibbon(_focusedWindow) && OnGallery())
                {
                    _skips++;
                    // The same key again: Shift+Tab when the key was Shift+Tab, even if Shift has been let go meanwhile.
                    bool press = shift && (GetKeyState(0x10) & 0x8000) == 0;
                    if (press)
                        keybd_event(0x10, 0, 0, UIntPtr.Zero);
                    keybd_event(0x09, 0, 0, UIntPtr.Zero);
                    keybd_event(0x09, 0, 2, UIntPtr.Zero); // KEYEVENTF_KEYUP
                    if (press)
                        keybd_event(0x10, 0, 2, UIntPtr.Zero);
                }
                else
                    _skips = 0;
            }
            catch (Exception ex)
            {
                _skips = 0;
                Log.Error("Ribbon Tab", ex);
            }
        }

        private static bool OnGallery()
        {
            return GalleryName.Equals(_focusedName, StringComparison.Ordinal)
                || (_focusedRole == ROLE_SYSTEM_TOOLBAR && GroupName.Equals(_focusedName, StringComparison.Ordinal));
        }

        private static bool IsRibbon(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero)
                return false;
            var cls = new StringBuilder(64);
            GetClassName(hwnd, cls, cls.Capacity);
            return cls.ToString() == "NetUIHWND"; // the ribbon's window
        }
    }
}
