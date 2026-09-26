using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace TinyKit.OutlookAddin.Common
{
    /// <summary>Wraps an Office window (Explorer/Inspector) HWND so WinForms dialogs are modal to it.</summary>
    internal sealed class WindowOwner : IWin32Window
    {
        [ComImport, Guid("00000114-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IOleWindow
        {
            void GetWindow(out IntPtr phwnd);
            void ContextSensitiveHelp([MarshalAs(UnmanagedType.Bool)] bool fEnterMode);
        }

        public IntPtr Handle { get; private set; }

        private WindowOwner(IntPtr handle)
        {
            Handle = handle;
        }

        /// <summary>Returns an owner for the given Office window object, or null if it has no HWND.</summary>
        public static IWin32Window From(object officeWindow)
        {
            var ole = officeWindow as IOleWindow;
            if (ole == null)
                return null;
            try
            {
                IntPtr hwnd;
                ole.GetWindow(out hwnd);
                return hwnd == IntPtr.Zero ? null : new WindowOwner(hwnd);
            }
            catch
            {
                return null;
            }
        }
    }
}
