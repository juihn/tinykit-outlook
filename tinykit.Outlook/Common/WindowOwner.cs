using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Common
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

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left, Top, Right, Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT rect);

        /// <summary>
        /// Shows a modeless window centred on the Office window it belongs to, on that window's screen (WinForms puts a
        /// modeless window on the main screen: CenterParent applies to ShowDialog only), and no larger than that screen.
        /// </summary>
        public static void ShowCentred(Form form, object officeWindow)
        {
            var owner = From(officeWindow);
            RECT r;
            if (owner != null && GetWindowRect(owner.Handle, out r))
            {
                form.StartPosition = FormStartPosition.Manual;
                Place(form, System.Drawing.Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom));
            }
            form.Show(owner);
        }

        /// <summary>
        /// Brings an open window to the front, first moving it onto the Office window's screen when it is on another one.
        /// </summary>
        public static void Activate(Form form, object officeWindow)
        {
            var owner = From(officeWindow);
            RECT r;
            if (owner != null && GetWindowRect(owner.Handle, out r) && form.WindowState == FormWindowState.Normal)
            {
                var ownerBounds = System.Drawing.Rectangle.FromLTRB(r.Left, r.Top, r.Right, r.Bottom);
                if (!Screen.FromRectangle(ownerBounds).Equals(Screen.FromControl(form)))
                    Place(form, ownerBounds);
            }
            form.Activate();
        }

        private static void Place(Form form, System.Drawing.Rectangle ownerBounds)
        {
            var area = Screen.FromRectangle(ownerBounds).WorkingArea;
            var size = new System.Drawing.Size(Math.Min(form.Width, area.Width), Math.Min(form.Height, area.Height));
            var x = ownerBounds.Left + (ownerBounds.Width - size.Width) / 2;
            var y = ownerBounds.Top + (ownerBounds.Height - size.Height) / 2;
            x = Math.Max(area.Left, Math.Min(x, area.Right - size.Width));
            y = Math.Max(area.Top, Math.Min(y, area.Bottom - size.Height));
            form.Bounds = new System.Drawing.Rectangle(x, y, size.Width, size.Height);
        }
    }
}
