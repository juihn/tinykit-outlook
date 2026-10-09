using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Automation;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// Puts the keyboard focus in a ribbon edit box found by its name (the ribbon has no API for that, and MSAA's
    /// accSelect leaves a ribbon box unfocused): UI Automation's SetFocus, on a worker thread, because UI Automation
    /// calls into the ribbon's provider on the UI thread and would wait for itself if made from there.
    /// </summary>
    internal static class RibbonFocus
    {
        private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr parent, EnumProc proc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hwnd, StringBuilder name, int count);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        /// <summary>
        /// Focuses the first edit box named <paramref name="name"/> in the ribbon of the window <paramref name="owner"/>,
        /// trying a few times while the ribbon draws a tab just shown. <paramref name="done"/> gets whether it worked
        /// (on the worker thread).
        /// </summary>
        public static void FocusEditBox(IntPtr owner, string name, Action<bool> done)
        {
            Task.Run(() =>
            {
                bool ok = false;
                try
                {
                    var condition = new AndCondition(
                        new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit),
                        new PropertyCondition(AutomationElement.NameProperty, name));
                    for (int attempt = 0; attempt < 10 && !ok; attempt++)
                    {
                        if (attempt > 0)
                            System.Threading.Thread.Sleep(150);
                        foreach (var hwnd in RibbonWindows(owner))
                        {
                            var box = AutomationElement.FromHandle(hwnd).FindFirst(TreeScope.Descendants, condition);
                            if (box == null || box.Current.IsOffscreen)
                                continue;
                            box.SetFocus();
                            ok = true;
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Ribbon focus", ex);
                }
                done(ok);
            });
        }

        private static List<IntPtr> RibbonWindows(IntPtr owner)
        {
            var found = new List<IntPtr>();
            EnumChildWindows(owner, (h, l) =>
            {
                var cls = new StringBuilder(64);
                GetClassName(h, cls, cls.Capacity);
                if (cls.ToString() == "NetUIHWND" && IsWindowVisible(h))
                    found.Add(h);
                return true;
            }, IntPtr.Zero);
            return found;
        }
    }
}
