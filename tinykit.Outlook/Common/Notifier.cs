using System;
using System.Reflection;
using System.Security;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// Informational messages as Windows notifications (toasts) instead of message boxes, so they don't block Outlook.
    /// They use the "urgent" scenario (shown even in Do Not Disturb) and are posted under Outlook (classic)'s app id.
    /// Errors and questions stay message boxes. If a toast can't be shown (e.g. Outlook's notifications are turned
    /// off in Windows Settings), the message box is used instead.
    /// </summary>
    internal static class Notifier
    {
        /// <summary>AppUserModelID of Outlook (classic); toasts need a registered app to show under.</summary>
        private const string AppId = "Microsoft.Office.OUTLOOK.EXE.15";

        // WinRT types are loaded by name, so building the add-in needs no Windows SDK reference.
        private const string WinRt = ", ContentType=WindowsRuntime";
        private static readonly Type XmlDocumentType = Type.GetType("Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom" + WinRt);
        private static readonly Type ToastType = Type.GetType("Windows.UI.Notifications.ToastNotification, Windows.UI.Notifications" + WinRt);
        private static readonly Type ManagerType = Type.GetType("Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications" + WinRt);
        private static readonly Type NotifierType = Type.GetType("Windows.UI.Notifications.ToastNotifier, Windows.UI.Notifications" + WinRt);

        /// <summary>Shows an informational message as a notification (a message box if that is not possible).</summary>
        public static void Info(Outlook.Explorer explorer, string message)
        {
            if (TryToast(message))
                return;
            MessageBox.Show(explorer == null ? null : WindowOwner.From(explorer), message, ThisAddIn.Title,
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static bool TryToast(string message)
        {
            try
            {
                if (XmlDocumentType == null || ToastType == null || ManagerType == null || NotifierType == null)
                    return false;

                var notifier = ManagerType.GetMethod("CreateToastNotifier", new[] { typeof(string) }).Invoke(null, new object[] { AppId });
                // NotificationSetting: 0 = Enabled; anything else means Windows would drop the toast silently.
                if (Convert.ToInt32(NotifierType.GetProperty("Setting").GetValue(notifier)) != 0)
                    return false;

                var xml = "<toast scenario=\"urgent\"><visual><binding template=\"ToastGeneric\">"
                    + "<text>" + SecurityElement.Escape(ThisAddIn.Title) + "</text>"
                    + "<text>" + SecurityElement.Escape(message) + "</text>"
                    + "</binding></visual></toast>";
                var doc = Activator.CreateInstance(XmlDocumentType);
                XmlDocumentType.GetMethod("LoadXml", new[] { typeof(string) }).Invoke(doc, new object[] { xml });
                var toast = Activator.CreateInstance(ToastType, doc);
                ToastType.GetProperty("Group").SetValue(toast, "tinykit");
                NotifierType.GetMethod("Show").Invoke(notifier, new[] { toast });
                return true;
            }
            catch (Exception ex)
            {
                Log.Error("Notifier", ex is TargetInvocationException && ex.InnerException != null ? ex.InnerException : ex);
                return false;
            }
        }
    }
}
