using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Settings
{
    /// <summary>The kind of folder shown: each kind has its own saved filters, view columns and quick filters.</summary>
    internal enum ItemKind
    {
        Mail,
        Contact,
        Task,
    }

    internal static class ItemKinds
    {
        public static readonly ItemKind[] All = { ItemKind.Mail, ItemKind.Contact, ItemKind.Task };

        /// <summary>The kind of the explorer's current folder, or null for other folders (calendar, notes, ...).</summary>
        public static ItemKind? Of(Outlook.Explorer explorer)
        {
            if (explorer == null)
                return null;
            try
            {
                switch (explorer.CurrentFolder.DefaultItemType)
                {
                    case Outlook.OlItemType.olMailItem:
                    case Outlook.OlItemType.olPostItem:
                        return ItemKind.Mail;
                    case Outlook.OlItemType.olContactItem:
                    case Outlook.OlItemType.olDistributionListItem:
                        return ItemKind.Contact;
                    case Outlook.OlItemType.olTaskItem:
                        return ItemKind.Task;
                }
            }
            catch (COMException)
            {
                // no current folder (e.g. while the explorer closes)
            }
            return null;
        }

        /// <summary>File name suffix: "Saved Filters - Mail.xml", "View Columns - Contacts.txt", ...</summary>
        public static string FileSuffix(ItemKind kind)
        {
            switch (kind)
            {
                case ItemKind.Contact: return "Contacts";
                case ItemKind.Task: return "Tasks";
                default: return "Mail";
            }
        }
    }
}
