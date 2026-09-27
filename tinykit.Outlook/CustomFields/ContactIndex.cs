using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using tinykit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>A contact found for an e-mail address.</summary>
    internal sealed class ContactEntry
    {
        /// <summary>The e-mail display name of the matching slot, else File As.</summary>
        public string Name;
        public string Department;
        public string EntryId;
        public string StoreId;
    }

    /// <summary>
    /// E-mail address → contact, over the default Contacts folder of every store and its subfolders
    /// (default store first; the first contact with an address wins).
    /// Rebuilt lazily after any contact is added, changed or removed.
    /// </summary>
    internal sealed class ContactIndex
    {
        private const string ContactSet = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/";
        private const string FileAs = ContactSet + "8005001F";
        private const string Department = "urn:schemas:contacts:department";
        private const int BatchRows = 2000;

        // (address, e-mail display name) per e-mail slot 1..3
        private static readonly string[][] Slots =
        {
            new[] { ContactSet + "8083001F", ContactSet + "8080001F" },
            new[] { ContactSet + "8093001F", ContactSet + "8090001F" },
            new[] { ContactSet + "80A3001F", ContactSet + "80A0001F" },
        };

        private readonly Outlook.NameSpace _session;
        private readonly List<Outlook.Items> _watched = new List<Outlook.Items>();
        private readonly HashSet<string> _watchedIds = new HashSet<string>(StringComparer.Ordinal);
        private Dictionary<string, ContactEntry> _entries;

        public ContactIndex(Outlook.NameSpace session)
        {
            _session = session;
        }

        /// <summary>
        /// The contact's name for <paramref name="smtp"/>: the e-mail display name of the matching slot,
        /// else File As. Null when the address is not in Contacts.
        /// </summary>
        public string Find(string smtp)
        {
            var entry = Lookup(smtp);
            return entry == null ? null : entry.Name;
        }

        /// <summary>The contact for <paramref name="smtp"/>, or null when the address is not in Contacts.</summary>
        public ContactEntry Lookup(string smtp)
        {
            if (string.IsNullOrEmpty(smtp))
                return null;
            if (_entries == null)
                Build();
            ContactEntry entry;
            return _entries.TryGetValue(smtp.Trim(), out entry) ? entry : null;
        }

        private void Build()
        {
            var entries = new Dictionary<string, ContactEntry>(StringComparer.OrdinalIgnoreCase);
            foreach (var folder in ContactFolders())
            {
                try
                {
                    Watch(folder);
                    var storeId = folder.StoreID;
                    var table = folder.GetTable();
                    table.Columns.RemoveAll();
                    table.Columns.Add("MessageClass");
                    table.Columns.Add("EntryID");
                    table.Columns.Add(FileAs);
                    table.Columns.Add(Department);
                    foreach (var slot in Slots)
                    {
                        table.Columns.Add(slot[0]);
                        table.Columns.Add(slot[1]);
                    }

                    // Columns: 0 MessageClass, 1 EntryID, 2 FileAs, 3 Department, then (address, display name) per slot.
                    const int firstSlot = 4;
                    while (!table.EndOfTable)
                    {
                        var rows = (Array)table.GetArray(BatchRows);
                        int c0 = rows.GetLowerBound(1);
                        for (int r = rows.GetLowerBound(0); r <= rows.GetUpperBound(0); r++)
                        {
                            var cls = rows.GetValue(r, c0) as string;
                            if (cls == null || !cls.StartsWith("IPM.Contact", StringComparison.OrdinalIgnoreCase))
                                continue;
                            var entryId = rows.GetValue(r, c0 + 1) as string;
                            var fileAs = (rows.GetValue(r, c0 + 2) as string ?? "").Trim();
                            var department = (rows.GetValue(r, c0 + 3) as string ?? "").Trim();
                            for (int s = 0; s < Slots.Length; s++)
                            {
                                var address = (rows.GetValue(r, c0 + firstSlot + 2 * s) as string ?? "").Trim();
                                if (address.IndexOf('@') < 0 || entries.ContainsKey(address))
                                    continue;
                                var display = (rows.GetValue(r, c0 + firstSlot + 1 + 2 * s) as string ?? "").Trim();
                                var name = display.Length > 0 ? display : fileAs;
                                if (name.Length > 0)
                                    entries[address] = new ContactEntry
                                    {
                                        Name = name,
                                        Department = department.Length > 0 ? department : null,
                                        EntryId = entryId,
                                        StoreId = storeId,
                                    };
                            }
                        }
                    }
                }
                catch (COMException ex)
                {
                    Log.Error("ContactIndex " + folder.FolderPath, ex);
                }
            }
            _entries = entries;
        }

        private IEnumerable<Outlook.MAPIFolder> ContactFolders()
        {
            var defaultStoreId = SafeGet(() => _session.DefaultStore.StoreID);
            var stores = new List<Outlook.Store>();
            foreach (Outlook.Store s in _session.Stores)
                stores.Add(s);

            foreach (var store in stores.OrderBy(s => SafeGet(() => s.StoreID) == defaultStoreId ? 0 : 1))
            {
                Outlook.MAPIFolder folder = null;
                try
                {
                    folder = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
                }
                catch (COMException)
                {
                    // store without a Contacts folder (PST archive, public folders, ...)
                }
                if (folder == null)
                    continue;
                foreach (var f in WithSubfolders(folder))
                    yield return f;
            }
        }

        // Contacts subfolders Outlook/Exchange maintain themselves, not the user's contacts.
        private static readonly HashSet<string> SystemFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Recipient Cache", "PeopleCentricConversation Buddies", "GAL Contacts", "Organizational Contacts",
            "Companies", "ExternalContacts", "PersonMetadata", "MeContact", "Skype for Business Contacts",
        };

        /// <summary>The folder and its contact subfolders, depth first (user contact groups like "olk/family").</summary>
        private static IEnumerable<Outlook.MAPIFolder> WithSubfolders(Outlook.MAPIFolder root)
        {
            var stack = new Stack<Outlook.MAPIFolder>();
            stack.Push(root);
            while (stack.Count > 0)
            {
                var folder = stack.Pop();
                yield return folder;

                var children = new List<Outlook.MAPIFolder>();
                try
                {
                    foreach (Outlook.MAPIFolder child in folder.Folders)
                    {
                        var name = child.Name ?? "";
                        if (child.DefaultItemType == Outlook.OlItemType.olContactItem
                            && !SystemFolders.Contains(name) && !name.StartsWith("{", StringComparison.Ordinal))
                            children.Add(child);
                    }
                }
                catch (COMException)
                {
                    // folder not accessible
                }
                for (int i = children.Count - 1; i >= 0; i--)
                    stack.Push(children[i]);
            }
        }

        private void Watch(Outlook.MAPIFolder folder)
        {
            if (!_watchedIds.Add(folder.EntryID))
                return;
            var items = folder.Items;
            items.ItemAdd += OnChanged;
            items.ItemChange += OnChanged;
            items.ItemRemove += OnRemoved;
            _watched.Add(items); // keep the RCW alive or the events stop
        }

        private void OnChanged(object item)
        {
            _entries = null;
        }

        private void OnRemoved()
        {
            _entries = null;
        }

        private static string SafeGet(Func<string> get)
        {
            try
            {
                return get();
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
