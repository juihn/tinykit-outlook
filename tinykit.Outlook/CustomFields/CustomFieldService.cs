using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    internal sealed class FillResult
    {
        public int Updated;
        public int Unchanged;
        public int Skipped;
        public int Failed;

        public override string ToString()
        {
            return Updated + " updated, " + Unchanged + " already up to date, " + Skipped + " skipped (not mail)"
                + (Failed > 0 ? ", " + Failed + " failed (see OutlookAddin.log)" : "");
        }
    }

    /// <summary>
    /// Fills the custom columns: automatically for mail arriving in each account's Inbox or Sent Items,
    /// and on demand for the selection or for the current folder's items that have none yet.
    /// New items are queued and stamped a moment later on the UI thread, a few per tick.
    /// </summary>
    internal sealed class CustomFieldService
    {
        private const int BatchPerTick = 20;
        private const int PrintableFlag = 0x4; // PDO_PRINT_SAVEAS: hide the columns from printouts, as before

        private readonly Outlook.Application _app;
        private readonly List<Outlook.Items> _watched = new List<Outlook.Items>();
        private readonly Queue<KeyValuePair<string, string>> _queue = new Queue<KeyValuePair<string, string>>();
        private readonly HashSet<string> _queued = new HashSet<string>(StringComparer.Ordinal);
        private readonly Timer _timer;
        private CustomFieldCalculator _calculator;

        public bool AutoFill { get; set; }

        public CustomFieldService(Outlook.Application app, bool autoFill)
        {
            _app = app;
            AutoFill = autoFill;
            _timer = new Timer { Interval = 1500 };
            _timer.Tick += (s, e) => ProcessQueue();
        }

        private CustomFieldCalculator Calculator
        {
            get { return _calculator ?? (_calculator = new CustomFieldCalculator(_app.Session)); }
        }

        /// <summary>Hooks ItemAdd on the Inbox and Sent Items of every account's delivery store.</summary>
        public void Start()
        {
            var stores = new HashSet<string>(StringComparer.Ordinal);
            foreach (Outlook.Account account in _app.Session.Accounts)
            {
                Outlook.Store store;
                try
                {
                    store = account.DeliveryStore;
                    if (store == null || !stores.Add(store.StoreID))
                        continue;
                }
                catch (COMException)
                {
                    continue;
                }
                Watch(store, Outlook.OlDefaultFolders.olFolderInbox);
                Watch(store, Outlook.OlDefaultFolders.olFolderSentMail);
            }
            _timer.Start();
        }

        private void Watch(Outlook.Store store, Outlook.OlDefaultFolders which)
        {
            try
            {
                var items = store.GetDefaultFolder(which).Items;
                items.ItemAdd += OnItemAdd;
                _watched.Add(items); // keep the RCW alive or the event stops firing
            }
            catch (COMException ex)
            {
                Log.Info("No " + which + " in " + store.DisplayName + ": " + ex.Message);
            }
        }

        private void OnItemAdd(object item)
        {
            if (!AutoFill)
                return;
            try
            {
                if (ItemView.From(item) == null)
                    return;
                dynamic d = item;
                string entryId = d.EntryID;
                string storeId = ((Outlook.MAPIFolder)d.Parent).StoreID;
                if (_queued.Add(entryId))
                    _queue.Enqueue(new KeyValuePair<string, string>(entryId, storeId));
            }
            catch (Exception ex)
            {
                Log.Error("CustomFields.OnItemAdd", ex);
            }
        }

        private void ProcessQueue()
        {
            for (int n = 0; n < BatchPerTick && _queue.Count > 0; n++)
            {
                var next = _queue.Dequeue();
                _queued.Remove(next.Key);
                try
                {
                    Fill(_app.Session.GetItemFromID(next.Key, next.Value));
                }
                catch (COMException ex)
                {
                    // moved or deleted in the meantime (e.g. a sent item moved to the Inbox — stamped there instead)
                    Log.Info("CustomFields: item gone before stamping: " + ex.Message);
                }
                catch (Exception ex)
                {
                    Log.Error("CustomFields.ProcessQueue", ex);
                }
            }
        }

        /// <summary>Computes and writes the columns; saves only when a value changed. Null for non-mail items.</summary>
        public bool? Fill(object item)
        {
            var view = ItemView.From(item);
            if (view == null)
                return null;
            return Write(view, Calculator.Compute(view));
        }

        public CustomFieldValues Compute(object item)
        {
            var view = ItemView.From(item);
            return view == null ? null : Calculator.Compute(view);
        }

        /// <summary>The item's stored domainRelated, or the computed one if it has none yet.</summary>
        public string DomainRelatedOf(object item)
        {
            var view = ItemView.From(item);
            if (view == null)
                return null;
            var stored = view.UserProperties.Find(CustomFieldNames.DomainRelated);
            var value = stored == null ? null : stored.Value as string;
            if (string.IsNullOrWhiteSpace(value) || value == CustomFieldValues.None)
                value = Calculator.Compute(view).DomainRelated;
            return value == CustomFieldValues.None ? null : value;
        }

        /// <summary>The item's stored nameRelated, or the computed one if it has none yet.</summary>
        public string NameRelatedOf(object item)
        {
            var view = ItemView.From(item);
            if (view == null)
                return null;
            var stored = view.UserProperties.Find(CustomFieldNames.NameRelated);
            var value = stored == null ? null : stored.Value as string;
            if (string.IsNullOrWhiteSpace(value) || value == CustomFieldValues.None)
                value = Calculator.Compute(view).NameRelated;
            return value == CustomFieldValues.None ? null : value;
        }

        public FillResult FillItems(IEnumerable<object> items)
        {
            var result = new FillResult();
            var cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                foreach (var item in items)
                {
                    try
                    {
                        var changed = Fill(item);
                        if (changed == null) result.Skipped++;
                        else if (changed.Value) result.Updated++;
                        else result.Unchanged++;
                    }
                    catch (Exception ex)
                    {
                        result.Failed++;
                        Log.Error("CustomFields.FillItems", ex);
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }
            finally
            {
                Cursor.Current = cursor;
            }
            return result;
        }

        private static bool Write(ItemView view, CustomFieldValues values)
        {
            bool changed = false;
            foreach (var pair in values.Pairs)
            {
                var prop = view.UserProperties.Find(pair.Key);
                if (prop == null)
                {
                    prop = view.UserProperties.Add(pair.Key, Outlook.OlUserPropertyType.olText, true);
                    HideFromPrint(prop);
                    changed = true;
                }
                if (!string.Equals(prop.Value as string, pair.Value, StringComparison.Ordinal))
                {
                    prop.Value = pair.Value;
                    changed = true;
                }
            }
            if (changed)
                view.Save();
            return changed;
        }

        // UserProperty flags live behind DISPID 107 (see stackoverflow.com/questions/701508).
        private static void HideFromPrint(Outlook.UserProperty prop)
        {
            try
            {
                const string member = "[DispID=107]";
                var type = prop.GetType();
                var flags = Convert.ToInt32(type.InvokeMember(member, BindingFlags.GetProperty, null, prop, null));
                type.InvokeMember(member, BindingFlags.SetProperty, null, prop, new object[] { flags & ~PrintableFlag });
            }
            catch (Exception ex)
            {
                Log.Error("HideFromPrint", ex);
            }
        }
    }
}
