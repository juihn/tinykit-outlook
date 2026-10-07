using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>
    /// Sent mail to Inbox: mail in an account's Sent Items is moved to the same account's Inbox, so a conversation reads
    /// in one place. It runs a few seconds after Outlook starts and after mail arrives in Sent Items (sent from Outlook,
    /// or sent from another mail client and synchronized), each time for all of Sent Items, so nothing is missed when
    /// many arrive at once. Gmail accounts are left out: there Sent Mail is a label, and moving out of it over IMAP
    /// deletes Gmail's sent copy or brings it back.
    /// </summary>
    internal sealed class SentToInbox
    {
        private const int DelayMs = 5000; // lets the custom fields be filled first (they are filled again in the Inbox anyway)

        private sealed class Pair
        {
            public Outlook.MAPIFolder Sent;
            public Outlook.MAPIFolder Inbox;
            public Outlook.Items Items; // kept alive or ItemAdd stops firing
        }

        private readonly Outlook.Application _app;
        private readonly List<Pair> _pairs = new List<Pair>();
        private readonly Timer _timer = new Timer { Interval = DelayMs };
        private bool _enabled;

        public SentToInbox(Outlook.Application app, bool enabled)
        {
            _app = app;
            _enabled = enabled;
            _timer.Tick += (s, e) =>
            {
                // An exception left to Outlook from a timer crashes it, and Outlook then disables the add-in.
                try
                {
                    Sweep();
                }
                catch (Exception ex)
                {
                    Log.Error("Sent to Inbox", ex);
                }
            };
        }

        /// <summary>On: the next sweep moves what is in Sent Items now.</summary>
        public bool Enabled
        {
            get { return _enabled; }
            set
            {
                _enabled = value;
                if (value)
                    Schedule();
            }
        }

        /// <summary>Watches the Sent Items of every account's delivery store (Gmail left out) and sweeps once.</summary>
        public void Start()
        {
            var stores = new HashSet<string>(StringComparer.Ordinal);
            foreach (Outlook.Account account in _app.Session.Accounts)
            {
                try
                {
                    var store = account.DeliveryStore;
                    if (store == null || !stores.Add(store.StoreID))
                        continue;
                    var sent = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderSentMail);
                    if (IsGmail(sent))
                    {
                        Log.Info("Sent to Inbox: " + sent.FolderPath + " left out (Gmail)");
                        continue;
                    }
                    var pair = new Pair
                    {
                        Sent = sent,
                        Inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox),
                        Items = sent.Items,
                    };
                    pair.Items.ItemAdd += item =>
                    {
                        try
                        {
                            Schedule();
                        }
                        catch (Exception ex)
                        {
                            Log.Error("Sent to Inbox: ItemAdd", ex);
                        }
                    };
                    _pairs.Add(pair);
                }
                catch (COMException ex)
                {
                    Log.Info("Sent to Inbox: account " + account.DisplayName + " left out: " + ex.Message);
                }
            }
            Schedule();
        }

        private static bool IsGmail(Outlook.MAPIFolder folder)
        {
            var path = folder.FolderPath ?? "";
            return path.IndexOf("[Gmail]", StringComparison.OrdinalIgnoreCase) >= 0
                || path.IndexOf("[Google Mail]", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static string PathOf(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.FolderPath;
            }
            catch (COMException)
            {
                return "(folder not reachable)";
            }
        }

        // A sweep a few seconds from now (again from now if one was waiting).
        private void Schedule()
        {
            if (!_enabled)
                return;
            _timer.Stop();
            _timer.Start();
        }

        private void Sweep()
        {
            _timer.Stop();
            if (!_enabled)
                return;
            foreach (var pair in _pairs)
            {
                int moved = 0;
                try
                {
                    var items = pair.Sent.Items;
                    for (int i = items.Count; i >= 1; i--) // from the end: moving takes the item out of the collection
                    {
                        var mail = items[i] as Outlook.MailItem;
                        if (mail == null)
                            continue; // meeting requests and other items stay
                        try
                        {
                            mail.Move(pair.Inbox);
                            moved++;
                        }
                        catch (COMException ex)
                        {
                            Log.Info("Sent to Inbox: \"" + mail.Subject + "\" not moved: " + ex.Message);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mail);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("Sent to Inbox: " + PathOf(pair.Sent), ex);
                }
                if (moved > 0)
                    Log.Info("Sent to Inbox: moved " + moved + " mail(s) from " + PathOf(pair.Sent) + " to " + PathOf(pair.Inbox));
            }
        }
    }
}
