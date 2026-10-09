using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>
    /// Mail in one default folder of each account (Sent Items, Junk Email) is moved to the same account's Inbox. It runs a
    /// few seconds after Outlook starts and after mail arrives in that folder, each time for the whole folder, so nothing
    /// is missed when many arrive at once.
    /// <list type="bullet">
    /// <item>Sent Items: a conversation reads in one place, also mail sent from another mail client and synchronized.
    ///   Gmail accounts too, where Sent Mail is a label: moving out of it over IMAP takes the label off.</item>
    /// <item>Junk Email: nothing is lost to a wrong junk verdict (Gmail's Spam too: moving out of it is "not spam"). The
    ///   moved mail gets 🅙 in domainMark, which filling the fields again keeps.</item>
    /// </list>
    /// </summary>
    internal sealed class FolderToInbox
    {
        private const int DelayMs = 5000; // lets the custom fields be filled first (they are filled again in the Inbox anyway)

        private sealed class Pair
        {
            public Outlook.MAPIFolder Source;
            public Outlook.MAPIFolder Inbox;
            public Outlook.Items Items; // kept alive or ItemAdd stops firing
        }

        private readonly Outlook.Application _app;
        private readonly Outlook.OlDefaultFolders _source;
        private readonly string _what; // for the log, e.g. "Sent to Inbox"
        private readonly string _domainMark; // written on each moved mail; null: none
        private readonly List<Pair> _pairs = new List<Pair>();
        private readonly Timer _timer = new Timer { Interval = DelayMs };
        private bool _enabled;

        public static FolderToInbox SentMail(Outlook.Application app, bool enabled)
        {
            return new FolderToInbox(app, enabled, Outlook.OlDefaultFolders.olFolderSentMail, "Sent to Inbox", null);
        }

        public static FolderToInbox JunkMail(Outlook.Application app, bool enabled)
        {
            return new FolderToInbox(app, enabled, Outlook.OlDefaultFolders.olFolderJunk, "Junk to Inbox", CustomFieldValues.FromJunk);
        }

        private FolderToInbox(Outlook.Application app, bool enabled, Outlook.OlDefaultFolders source, string what, string domainMark)
        {
            _app = app;
            _enabled = enabled;
            _source = source;
            _what = what;
            _domainMark = domainMark;
            _timer.Tick += (s, e) =>
            {
                // An exception left to Outlook from a timer crashes it, and Outlook then disables the add-in.
                try
                {
                    Sweep();
                }
                catch (Exception ex)
                {
                    Log.Error(_what, ex);
                }
            };
        }

        /// <summary>On: the next sweep moves what is in the folder now.</summary>
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

        /// <summary>Watches the folder of every account's delivery store and sweeps once.</summary>
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
                    var source = store.GetDefaultFolder(_source);
                    var pair = new Pair
                    {
                        Source = source,
                        Inbox = store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox),
                        Items = source.Items,
                    };
                    pair.Items.ItemAdd += item =>
                    {
                        try
                        {
                            Schedule();
                        }
                        catch (Exception ex)
                        {
                            Log.Error(_what + ": ItemAdd", ex);
                        }
                    };
                    _pairs.Add(pair);
                }
                catch (COMException ex)
                {
                    Log.Info(_what + ": account " + account.DisplayName + " left out: " + ex.Message);
                }
            }
            Schedule();
        }

        // The mark in domainMark (🅙 for junk), which filling the fields keeps; a failure only leaves the mail unmarked.
        private void MarkMoved(Outlook.MailItem mail)
        {
            try
            {
                var prop = mail.UserProperties.Find(CustomFieldNames.DomainMark);
                if (prop == null)
                {
                    prop = mail.UserProperties.Add(CustomFieldNames.DomainMark, Outlook.OlUserPropertyType.olText, true);
                    CustomFieldService.HideFromPrint(prop);
                }
                prop.Value = _domainMark;
                mail.Save();
            }
            catch (Exception ex)
            {
                Log.Info(_what + ": \"" + mail.Subject + "\" not marked: " + ex.Message);
            }
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
                    var items = pair.Source.Items;
                    for (int i = items.Count; i >= 1; i--) // from the end: moving takes the item out of the collection
                    {
                        var mail = items[i] as Outlook.MailItem;
                        if (mail == null)
                            continue; // meeting requests and other items stay
                        try
                        {
                            var inInbox = mail.Move(pair.Inbox);
                            moved++;
                            if (_domainMark != null && inInbox != null)
                            {
                                MarkMoved(inInbox);
                                Marshal.ReleaseComObject(inInbox);
                            }
                        }
                        catch (COMException ex)
                        {
                            Log.Info(_what + ": \"" + mail.Subject + "\" not moved: " + ex.Message);
                        }
                        finally
                        {
                            Marshal.ReleaseComObject(mail);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Error(_what + ": " + PathOf(pair.Source), ex);
                }
                if (moved > 0)
                    Log.Info(_what + ": moved " + moved + " mail(s) from " + PathOf(pair.Source) + " to " + PathOf(pair.Inbox));
            }
        }
    }
}
