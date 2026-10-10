using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>
    /// Mail in one default folder of each account (Sent Items, Junk Email) is moved to the same account's Inbox, a few
    /// seconds after it arrives there.
    /// <list type="bullet">
    /// <item>Sent Items: the whole folder, also after Outlook starts and when turned on, so nothing is missed when many
    ///   arrive at once; a conversation reads in one place, also mail sent from another mail client and synchronized.
    ///   Gmail accounts too, where Sent Mail is a label: moving out of it over IMAP takes the label off.</item>
    /// <item>Junk Email: only mail arriving while Outlook runs with the option on (what was there before stays), so nothing
    ///   new is lost to a wrong junk verdict (Gmail's Spam too: moving out of it is "not spam"). The moved mail gets
    ///   🅙 in domainMark, which filling the fields again keeps.</item>
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
            public readonly List<string> Arrived = new List<string>(); // new-only: entry ids since the last sweep
        }

        private readonly Outlook.Application _app;
        private readonly Outlook.OlDefaultFolders _source;
        private readonly string _what; // for the log, e.g. "Sent to Inbox"
        private readonly string _domainMark; // written on each moved mail; null: none
        private readonly bool _newOnly; // only mail arriving from now on, not what is in the folder already
        private readonly List<Pair> _pairs = new List<Pair>();
        private readonly Timer _timer = new Timer { Interval = DelayMs };
        private bool _enabled;

        public static FolderToInbox SentMail(Outlook.Application app, bool enabled)
        {
            return new FolderToInbox(app, enabled, Outlook.OlDefaultFolders.olFolderSentMail, "Sent to Inbox", null, false);
        }

        public static FolderToInbox JunkMail(Outlook.Application app, bool enabled)
        {
            return new FolderToInbox(app, enabled, Outlook.OlDefaultFolders.olFolderJunk, "Junk to Inbox", CustomFieldValues.FromJunk, true);
        }

        private FolderToInbox(Outlook.Application app, bool enabled, Outlook.OlDefaultFolders source, string what, string domainMark, bool newOnly)
        {
            _app = app;
            _enabled = enabled;
            _source = source;
            _what = what;
            _domainMark = domainMark;
            _newOnly = newOnly;
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

        /// <summary>On: the next sweep moves what is in the folder now (new-only: what arrives from now on).</summary>
        public bool Enabled
        {
            get { return _enabled; }
            set
            {
                _enabled = value;
                if (value && !_newOnly)
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
                            if (_newOnly)
                            {
                                var mail = item as Outlook.MailItem;
                                if (!_enabled || mail == null)
                                    return; // meeting requests and other items stay
                                pair.Arrived.Add(mail.EntryID);
                            }
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
            if (!_newOnly)
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
        /// <summary>Stops the timer (the add-in is being unloaded: a tick after that would end Outlook).</summary>
        public void Stop()
        {
            _enabled = false;
            _timer.Stop();
            _timer.Dispose();
        }

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
                    if (_newOnly)
                    {
                        var ids = pair.Arrived.ToList();
                        pair.Arrived.Clear();
                        var sourceId = pair.Source.EntryID;
                        foreach (var id in ids)
                        {
                            Outlook.MailItem mail;
                            try
                            {
                                mail = _app.Session.GetItemFromID(id, pair.Source.StoreID) as Outlook.MailItem;
                            }
                            catch (COMException)
                            {
                                continue; // gone meanwhile
                            }
                            if (mail == null)
                                continue;
                            var parent = mail.Parent as Outlook.MAPIFolder;
                            if (parent != null && parent.EntryID == sourceId) // not moved elsewhere meanwhile
                            {
                                if (MoveOne(mail, pair))
                                    moved++;
                            }
                            else
                                Marshal.ReleaseComObject(mail);
                        }
                    }
                    else
                    {
                        var items = pair.Source.Items;
                        for (int i = items.Count; i >= 1; i--) // from the end: moving takes the item out of the collection
                        {
                            var mail = items[i] as Outlook.MailItem;
                            if (mail != null && MoveOne(mail, pair)) // meeting requests and other items stay
                                moved++;
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

        // Moves one mail to the Inbox (marking it when asked) and releases it; false when Outlook refused.
        private bool MoveOne(Outlook.MailItem mail, Pair pair)
        {
            try
            {
                var inInbox = mail.Move(pair.Inbox);
                if (_domainMark != null && inInbox != null)
                {
                    MarkMoved(inInbox);
                    Marshal.ReleaseComObject(inInbox);
                }
                return true;
            }
            catch (COMException ex)
            {
                Log.Info(_what + ": \"" + mail.Subject + "\" not moved: " + ex.Message);
                return false;
            }
            finally
            {
                Marshal.ReleaseComObject(mail);
            }
        }
    }
}
