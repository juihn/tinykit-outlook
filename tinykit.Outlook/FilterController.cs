using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.CustomFields;
using tinykit.OutlookAddin.Filtering;
using tinykit.OutlookAddin.Formatting;
using tinykit.OutlookAddin.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin
{
    /// <summary>Filter state and actions behind the ribbon; the ribbon only translates callbacks.</summary>
    internal sealed class FilterController
    {
        public const int MaxSavedFilters = 20;

        private const string QuickSourcePrefix = "quick:";
        private const string SavedSourcePrefix = "saved:";

        // Text of the Quick Filter input box, as last reported by the ribbon (Enter or focus leaving the box).
        private string _inputText = "";
        private bool _syncing;

        /// <summary>Saved filters of one kind of folder and the state of their file.</summary>
        private sealed class KindState
        {
            public FilterSettings Settings;
            // Content hash of the file as last loaded or written by the add-in; any other content means the file
            // was edited outside (e.g. in VS Code) and must be reloaded before use.
            public string KnownHash;
            public string ReportedBadHash;
        }

        private readonly Dictionary<ItemKind, KindState> _states = new Dictionary<ItemKind, KindState>();
        private ItemKind _kind = ItemKind.Mail;

        /// <summary>
        /// Makes the kind of the explorer's folder (mail, contacts, tasks) current: <see cref="Settings"/> and the
        /// ribbon then use that kind's saved filters. Returns null (current kind unchanged) for other folders.
        /// </summary>
        public ItemKind? Focus(Outlook.Explorer explorer)
        {
            var kind = ItemKinds.Of(explorer);
            if (kind.HasValue)
                _kind = kind.Value;
            return kind;
        }

        /// <summary>The kind made current by the last <see cref="Focus"/>.</summary>
        public ItemKind Kind
        {
            get { return _kind; }
        }

        /// <summary>Saved filters of the current kind of folder.</summary>
        public FilterSettings Settings
        {
            get { return State(_kind).Settings; }
        }

        private KindState State(ItemKind kind)
        {
            KindState state;
            if (_states.TryGetValue(kind, out state))
                return state;
            state = new KindState();
            var path = SettingsPaths.SavedFiltersFile(kind);
            try
            {
                state.Settings = FilterSettings.Load(path, kind);
                state.KnownHash = FileHash(path);
            }
            catch (Exception ex)
            {
                Log.Error("Load " + Path.GetFileName(path), ex);
                state.Settings = new FilterSettings { Kind = kind };
            }
            _states[kind] = state;
            return state;
        }

        public FilterHistory History { get; private set; }
        public ViewFilterService Views { get; private set; }

        /// <summary>Custom-column stamping; set by ThisAddIn at startup.</summary>
        public CustomFieldService Fields { get; set; }

        /// <summary>Moves mail in Sent Items to the same account's Inbox (Move Sent Mail to Inbox).</summary>
        public FolderToInbox SentMail { get; set; }

        /// <summary>Moves mail in Junk Email to the same account's Inbox (Move Junk Mail to Inbox).</summary>
        public FolderToInbox JunkMail { get; set; }

        /// <summary>Set by the ribbon once loaded; refreshes every control.</summary>
        public Action Invalidate = delegate { };

        public FilterController()
        {
            Directory.CreateDirectory(SettingsPaths.LocalFolder);
            SettingsPaths.MigrateLegacyFiles();
            History = FilterHistory.Load(SettingsPaths.HistoryFile);
            Views = new ViewFilterService(SettingsPaths.ViewStateFile);
            Views.IsAddinFilter = f => !string.IsNullOrWhiteSpace(f)
                && _states.Values.Any(st => st.Settings.Filters.Any(s => Dasl.SameFilter(s.Sql, f))
                                         || Dasl.SameFilter(OthersSql(st.Settings) ?? "", f));
            State(ItemKind.Mail); // mail settings hold autoFillFields, needed at startup
            MeSymbols.Current = () =>
            {
                var mail = State(ItemKind.Mail).Settings;
                return Tuple.Create(mail.MeSent, mail.MeTo, mail.MeCc, mail.ContactMark);
            };
        }

        // ---------- Saved Filters - <kind>.xml sync ----------

        private FileSystemWatcher _watcher;
        private volatile bool _fileTouched;
        private Timer _fileTimer;
        private Func<Outlook.Explorer> _activeExplorer = () => null;

        /// <summary>Reloads the saved filters files automatically whenever one is saved outside the add-in.</summary>
        public void WatchSettingsFile(Func<Outlook.Explorer> activeExplorer)
        {
            _activeExplorer = activeExplorer;
            _watcher = new FileSystemWatcher(SettingsPaths.SettingsFolder, "Saved Filters - *.xml")
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size,
            };
            // Raised on a thread-pool thread: only set a flag, the UI-thread timer does the work.
            _watcher.Changed += (s, e) => _fileTouched = true;
            _watcher.Created += (s, e) => _fileTouched = true;
            _watcher.Renamed += (s, e) => _fileTouched = true;
            _watcher.EnableRaisingEvents = true;

            _fileTimer = new Timer { Interval = 1000 };
            _fileTimer.Tick += (s, e) =>
            {
                if (!_fileTouched)
                    return;
                _fileTouched = false;
                try
                {
                    var explorer = _activeExplorer();
                    foreach (var kind in _states.Keys.ToList())
                        ReloadIfChanged(kind, explorer);
                }
                catch (Exception ex)
                {
                    Log.Error("Auto reload", ex);
                }
            };
            _fileTimer.Start();
        }

        /// <summary>Reloads the current kind's saved filters file if it was changed outside the add-in.</summary>
        /// <summary>Stops watching the saved filters files (the add-in is being unloaded).</summary>
        public void StopWatching()
        {
            if (_watcher != null)
            {
                _watcher.EnableRaisingEvents = false;
                _watcher.Dispose();
            }
            if (_fileTimer != null)
            {
                _fileTimer.Stop();
                _fileTimer.Dispose();
            }
        }

        private bool ReloadIfChanged(Outlook.Explorer explorer)
        {
            return ReloadIfChanged(_kind, explorer);
        }

        /// <summary>
        /// Reloads a saved filters file if it differs from what the add-in last loaded or wrote.
        /// An invalid file is reported once and the current saved filters are kept.
        /// </summary>
        private bool ReloadIfChanged(ItemKind kind, Outlook.Explorer explorer)
        {
            var state = State(kind);
            var path = SettingsPaths.SavedFiltersFile(kind);
            var hash = FileHash(path);
            if (hash == null)
            {
                _fileTouched = true; // locked while being written: try again on the next tick
                return false;
            }
            if (hash == state.KnownHash)
                return false;

            FilterSettings loaded;
            try
            {
                loaded = FilterSettings.Load(path, kind);
            }
            catch (Exception ex)
            {
                if (hash != state.ReportedBadHash)
                {
                    state.ReportedBadHash = hash;
                    Log.Error(Path.GetFileName(path), ex);
                    MessageBox.Show(explorer == null ? null : WindowOwner.From(explorer),
                        Path.GetFileName(path) + " could not be loaded; the previous saved filters are kept until it is fixed.\n\n" + ex.Message,
                        ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return false;
            }
            state.KnownHash = hash;
            UseSettings(explorer, kind, loaded);
            return true;
        }

        /// <summary>
        /// Switches to newly loaded settings of <paramref name="kind"/>; when the explorer shows that kind of folder,
        /// re-applies the shown saved filter if its SQL changed.
        /// </summary>
        private void UseSettings(Outlook.Explorer explorer, ItemKind kind, FilterSettings loaded)
        {
            var shown = explorer != null && ItemKinds.Of(explorer) == kind;
            var previous = _kind;
            _kind = kind;
            try
            {
                var active = shown ? ActiveSavedFilter(explorer) : null;
                State(kind).Settings = loaded;
                if (kind == ItemKind.Mail && Fields != null)
                    Fields.AutoFill = loaded.AutoFillFields;
                if (SentMail != null && kind == ItemKind.Mail)
                    SentMail.Enabled = loaded.MoveSentToInbox;
                if (JunkMail != null && kind == ItemKind.Mail)
                    JunkMail.Enabled = loaded.MoveJunkToInbox;
                if (shown)
                {
                    if (active != null)
                    {
                        var now = loaded.Find(active.Name);
                        if (now == null)
                            Views.Clear(explorer);
                        else if (!Dasl.SameFilter(now.Sql, active.Sql))
                            Views.Apply(explorer, Dasl.StripSqlPrefix(now.Sql), SavedSourcePrefix + now.Name);
                    }
                    if (Views.ActiveSource(explorer) == OthersSource)
                    {
                        var others = OthersSql(loaded);
                        if (others == null)
                            Views.Clear(explorer);
                        else
                            Views.Apply(explorer, others, OthersSource);
                    }
                    if (loaded.AutoApplyFormats)
                        SyncFormats(explorer, false);
                }
            }
            finally
            {
                if (!shown)
                    _kind = previous;
            }
            Invalidate();
        }

        private static string FileHash(string path)
        {
            try
            {
                using (var sha = SHA1.Create())
                    return Convert.ToBase64String(sha.ComputeHash(File.ReadAllBytes(path)));
            }
            catch (IOException)
            {
                return null;
            }
            catch (UnauthorizedAccessException)
            {
                return null;
            }
        }

        /// <summary>
        /// The saved filters shown on the ribbon: those with SQL (empty ones are placeholders), at most 20; the ones with
        /// an icon first, each part in file order.
        /// </summary>
        public IList<SavedFilter> VisibleFilters
        {
            get
            {
                return Settings.Filters.Where(f => !string.IsNullOrWhiteSpace(f.Sql))
                    .OrderBy(f => string.IsNullOrEmpty(f.Icon) ? 1 : 0) // stable: file order within each part
                    .Take(MaxSavedFilters).ToList();
            }
        }

        /// <summary>The ribbon's icon slots (filters with an icon); the visible filters start with them.</summary>
        public const int MaxIconFilters = 6;

        public SavedFilter FilterAt(int index)
        {
            var visible = VisibleFilters;
            return index >= 0 && index < visible.Count ? visible[index] : null;
        }

        // ---------- Quick filters ----------

        public string InputText
        {
            get { return _inputText; }
        }

        public void SetInputText(string text)
        {
            _inputText = text ?? "";
        }

        /// <summary>The quick filter fields of a kind of folder, in ribbon order (row 1 left, right; row 2 left, right).</summary>
        public static QuickKind[] QuickKindsFor(ItemKind? kind)
        {
            switch (kind)
            {
                case ItemKind.Mail: return new[] { QuickKind.From, QuickKind.Subject, QuickKind.Name, QuickKind.Domain };
                case ItemKind.Contact: return new[] { QuickKind.FileAs, QuickKind.Email, QuickKind.Company, QuickKind.Department };
                default: return new QuickKind[0]; // tasks and other folders: no quick filter
            }
        }

        /// <summary>
        /// Quick filter button: filters by the input box text if there is any, otherwise by the value of the first
        /// selected mail or contact.
        /// </summary>
        public void QuickButton(Outlook.Explorer explorer, QuickKind kind)
        {
            var typed = _inputText.Trim();
            if (typed.Length > 0)
            {
                ApplyQuick(explorer, kind, typed);
                return;
            }
            string value, what;
            if (_kind == ItemKind.Contact)
            {
                what = "contact";
                var contact = FirstSelected(explorer) as Outlook.ContactItem;
                if (contact == null)
                    throw new UserMessageException("Type a value in the box, or select a contact first.");
                value = ContactValue(contact, kind);
            }
            else
            {
                what = "mail";
                var info = MailInfo.FromFirstSelected(explorer,
                    Fields == null ? null : (Func<object, string>)Fields.DomainRelatedOf,
                    Fields == null ? null : (Func<object, string>)Fields.NameRelatedOf);
                if (info == null)
                    throw new UserMessageException("Type a value in the box, or select a mail first.");
                value = info.ValueFor(kind);
            }
            if (string.IsNullOrWhiteSpace(value))
                throw new UserMessageException("The selected " + what + " has no " + kind + " value.");
            ApplyQuick(explorer, kind, value);
        }

        private static object FirstSelected(Outlook.Explorer explorer)
        {
            try
            {
                var selection = explorer.Selection;
                return selection.Count > 0 ? selection[1] : null;
            }
            catch (COMException)
            {
                return null;
            }
        }

        private static string ContactValue(Outlook.ContactItem c, QuickKind kind)
        {
            switch (kind)
            {
                case QuickKind.FileAs: return c.FileAs;
                case QuickKind.Email:
                    return new[] { c.Email1Address, c.Email2Address, c.Email3Address }
                        .FirstOrDefault(a => !string.IsNullOrWhiteSpace(a) && a.Contains("@"));
                case QuickKind.Company: return c.CompanyName;
                case QuickKind.Department: return c.Department;
            }
            return null;
        }

        /// <summary>
        /// History item picked from a button's drop-down: filters by it and shows it in the input box.
        /// Ctrl+click removes the item from that list instead.
        /// </summary>
        public void QuickFromHistory(Outlook.Explorer explorer, QuickKind kind, int index)
        {
            var list = History.Get(kind);
            if (index < 0 || index >= list.Count)
                return;
            var value = list[index].Value;
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                History.Remove(kind, value);
                Invalidate();
                return;
            }
            _inputText = value;
            ApplyQuick(explorer, kind, value);
        }

        private void ApplyQuick(Outlook.Explorer explorer, QuickKind kind, string value)
        {
            value = value.Trim();
            Views.Apply(explorer, Dasl.Build(kind, value), QuickSourcePrefix + kind);
            History.Touch(kind, value);
            Invalidate();
        }

        private const string CustomSource = "custom";
        private CustomFilterForm _customFilter;

        /// <summary>
        /// Custom Filter: a window (one at a time) for the text to find and the fields to look in, for the current
        /// mail, calendar, contact or task folder. It filters the view it was opened for; Apply with no conditions
        /// restores the view's own filter.
        /// </summary>
        public void ShowCustomFilter(Outlook.Explorer explorer)
        {
            var folder = explorer.CurrentFolder;
            var kind = CustomFilter.KindOf(folder);
            if (kind == null)
                throw new UserMessageException("Custom Filter works in mail, calendar, contact and task folders.");
            if (_customFilter != null && !_customFilter.IsDisposed)
                _customFilter.Close(); // opened for another folder: start again for this one

            var entryId = folder.EntryID;
            var folderName = folder.Name;
            _customFilter = new CustomFilterForm(kind.Value, folderName, sql =>
            {
                Outlook.MAPIFolder current;
                try
                {
                    current = explorer.CurrentFolder;
                }
                catch (COMException)
                {
                    throw new UserMessageException("The Outlook window this filter belongs to is closed.");
                }
                if (current == null || current.EntryID != entryId)
                    throw new UserMessageException("This filter is for \"" + folderName + "\": go back there to apply it, "
                        + "or press Custom Filter again in this folder.");
                if (sql.Length == 0)
                {
                    if (Views.HasAddinFilter(explorer, true))
                        Views.Clear(explorer, true);
                    Invalidate();
                    return "all items shown " + DateTime.Now.ToString("HH:mm:ss");
                }
                Views.Apply(explorer, sql, CustomSource, true);
                Invalidate();
                return "applied " + DateTime.Now.ToString("HH:mm:ss");
            });
            _customFilter.Icon = OfficeImage.IconFromImageMso(explorer.CommandBars, "ApplyFilter");
            WindowOwner.ShowCentred(_customFilter, explorer);
        }

        private const string MeProperty = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/me";

        /// <summary>
        /// me Column Symbols (Table View group's dialog button): picks the symbols for mail I sent, me in To and me only in
        /// Cc. A changed symbol is also changed in the saved filters that compare me with it, and in the mail that has it
        /// (all mail folders, asked first), so filters and the column stay consistent. Sent mail's nameRelated starts with
        /// the sent symbol and one space; mail whose nameRelated has another sent symbol or no space is set right too.
        /// </summary>
        public void EditMeSymbols(Outlook.Explorer explorer)
        {
            ReloadIfChanged(ItemKind.Mail, explorer);
            var settings = State(ItemKind.Mail).Settings;
            string contact, sent, to, cc;
            var viewFont = ConditionalFormatService.GetViewFont(explorer); // null when not a table view
            using (var form = new MeSymbolsForm(settings.ContactMark, settings.MeSent, settings.MeTo, settings.MeCc,
                viewFont == null ? null : viewFont.Item1, viewFont == null ? 0 : viewFont.Item2))
            {
                form.Icon = OfficeImage.IconFromImageMso(explorer.CommandBars, "TableInsert");
                if (form.ShowDialog(WindowOwner.From(explorer)) != DialogResult.OK)
                    return;
                contact = form.Contact;
                sent = form.Sent;
                to = form.To;
                cc = form.Cc;
            }
            if (sent == to || sent == cc || to == cc)
                throw new UserMessageException("Sent Mail, To and Cc need three different symbols.");
            var changes = new Dictionary<string, string>(StringComparer.Ordinal);
            if (sent != settings.MeSent) changes[settings.MeSent] = sent;
            if (to != settings.MeTo) changes[settings.MeTo] = to;
            if (cc != settings.MeCc) changes[settings.MeCc] = cc;

            // Mail to change, found before anything changes (so swapping two symbols works): me values with an old symbol,
            // and nameRelated values starting with a sent symbol or a contact mark that are not in the chosen form
            // ("<sent> <contact> name" for sent mail, "<contact> name" for received mail).
            var session = explorer.Application.Session;
            var updates = new Dictionary<string, MailFix>(StringComparer.Ordinal); // by entry id
            int meCount = 0, nameCount = 0;
            Cursor.Current = Cursors.WaitCursor;
            foreach (Outlook.Store store in session.Stores)
            {
                List<Search.FindFolder> folders;
                try
                {
                    folders = Search.FindItems.FoldersOf(store, new[] { Search.FindKind.Mail });
                }
                catch (COMException)
                {
                    continue;
                }
                foreach (var f in folders)
                {
                    foreach (var change in changes)
                    {
                        foreach (var id in EntryIds(f, Dasl.PropertyEquals(MeProperty, change.Key), null))
                        {
                            Fix(updates, id.Item1, f.StoreId).Me = change.Value;
                            meCount++;
                        }
                    }
                    var starts = MeSymbols.SentChoices.Select(c => c.Text).Concat(new[] { settings.MeSent })
                        .Concat(MeSymbols.KnownContactMarks).Concat(new[] { settings.ContactMark }).Distinct();
                    foreach (var symbol in starts)
                    {
                        foreach (var id in EntryIds(f, Dasl.Like(NameRelatedProperty, symbol + "%"), NameRelatedProperty + "/0x0000001F"))
                        {
                            var value = id.Item2 ?? "";
                            // Sent mail to someone in Contacts stored without the mark (the name as picked from an address
                            // book): recomputed, so it gets the mark and the contact's name.
                            var plain = SentNameWithoutMark(value, settings);
                            if (plain != null && !updates.ContainsKey(id.Item1) && Fields.Calculator.Contacts.AddressNamed(plain) != null
                                && !ComputesTo(session, id.Item1, f.StoreId, value))
                            {
                                Fix(updates, id.Item1, f.StoreId).Recompute = true;
                                nameCount++;
                                continue;
                            }
                            var wanted = NameRelatedIn(value, settings, sent, contact);
                            if (wanted == null || wanted == value || (updates.ContainsKey(id.Item1) && updates[id.Item1].NameRelated != null))
                                continue;
                            Fix(updates, id.Item1, f.StoreId).NameRelated = wanted;
                            nameCount++;
                        }
                    }
                    // The old form of a name from Contacts, "[name]": recomputed per mail (the sender's contact may have
                    // been renamed since), so it gets the chosen mark or, for a sender no longer in Contacts, the plain name.
                    foreach (var id in EntryIds(f, Dasl.Like(NameRelatedProperty, "[%"), NameRelatedProperty + "/0x0000001F"))
                    {
                        var value = (id.Item2 ?? "").Trim();
                        if (!value.EndsWith("]", StringComparison.Ordinal) || updates.ContainsKey(id.Item1)
                            || ComputesTo(session, id.Item1, f.StoreId, value))
                            continue; // e.g. a sender whose own display name is in brackets
                        Fix(updates, id.Item1, f.StoreId).Recompute = true;
                        nameCount++;
                    }
                }
            }
            Cursor.Current = Cursors.Default;
            if (changes.Count == 0 && updates.Count == 0 && contact == settings.ContactMark)
                return; // nothing chosen differently, and every nameRelated is already in the chosen form
            if (updates.Count > 0 && MessageBox.Show(WindowOwner.From(explorer),
                    "Change " + updates.Count + " mail(s)?\n"
                    + (meCount > 0 ? "\n- the me column of " + meCount + " mail(s) to the new symbol(s)" : "")
                    + (nameCount > 0 ? "\n- the nameRelated of " + nameCount + " mail(s) to the marks \"" + sent + " \" (sent) and \"" + contact + " \" (Contacts)" : "")
                    + "\n\n(No: only mail filled from now on gets them.)",
                    ThisAddIn.Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                updates.Clear();

            settings.ContactMark = contact;
            settings.MeSent = sent;
            settings.MeTo = to;
            settings.MeCc = cc;
            SaveSettings(ItemKind.Mail);
            var filters = changes.Count > 0 ? ChangeMeInSavedFilters(changes) : 0;

            int done = 0, failed = 0;
            Cursor.Current = Cursors.WaitCursor;
            foreach (var u in updates)
            {
                try
                {
                    dynamic item = session.GetItemFromID(u.Key, u.Value.StoreId);
                    bool changed = false;
                    if (u.Value.Me != null)
                    {
                        var prop = item.UserProperties.Find(CustomFieldNames.Me);
                        if (prop != null) { prop.Value = u.Value.Me; changed = true; }
                    }
                    if (u.Value.Recompute)
                    {
                        // after the settings are saved above, so the computed name has the chosen mark
                        CustomFieldValues values = Fields.Compute((object)item);
                        var prop = item.UserProperties.Find(CustomFieldNames.NameRelated);
                        if (values != null && prop != null && !string.Equals(prop.Value as string, values.NameRelated, StringComparison.Ordinal))
                        {
                            prop.Value = values.NameRelated;
                            changed = true;
                        }
                    }
                    else if (u.Value.NameRelated != null)
                    {
                        var prop = item.UserProperties.Find(CustomFieldNames.NameRelated);
                        if (prop != null) { prop.Value = u.Value.NameRelated; changed = true; }
                    }
                    if (changed)
                    {
                        item.Save();
                        done++;
                    }
                    Marshal.ReleaseComObject(item);
                }
                catch (Exception ex)
                {
                    failed++;
                    Log.Info("me symbols: an item not changed: " + ex.Message);
                }
                if ((done + failed) % 25 == 0)
                    Application.DoEvents();
            }
            Cursor.Current = Cursors.Default;
            var summary = "Column symbols now: contact " + contact + ", sent " + sent + ", To " + to + ", Cc " + cc + ". " + done + " mail(s) changed"
                + (failed > 0 ? ", " + failed + " could not be" : "") + "; " + filters + " saved filter(s) updated.";
            Log.Info(summary);
            Notifier.Info(explorer, summary);
        }

        /// <summary>The name in a sent mail's nameRelated that has no contact mark ("→ name"); null otherwise.</summary>
        private static string SentNameWithoutMark(string value, FilterSettings settings)
        {
            var sentSymbol = MeSymbols.SentChoices.Select(c => c.Text).Concat(new[] { settings.MeSent })
                .FirstOrDefault(x => value.StartsWith(x, StringComparison.Ordinal));
            if (sentSymbol == null)
                return null;
            var rest = value.Substring(sentSymbol.Length).Trim();
            var marks = MeSymbols.KnownContactMarks.Concat(new[] { settings.ContactMark });
            return rest.Length == 0 || marks.Any(m => rest.StartsWith(m, StringComparison.Ordinal)) ? null : rest;
        }

        /// <summary>
        /// A nameRelated value in the chosen form, or null when it does not start with a sent symbol or a contact mark: the
        /// sent symbol (any) becomes "<paramref name="sent"/> ", and a contact mark (any) right after it, or at the start of
        /// a received mail's value, becomes <paramref name="contact"/> and one space before the name.
        /// </summary>
        private static string NameRelatedIn(string value, FilterSettings settings, string sent, string contact)
        {
            var sentSymbols = MeSymbols.SentChoices.Select(c => c.Text).Concat(new[] { settings.MeSent }).ToList();
            var marks = MeSymbols.KnownContactMarks.Concat(new[] { settings.ContactMark }).ToList();
            var rest = value;
            var sentSymbol = sentSymbols.FirstOrDefault(x => rest.StartsWith(x, StringComparison.Ordinal));
            if (sentSymbol != null)
                rest = rest.Substring(sentSymbol.Length).TrimStart();
            var mark = marks.FirstOrDefault(x => rest.StartsWith(x, StringComparison.Ordinal));
            if (mark != null)
                rest = contact + " " + rest.Substring(mark.Length).TrimStart();
            if (sentSymbol == null && mark == null)
                return null;
            return sentSymbol != null ? sent + " " + rest : rest;
        }

        // Whether the mail's nameRelated computed now is the value it has (then there is nothing to recompute).
        private bool ComputesTo(Outlook.NameSpace session, string entryId, string storeId, string value)
        {
            object item = null;
            try
            {
                item = session.GetItemFromID(entryId, storeId);
                var values = Fields.Compute(item);
                return values != null && string.Equals(values.NameRelated, value, StringComparison.Ordinal);
            }
            catch (COMException)
            {
                return true; // gone or not reachable: leave it
            }
            finally
            {
                if (item != null)
                    Marshal.ReleaseComObject(item);
            }
        }

        private const string NameRelatedProperty = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/nameRelated";

        /// <summary>What one mail gets: a new me value and/or a new nameRelated (null: unchanged).</summary>
        private sealed class MailFix
        {
            public string StoreId;
            public string Me;
            public string NameRelated;
            public bool Recompute; // nameRelated in the old "[name]" form: computed again
        }

        private static MailFix Fix(Dictionary<string, MailFix> updates, string entryId, string storeId)
        {
            MailFix fix;
            if (!updates.TryGetValue(entryId, out fix))
                updates[entryId] = fix = new MailFix { StoreId = storeId };
            return fix;
        }

        // The entry ids (and the given column's value, when one is asked for) of a folder's items matching a filter.
        private static List<Tuple<string, string>> EntryIds(Search.FindFolder f, string filter, string column)
        {
            var found = new List<Tuple<string, string>>();
            try
            {
                var table = f.Folder.GetTable("@SQL=" + filter, Outlook.OlTableContents.olUserItems);
                table.Columns.RemoveAll();
                table.Columns.Add("EntryID");
                if (column != null)
                    table.Columns.Add(column);
                while (!table.EndOfTable)
                {
                    var row = table.GetNextRow();
                    found.Add(Tuple.Create((string)row[1], column == null ? null : row[2] as string));
                }
            }
            catch (Exception ex)
            {
                Log.Info("me symbols: " + f.Path + " skipped: " + ex.Message);
            }
            return found;
        }

        // In every kind's saved filters: "me" = 'old' becomes "me" = 'new' (all at once, so swapped symbols do not mix).
        private int ChangeMeInSavedFilters(IDictionary<string, string> changes)
        {
            int count = 0;
            foreach (ItemKind kind in Enum.GetValues(typeof(ItemKind)))
            {
                var state = State(kind);
                bool changed = false;
                foreach (var f in state.Settings.Filters)
                {
                    if (string.IsNullOrEmpty(f.Sql))
                        continue;
                    var sql = f.Sql;
                    var marks = new Dictionary<string, string>();
                    int n = 0;
                    foreach (var change in changes)
                    {
                        var mark = "\u0001" + (n++) + "\u0001";
                        sql = sql.Replace(Dasl.PropertyEquals(MeProperty, change.Key), mark);
                        marks[mark] = Dasl.PropertyEquals(MeProperty, change.Value);
                    }
                    foreach (var m in marks)
                        sql = sql.Replace(m.Key, m.Value);
                    if (sql != f.Sql)
                    {
                        f.Sql = sql;
                        changed = true;
                        count++;
                    }
                }
                if (changed)
                    SaveSettings(kind);
            }
            return count;
        }

        /// <summary>Custom Shortcuts window: the add-in's keyboard shortcuts, each turned on or off.</summary>
        public void ShowCustomShortcuts(Outlook.Explorer explorer)
        {
            using (var form = new CustomShortcutsForm(Globals.ThisAddIn.Shortcuts))
            {
                form.Icon = OfficeImage.IconFromImageMso(explorer.CommandBars, "AdpPrimaryKey");
                form.ShowDialog(WindowOwner.From(explorer));
            }
        }

        private Search.FindItemsForm _findItems;

        /// <summary>Find Items: one window for the whole session; pressing the button again brings it to the front.</summary>
        public void ShowFindItems(Outlook.Explorer explorer)
        {
            if (_findItems != null && !_findItems.IsDisposed)
            {
                WindowOwner.Activate(_findItems, explorer);
                return;
            }
            _findItems = new Search.FindItemsForm(explorer.Application);
            _findItems.Icon = OfficeImage.IconFromImageMso(explorer.CommandBars, "ResearchPane");
            WindowOwner.ShowCentred(_findItems, explorer);
        }

        public void Clear(Outlook.Explorer explorer)
        {
            if (!Views.Clear(explorer))
                throw new UserMessageException("The current folder is not shown in a table (list) view.");
            _inputText = "";
            Invalidate();
        }

        // ---------- Saved filters ----------

        /// <summary>
        /// Saved filters left out of the Add to menu: the first-install mail filters, which are computed from the mail
        /// itself (flag status, me, nameRelated) rather than lists of domains or subjects.
        /// </summary>
        private static readonly HashSet<string> NotAddTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Flagged", "Sent", "Unknown",
        };

        /// <summary>The saved filters the Add to menu lists (placeholders without SQL included), in file order.</summary>
        public List<string> AddToTargets
        {
            get
            {
                return _kind != ItemKind.Mail ? new List<string>()
                    : Settings.Filters.Select(f => f.Name).Where(n => !NotAddTargets.Contains(n))
                        .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            }
        }

        /// <summary>Whether Add/New... lists this filter (mail folders; not Flagged, Sent or Unknown).</summary>
        public bool IsAddTarget(string name)
        {
            return AddToTargets.Contains(name, StringComparer.OrdinalIgnoreCase);
        }

        private static readonly string DomainRelatedProperty = CustomFieldNames.Dasl(CustomFieldNames.DomainRelated);

        private const string SubjectProperty = "http://schemas.microsoft.com/mapi/proptag/0x0037001F"; // PR_SUBJECT

        public bool IsSavedActive(Outlook.Explorer explorer, int index)
        {
            var f = FilterAt(index);
            return f != null && Views.ActiveSource(explorer) == SavedSourcePrefix + f.Name;
        }

        public void ToggleSaved(Outlook.Explorer explorer, int index, bool on)
        {
            var f = FilterAt(index);
            if (f == null)
                return;
            if (on)
            {
                Views.Apply(explorer, Dasl.StripSqlPrefix(f.Sql), SavedSourcePrefix + f.Name);
            }
            else
            {
                Views.Clear(explorer);
            }
            Invalidate();
        }

        // ---------- Others: the mail no saved filter matches ----------

        private const string OthersSource = "others";

        /// <summary>
        /// NOT ((filter 1) OR (filter 2) ...) over the saved filters shown on the ribbon (placeholders without SQL are
        /// left out); null when there are none.
        /// </summary>
        private static string OthersSql(FilterSettings settings)
        {
            var parts = settings.Filters.Where(f => !string.IsNullOrWhiteSpace(f.Sql)).Take(MaxSavedFilters)
                .Select(f => "(" + Dasl.StripSqlPrefix(f.Sql).Trim() + ")").ToList();
            return parts.Count == 0 ? null : "NOT (" + string.Join(" OR ", parts) + ")";
        }

        public bool IsOthersVisible
        {
            get { return OthersSql(Settings) != null; }
        }

        public bool IsOthersActive(Outlook.Explorer explorer)
        {
            return Views.ActiveSource(explorer) == OthersSource;
        }

        /// <summary>Others toggle: shows only the items none of the saved filters match; off restores the view's filter.</summary>
        public void ToggleOthers(Outlook.Explorer explorer, bool on)
        {
            ReloadIfChanged(explorer); // use the saved filters as they are in the file now
            var sql = OthersSql(Settings);
            if (on && sql != null)
                Views.Apply(explorer, sql, OthersSource);
            else
                Views.Clear(explorer);
            Invalidate();
        }

        /// <summary>The saved filter currently applied to the explorer's view, or null.</summary>
        public SavedFilter ActiveSavedFilter(Outlook.Explorer explorer)
        {
            var source = Views.ActiveSource(explorer);
            if (source == null || !source.StartsWith(SavedSourcePrefix, StringComparison.Ordinal))
                return null;
            return Settings.Find(source.Substring(SavedSourcePrefix.Length));
        }

        public bool IsActiveFormatEnabled(Outlook.Explorer explorer)
        {
            var f = ActiveSavedFilter(explorer);
            return f != null && f.FormatEnabled && f.Format != null && Settings.FormatsOn;
        }

        /// <summary>The saved filters whose formats go into the views (none while all formats are off).</summary>
        private IList<SavedFilter> FormattedFilters
        {
            get { return Settings.FormatsOn ? VisibleFilters : new List<SavedFilter>(); }
        }

        /// <summary>Turns the conditional formatting of the applied saved filter on or off.</summary>
        public void SetActiveFormatEnabled(Outlook.Explorer explorer, bool on)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            var f = RequireActiveSavedFilter(explorer, "Format turns the conditional formatting of the applied saved filter on or off.");
            if (on && f.Format == null && !EditFormat(explorer, f))
            {
                Invalidate(); // cancelled: un-press the toggle
                return;
            }
            f.FormatEnabled = on;
            if (on)
                Settings.FormatsOn = true; // turning one on also ends "All Formats Off"
            SaveSettings();
            KeepSelection(explorer, () => SyncFormats(explorer, true));
        }

        /// <summary>
        /// Format button (toggle): Ctrl+click or no format defined yet → the Format dialog (a new format is turned on
        /// right away); otherwise it turns the applied saved filter's format on or off.
        /// </summary>
        public void FormatButton(Outlook.Explorer explorer, bool pressed)
        {
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                Invalidate(); // Ctrl+click only edits: keep the toggle state as it was
                EditActiveFormat(explorer);
                return;
            }
            var f = RequireActiveSavedFilter(explorer, "Format turns the conditional formatting of the applied saved filter on or off.");
            if (f.Format == null)
            {
                SetActiveFormatEnabled(explorer, true); // opens the dialog first; cancelled → stays off
                return;
            }
            SetActiveFormatEnabled(explorer, pressed);
        }

        /// <summary>Format dialog: edits the format (style, strikeout, underline, color) of the applied saved filter.</summary>
        public void EditActiveFormat(Outlook.Explorer explorer)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            var f = RequireActiveSavedFilter(explorer, "Format edits the conditional format of the applied saved filter.");
            if (!EditFormat(explorer, f))
                return;
            SaveSettings();
            if (f.FormatEnabled)
                KeepSelection(explorer, () => SyncFormats(explorer, true));
            Invalidate();
        }

        private SavedFilter RequireActiveSavedFilter(Outlook.Explorer explorer, string purpose)
        {
            var f = ActiveSavedFilter(explorer);
            if (f == null)
                throw new UserMessageException("Apply a saved filter first.\n" + purpose);
            return f;
        }

        /// <summary>Shows the format dialog for <paramref name="f"/>; false if cancelled.</summary>
        private static bool EditFormat(Outlook.Explorer explorer, SavedFilter f)
        {
            using (var dialog = new FormatDialog(f.Name, f.Format))
            {
                if (dialog.ShowDialog(WindowOwner.From(explorer)) != DialogResult.OK)
                    return false;
                f.Format = dialog.Result;
                return true;
            }
        }

        /// <summary>
        /// Add to &gt; filter: a dialog chooses domainRelated, From address, subject, or several of them; subjects become
        /// patterns in which numbers, dates and month/weekday names are % (editable), and each line is added to the saved
        /// filter as a condition (several fields: all of them must match). The first time in a session "Issue" starts with
        /// domainRelated and every other filter with Subject; later the last choice for that filter.
        /// </summary>
        public void AddSelectionToFilter(Outlook.Explorer explorer, string filterName)
        {
            if (_kind != ItemKind.Mail)
                throw new UserMessageException("\"Add to\" works in mail folders.");
            var mails = SelectedValues(explorer, item => DomainRelatedOf(item) + "\t" + SenderAddressOf(item) + "\t" + Dasl.SubjectPattern(SubjectOf(item)))
                .Select(v => v.Split(new[] { '\t' }, 3))
                .Select(p => new MailKey
                {
                    Domain = p[0].Trim(),
                    From = p.Length > 1 ? p[1].Trim() : "",
                    Pattern = p.Length > 2 ? p[2].Trim() : "",
                })
                .Where(m => m.Domain.Length > 0 || m.From.Length > 0 || m.Pattern.Length > 0).ToList();
            if (mails.Count == 0)
                throw new UserMessageException("Select the mails to add to \"" + filterName + "\" first.");

            var defaultField = string.Equals(filterName, "Issue", StringComparison.OrdinalIgnoreCase)
                ? ConditionField.DomainRelated : ConditionField.Subject;
            ConditionField field;
            List<MailKey> values;
            using (var dialog = new AddConditionsDialog(filterName, mails, defaultField))
            {
                if (dialog.ShowDialog(WindowOwner.From(explorer)) != DialogResult.OK)
                    return;
                field = dialog.Field;
                values = dialog.Values;
            }
            var fields = AddConditionsDialog.Fields.Where(f => (field & f) != 0).ToList();
            var what = string.Join(" and ", fields.Select(f => f == ConditionField.DomainRelated ? "domainRelated" : f == ConditionField.From ? "From address" : "subject"));
            AddSelectionTo(explorer, filterName, what, values, v =>
            {
                var parts = fields.Select(f => ConditionFor(f, v)).ToList();
                return parts.Count == 1 ? parts[0] : "(" + string.Join(" AND ", parts) + ")";
            }, null);
        }

        /// <summary>
        /// Add/New... > New...: asks for the name of a new saved filter, then builds it from the selected mails as Add to
        /// does (the conditions window); the filter is created only when conditions are added.
        /// </summary>
        public void NewFilterFromSelection(Outlook.Explorer explorer)
        {
            if (_kind != ItemKind.Mail)
                throw new UserMessageException("New saved filters from mails work in mail folders.");
            Outlook.Selection selection = null;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
            }
            if (selection == null || selection.Count == 0)
                throw new UserMessageException("Select the mails for the new saved filter first.");

            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            var name = Microsoft.VisualBasic.Interaction.InputBox(
                "Name of the new saved filter (its ribbon button label):", ThisAddIn.Title + " – New saved filter", "").Trim();
            if (name.Length == 0)
                return; // cancelled
            if (Settings.Find(name) != null)
                throw new UserMessageException("A saved filter named \"" + name + "\" already exists. Use Add/New... > " + name + " to add to it.");
            AddSelectionToFilter(explorer, name);
        }

        private static string ConditionFor(ConditionField field, MailKey v)
        {
            switch (field)
            {
                case ConditionField.DomainRelated:
                    return Dasl.PropertyEquals(DomainRelatedProperty, v.Domain);
                case ConditionField.From:
                    // the sender's SMTP address, or its raw address (SMTP, or an Exchange DN) as the F quick filter does
                    return "(" + Dasl.PropertyEquals(Dasl.SenderSmtp, v.From) + " OR " + Dasl.PropertyEquals(Dasl.SenderEmail, v.From) + ")";
                default:
                    return Dasl.PropertyMatches(SubjectProperty, v.Pattern);
            }
        }

        /// <summary>The sender's SMTP address (PR_SENDER_SMTP_ADDRESS), else the raw sender address; "" if none.</summary>
        private static string SenderAddressOf(object item)
        {
            dynamic d = item;
            try
            {
                var smtp = ((Outlook.PropertyAccessor)d.PropertyAccessor).GetProperty(Dasl.SenderSmtp) as string;
                if (!string.IsNullOrWhiteSpace(smtp))
                    return smtp.Trim();
            }
            catch (COMException)
            {
                // property not present
            }
            try
            {
                return ((string)d.SenderEmailAddress ?? "").Trim();
            }
            catch (Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Appends one OR-ed condition per value to the saved filter <paramref name="filterName"/> (created if missing,
        /// right after <paramref name="insertAfter"/> when given). Values the filter already covers are skipped.
        /// </summary>
        private void AddSelectionTo<T>(Outlook.Explorer explorer, string filterName, string what, List<T> values,
            Func<T, string> clauseFor, string insertAfter)
        {
            if (_kind != ItemKind.Mail)
                throw new UserMessageException("\"Add to " + filterName + "\" works in mail folders.");
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            if (values.Count == 0)
                throw new UserMessageException("Select the mails whose " + what + " should be added to \"" + filterName + "\".");

            var f = Settings.Find(filterName);
            if (f == null)
            {
                if (Settings.Filters.Count >= MaxSavedFilters)
                    throw new UserMessageException("The ribbon shows at most " + MaxSavedFilters + " saved filters. Remove one in " + SettingsPaths.SavedFiltersName(_kind) + " first.");
                f = new SavedFilter { Name = filterName, Sql = "" };
                var after = insertAfter == null ? null : Settings.Find(insertAfter);
                if (after != null)
                    Settings.Filters.Insert(Settings.Filters.IndexOf(after) + 1, f);
                else
                    Settings.Filters.Add(f);
            }

            var folder = explorer.CurrentFolder;
            int added = 0;
            foreach (var value in values)
            {
                var clause = clauseFor(value);
                var sql = Dasl.StripSqlPrefix(f.Sql);
                if (sql.Length > 0 && Covers(folder, sql, clause))
                    continue;
                f.Sql = sql.Length == 0 ? clause : sql + " OR\n" + LastLineIndent(sql) + clause;
                added++;
            }
            if (added == 0)
                throw new UserMessageException("\"" + filterName + "\" already includes the selected " + what
                    + (values.Count > 1 ? " values." : "."));

            SaveSettings();
            // Re-applying the view resets the selection; keep it so the next "Add to ..." acts on
            // what the user sees selected, not on whatever Outlook selects after the refresh.
            KeepSelection(explorer, () =>
            {
                if (Views.ActiveSource(explorer) == SavedSourcePrefix + f.Name)
                    Views.Apply(explorer, Dasl.StripSqlPrefix(f.Sql), SavedSourcePrefix + f.Name);
                if (f.FormatEnabled && f.Format != null)
                    SyncFormats(explorer, true);
            });
            Invalidate();
        }

        /// <summary>Runs <paramref name="action"/> (which may re-apply the view) and restores the selection afterwards.</summary>
        private static void KeepSelection(Outlook.Explorer explorer, Action action)
        {
            var ids = new List<KeyValuePair<string, string>>();
            try
            {
                var selection = explorer.Selection;
                for (int i = 1; i <= selection.Count; i++)
                {
                    object item = selection[i];
                    try
                    {
                        dynamic d = item;
                        ids.Add(new KeyValuePair<string, string>((string)d.EntryID, ((Outlook.MAPIFolder)d.Parent).StoreID));
                    }
                    catch (Exception)
                    {
                        // item without an EntryID (e.g. a conversation header)
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(item);
                    }
                }
            }
            catch (COMException)
            {
                // no selection in this pane
            }

            action();

            if (ids.Count > 0)
                RestoreSelectionLater(explorer, ids);
        }

        /// <summary>
        /// View.Apply refreshes the list asynchronously and selects its own item afterwards, so the selection is
        /// restored from a UI-thread timer, retrying until it sticks (about 3 seconds at most).
        /// </summary>
        private static void RestoreSelectionLater(Outlook.Explorer explorer, List<KeyValuePair<string, string>> ids)
        {
            const int maxTries = 10;
            int tries = 0;
            var timer = new Timer { Interval = 300 };
            timer.Tick += (s, e) =>
            {
                tries++;
                bool done = true;
                try
                {
                    if (!SelectionIs(explorer, ids))
                    {
                        explorer.ClearSelection();
                        foreach (var id in ids)
                        {
                            var item = explorer.Session.GetItemFromID(id.Key, id.Value);
                            if (explorer.IsItemSelectableInView(item))
                                explorer.AddToSelection(item);
                        }
                        done = SelectionIs(explorer, ids);
                    }
                }
                catch (Exception ex)
                {
                    Log.Error("RestoreSelection", ex);
                }
                // Keep checking a few more ticks even after success: a late refresh may still reselect.
                if ((done && tries >= 3) || tries >= maxTries)
                {
                    timer.Stop();
                    timer.Dispose();
                }
            };
            timer.Start();
        }

        private static bool SelectionIs(Outlook.Explorer explorer, List<KeyValuePair<string, string>> ids)
        {
            var selection = explorer.Selection;
            if (selection.Count != ids.Count)
                return false;
            var wanted = new HashSet<string>(ids.Select(i => i.Key));
            for (int i = 1; i <= selection.Count; i++)
            {
                object item = selection[i];
                try
                {
                    if (!wanted.Contains((string)((dynamic)item).EntryID))
                        return false;
                }
                catch (Exception)
                {
                    return false;
                }
                finally
                {
                    Marshal.ReleaseComObject(item);
                }
            }
            return true;
        }

        /// <summary>Leading whitespace of the last line, so an appended condition lines up with the previous one.</summary>
        private static string LastLineIndent(string sql)
        {
            var last = sql.Substring(sql.LastIndexOf('\n') + 1);
            return last.Substring(0, last.Length - last.TrimStart(' ', '\t').Length);
        }

        /// <summary>True if some item of the folder matches both the filter and the clause.</summary>
        private static bool Covers(Outlook.MAPIFolder folder, string sql, string clause)
        {
            try
            {
                return folder != null && folder.Items.Restrict("@SQL=(" + sql + ") AND " + clause).Count > 0;
            }
            catch (COMException)
            {
                return false; // unparsable existing SQL: just append
            }
        }

        /// <summary>Distinct non-empty values of the selected items, in selection order.</summary>
        private static List<string> SelectedValues(Outlook.Explorer explorer, Func<object, string> valueOf)
        {
            var values = new List<string>();
            Outlook.Selection selection;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
                return values;
            }
            for (int i = 1; i <= selection.Count; i++)
            {
                object item = selection[i];
                try
                {
                    string value = null;
                    try
                    {
                        value = valueOf(item);
                    }
                    catch (Exception)
                    {
                        // item type without that value
                    }
                    if (!string.IsNullOrWhiteSpace(value) && !values.Contains(value))
                        values.Add(value);
                }
                finally
                {
                    Marshal.ReleaseComObject(item);
                }
            }
            return values;
        }

        private static string SubjectOf(object item)
        {
            return ((dynamic)item).Subject as string;
        }

        private string DomainRelatedOf(object item)
        {
            return Fields == null ? null : Fields.DomainRelatedOf(item);
        }

        // ---------- Conditional formatting ----------

        /// <summary>
        /// Brings the current view's formatting rules in line with the saved filters (rewrites them all when
        /// <paramref name="force"/>). <paramref name="reason"/> goes to the log with each rewrite (the caller by default).
        /// </summary>
        public void SyncFormats(Outlook.Explorer explorer, bool force, [CallerMemberName] string reason = "")
        {
            if (_syncing)
                return;
            _syncing = true;
            try
            {
                if (!ConditionalFormatService.Sync(explorer, FormattedFilters, force, Views.OwnFilter(explorer), reason) && force)
                    throw new UserMessageException("Conditional formatting needs a table (list) view.");
            }
            finally
            {
                _syncing = false;
                Invalidate();
            }
        }

        /// <summary>View Font button: font name and size for the whole current table view.</summary>
        public void EditViewFont(Outlook.Explorer explorer)
        {
            var current = ConditionalFormatService.GetViewFont(explorer);
            if (current == null)
                throw new UserMessageException("View Font needs a table (list) view.");
            using (var dialog = new ViewFontDialog(explorer.CurrentView is Outlook.View v ? v.Name : "", current.Item1, current.Item2))
            {
                if (dialog.ShowDialog(WindowOwner.From(explorer)) != DialogResult.OK)
                    return;
                KeepSelection(explorer, () => ConditionalFormatService.SetViewFont(explorer, dialog.FontName, dialog.FontSize,
                    dialog.ApplyToHeaders, dialog.ApplyToRules, Views.OwnFilter(explorer)));
            }
        }

        /// <summary>Sort by Company/Dept (contact folders): Company, then Department, both ascending; saved in the view.</summary>
        public void SortByCompany(Outlook.Explorer explorer)
        {
            KeepSelection(explorer, () => ViewColumnsService.SortBy(explorer, Views.OwnFilter(explorer), "Company", "Department"));
        }

        /// <summary>
        /// Apply Predefined Columns button: switches to the TinyKit view and gives it the columns of the View Columns file of
        /// the folder's kind (mail, contacts, tasks; created with defaults on first use). Ctrl+click opens that file instead.
        /// </summary>
        public void ViewColumnsButton(Outlook.Explorer explorer)
        {
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                OpenInEditor(ColumnsFile(explorer));
                return;
            }
            ApplyPredefinedColumns(explorer);
        }

        // The View Columns file of the folder's kind, made with the default columns when it does not exist yet.
        private string ColumnsFile(Outlook.Explorer explorer)
        {
            var kind = Focus(explorer);
            if (kind == null)
                throw new UserMessageException("Apply Predefined Columns works in mail, contact and task folders.");
            var path = SettingsPaths.ViewColumnsFile(kind.Value);
            if (!File.Exists(path))
                ViewColumns.CreateDefault(path, kind.Value);
            return path;
        }

        /// <summary>Switches to the TinyKit view and gives it the columns of the folder kind's View Columns file.</summary>
        public void ApplyPredefinedColumns(Outlook.Explorer explorer)
        {
            var path = ColumnsFile(explorer);
            var name = Path.GetFileName(path);

            List<ViewColumn> columns;
            try
            {
                columns = ViewColumns.Load(path);
            }
            catch (FormatException ex)
            {
                throw new UserMessageException(name + ": " + ex.Message + "\n\nCtrl+click Apply Predefined Columns to edit the file.");
            }
            if (columns.Count == 0)
                throw new UserMessageException(name + " has no columns (every line is empty or a # comment).\n\nCtrl+click Apply Predefined Columns to edit the file.");

            // The columns go into the TinyKit table view (made if needed, also from a card or other non-table view); a quick
            // or saved filter shown now is cleared first, as switching views would drop it anyway.
            if (Views.HasAddinFilter(explorer, true))
                Views.Clear(explorer, true);
            List<string> problems = null;
            KeepSelection(explorer, () => problems = ViewColumnsService.Apply(explorer, columns, Views.OwnFilter(explorer)));
            Invalidate();
            if (problems != null && problems.Count > 0)
                MessageBox.Show(WindowOwner.From(explorer),
                    "The other columns were applied, but:\n\n" + string.Join("\n", problems) + "\n\nCtrl+click Apply Predefined Columns to edit " + path + ".",
                    ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        /// <summary>Automatic column sizing of the current table view; null when the folder is not shown as a table.</summary>
        public bool? AutomaticColumnSizing(Outlook.Explorer explorer)
        {
            return ViewColumnsService.GetAutomaticColumnSizing(explorer);
        }

        /// <summary>
        /// Turns Automatic column sizing on or off. Turned off in the TinyKit view, the predefined columns are applied again,
        /// so the columns get the widths of the View Columns file (not whatever automatic sizing had made of them).
        /// </summary>
        public void SetAutomaticColumnSizing(Outlook.Explorer explorer, bool on)
        {
            KeepSelection(explorer, () => ViewColumnsService.SetAutomaticColumnSizing(explorer, on, Views.OwnFilter(explorer)));
            var view = explorer.CurrentView as Outlook.View;
            if (!on && view != null && view.Name == ViewColumnsService.ViewName)
                ApplyPredefinedColumns(explorer);
            Invalidate();
        }

        /// <summary>Kept in the mail saved filters file (sendDelaySeconds): seconds a sent mail waits in the Outbox.</summary>
        public int SendDelaySeconds
        {
            get { return State(ItemKind.Mail).Settings.SendDelaySeconds; }
        }

        public void SetSendDelaySeconds(int seconds)
        {
            ReloadIfChanged(ItemKind.Mail, _activeExplorer());
            State(ItemKind.Mail).Settings.SendDelaySeconds = seconds;
            SaveSettings(ItemKind.Mail);
        }

        /// <summary>
        /// Application.ItemSend: a mail waits <see cref="SendDelaySeconds"/> in the Outbox (Do not deliver before), so it can
        /// still be opened there, changed and sent again, or deleted. A later delivery time already set is kept.
        /// </summary>
        public void DelaySending(object item)
        {
            var mail = item as Outlook.MailItem;
            var seconds = SendDelaySeconds;
            if (mail == null || seconds <= 0)
                return;
            var at = DateTime.Now.AddSeconds(seconds);
            var set = mail.DeferredDeliveryTime;
            if (set.Year < 4500 && set >= at)
                return;
            mail.DeferredDeliveryTime = at;
            Log.Info("Send delay: \"" + mail.Subject + "\" goes at " + at.ToString("HH:mm:ss"));
        }

        /// <summary>
        /// All Formats Off: takes every saved filter's format out of the views. Each filter's Format on/off setting is
        /// kept, so Refresh Formats brings the same formats back.
        /// </summary>
        public void AllFormatsOff(Outlook.Explorer explorer)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            if (!Settings.FormatsOn)
                throw new UserMessageException("All formats are already off. Refresh Formats turns them back on.");
            Settings.FormatsOn = false;
            SaveSettings();
            KeepSelection(explorer, () => SyncFormats(explorer, true));
        }

        /// <summary>Refresh Formats: turns the formats on (after All Formats Off) and rewrites them into this view.</summary>
        public void RefreshFormats(Outlook.Explorer explorer)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            if (!Settings.FormatsOn)
            {
                Settings.FormatsOn = true;
                SaveSettings();
            }
            KeepSelection(explorer, () => SyncFormats(explorer, true));
        }

        /// <summary>Remove Formats: deletes the add-in's rules from this view after a warning; settings are unchanged.</summary>
        public void RemoveFormats(Outlook.Explorer explorer)
        {
            var answer = MessageBox.Show(WindowOwner.From(explorer),
                "Remove all of this add-in's conditional formatting rules ([TK] ...) from the current view?\n\n"
                + "Your own rules and Outlook's built-in rules are not touched, and the Format on/off settings of the saved filters stay as they are, "
                + "so the rules come back with Refresh Formats or when the view is next synced (Auto-apply formats).\n"
                + "To stop formatting for good, use All Formats Off or turn the filter's Format off.",
                ThisAddIn.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (answer != DialogResult.OK)
                return;
            if (!KeepSelectionResult(explorer, () => ConditionalFormatService.Remove(explorer, Views.OwnFilter(explorer))))
                throw new UserMessageException("Conditional formatting needs a table (list) view.");
        }

        private static bool KeepSelectionResult(Outlook.Explorer explorer, Func<bool> action)
        {
            bool result = false;
            KeepSelection(explorer, () => result = action());
            return result;
        }

        public bool AutoApplyFormats
        {
            get { return Settings.AutoApplyFormats; }
        }

        public void SetAutoApplyFormats(Outlook.Explorer explorer, bool on)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            Settings.AutoApplyFormats = on;
            SaveSettings();
            if (on)
                SyncFormats(explorer, false);
        }

        /// <summary>
        /// Folder or view switched: switch to the saved filters of the folder's kind, keep formats in sync (quietly)
        /// and refresh the ribbon (which groups show depends on the kind).
        /// </summary>
        public void OnViewChanged(Outlook.Explorer explorer, string trigger = "folder/view switch")
        {
            try
            {
                if (Focus(explorer) != null && Settings.AutoApplyFormats)
                {
                    Outlook.View view;
                    try
                    {
                        view = ViewFilterService.GetTableView(explorer);
                    }
                    catch (UserMessageException)
                    {
                        view = null; // not a table view: nothing to format
                    }
                    if (view != null)
                    {
                        // A rule's Filter set through the object model does not survive an Outlook restart, so each
                        // view's rules are rewritten on its first visit per session; later visits only fix differences.
                        var first = _syncedViews.Add(ViewFilterService.ViewKey(explorer, view));
                        SyncFormats(explorer, first, trigger + (first ? ", first visit" : ""));
                    }
                }
            }
            catch (UserMessageException ex)
            {
                // e.g. Outlook dropped some of the user's own rules while saving (the sync's guard)
                Log.Info("Formats (" + trigger + "): " + ex.Message);
            }
            catch (Exception ex)
            {
                Log.Error("OnViewChanged", ex);
            }
            Invalidate();
        }

        private readonly HashSet<string> _syncedViews = new HashSet<string>(StringComparer.Ordinal);

        // ---------- Custom fields ----------

        private const int ConfirmAbove = 200;

        /// <summary>
        /// Fill Fields button: fills the current folder's items that have no fields yet; Shift+click recomputes the
        /// selected items instead, Ctrl+click removes their fields.
        /// </summary>
        public void FillFieldsButton(Outlook.Explorer explorer)
        {
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                ClearSelectedFields(explorer);
            else if ((Control.ModifierKeys & Keys.Shift) == Keys.Shift)
                FillSelectedFields(explorer);
            else
                FillMissingFields(explorer);
        }

        /// <summary>Fills (recomputes) the custom mail fields of the selected items.</summary>
        public void FillSelectedFields(Outlook.Explorer explorer)
        {
            Outlook.Selection selection;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
                selection = null;
            }
            if (selection == null || selection.Count == 0)
                throw new UserMessageException("Select the mail items to fill first.");

            int count = selection.Count;
            if (!ConfirmMany(explorer, count, "selected"))
                return;
            var items = new List<object>(count);
            for (int i = 1; i <= count; i++)
                items.Add(selection[i]);
            Report(explorer, Fields.FillItems(items));
        }

        /// <summary>Removes the custom mail fields from the selected items (asked first); Fill Fields fills them again.</summary>
        public void ClearSelectedFields(Outlook.Explorer explorer)
        {
            Outlook.Selection selection;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
                selection = null;
            }
            if (selection == null || selection.Count == 0)
                throw new UserMessageException("Select the mail items to clear first.");

            int count = selection.Count;
            if (MessageBox.Show(WindowOwner.From(explorer),
                    "Clear the custom fields (" + string.Join(", ", CustomFieldNames.All) + ") of " + count + " selected item(s)?\n\n"
                    + "They stay empty until filled again (Fill Fields).",
                    ThisAddIn.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                return;
            var items = new List<object>(count);
            for (int i = 1; i <= count; i++)
                items.Add(selection[i]);
            var r = Fields.ClearItems(items);
            var message = "Custom fields: " + r[0] + " cleared, " + r[1] + " had none, " + r[2] + " skipped (not mail)"
                + (r[3] > 0 ? ", " + r[3] + " failed (see OutlookAddin.log)" : "");
            Log.Info(message);
            if (r[3] > 0)
                MessageBox.Show(WindowOwner.From(explorer), message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                Notifier.Info(explorer, message);
        }

        /// <summary>Fills the items of the current folder that have no domainRelated or no domainMark yet.</summary>
        public void FillMissingFields(Outlook.Explorer explorer)
        {
            var folder = explorer.CurrentFolder;
            if (folder == null)
                return;
            var missing = folder.Items.Restrict("@SQL=\"" + CustomFieldNames.Dasl(CustomFieldNames.DomainRelated) + "\" IS NULL OR \""
                + CustomFieldNames.Dasl(CustomFieldNames.DomainMark) + "\" IS NULL");
            int count = missing.Count;
            if (count == 0)
                throw new UserMessageException("Every item in \"" + folder.Name + "\" already has its fields.\n"
                    + "Shift+click Fill Fields to recompute the selected items.");
            if (!ConfirmMany(explorer, count, "unfilled in \"" + folder.Name + "\""))
                return;
            var items = new List<object>(count);
            foreach (var item in missing)
                items.Add(item);
            Report(explorer, Fields.FillItems(items));
        }

        /// <summary>
        /// Add Known Domain: adds the base domains of the selected mails' senders to Known Domains.txt, then refills
        /// the selected mails and this folder's mails from those domains that are still marked unknown.
        /// Ctrl+click opens Known Domains.txt for editing instead.
        /// </summary>
        public void AddKnownDomainButton(Outlook.Explorer explorer)
        {
            var known = Fields.Known;
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                known.EnsureExists();
                OpenInEditor(known.Path);
                return;
            }

            var domains = SelectedValues(explorer, Fields.SenderBaseDomainOf);
            if (domains.Count == 0)
                throw new UserMessageException("Select the mails whose sender's domain should become known.\n"
                    + "(Mail you sent is skipped.) Ctrl+click Add Known Domain to edit " + Path.GetFileName(known.Path) + ".");
            var added = known.Add(domains);

            // The selected mails, and this folder's mail from those domains that is not marked "-" yet.
            var items = new List<object>();
            var selection = explorer.Selection;
            for (int i = 1; i <= selection.Count; i++)
                items.Add(selection[i]);
            var ids = new HashSet<string>(items.Select(it => (string)((dynamic)it).EntryID));
            var domainMark = CustomFieldNames.Dasl(CustomFieldNames.DomainMark);
            var fromDomain = string.Join(" OR ", domains.SelectMany(d => new[]
            {
                Dasl.Like(Dasl.SenderSmtp, "%@" + d), Dasl.Like(Dasl.SenderSmtp, "%." + d),
                Dasl.Like(Dasl.SenderEmail, "%@" + d), Dasl.Like(Dasl.SenderEmail, "%." + d),
            }));
            try
            {
                var more = explorer.CurrentFolder.Items.Restrict("@SQL=(" + fromDomain + ") AND (\"" + domainMark
                    + "\" IS NULL OR \"" + domainMark + "\" <> '" + CustomFieldValues.None + "')");
                foreach (var item in more)
                {
                    if (ids.Add((string)((dynamic)item).EntryID))
                        items.Add(item);
                    else
                        Marshal.ReleaseComObject(item);
                }
            }
            catch (COMException ex)
            {
                Log.Error("AddKnownDomain restrict", ex); // the selection is still refilled
            }
            var result = Fields.FillItems(items);

            var text = (added.Count > 0 ? "Added to known domains: " + string.Join(", ", added) : "Already known: " + string.Join(", ", domains))
                + "\n\nRefilled " + (result.Updated + result.Unchanged) + " mail(s)"
                + (result.Failed > 0 ? ", " + result.Failed + " failed (see OutlookAddin.log)" : "") + ".";
            if (result.Failed > 0)
                MessageBox.Show(WindowOwner.From(explorer), text, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                Notifier.Info(explorer, text);
        }

        /// <summary>Recipients Report: the first selected mail's recipients by domain and department, in a window.</summary>
        public void ShowRecipientsReport(Outlook.Explorer explorer)
        {
            ShowRecipientsReport(FirstSelected(explorer), explorer);
        }

        /// <summary>Recipients Report of the given item, owned by the given Explorer or Inspector.</summary>
        public void ShowRecipientsReport(object item, object ownerWindow)
        {
            if (item == null || ItemView.ForReport(item) == null)
                throw new UserMessageException("Select a mail, meeting request or calendar item first.");
            dynamic d = item;
            string entryId = d.EntryID;
            var app = Globals.ThisAddIn.Application;
            Func<object> read;
            if (string.IsNullOrEmpty(entryId))
                read = () => item; // not in a store (e.g. a .msg file opened from disk)
            else
            {
                string storeId = ((Outlook.MAPIFolder)d.Parent).StoreID;
                read = () => app.Session.GetItemFromID(entryId, storeId);
            }
            // Refresh re-reads the item, so edits (e.g. a contact added meanwhile) show up.
            var form = new Reports.RecipientsReportForm(app, () => Reports.RecipientsReport.Build(read(), Fields.Calculator));
            form.Show(WindowOwner.From(ownerWindow));
        }

        public const string DefaultContactForm = "IPM.Contact";

        // PR_DEF_POST_MSGCLASS: the form a folder's "New" uses (Properties > General > "When posting to this folder, use").
        private const string DefaultPostClassProp = "http://schemas.microsoft.com/mapi/proptag/0x36E5001F";

        /// <summary>"myContactForm" for IPM.Contact.myContactForm; "Contact" for IPM.Contact.</summary>
        public static string FormName(string messageClass)
        {
            int dot = messageClass.LastIndexOf('.');
            return dot < 0 ? messageClass : messageClass.Substring(dot + 1);
        }

        /// <summary>
        /// The custom contact form (IPM.Contact.*) set as the default form of the current contact folder, or else of
        /// the default Contacts folder; null when neither uses one.
        /// </summary>
        public string CustomContactForm(Outlook.Explorer explorer)
        {
            Outlook.MAPIFolder current = null;
            try
            {
                current = explorer.CurrentFolder;
            }
            catch (COMException)
            {
            }
            return CustomFormOf(current)
                ?? CustomFormOf(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts));
        }

        internal static string CustomFormOf(Outlook.MAPIFolder folder)
        {
            if (folder == null || folder.DefaultItemType != Outlook.OlItemType.olContactItem)
                return null;
            try
            {
                var form = folder.PropertyAccessor.GetProperty(DefaultPostClassProp) as string;
                return form != null && form.StartsWith(DefaultContactForm + ".", StringComparison.OrdinalIgnoreCase) ? form : null;
            }
            catch (COMException)
            {
                return null; // not set: the folder uses IPM.Contact
            }
        }

        /// <summary>Sets the message class of the selected contacts, so they open with that form.</summary>
        public void SetContactForm(Outlook.Explorer explorer, string messageClass)
        {
            Outlook.Selection selection;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
                selection = null;
            }
            if (selection == null || selection.Count == 0)
                throw new UserMessageException("Select the contacts first.");

            int changed = 0, already = 0, skipped = 0;
            var selected = new List<string>(); // entry IDs, to select them again after the refresh
            string storeId = null;
            var cursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                for (int i = 1; i <= selection.Count; i++)
                {
                    object item = selection[i];
                    try
                    {
                        var contact = item as Outlook.ContactItem;
                        if (contact == null)
                            skipped++;
                        else
                        {
                            selected.Add(contact.EntryID);
                            if (storeId == null)
                                storeId = ((Outlook.MAPIFolder)contact.Parent).StoreID;
                            if (string.Equals(contact.MessageClass, messageClass, StringComparison.OrdinalIgnoreCase))
                                already++;
                            else
                            {
                                contact.MessageClass = messageClass;
                                contact.Save();
                                changed++;
                            }
                        }
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
                Marshal.ReleaseComObject(selection);
            }
            if (changed > 0)
                RefreshForms(explorer, selected, storeId);
            Notifier.Info(explorer, FormName(messageClass) + " form: " + changed + (changed == 1 ? " contact" : " contacts") + " changed"
                + (already > 0 ? ", " + already + " already had it" : "") + (skipped > 0 ? ", " + skipped + " other items skipped" : "") + ".");
        }

        /// <summary>
        /// Outlook keeps its in-memory copy of an item, with the form it was opened with, while anything refers to it
        /// (the reading pane, the selection), so a changed message class shows only after a restart. Dropping the
        /// selection and visiting a mail folder (other folder types do not always do it) lets Outlook reload the items;
        /// a moment later this comes back and selects the same items again. The window does not paint meanwhile.
        /// </summary>
        private static void RefreshForms(Outlook.Explorer explorer, List<string> entryIds, string storeId)
        {
            Outlook.MAPIFolder folder;
            var noPaint = RedrawLock.Suspend(explorer);
            try
            {
                folder = explorer.CurrentFolder;
                explorer.ClearSelection();
                GC.Collect();
                GC.WaitForPendingFinalizers();
                explorer.CurrentFolder = Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            }
            catch (COMException ex)
            {
                if (noPaint != null)
                    noPaint.Dispose();
                Log.Info("Contact form refresh: could not leave the folder: " + ex.Message);
                return;
            }

            var back = new Timer { Interval = 300 };
            back.Tick += (s, e) =>
            {
                back.Stop();
                back.Dispose();
                try
                {
                    explorer.CurrentFolder = folder;
                    var session = Globals.ThisAddIn.Application.Session;
                    foreach (var id in entryIds)
                    {
                        try
                        {
                            explorer.AddToSelection(session.GetItemFromID(id, storeId));
                        }
                        catch (COMException)
                        {
                            // filtered out of the view, or deleted meanwhile
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Info("Contact form refresh: could not come back: " + ex.Message);
                }
                finally
                {
                    if (noPaint != null)
                        noPaint.Dispose();
                }
            };
            back.Start();
        }

        /// <summary>
        /// Copy Items Text: one line per selected mail, <c>'yy.MM.dd요일 HH:mm &lt;sender&gt; subject</c>, to the clipboard.
        /// The lines follow the order of the view (Outlook's Selection comes in its own order).
        /// Shift+click appends the clipboard's current text after the new lines.
        /// </summary>
        public void CopySelectedItems(Outlook.Explorer explorer)
        {
            bool append = (Control.ModifierKeys & Keys.Shift) == Keys.Shift;
            Outlook.Selection selection;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
                selection = null;
            }
            if (selection == null || selection.Count == 0)
                throw new UserMessageException("Select the mails to copy first.");

            var lines = new List<KeyValuePair<string, string>>(); // record key, line
            int skipped = 0;
            for (int i = 1; i <= selection.Count; i++)
            {
                object item = selection[i];
                try
                {
                    var line = ItemLine(item);
                    if (line == null) skipped++;
                    else
                        lines.Add(new KeyValuePair<string, string>(RecordKeyOf(item), line));
                }
                finally
                {
                    Marshal.ReleaseComObject(item);
                }
            }
            int copied = lines.Count;
            if (copied == 0)
                throw new UserMessageException("Select the mails to copy first. (Contacts and tasks are skipped.)");

            var order = ViewOrder(explorer, lines.Select(l => l.Key));
            if (order != null)
                lines = lines.Select((l, i) => new { l, i })
                    .OrderBy(x => { int at; return x.l.Key != null && order.TryGetValue(x.l.Key, out at) ? at : int.MaxValue; })
                    .ThenBy(x => x.i).Select(x => x.l).ToList();
            var text = new System.Text.StringBuilder();
            foreach (var l in lines)
                text.Append(l.Value).Append(Environment.NewLine);
            if (append && Clipboard.ContainsText())
                text.Append(Clipboard.GetText()).Append(Environment.NewLine);
            Clipboard.SetText(text.ToString());
            Notifier.Info(explorer, copied + (copied == 1 ? " mail" : " mails") + " copied to the clipboard"
                + (append ? ", before its previous text" : "") + (skipped > 0 ? " (" + skipped + " other items skipped)" : "") + ".");
        }

        // PR_RECORD_KEY: the same in a view's table and on the item. (A table's EntryID column can be the short-term
        // ID, e.g. in outlook.com stores, which never equals the item's EntryID.)
        private const string RecordKeyProp = "http://schemas.microsoft.com/mapi/proptag/0x0FF90102";

        private static string RecordKeyOf(object item)
        {
            try
            {
                var props = ((dynamic)item).PropertyAccessor as Outlook.PropertyAccessor;
                return props == null ? null : props.BinaryToString(props.GetProperty(RecordKeyProp)) as string;
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>
        /// The position of each wanted record key in the current view (its sort and filter), reading the view's rows
        /// until all are found; null when the view has no table.
        /// </summary>
        private static Dictionary<string, int> ViewOrder(Outlook.Explorer explorer, IEnumerable<string> recordKeys)
        {
            var wanted = new HashSet<string>(recordKeys.Where(k => k != null), StringComparer.OrdinalIgnoreCase);
            if (wanted.Count == 0)
                return null;
            try
            {
                var view = explorer.CurrentView as Outlook.TableView;
                if (view == null)
                    return null;
                var table = view.GetTable();
                table.Columns.RemoveAll();
                table.Columns.Add(RecordKeyProp);
                var order = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                for (int at = 0; !table.EndOfTable && order.Count < wanted.Count; at++)
                {
                    var key = table.GetNextRow().BinaryToString(RecordKeyProp);
                    if (key != null && wanted.Contains(key) && !order.ContainsKey(key))
                        order[key] = at;
                }
                return order;
            }
            catch (COMException ex)
            {
                Log.Info("Copy Items Text: view order unavailable, using the selection's order: " + ex.Message);
                return null;
            }
        }

        /// <summary>
        /// Open Contact of MailItem: opens the contact that has the address of the person in the first selected
        /// mail's nameRelated (mail I sent: the first recipient; otherwise the sender), from any Contacts folder as
        /// Recipients Report finds them. Without one, a new contact with that name and address is opened (not saved) in the
        /// Contacts folder of the mail's own account, else of the default account.
        /// </summary>
        public void OpenNameRelatedContact(Outlook.Explorer explorer)
        {
            Outlook.Selection selection = explorer.Selection;
            var mail = selection.Count == 0 ? null : selection[1] as Outlook.MailItem;
            if (mail == null)
                throw new UserMessageException("Select a mail first.");
            // The person in nameRelated, resolved as it is: mail I sent -> its first recipient; otherwise the sender (the
            // principal for "on behalf of", Exchange senders by the GAL).
            string name;
            var smtp = Fields.Calculator.NameRelatedSmtpOf(ItemView.From(mail), out name);
            if (string.IsNullOrEmpty(smtp))
                throw new UserMessageException("The person in this mail's nameRelated has no e-mail address.");

            var session = explorer.Application.Session;
            var entry = Fields.Calculator.Contacts.Lookup(smtp);
            if (entry != null)
            {
                ((dynamic)session.GetItemFromID(entry.EntryId, entry.StoreId)).Display();
                return;
            }

            var folder = ContactsFolderOf(mail) ?? session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            // Outlook's own contact form: a new item of a custom form must not be read or written before it is saved.
            var contact = (Outlook.ContactItem)folder.Items.Add(Outlook.OlItemType.olContactItem);
            string first, last;
            Reports.RecipientsReportForm.SplitName(SenderNameFor(name, smtp), out first, out last);
            contact.FirstName = first;
            contact.LastName = last;
            contact.Email1Address = smtp;
            contact.Display();
        }

        // The sender's name, or the part of the address before @ when the name is the address itself.
        private static string SenderNameFor(string name, string smtp)
        {
            name = (name ?? "").Trim().Trim('\'', '"');
            if (name.Length == 0 || name.IndexOf('@') > 0)
                name = smtp.Substring(0, Math.Max(0, smtp.IndexOf('@'))).Replace('.', ' ').Replace('_', ' ');
            return name;
        }

        // The Contacts folder of the account (store) the mail is in; null when that store has none (e.g. a .pst archive).
        private static Outlook.MAPIFolder ContactsFolderOf(Outlook.MailItem mail)
        {
            try
            {
                var parent = mail.Parent as Outlook.MAPIFolder;
                return parent == null ? null : parent.Store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            }
            catch (COMException)
            {
                return null;
            }
        }

        // A mail or meeting request: 'yy.MM.dd요일 HH:mm <sender> subject; null for other items.
        private static string ItemLine(object item)
        {
            if (ItemView.From(item) == null)
                return null;
            dynamic d = item;
            DateTime received = d.ReceivedTime;
            string sender = d.SenderName;
            string subject = d.Subject;
            return "'" + received.ToString("yy.MM.dd", System.Globalization.CultureInfo.InvariantCulture)
                + received.ToString("ddd", System.Globalization.CultureInfo.GetCultureInfo("ko-KR"))
                + " " + received.ToString("HH:mm", System.Globalization.CultureInfo.InvariantCulture)
                + " <" + sender + "> " + subject;
        }

        /// <summary>Kept in the mail saved filters file (clearInboxFiltersOnExit).</summary>
        public bool ClearInboxFiltersOnExit
        {
            get { return State(ItemKind.Mail).Settings.ClearInboxFiltersOnExit; }
        }

        public void SetClearInboxFiltersOnExit(bool on)
        {
            ReloadIfChanged(ItemKind.Mail, _activeExplorer());
            State(ItemKind.Mail).Settings.ClearInboxFiltersOnExit = on;
            SaveSettings(ItemKind.Mail);
        }

        /// <summary>
        /// Outlook is closing (its last window): clears the add-in's filters in the Inbox of every account, as Clear
        /// Filter does, so Outlook opens with unfiltered Inboxes. Does nothing when the option is off.
        /// </summary>
        public void ClearInboxFiltersAtExit(Outlook.Application app, Outlook.Explorer closing)
        {
            if (!ClearInboxFiltersOnExit)
                return;
            var stores = new HashSet<string>(StringComparer.Ordinal);
            int views = 0;
            foreach (Outlook.Account account in app.Session.Accounts)
            {
                try
                {
                    var store = account.DeliveryStore;
                    if (store == null || !stores.Add(store.StoreID))
                        continue;
                    views += Views.ClearFolder(store.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox), closing);
                }
                catch (COMException ex)
                {
                    Log.Error("Clear Inbox filters at exit (" + SafeName(account) + ")", ex);
                }
            }
            if (views > 0)
                Log.Info("Cleared the filters of " + views + " Inbox view(s) at exit.");
        }

        private static string SafeName(Outlook.Account account)
        {
            try
            {
                return account.DisplayName;
            }
            catch (COMException)
            {
                return "?";
            }
        }

        /// <summary>Kept in the mail saved filters file (autoFillFields).</summary>
        public bool AutoFillFields
        {
            get { return State(ItemKind.Mail).Settings.AutoFillFields; }
        }

        public void SetAutoFillFields(bool on)
        {
            ReloadIfChanged(ItemKind.Mail, _activeExplorer());
            State(ItemKind.Mail).Settings.AutoFillFields = on;
            if (Fields != null)
                Fields.AutoFill = on;
            SaveSettings(ItemKind.Mail);
        }

        /// <summary>Kept in the mail saved filters file (moveSentToInbox).</summary>
        public bool MoveSentToInbox
        {
            get { return State(ItemKind.Mail).Settings.MoveSentToInbox; }
        }

        public void SetMoveSentToInbox(bool on)
        {
            ReloadIfChanged(ItemKind.Mail, _activeExplorer());
            State(ItemKind.Mail).Settings.MoveSentToInbox = on;
            if (SentMail != null)
                SentMail.Enabled = on;
            SaveSettings(ItemKind.Mail);
        }

        /// <summary>Kept in the mail saved filters file (moveJunkToInbox).</summary>
        public bool MoveJunkToInbox
        {
            get { return State(ItemKind.Mail).Settings.MoveJunkToInbox; }
        }

        public void SetMoveJunkToInbox(bool on)
        {
            ReloadIfChanged(ItemKind.Mail, _activeExplorer());
            State(ItemKind.Mail).Settings.MoveJunkToInbox = on;
            if (JunkMail != null)
                JunkMail.Enabled = on;
            SaveSettings(ItemKind.Mail);
        }

        private static bool ConfirmMany(Outlook.Explorer explorer, int count, string what)
        {
            return count <= ConfirmAbove
                || MessageBox.Show(WindowOwner.From(explorer),
                    "Fill the custom fields of " + count + " " + what + " items?\nOutlook is busy until it finishes.",
                    ThisAddIn.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK;
        }

        private static void Report(Outlook.Explorer explorer, FillResult result)
        {
            if (result.Failed > 0)
                MessageBox.Show(WindowOwner.From(explorer), "Custom fields: " + result, ThisAddIn.Title,
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                Notifier.Info(explorer, "Custom fields: " + result);
        }

        // ---------- Settings ----------

        /// <summary>Opens the saved filters file of the current kind of folder.</summary>
        public void OpenSettingsFile()
        {
            var path = SettingsPaths.SavedFiltersFile(_kind);
            if (!File.Exists(path))
                SaveSettings();
            OpenInEditor(path);
        }

        /// <summary>Opens a settings file in VS Code, or Notepad when VS Code is not installed.</summary>
        private static void OpenInEditor(string path)
        {
            Process.Start(FindVsCode() ?? "notepad.exe", "\"" + path + "\"");
        }

        /// <summary>Code.exe of a user or machine VS Code install, or the one next to "code.cmd" on PATH; null if none.</summary>
        private static string FindVsCode()
        {
            var candidates = new List<string>
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), @"Programs\Microsoft VS Code\Code.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), @"Microsoft VS Code\Code.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), @"Microsoft VS Code\Code.exe"),
            };
            foreach (var dir in (Environment.GetEnvironmentVariable("PATH") ?? "").Split(';'))
            {
                try
                {
                    if (dir.Length > 0 && File.Exists(Path.Combine(dir, "code.cmd")))
                        candidates.Add(Path.GetFullPath(Path.Combine(dir, @"..\Code.exe")));
                }
                catch (ArgumentException)
                {
                    // malformed PATH entry
                }
            }
            return candidates.FirstOrDefault(File.Exists);
        }

        public void ReloadSettings(Outlook.Explorer explorer)
        {
            var path = SettingsPaths.SavedFiltersFile(_kind);
            FilterSettings loaded;
            try
            {
                loaded = FilterSettings.Load(path, _kind);
            }
            catch (Exception ex)
            {
                throw new UserMessageException(Path.GetFileName(path) + " could not be loaded; the previous saved filters are kept.\n\n" + ex.Message);
            }
            State(_kind).KnownHash = FileHash(path);
            UseSettings(explorer, _kind, loaded);
            var withSql = Settings.Filters.Count(f => !string.IsNullOrWhiteSpace(f.Sql));
            if (withSql > MaxSavedFilters)
                Notifier.Info(explorer, "Only the first " + MaxSavedFilters + " of " + withSql + " saved filters are shown on the ribbon.");
        }

        /// <summary>Saves the current view's filter (e.g. built with View Settings &gt; Filter) as a saved filter.</summary>
        public void SaveCurrentViewFilter(Outlook.Explorer explorer)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            var view = ViewFilterService.RequireTableView(explorer);
            var sql = CurrentFilterText(explorer, view);
            if (sql.Length == 0)
                throw new UserMessageException("The current view has no filter.\n"
                    + "Define one in View Settings > Filter (or apply a quick filter) first.");

            var name = Microsoft.VisualBasic.Interaction.InputBox(
                "Name for this saved filter (shown as its ribbon button):\n\n" + Truncate(sql, 400),
                ThisAddIn.Title, "").Trim();
            if (name.Length == 0)
                return;

            var existing = Settings.Find(name);
            if (existing != null)
            {
                if (MessageBox.Show(WindowOwner.From(explorer), "Replace the SQL of saved filter \"" + name + "\"?",
                        ThisAddIn.Title, MessageBoxButtons.OKCancel, MessageBoxIcon.Question) != DialogResult.OK)
                    return;
                existing.Sql = sql;
            }
            else
            {
                if (Settings.Filters.Count >= MaxSavedFilters)
                    throw new UserMessageException("The ribbon shows at most " + MaxSavedFilters + " saved filters. Remove one in " + SettingsPaths.SavedFiltersName(_kind) + " first.");
                Settings.Filters.Add(new SavedFilter { Name = name, Sql = sql });
            }
            SaveSettings();
            Invalidate();
        }

        /// <summary>
        /// The filter shown in the explorer. Tools that filter through the folder's view object
        /// (Folder.CurrentView / Folder.Views) are covered as well, and what each source returned is logged.
        /// </summary>
        private static string CurrentFilterText(Outlook.Explorer explorer, Outlook.View view)
        {
            var sources = new List<KeyValuePair<string, Func<string>>>
            {
                new KeyValuePair<string, Func<string>>("Explorer.CurrentView", () => view.Filter),
                new KeyValuePair<string, Func<string>>("Folder.CurrentView", () => ((Outlook.View)explorer.CurrentFolder.CurrentView).Filter),
                new KeyValuePair<string, Func<string>>("Folder.Views[name]", () => explorer.CurrentFolder.Views[view.Name].Filter),
            };
            var seen = new List<string>();
            foreach (var source in sources)
            {
                string text;
                try
                {
                    text = (source.Value() ?? "").Trim();
                }
                catch (Exception ex)
                {
                    text = "";
                    seen.Add(source.Key + ": " + ex.Message);
                    continue;
                }
                seen.Add(source.Key + ": " + text.Length + " chars");
                if (text.Length > 0)
                    return text;
            }
            Log.Info("Save View Filter found no filter in view \"" + view.Name + "\" (" + string.Join("; ", seen) + ")");
            return "";
        }

        private void SaveSettings()
        {
            SaveSettings(_kind);
        }

        private void SaveSettings(ItemKind kind)
        {
            var state = State(kind);
            var path = SettingsPaths.SavedFiltersFile(kind);
            state.Settings.Save(path);
            state.KnownHash = FileHash(path);
        }

        public static string Truncate(string s, int max)
        {
            return s == null || s.Length <= max ? s : s.Substring(0, max - 1) + "…";
        }
    }
}
