using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

        public FilterSettings Settings { get; private set; }
        public FilterHistory History { get; private set; }
        public ViewFilterService Views { get; private set; }

        /// <summary>Custom-column stamping; set by ThisAddIn at startup.</summary>
        public CustomFieldService Fields { get; set; }

        /// <summary>Set by the ribbon once loaded; refreshes every control.</summary>
        public Action Invalidate = delegate { };

        public FilterController()
        {
            Directory.CreateDirectory(SettingsPaths.LocalFolder);
            SettingsPaths.MigrateLegacyFiles();
            History = FilterHistory.Load(SettingsPaths.HistoryFile);
            Views = new ViewFilterService(SettingsPaths.ViewStateFile);
            Views.IsAddinFilter = f => !string.IsNullOrWhiteSpace(f)
                && Settings != null && Settings.Filters.Any(s => Dasl.SameFilter(s.Sql, f));
            try
            {
                Settings = FilterSettings.Load(SettingsPaths.SavedFiltersFile);
                _knownFileHash = FileHash(SettingsPaths.SavedFiltersFile);
            }
            catch (Exception ex)
            {
                Log.Error("Load settings", ex);
                Settings = new FilterSettings();
            }
        }

        // ---------- Saved Filters.xml sync ----------

        // Content hash of Saved Filters.xml as last loaded or written by the add-in; any other content means
        // the file was edited outside (e.g. in VS Code) and must be reloaded before use.
        private string _knownFileHash;
        private string _reportedBadHash;
        private FileSystemWatcher _watcher;
        private volatile bool _fileTouched;
        private Timer _fileTimer;
        private Func<Outlook.Explorer> _activeExplorer = () => null;

        /// <summary>Reloads Saved Filters.xml automatically whenever it is saved outside the add-in.</summary>
        public void WatchSettingsFile(Func<Outlook.Explorer> activeExplorer)
        {
            _activeExplorer = activeExplorer;
            _watcher = new FileSystemWatcher(SettingsPaths.SettingsFolder, Path.GetFileName(SettingsPaths.SavedFiltersFile))
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
                    ReloadIfChanged(_activeExplorer());
                }
                catch (Exception ex)
                {
                    Log.Error("Auto reload", ex);
                }
            };
            _fileTimer.Start();
        }

        /// <summary>
        /// Reloads Saved Filters.xml if it differs from what the add-in last loaded or wrote.
        /// An invalid file is reported once and the current saved filters are kept.
        /// </summary>
        private bool ReloadIfChanged(Outlook.Explorer explorer)
        {
            var path = SettingsPaths.SavedFiltersFile;
            var hash = FileHash(path);
            if (hash == null)
            {
                _fileTouched = true; // locked while being written: try again on the next tick
                return false;
            }
            if (hash == _knownFileHash)
                return false;

            FilterSettings loaded;
            try
            {
                loaded = FilterSettings.Load(path);
            }
            catch (Exception ex)
            {
                if (hash != _reportedBadHash)
                {
                    _reportedBadHash = hash;
                    Log.Error("Saved Filters.xml", ex);
                    MessageBox.Show(explorer == null ? null : WindowOwner.From(explorer),
                        "Saved Filters.xml could not be loaded; the previous saved filters are kept until it is fixed.\n\n" + ex.Message,
                        ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                return false;
            }
            _knownFileHash = hash;
            UseSettings(explorer, loaded);
            return true;
        }

        /// <summary>Switches to newly loaded settings, re-applying the shown saved filter if its SQL changed.</summary>
        private void UseSettings(Outlook.Explorer explorer, FilterSettings loaded)
        {
            var active = explorer == null ? null : ActiveSavedFilter(explorer);
            Settings = loaded;
            if (Fields != null)
                Fields.AutoFill = Settings.AutoFillFields;
            if (explorer != null)
            {
                if (active != null)
                {
                    var now = Settings.Find(active.Name);
                    if (now == null)
                        Views.Clear(explorer);
                    else if (!Dasl.SameFilter(now.Sql, active.Sql))
                        Views.Apply(explorer, Dasl.StripSqlPrefix(now.Sql), SavedSourcePrefix + now.Name);
                }
                if (Settings.AutoApplyFormats)
                    SyncFormats(explorer, false);
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

        public IList<SavedFilter> VisibleFilters
        {
            get { return Settings.Filters.Take(MaxSavedFilters).ToList(); }
        }

        public SavedFilter FilterAt(int index)
        {
            return index >= 0 && index < Settings.Filters.Count && index < MaxSavedFilters ? Settings.Filters[index] : null;
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

        /// <summary>
        /// F/N/S/D button: filters by the input box text if there is any, otherwise by the value of the first
        /// selected mail.
        /// </summary>
        public void QuickButton(Outlook.Explorer explorer, QuickKind kind)
        {
            var typed = _inputText.Trim();
            if (typed.Length > 0)
            {
                ApplyQuick(explorer, kind, typed);
                return;
            }
            var info = MailInfo.FromFirstSelected(explorer,
                Fields == null ? null : (Func<object, string>)Fields.DomainRelatedOf,
                Fields == null ? null : (Func<object, string>)Fields.NameRelatedOf);
            if (info == null)
                throw new UserMessageException("Type a value in the box, or select a mail first.");
            var value = info.ValueFor(kind);
            if (string.IsNullOrWhiteSpace(value))
                throw new UserMessageException("The selected mail has no " + kind + " value.");
            ApplyQuick(explorer, kind, value);
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

        public void Clear(Outlook.Explorer explorer)
        {
            if (!Views.Clear(explorer))
                throw new UserMessageException("The current folder is not shown in a table (list) view.");
            _inputText = "";
            Invalidate();
        }

        // ---------- Saved filters ----------

        /// <summary>Saved filter that "Add to Delete" extends with subjects (created on first use).</summary>
        public const string DeleteFilterName = "Delete";

        /// <summary>Saved filter that "Add to Issue" extends with domainRelated values (created after Delete).</summary>
        public const string IssueFilterName = "Issue";

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

        /// <summary>Add to Delete: the selected mails' subjects become exact-subject conditions of "Delete".</summary>
        public void AddSelectionToDelete(Outlook.Explorer explorer)
        {
            AddSelectionTo(explorer, DeleteFilterName, "subject", SelectedValues(explorer, SubjectOf),
                v => Dasl.PropertyEquals(SubjectProperty, v), null);
        }

        /// <summary>Add to Issue: the selected mails' domainRelated values become conditions of "Issue".</summary>
        public void AddSelectionToIssue(Outlook.Explorer explorer)
        {
            AddSelectionTo(explorer, IssueFilterName, "domainRelated", SelectedValues(explorer, DomainRelatedOf),
                v => Dasl.PropertyEquals(DomainRelatedProperty, v), DeleteFilterName);
        }

        /// <summary>
        /// Appends one OR-ed condition per value to the saved filter <paramref name="filterName"/> (created if missing,
        /// right after <paramref name="insertAfter"/> when given). Values the filter already covers are skipped.
        /// </summary>
        private void AddSelectionTo(Outlook.Explorer explorer, string filterName, string what, List<string> values,
            Func<string, string> clauseFor, string insertAfter)
        {
            ReloadIfChanged(explorer); // never overwrite edits made outside the add-in
            if (values.Count == 0)
                throw new UserMessageException("Select the mails whose " + what + " should be added to \"" + filterName + "\".");

            var f = Settings.Find(filterName);
            if (f == null)
            {
                if (Settings.Filters.Count >= MaxSavedFilters)
                    throw new UserMessageException("The ribbon shows at most " + MaxSavedFilters + " saved filters. Remove one in Saved Filters.xml first.");
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

        public void SyncFormats(Outlook.Explorer explorer, bool force)
        {
            if (_syncing)
                return;
            _syncing = true;
            try
            {
                if (!ConditionalFormatService.Sync(explorer, FormattedFilters, force, Views.OwnFilter(explorer)) && force)
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

        /// <summary>
        /// View Columns button: replaces the current table view's columns with View Columns.txt (created with defaults
        /// on first use). Ctrl+click opens the file for editing instead.
        /// </summary>
        public void ViewColumnsButton(Outlook.Explorer explorer)
        {
            var path = SettingsPaths.ViewColumnsFile;
            if (!File.Exists(path))
                ViewColumns.CreateDefault(path);
            if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
            {
                OpenInEditor(path);
                return;
            }

            List<ViewColumn> columns;
            try
            {
                columns = ViewColumns.Load(path);
            }
            catch (FormatException ex)
            {
                throw new UserMessageException("View Columns.txt: " + ex.Message + "\n\nCtrl+click View Columns to edit the file.");
            }
            if (columns.Count == 0)
                throw new UserMessageException("View Columns.txt has no columns (every line is empty or a # comment).\n\nCtrl+click View Columns to edit the file.");

            List<string> problems = null;
            KeepSelection(explorer, () => problems = ViewColumnsService.Apply(explorer, columns, Views.OwnFilter(explorer)));
            if (problems != null && problems.Count > 0)
                MessageBox.Show(WindowOwner.From(explorer),
                    "The other columns were applied, but:\n\n" + string.Join("\n", problems) + "\n\nCtrl+click View Columns to edit " + path + ".",
                    ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
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

        /// <summary>Folder or view switched: keep formats in sync (quietly) and refresh pressed states.</summary>
        public void OnViewChanged(Outlook.Explorer explorer)
        {
            try
            {
                if (Settings.AutoApplyFormats)
                {
                    // A rule's Filter set through the object model does not survive an Outlook restart, so each
                    // view's rules are rewritten on its first visit per session; later visits only fix differences.
                    var view = ViewFilterService.GetTableView(explorer);
                    var first = view != null && _syncedViews.Add(ViewFilterService.ViewKey(explorer, view));
                    SyncFormats(explorer, first);
                }
            }
            catch (UserMessageException)
            {
                // not a table view: nothing to format
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

        /// <summary>Fills domainRelated/nameRelated/me/tos/ccs for the selected items.</summary>
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

        /// <summary>Fills the items of the current folder that have no domainRelated yet.</summary>
        public void FillMissingFields(Outlook.Explorer explorer)
        {
            var folder = explorer.CurrentFolder;
            if (folder == null)
                return;
            var missing = folder.Items.Restrict("@SQL=\"" + CustomFieldNames.Dasl(CustomFieldNames.DomainRelated) + "\" IS NULL");
            int count = missing.Count;
            if (count == 0)
                throw new UserMessageException("Every item in \"" + folder.Name + "\" already has its fields.");
            if (!ConfirmMany(explorer, count, "unfilled in \"" + folder.Name + "\""))
                return;
            var items = new List<object>(count);
            foreach (var item in missing)
                items.Add(item);
            Report(explorer, Fields.FillItems(items));
        }

        public bool AutoFillFields
        {
            get { return Settings.AutoFillFields; }
        }

        public void SetAutoFillFields(bool on)
        {
            ReloadIfChanged(_activeExplorer());
            Settings.AutoFillFields = on;
            if (Fields != null)
                Fields.AutoFill = on;
            SaveSettings();
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
            MessageBox.Show(WindowOwner.From(explorer), "Custom fields: " + result, ThisAddIn.Title,
                MessageBoxButtons.OK, result.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
        }

        // ---------- Settings ----------

        public void OpenSettingsFile()
        {
            if (!File.Exists(SettingsPaths.SavedFiltersFile))
                SaveSettings();
            OpenInEditor(SettingsPaths.SavedFiltersFile);
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
            FilterSettings loaded;
            try
            {
                loaded = FilterSettings.Load(SettingsPaths.SavedFiltersFile);
            }
            catch (Exception ex)
            {
                throw new UserMessageException("Saved Filters.xml could not be loaded; the previous saved filters are kept.\n\n" + ex.Message);
            }
            _knownFileHash = FileHash(SettingsPaths.SavedFiltersFile);
            UseSettings(explorer, loaded);
            if (Settings.Filters.Count > MaxSavedFilters)
                MessageBox.Show(WindowOwner.From(explorer),
                    "Only the first " + MaxSavedFilters + " of " + Settings.Filters.Count + " saved filters are shown on the ribbon.",
                    ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
                    throw new UserMessageException("The ribbon shows at most " + MaxSavedFilters + " saved filters. Remove one in Saved Filters.xml first.");
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
            Settings.Save(SettingsPaths.SavedFiltersFile);
            _knownFileHash = FileHash(SettingsPaths.SavedFiltersFile);
        }

        public static string Truncate(string s, int max)
        {
            return s == null || s.Length <= max ? s : s.Substring(0, max - 1) + "…";
        }
    }
}
