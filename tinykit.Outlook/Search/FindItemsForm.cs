using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Search
{
    /// <summary>
    /// Find Items window (resizable, stays open): the text to find, Message Body, Filter (kinds of items and accounts) and
    /// Search in one row; the results (Type, Account, Folder, Date, Subject, Field 1-3) and, under them, the selected
    /// item's details. Enter searches, Esc stops a search or empties the text (and closes the window when it is empty);
    /// double-click or Enter in the results opens the item. Size, splitter, Message Body, kinds and accounts are kept.
    /// </summary>
    internal sealed class FindItemsForm : Form
    {
        private const int MaxResults = 5000;
        private const string Title = "Find Items";

        private static string _lastText = "";

        private readonly Outlook.Application _app;
        private readonly TextBox _text;
        private readonly CheckBox _body;
        private readonly Button _filter;
        private readonly Button _search;
        private readonly ContextMenuStrip _filterMenu = new ContextMenuStrip();
        private readonly ListView _list;
        private readonly TextBox _details;
        private readonly SplitContainer _split;
        private readonly HashSet<FindKind> _kinds = new HashSet<FindKind>(Enum.GetValues(typeof(FindKind)).Cast<FindKind>());
        private readonly HashSet<string> _excludedAccounts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly List<FindResult> _results = new List<FindResult>();
        private int _sortColumn = 3; // Date
        private bool _sortDescending = true;
        private bool _searching;
        private bool _stop;
        private int _savedSplit = -1;
        private bool _bodyChecked;

        public FindItemsForm(Outlook.Application app)
        {
            _app = app;
            Text = Title;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            ClientSize = new Size(1000, 640);
            MinimumSize = new Size(560, 320);
            LoadSettings();

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 2, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            // Text, Message Body, Filter, Search.
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 4, RowCount = 1, Margin = new Padding(0, 0, 0, 6) };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            for (int i = 0; i < 3; i++)
                top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _text = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 1, 6, 0), Text = _lastText };
            _text.TextChanged += (s, e) => _lastText = _text.Text;
            _text.KeyDown += OnTextKeyDown;
            _body = new CheckBox { Text = "Message Body / Notes", AutoSize = true, Checked = _bodyChecked, Margin = new Padding(0, 3, 6, 0), Anchor = AnchorStyles.Left };
            _filter = new Button { Text = "Filter ▾", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, 0, 6, 0) };
            _filter.Click += (s, e) => ShowFilterMenu();
            // Ticking an item keeps the menu open, so several can be changed at once.
            _filterMenu.Closing += (s, e) => { if (e.CloseReason == ToolStripDropDownCloseReason.ItemClicked) e.Cancel = true; };
            _search = new Button { Text = "Search", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0) };
            _search.Click += (s, e) => { if (_searching) _stop = true; else RunSearch(); };
            top.Controls.Add(_text, 0, 0);
            top.Controls.Add(_body, 1, 0);
            top.Controls.Add(_filter, 2, 0);
            top.Controls.Add(_search, 3, 0);
            layout.Controls.Add(top, 0, 0);

            _list = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                HideSelection = false,
                MultiSelect = false,
                VirtualMode = true,
            };
            foreach (var column in new[] { "Type", "Account", "Folder", "Date", "Subject", "Field 1", "Field 2", "Field 3" })
                _list.Columns.Add(column);
            _list.RetrieveVirtualItem += (s, e) => e.Item = RowOf(_results[e.ItemIndex]);
            _list.SelectedIndexChanged += (s, e) => ShowDetails();
            _list.DoubleClick += (s, e) => OpenSelected();
            _list.KeyDown += (s, e) =>
            {
                if (e.KeyCode == Keys.Enter)
                {
                    OpenSelected();
                    e.Handled = e.SuppressKeyPress = true;
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    EscapePressed();
                    e.Handled = e.SuppressKeyPress = true;
                }
            };
            _list.ColumnClick += (s, e) => SortBy(e.Column);
            _details = new TextBox
            {
                Dock = DockStyle.Fill,
                Multiline = true,
                ReadOnly = true,
                ScrollBars = ScrollBars.Vertical,
                WordWrap = true,
                BackColor = SystemColors.Window,
            };
            _split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, Margin = new Padding(0) };
            _split.Panel1.Controls.Add(_list);
            _split.Panel2.Controls.Add(_details);
            layout.Controls.Add(_split, 0, 1);

            Controls.Add(layout);
            SizeColumns();
        }

        private void OnTextKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                if (!_searching)
                    RunSearch();
                e.Handled = e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                EscapePressed();
                e.Handled = e.SuppressKeyPress = true;
            }
        }

        // Esc stops a search; otherwise it empties the text, and with the text already empty it closes the window.
        private void EscapePressed()
        {
            if (_searching)
                _stop = true;
            else if (_text.Text.Length == 0)
                Close();
            else
            {
                _text.Text = "";
                _text.Focus();
            }
        }

        private void ShowFilterMenu()
        {
            _filterMenu.Items.Clear();
            _filterMenu.Items.Add(new ToolStripLabel("Item Types") { Font = new Font(Font, FontStyle.Bold) });
            foreach (FindKind kind in Enum.GetValues(typeof(FindKind)))
            {
                var k = kind;
                var item = new ToolStripMenuItem(FindItems.LabelOf(k)) { CheckOnClick = true, Checked = _kinds.Contains(k) };
                item.CheckedChanged += (s, e) => { if (item.Checked) _kinds.Add(k); else _kinds.Remove(k); };
                _filterMenu.Items.Add(item);
            }
            _filterMenu.Items.Add(new ToolStripSeparator());
            _filterMenu.Items.Add(new ToolStripLabel("Accounts") { Font = new Font(Font, FontStyle.Bold) });
            foreach (var name in AccountNames())
            {
                var n = name;
                var item = new ToolStripMenuItem(n) { CheckOnClick = true, Checked = !_excludedAccounts.Contains(n) };
                item.CheckedChanged += (s, e) => { if (item.Checked) _excludedAccounts.Remove(n); else _excludedAccounts.Add(n); };
                _filterMenu.Items.Add(item);
            }
            _filterMenu.Show(_filter, new Point(0, _filter.Height));
        }

        private List<string> AccountNames()
        {
            var names = new List<string>();
            foreach (Outlook.Store store in _app.Session.Stores)
            {
                try
                {
                    names.Add(store.DisplayName);
                }
                catch (COMException)
                {
                }
            }
            return names;
        }

        private void RunSearch()
        {
            var text = _text.Text.Trim();
            if (text.Length == 0)
            {
                _text.Focus();
                return;
            }
            if (_kinds.Count == 0)
            {
                MessageBox.Show(this, "Tick at least one item type under Filter.", ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            _searching = true;
            _stop = false;
            _search.Text = "Stop";
            _results.Clear();
            _list.VirtualListSize = 0;
            _details.Text = "";
            var watch = Stopwatch.StartNew();
            int folderCount = 0;
            string status;
            try
            {
                var folders = new List<FindFolder>();
                foreach (Outlook.Store store in _app.Session.Stores)
                {
                    string name;
                    try
                    {
                        name = store.DisplayName;
                    }
                    catch (COMException)
                    {
                        continue;
                    }
                    if (_excludedAccounts.Contains(name))
                        continue;
                    Text = Title + " — listing folders of " + name + "…";
                    Application.DoEvents();
                    if (_stop || IsDisposed)
                        break;
                    folders.AddRange(FindItems.FoldersOf(store, _kinds));
                }
                foreach (var folder in folders)
                {
                    if (_stop || IsDisposed || _results.Count >= MaxResults)
                        break;
                    folderCount++;
                    Text = Title + " — " + folderCount + "/" + folders.Count + " " + folder.Account + " › " + folder.Path + "  (" + _results.Count + " found)";
                    Application.DoEvents();
                    if (_stop || IsDisposed)
                        break;
                    try
                    {
                        _results.AddRange(FindItems.Search(folder, FindItems.FilterFor(folder.Kind, text, _body.Checked),
                            MaxResults - _results.Count, () => { Application.DoEvents(); return !_stop && !IsDisposed; }));
                    }
                    catch (Exception ex) // one folder that cannot be read (e.g. offline) does not stop the search
                    {
                        Log.Info("Find Items: " + folder.Account + " › " + folder.Path + " skipped: " + ex.Message);
                    }
                }
                if (IsDisposed)
                    return;
                status = _results.Count + (_results.Count >= MaxResults ? "+ (the first " + MaxResults + ")" : "") + " found in "
                    + folderCount + " folders, " + (watch.ElapsedMilliseconds / 1000.0).ToString("0.0", CultureInfo.InvariantCulture) + " s"
                    + (_stop && _results.Count < MaxResults ? " (stopped)" : "");
            }
            catch (Exception ex)
            {
                Log.Error("Find Items", ex);
                status = "error: " + ex.Message;
            }
            finally
            {
                _searching = false;
                if (!IsDisposed)
                    _search.Text = "Search";
            }
            if (IsDisposed)
                return;
            Text = Title + " — " + status;
            Sort();
            if (_results.Count > 0)
            {
                _list.SelectedIndices.Clear();
                _list.SelectedIndices.Add(0);
                _list.EnsureVisible(0);
            }
        }

        private static ListViewItem RowOf(FindResult r)
        {
            var date = r.DateHasTime ? FindItems.DateTimeText(r.Date) : FindItems.DateText(r.Date);
            var row = new ListViewItem(FindItems.LabelOf(r.Kind));
            row.SubItems.Add(r.Account ?? "");
            row.SubItems.Add(r.Folder ?? "");
            row.SubItems.Add(date ?? "");
            row.SubItems.Add(r.Subject ?? "");
            row.SubItems.Add(r.Field1 ?? "");
            row.SubItems.Add(r.Field2 ?? "");
            row.SubItems.Add(Clip(r.Field3));
            return row;
        }

        private static string Clip(string s)
        {
            if (string.IsNullOrEmpty(s))
                return "";
            return s.Length > 200 ? s.Substring(0, 200) + "…" : s;
        }

        private FindResult Selected
        {
            get { return _list.SelectedIndices.Count == 0 ? null : _results[_list.SelectedIndices[0]]; }
        }

        private void ShowDetails()
        {
            var r = Selected;
            if (r == null)
            {
                _details.Text = "";
                return;
            }
            try
            {
                _details.Text = FindItems.Details(_app.Session, r);
            }
            catch (Exception ex)
            {
                Log.Error("Find Items details", ex);
                _details.Text = "(could not read the item: " + ex.Message + ")";
            }
        }

        private void OpenSelected()
        {
            var r = Selected;
            if (r == null)
                return;
            try
            {
                FindItems.Open(_app.Session, r);
            }
            catch (Exception ex)
            {
                Log.Error("Find Items open", ex);
                MessageBox.Show(this, ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SortBy(int column)
        {
            if (column == _sortColumn)
                _sortDescending = !_sortDescending;
            else
            {
                _sortColumn = column;
                _sortDescending = column == 3; // dates: newest first
            }
            var selected = Selected;
            Sort();
            if (selected != null)
            {
                int i = _results.IndexOf(selected);
                _list.SelectedIndices.Clear();
                _list.SelectedIndices.Add(i);
                _list.EnsureVisible(i);
            }
        }

        private void Sort()
        {
            Comparison<FindResult> compare;
            if (_sortColumn == 3)
                compare = (a, b) => Nullable.Compare(a.Date, b.Date);
            else
            {
                Func<FindResult, string> key = r => RowOf(r).SubItems[_sortColumn].Text;
                compare = (a, b) => string.Compare(key(a), key(b), StringComparison.CurrentCultureIgnoreCase);
            }
            // Stable: equal items keep their order.
            var ordered = _results.Select((r, i) => new { r, i }).ToList();
            ordered.Sort((x, y) =>
            {
                int c = compare(x.r, y.r);
                if (_sortDescending)
                    c = -c;
                return c != 0 ? c : x.i.CompareTo(y.i);
            });
            _results.Clear();
            _results.AddRange(ordered.Select(o => o.r));
            _list.VirtualListSize = _results.Count;
            _list.Invalidate();
        }

        private void SizeColumns()
        {
            var widths = new[] { 70, 150, 160, 120, 300, 160, 160, 300 };
            for (int i = 0; i < widths.Length; i++)
                _list.Columns[i].Width = LogicalToDeviceUnits(widths[i]);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (_savedSplit > 0 && _savedSplit < _split.Height - _split.Panel2MinSize)
                _split.SplitterDistance = _savedSplit;
            else
                _split.SplitterDistance = _split.Height * 3 / 5;
            _text.Focus();
            _text.SelectAll();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            _stop = true;
            SaveSettings();
            base.OnFormClosing(e);
        }

        // Find Items.txt: size=W H, split=N, body=0|1, kinds=Mail,Contact,..., exclude=<account> (one line each).
        private void LoadSettings()
        {
            try
            {
                var path = SettingsPaths.FindItemsSettingsFile;
                if (!File.Exists(path))
                {
                    // First time: leave public folders out.
                    foreach (Outlook.Store store in _app.Session.Stores)
                    {
                        try
                        {
                            if (store.ExchangeStoreType == Outlook.OlExchangeStoreType.olExchangePublicFolder)
                                _excludedAccounts.Add(store.DisplayName);
                        }
                        catch (COMException)
                        {
                        }
                    }
                    return;
                }
                foreach (var line in File.ReadAllLines(path))
                {
                    int eq = line.IndexOf('=');
                    if (eq <= 0)
                        continue;
                    var key = line.Substring(0, eq).Trim();
                    var value = line.Substring(eq + 1).Trim();
                    switch (key)
                    {
                        case "size":
                            var parts = value.Split(' ');
                            int w, h;
                            if (parts.Length == 2 && int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out w)
                                && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out h) && w > 0 && h > 0)
                            {
                                var screen = Screen.PrimaryScreen.WorkingArea;
                                Size = new Size(Math.Min(w, screen.Width), Math.Min(h, screen.Height));
                            }
                            break;
                        case "split":
                            int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _savedSplit);
                            break;
                        case "body":
                            _bodyChecked = value == "1";
                            break;
                        case "kinds":
                            _kinds.Clear();
                            foreach (var name in value.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                FindKind kind;
                                if (Enum.TryParse(name.Trim(), out kind))
                                    _kinds.Add(kind);
                            }
                            break;
                        case "exclude":
                            _excludedAccounts.Add(value);
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Find Items settings", ex);
            }
        }

        private void SaveSettings()
        {
            try
            {
                var size = WindowState == FormWindowState.Normal ? Size : RestoreBounds.Size;
                var lines = new List<string>
                {
                    "size=" + size.Width.ToString(CultureInfo.InvariantCulture) + " " + size.Height.ToString(CultureInfo.InvariantCulture),
                    "split=" + _split.SplitterDistance.ToString(CultureInfo.InvariantCulture),
                    "body=" + (_body.Checked ? "1" : "0"),
                    "kinds=" + string.Join(",", _kinds.OrderBy(k => k)),
                };
                lines.AddRange(_excludedAccounts.OrderBy(a => a).Select(a => "exclude=" + a));
                Directory.CreateDirectory(SettingsPaths.LocalFolder);
                File.WriteAllLines(SettingsPaths.FindItemsSettingsFile, lines);
            }
            catch (Exception ex)
            {
                Log.Error("Find Items settings", ex);
            }
        }
    }
}
