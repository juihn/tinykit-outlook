using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.CustomFields;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Reports
{
    /// <summary>
    /// Recipients Report window: the mail's recipients grouped by domain and department.
    /// Click a person: open the contact (or search LinkedIn); Ctrl+click: new contact; Shift+click: add the address to
    /// the clipboard.
    /// </summary>
    internal sealed class RecipientsReportForm : Form
    {
        private const string PersonHint = "Click: open the contact (or search LinkedIn)   ·   Ctrl+click: new contact   ·   Shift+click: add the address to the clipboard";

        // Remembered for the session, so the next report opens where the last one was.
        private static Rectangle? _lastBounds;
        private static bool _showName, _showUser;

        private readonly Func<RecipientsReport> _build;
        private readonly Outlook.Application _app;
        private RecipientsReport _report;

        private readonly Label _title;
        private readonly Panel _senderRow;
        private readonly Label _counts;
        private readonly CheckBox _nameBox, _userBox;
        private readonly Panel _body;
        private readonly TextBox _search, _result;
        private readonly Label _status;
        private readonly List<Control> _layoutItems = new List<Control>();

        public RecipientsReportForm(Outlook.Application app, Func<RecipientsReport> build)
        {
            _app = app;
            _build = build;
            Text = "Recipients Report";
            Font = new Font("Segoe UI", 9f);
            KeyPreview = true;
            ShowIcon = false;
            MinimumSize = new Size(420, 300);
            StartPosition = FormStartPosition.Manual;
            var bounds = _lastBounds ?? new Rectangle(Cursor.Position.X - 300, Cursor.Position.Y + 20, 620, 520);
            Bounds = Screen.GetWorkingArea(bounds).IntersectsWith(bounds) ? bounds : new Rectangle(100, 100, 620, 520);

            // Top: subject (click opens the mail) and the sender.
            var top = new Panel { Dock = DockStyle.Top, Padding = new Padding(8, 8, 8, 4), Height = 72 };
            _title = new Label
            {
                UseMnemonic = false,
                Font = new Font(Font.FontFamily, 11f, FontStyle.Bold),
                Cursor = Cursors.Hand,
                AutoEllipsis = true,
                Dock = DockStyle.Top,
                Height = 26,
            };
            _title.Click += (s, e) => Safe(() => ((dynamic)_report.Item).Display());
            _title.MouseEnter += (s, e) => _status.Text = "Click to open the mail";
            _title.MouseLeave += (s, e) => _status.Text = "";
            _senderRow = new Panel { Dock = DockStyle.Top, Height = 30 };
            top.Controls.Add(_senderRow);
            top.Controls.Add(_title);

            // Options: counts and what each person shows.
            var options = new Panel { Dock = DockStyle.Top, Height = 26, Padding = new Padding(8, 2, 8, 2) };
            _counts = new Label { UseMnemonic = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            _nameBox = new CheckBox { Text = "Recipient name", AutoSize = true, Dock = DockStyle.Right, Checked = _showName, ForeColor = Color.DimGray };
            _userBox = new CheckBox { Text = "Username", AutoSize = true, Dock = DockStyle.Right, Checked = _showUser, ForeColor = Color.SlateBlue };
            _nameBox.CheckedChanged += (s, e) => { _showName = _nameBox.Checked; Relayout(); };
            _userBox.CheckedChanged += (s, e) => { _showUser = _userBox.Checked; Relayout(); };
            options.Controls.Add(_counts);
            options.Controls.Add(_nameBox);
            options.Controls.Add(_userBox);

            // Bottom: search, status and buttons.
            // Auto-sized rows and buttons, so the texts fit at any display scaling.
            var bottom = new TableLayoutPanel
            {
                Dock = DockStyle.Bottom,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(6, 4, 6, 6),
                ColumnCount = 3,
                RowCount = 2,
            };
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            bottom.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _search = new TextBox { Width = 160, Anchor = AnchorStyles.Left, Margin = new Padding(2, 4, 4, 4) };
            var searchButton = new Button { Text = "Search", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 0, 8, 0) };
            _result = new TextBox
            {
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                BackColor = SystemColors.Control,
                Anchor = AnchorStyles.Left | AnchorStyles.Right,
                Margin = new Padding(8, 4, 2, 4),
            };
            _status = new Label { UseMnemonic = false, AutoSize = false, ForeColor = Color.DimGray, AutoEllipsis = true,
                Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            var buttons = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                WrapContents = false,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0),
            };
            Func<string, Button> button = text => new Button { Text = text, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(8, 0, 8, 0), Margin = new Padding(6, 3, 0, 3) };
            var copy = button("Copy Contents");
            var refresh = button("Refresh");
            var close = button("Close");
            searchButton.Click += (s, e) => Search();
            _search.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) { Search(); e.SuppressKeyPress = true; } };
            copy.Click += (s, e) => { Clipboard.SetText(ContentsText()); _status.Text = "The report was copied to the clipboard."; };
            refresh.Click += (s, e) => Populate();
            close.Click += (s, e) => Close();
            CancelButton = close;
            buttons.Controls.AddRange(new Control[] { close, refresh, copy }); // right to left
            bottom.Controls.Add(_search, 0, 0);
            bottom.Controls.Add(searchButton, 1, 0);
            bottom.Controls.Add(_result, 2, 0);
            var row2 = new TableLayoutPanel { ColumnCount = 2, RowCount = 1, Dock = DockStyle.Fill, AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0) };
            row2.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            row2.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            row2.Controls.Add(_status, 0, 0);
            row2.Controls.Add(buttons, 1, 0);
            bottom.Controls.Add(row2, 0, 1);
            bottom.SetColumnSpan(row2, 3);

            // Body: domains, departments and people.
            _body = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = SystemColors.Window };
            _body.Resize += (s, e) => Relayout();

            Controls.Add(_body);
            Controls.Add(bottom);
            Controls.Add(options);
            Controls.Add(top);

            Populate();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            _lastBounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;
            base.OnFormClosed(e);
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape && ActiveControl != _search)
                Close();
            base.OnKeyDown(e);
        }

        private void Populate()
        {
            _report = _build();
            if (_report == null)
            {
                Close();
                return;
            }
            _title.Text = _report.Subject.Length > 0 ? _report.Subject : "(no subject)";

            _senderRow.Controls.Clear();
            var senderLabel = new Label { Text = "Sender:", AutoSize = true, Location = new Point(0, 6) };
            var sender = new PersonChip(_report.Sender, true) { Location = new Point(senderLabel.PreferredWidth + 4, 2) };
            WireChip(sender);
            _senderRow.Controls.Add(senderLabel);
            _senderRow.Controls.Add(sender);

            _body.SuspendLayout();
            foreach (var c in _layoutItems)
                c.Dispose();
            _layoutItems.Clear();
            foreach (var domain in _report.Domains)
            {
                _layoutItems.Add(new Label
                {
                    UseMnemonic = false,
                    Text = domain.Name + "  (" + domain.Count + ")",
                    AutoSize = true,
                    Font = new Font(Font.FontFamily, 10.5f, FontStyle.Bold | FontStyle.Underline),
                    ForeColor = Color.SlateBlue,
                    Tag = "domain",
                });
                foreach (var department in domain.Departments)
                {
                    if (department.Name != null)
                        _layoutItems.Add(new Label { UseMnemonic = false, Text = "□ " + department.Name, AutoSize = true, ForeColor = Color.DarkOrange, Tag = "department" });
                    foreach (var person in department.People)
                    {
                        var chip = new PersonChip(person, false) { Tag = department.Name == null ? "person" : "person-in-department" };
                        WireChip(chip);
                        _layoutItems.Add(chip);
                    }
                }
            }
            _body.Controls.AddRange(_layoutItems.ToArray());
            _body.ResumeLayout();

            _counts.Text = _report.Domains.Count + " domains, " + _report.RecipientCount + " recipients"
                + (_report.IncludesSender ? " (other than the sender)" : "");
            Relayout();
        }

        /// <summary>Places the headers and wraps the people across the width.</summary>
        private void Relayout()
        {
            if (_body == null)
                return;
            foreach (var chip in _layoutItems.OfType<PersonChip>())
                chip.SetParts(_showName, _showUser);
            var sender = _senderRow.Controls.OfType<PersonChip>().FirstOrDefault();
            if (sender != null)
                sender.SetParts(true, true);

            _body.SuspendLayout();
            int width = _body.ClientSize.Width - 8;
            int y = 6 + _body.AutoScrollPosition.Y;
            int x = 0, rowHeight = 0, indent = 0;
            foreach (var c in _layoutItems)
            {
                var kind = (string)c.Tag;
                if (kind == "domain" || kind == "department")
                {
                    if (x > 0) { y += rowHeight + 2; x = 0; }
                    if (kind == "domain" && c != _layoutItems[0]) y += 8;
                    c.Location = new Point(kind == "domain" ? 8 : 16, y);
                    y += c.PreferredSize.Height + 2;
                    rowHeight = 0;
                    indent = kind == "domain" ? 14 : 30;
                    continue;
                }
                if (x == 0) x = indent;
                var size = c.Size;
                if (x > indent && x + size.Width > width)
                {
                    y += rowHeight + 2;
                    x = indent;
                    rowHeight = 0;
                }
                c.Location = new Point(x, y);
                x += size.Width + 6;
                rowHeight = Math.Max(rowHeight, size.Height);
            }
            _body.ResumeLayout();
        }

        private void WireChip(PersonChip chip)
        {
            chip.MouseEnter += (s, e) => _status.Text = PersonHint;
            chip.MouseLeave += (s, e) => _status.Text = "";
            chip.MouseClick += (s, e) =>
            {
                if (e.Button == MouseButtons.Left)
                    Safe(() => PersonClicked(chip.Person));
            };
        }

        private void PersonClicked(ReportPerson p)
        {
            if ((ModifierKeys & Keys.Shift) == Keys.Shift)
                AddToClipboard(p.Smtp);
            else if ((ModifierKeys & Keys.Control) == Keys.Control)
                NewContact(p);
            else if (p.Contact != null)
                ((dynamic)_app.Session.GetItemFromID(p.Contact.EntryId, p.Contact.StoreId)).Display();
            else
                Process.Start("https://www.google.com/search?q=" + Uri.EscapeDataString("linkedin " + SearchName(p)));
        }

        /// <summary>Adds the address to the addresses already on the clipboard ("a@x; b@y"), or starts a new list.</summary>
        private void AddToClipboard(string smtp)
        {
            var current = Clipboard.ContainsText() ? Clipboard.GetText().Trim() : "";
            var parts = current.Split(';').Select(t => t.Trim()).Where(t => t.Length > 0).ToList();
            var isList = parts.Count > 0 && parts.All(t => t.IndexOf('@') > 0 && t.IndexOf(' ') < 0);
            var list = isList ? parts : new List<string>();
            if (!list.Contains(smtp, StringComparer.OrdinalIgnoreCase))
                list.Add(smtp);
            Clipboard.SetText(string.Join("; ", list));
            _status.Text = "Clipboard: " + string.Join("; ", list);
        }

        /// <summary>A new contact with the person's name split into first/last name and the address as E-mail.</summary>
        private void NewContact(ReportPerson p)
        {
            var contact = (Outlook.ContactItem)_app.CreateItem(Outlook.OlItemType.olContactItem);
            var name = (p.Name ?? p.Description ?? "").Trim();
            if (name.IndexOf('@') > 0)
                name = name.Substring(0, name.IndexOf('@')).Replace('.', ' ').Replace('_', ' ');
            string first, last;
            SplitName(name, out first, out last);
            contact.FirstName = first;
            contact.LastName = last;
            contact.Email1Address = p.Smtp;
            contact.Display();
        }

        /// <summary>"Last, First" → as written; Korean/CJK names → family name first; others → first name first.</summary>
        internal static void SplitName(string name, out string first, out string last)
        {
            first = last = "";
            name = (name ?? "").Trim();
            if (name.Length == 0)
                return;
            int comma = name.IndexOf(',');
            if (comma > 0)
            {
                last = name.Substring(0, comma).Trim();
                first = name.Substring(comma + 1).Trim();
                return;
            }
            bool asian = name.Any(ch => ch >= 0x1100);
            int space = name.IndexOf(' ');
            if (asian)
            {
                if (space < 0) { last = name.Substring(0, 1); first = name.Substring(1); }
                else { last = name.Substring(0, space); first = name.Substring(space + 1).Trim(); }
            }
            else
            {
                if (space < 0) { first = name; }
                else { first = name.Substring(0, space); last = name.Substring(space + 1).Trim(); }
            }
        }

        private static string SearchName(ReportPerson p)
        {
            var name = p.Name ?? p.Description ?? p.UserName;
            return name.IndexOf('@') > 0 ? p.UserName.Replace('.', ' ').Replace('_', ' ') : name;
        }

        private void Search()
        {
            var text = _search.Text.Trim();
            if (text.Length == 0)
            {
                _result.Text = "";
                return;
            }
            var found = _report.Domains.SelectMany(d => d.Departments).SelectMany(d => d.People)
                .Where(p => Contains(p.UserName, text) || Contains(p.Name, text) || Contains(p.Description, text))
                .Select(p => p.Smtp).ToList();
            _result.Text = found.Count == 0 ? "(not found)" : string.Join("; ", found);
        }

        private static bool Contains(string s, string part)
        {
            return s != null && s.IndexOf(part, StringComparison.CurrentCultureIgnoreCase) >= 0;
        }

        /// <summary>The report as text: subject, sender, then domain / department / person lines.</summary>
        private string ContentsText()
        {
            var sb = new StringBuilder();
            sb.AppendLine(_report.Subject);
            sb.AppendLine("Sender: " + Line(_report.Sender));
            sb.AppendLine(_counts.Text);
            foreach (var domain in _report.Domains)
            {
                sb.AppendLine();
                sb.AppendLine(domain.Name + " (" + domain.Count + ")");
                foreach (var department in domain.Departments)
                {
                    var indent = "  ";
                    if (department.Name != null)
                    {
                        sb.AppendLine("  " + department.Name);
                        indent = "    ";
                    }
                    foreach (var p in department.People)
                        sb.AppendLine(indent + p.Type + "  " + Line(p));
                }
            }
            return sb.ToString();
        }

        private static string Line(ReportPerson p)
        {
            return p.Description + (p.Name != null ? " (" + p.Name + ")" : "")
                + (p.Smtp != null && !string.Equals(p.Smtp, p.Description, StringComparison.OrdinalIgnoreCase) ? " <" + p.Smtp + ">" : "");
        }

        private void Safe(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log.Error("RecipientsReport", ex);
                MessageBox.Show(this, ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    /// <summary>
    /// A person: a bullet (blue for To, gray for Cc/Bcc; none for the sender), the description on green when the
    /// address is in Contacts (light blue otherwise), then optionally the display name and the user name.
    /// </summary>
    internal sealed class PersonChip : Control
    {
        private static readonly Font SmallFont = new Font("Segoe UI", 8f);
        private readonly bool _isSender;
        private bool _showName, _showUser;
        private readonly ToolTip _tip = new ToolTip();

        public ReportPerson Person { get; private set; }

        public PersonChip(ReportPerson person, bool isSender)
        {
            Person = person;
            _isSender = isSender;
            DoubleBuffered = true;
            Cursor = Cursors.Hand;
            Font = isSender ? new Font("Segoe UI", 9.5f, FontStyle.Bold) : new Font("Segoe UI", 9f);
            _tip.SetToolTip(this, person.Smtp + (person.Contact != null ? "  (in Contacts)" : "") + (isSender ? "" : "  · " + person.Type));
            SetParts(false, false);
        }

        public void SetParts(bool showName, bool showUser)
        {
            _showName = showName;
            _showUser = showUser;
            Size = Measure();
            Invalidate();
        }

        private int Bullet { get { return _isSender ? 0 : 12; } }

        private Size Measure()
        {
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
            int w = Bullet + 3 + TextRenderer.MeasureText(Person.Description, Font, Size.Empty, flags).Width + 3;
            int h = TextRenderer.MeasureText("Ag", Font, Size.Empty, flags).Height + 4;
            if (_showName && Person.Name != null)
                w += 4 + TextRenderer.MeasureText(Person.Name, SmallFont, Size.Empty, flags).Width;
            if (_showUser && Person.UserName != null)
                w += 4 + TextRenderer.MeasureText(_isSender ? Person.Smtp : Person.UserName, SmallFont, Size.Empty, flags).Width;
            return new Size(w + 2, h);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            var flags = TextFormatFlags.NoPadding | TextFormatFlags.SingleLine;
            if (!_isSender)
                using (var b = new SolidBrush(Person.IsTo ? Color.DodgerBlue : Color.DarkGray))
                    g.FillEllipse(b, 1, Height / 2 - 4, 8, 8);

            int x = Bullet;
            var descWidth = TextRenderer.MeasureText(Person.Description, Font, Size.Empty, flags).Width + 6;
            using (var bg = new SolidBrush(Person.Contact != null ? Color.GreenYellow : Color.AliceBlue))
                g.FillRectangle(bg, x, 0, descWidth, Height);
            TextRenderer.DrawText(g, Person.Description, Font, new Point(x + 3, 2), ForeColor, flags);
            x += descWidth + 4;

            int smallY = Height - TextRenderer.MeasureText("Ag", SmallFont, Size.Empty, flags).Height - 2;
            if (_showName && Person.Name != null)
            {
                TextRenderer.DrawText(g, Person.Name, SmallFont, new Point(x, smallY), Color.DimGray, flags);
                x += TextRenderer.MeasureText(Person.Name, SmallFont, Size.Empty, flags).Width + 4;
            }
            if (_showUser && Person.UserName != null)
                TextRenderer.DrawText(g, _isSender ? Person.Smtp : Person.UserName, SmallFont, new Point(x, smallY), Color.SlateBlue, flags);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _tip.Dispose();
            base.Dispose(disposing);
        }
    }
}
