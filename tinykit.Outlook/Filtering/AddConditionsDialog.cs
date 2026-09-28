using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>What the selected mails are added to a saved filter by; several = all of them together (AND).</summary>
    [Flags]
    internal enum ConditionField
    {
        None = 0,
        DomainRelated = 1,
        From = 2,
        Subject = 4,
    }

    /// <summary>One selected mail: its domainRelated, From address and subject pattern (any may be empty).</summary>
    internal sealed class MailKey
    {
        public string Domain = "";
        public string From = "";
        public string Pattern = "";

        public string Get(ConditionField field)
        {
            return field == ConditionField.DomainRelated ? Domain : field == ConditionField.From ? From : Pattern;
        }

        public void Set(ConditionField field, string value)
        {
            if (field == ConditionField.DomainRelated) Domain = value;
            else if (field == ConditionField.From) From = value;
            else Pattern = value;
        }
    }

    /// <summary>
    /// Asks by which of domainRelated, From address and subject the selected mails go into a saved filter, and shows the
    /// conditions so they can be edited before they are added: one per line (% = any text in a subject); with several
    /// fields checked, the values of one mail on one line separated by tabs, all of which must match.
    /// </summary>
    internal sealed class AddConditionsDialog : Form
    {
        /// <summary>The fields in the order of the check boxes and of the tab-separated values.</summary>
        public static readonly ConditionField[] Fields = { ConditionField.DomainRelated, ConditionField.From, ConditionField.Subject };

        // The choice is remembered per saved filter for the session: the same kind of mail tends to go into the same filter.
        private static readonly Dictionary<string, ConditionField> LastField = new Dictionary<string, ConditionField>(StringComparer.OrdinalIgnoreCase);

        private readonly string _filterName;
        private readonly IList<MailKey> _mails;
        private readonly Dictionary<ConditionField, CheckBox> _boxes = new Dictionary<ConditionField, CheckBox>();
        private readonly Label _hint;
        private readonly TextBox _values;
        private readonly Button _ok;
        private readonly Dictionary<ConditionField, string> _texts = new Dictionary<ConditionField, string>();
        private ConditionField _shown;

        /// <param name="defaultField">The choice the first time for this filter in a session.</param>
        public AddConditionsDialog(string filterName, IList<MailKey> mails, ConditionField defaultField)
        {
            _filterName = filterName;
            _mails = mails;
            Text = "Add to " + filterName;
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(600, 324);

            var intro = new Label
            {
                Text = "Add the selected mails to the saved filter \"" + filterName + "\" by:",
                Location = new Point(12, 12),
                AutoSize = true,
                UseMnemonic = false,
            };
            var captions = new Dictionary<ConditionField, string>
            {
                { ConditionField.DomainRelated, "domainRelated  (mail from these domains)" },
                { ConditionField.From, "From address  (mail from these senders)" },
                { ConditionField.Subject, "Subject  (numbers, dates, month and weekday names become %)" },
            };
            int y = 38;
            foreach (var field in Fields)
            {
                _boxes[field] = new CheckBox
                {
                    Text = captions[field],
                    Location = new Point(24, y),
                    AutoSize = true,
                    Enabled = mails.Any(m => m.Get(field).Length > 0),
                };
                y += 24;
            }
            _hint = new Label { Location = new Point(12, y + 8), AutoSize = true, ForeColor = SystemColors.GrayText, UseMnemonic = false };
            _values = new TextBox
            {
                Location = new Point(12, y + 30),
                Size = new Size(576, 324 - (y + 30) - 48),
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            _ok = new Button { Text = "Add", DialogResult = DialogResult.OK, Location = new Point(432, 288), Size = new Size(75, 26) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(513, 288), Size = new Size(75, 26) };
            AcceptButton = _ok;
            CancelButton = cancel;
            Controls.Add(intro);
            Controls.AddRange(_boxes.Values.Cast<Control>().ToArray());
            Controls.AddRange(new Control[] { _hint, _values, _ok, cancel });

            ConditionField start;
            if (!LastField.TryGetValue(filterName, out start))
                start = defaultField;
            foreach (var field in Fields)
                if (!_boxes[field].Enabled)
                    start &= ~field;
            if (start == ConditionField.None)
                start = Fields.FirstOrDefault(f => _boxes[f].Enabled);
            foreach (var field in Fields)
                _boxes[field].Checked = (start & field) != 0;
            _shown = Field;
            Show(_shown);
            foreach (var box in _boxes.Values)
                box.CheckedChanged += (s, e) => Switch();
        }

        /// <summary>The checked fields.</summary>
        public ConditionField Field
        {
            get { return Fields.Where(f => _boxes[f].Checked).Aggregate(ConditionField.None, (a, f) => a | f); }
        }

        /// <summary>
        /// The conditions, one per non-empty line: the values of the checked fields in check-box order, split at tabs
        /// (lines without a value for every checked field are left out).
        /// </summary>
        public List<MailKey> Values
        {
            get
            {
                var checkedFields = Fields.Where(f => (Field & f) != 0).ToList();
                var result = new List<MailKey>();
                foreach (var line in _values.Lines.Select(l => l.Trim()).Where(l => l.Length > 0).Distinct(StringComparer.Ordinal))
                {
                    var parts = checkedFields.Count == 1 ? new[] { line } : line.Split(new[] { '\t' }, checkedFields.Count);
                    if (parts.Length != checkedFields.Count || parts.Any(p => p.Trim().Length == 0))
                        continue;
                    var key = new MailKey();
                    for (int i = 0; i < parts.Length; i++)
                        key.Set(checkedFields[i], parts[i].Trim());
                    result.Add(key);
                }
                return result;
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
                LastField[_filterName] = Field;
            base.OnFormClosed(e);
        }

        /// <summary>Keeps each choice's (possibly edited) text when switching between them.</summary>
        private void Switch()
        {
            var now = Field;
            if (now == _shown)
                return;
            _texts[_shown] = _values.Text;
            _shown = now;
            Show(now);
        }

        private void Show(ConditionField field)
        {
            string text;
            if (!_texts.TryGetValue(field, out text))
                _texts[field] = text = Initial(field);
            _values.Text = text;
            _values.Enabled = field != ConditionField.None;
            _ok.Enabled = field != ConditionField.None;
            var names = Fields.Where(f => (field & f) != 0).Select(f => f == ConditionField.DomainRelated ? "domainRelated" : f == ConditionField.From ? "from" : "subject").ToList();
            _hint.Text = names.Count == 0 ? "Check domainRelated, From address, Subject, or several of them."
                : names.Count == 1 ? "One condition per line; edit as needed.  % = any text (subject)."
                : "One mail per line: " + string.Join(" <Tab> ", names) + "; all must match.  % = any text (subject).";
        }

        /// <summary>The selected mails' values for the checked fields: one line each, distinct.</summary>
        private string Initial(ConditionField field)
        {
            var checkedFields = Fields.Where(f => (field & f) != 0).ToList();
            if (checkedFields.Count == 0)
                return "";
            var lines = _mails.Where(m => checkedFields.All(f => m.Get(f).Length > 0))
                .Select(m => string.Join("\t", checkedFields.Select(m.Get)))
                .Distinct(checkedFields.Count == 1 && checkedFields[0] != ConditionField.Subject
                    ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal);
            return string.Join(Environment.NewLine, lines);
        }
    }
}
