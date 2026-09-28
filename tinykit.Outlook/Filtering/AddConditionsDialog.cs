using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>What the selected mails are added to a saved filter by; both = the two conditions together (AND).</summary>
    [Flags]
    internal enum ConditionField
    {
        None = 0,
        DomainRelated = 1,
        Subject = 2,
        Both = DomainRelated | Subject,
    }

    /// <summary>One selected mail: its domainRelated and its subject pattern (either may be null).</summary>
    internal sealed class MailKey
    {
        public string Domain;
        public string Pattern;
    }

    /// <summary>
    /// Asks whether the selected mails go into a saved filter by domainRelated, by subject, or by both together, and shows
    /// the conditions (one per line, % = any text; domainRelated and subject separated by a tab when both are checked) so
    /// they can be edited before they are added.
    /// </summary>
    internal sealed class AddConditionsDialog : Form
    {
        // The choice is remembered per saved filter for the session: the same kind of mail tends to go into the same filter.
        private static readonly Dictionary<string, ConditionField> LastField = new Dictionary<string, ConditionField>(StringComparer.OrdinalIgnoreCase);

        private readonly string _filterName;
        private readonly CheckBox _byDomain;
        private readonly CheckBox _bySubject;
        private readonly Label _hint;
        private readonly TextBox _values;
        private readonly Button _ok;
        private readonly Dictionary<ConditionField, string> _texts = new Dictionary<ConditionField, string>();
        private ConditionField _shown;

        /// <param name="defaultField">The choice the first time for this filter in a session.</param>
        public AddConditionsDialog(string filterName, IList<MailKey> mails, ConditionField defaultField)
        {
            _filterName = filterName;
            Text = "Add to " + filterName;
            Font = new Font("Segoe UI", 9f);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(560, 300);

            var domains = mails.Select(m => m.Domain).Where(d => !string.IsNullOrEmpty(d)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            var patterns = mails.Select(m => m.Pattern).Where(p => !string.IsNullOrEmpty(p)).Distinct(StringComparer.Ordinal).ToList();
            var pairs = mails.Where(m => !string.IsNullOrEmpty(m.Domain) && !string.IsNullOrEmpty(m.Pattern))
                .Select(m => m.Domain + "\t" + m.Pattern).Distinct(StringComparer.Ordinal).ToList();
            _texts[ConditionField.DomainRelated] = string.Join(Environment.NewLine, domains);
            _texts[ConditionField.Subject] = string.Join(Environment.NewLine, patterns);
            _texts[ConditionField.Both] = string.Join(Environment.NewLine, pairs);
            _texts[ConditionField.None] = "";

            var intro = new Label
            {
                Text = "Add the selected mails to the saved filter \"" + filterName + "\" by:",
                Location = new Point(12, 12),
                AutoSize = true,
                UseMnemonic = false,
            };
            _byDomain = new CheckBox
            {
                Text = "domainRelated  (mail from these domains)",
                Location = new Point(24, 38),
                AutoSize = true,
                Enabled = domains.Count > 0,
            };
            _bySubject = new CheckBox
            {
                Text = "Subject  (numbers, dates, month and weekday names become %)",
                Location = new Point(24, 62),
                AutoSize = true,
                Enabled = patterns.Count > 0,
            };
            _hint = new Label { Location = new Point(12, 94), AutoSize = true, ForeColor = SystemColors.GrayText };
            _values = new TextBox
            {
                Location = new Point(12, 116),
                Size = new Size(536, 136),
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                AcceptsTab = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            _ok = new Button { Text = "Add", DialogResult = DialogResult.OK, Location = new Point(392, 264), Size = new Size(75, 26) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(473, 264), Size = new Size(75, 26) };
            AcceptButton = _ok;
            CancelButton = cancel;
            Controls.AddRange(new Control[] { intro, _byDomain, _bySubject, _hint, _values, _ok, cancel });

            ConditionField start;
            if (!LastField.TryGetValue(filterName, out start))
                start = defaultField;
            if (domains.Count == 0) start &= ~ConditionField.DomainRelated;
            if (patterns.Count == 0) start &= ~ConditionField.Subject;
            if (start == ConditionField.None)
                start = domains.Count > 0 ? ConditionField.DomainRelated : ConditionField.Subject;
            _byDomain.Checked = (start & ConditionField.DomainRelated) != 0;
            _bySubject.Checked = (start & ConditionField.Subject) != 0;
            _shown = Field;
            Show(_shown);
            _byDomain.CheckedChanged += (s, e) => Switch();
            _bySubject.CheckedChanged += (s, e) => Switch();
        }

        /// <summary>What was chosen: domainRelated, Subject, or Both.</summary>
        public ConditionField Field
        {
            get
            {
                return (_byDomain.Checked ? ConditionField.DomainRelated : ConditionField.None)
                     | (_bySubject.Checked ? ConditionField.Subject : ConditionField.None);
            }
        }

        /// <summary>
        /// The conditions: for domainRelated or Subject, one value per line; for Both, (domainRelated, subject pattern)
        /// from lines split at their first tab (lines without one are left out).
        /// </summary>
        public List<MailKey> Values
        {
            get
            {
                var lines = _values.Lines.Select(l => l.Trim()).Where(l => l.Length > 0).Distinct(StringComparer.Ordinal);
                var field = Field;
                if (field == ConditionField.DomainRelated)
                    return lines.Select(l => new MailKey { Domain = l }).ToList();
                if (field == ConditionField.Subject)
                    return lines.Select(l => new MailKey { Pattern = l }).ToList();
                return lines.Select(l => l.Split(new[] { '\t' }, 2))
                    .Where(p => p.Length == 2 && p[0].Trim().Length > 0 && p[1].Trim().Length > 0)
                    .Select(p => new MailKey { Domain = p[0].Trim(), Pattern = p[1].Trim() }).ToList();
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
            _values.Text = _texts[field];
            _values.Enabled = field != ConditionField.None;
            _ok.Enabled = field != ConditionField.None;
            _hint.Text = field == ConditionField.Both
                ? "One mail per line: domainRelated <Tab> subject; both must match.  % = any text."
                : field == ConditionField.None
                    ? "Check domainRelated, Subject, or both."
                    : "One condition per line; edit as needed.  % = any text.";
        }
    }
}
