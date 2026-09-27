using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>What the selected mails are added to a saved filter by.</summary>
    internal enum ConditionField
    {
        DomainRelated,
        Subject,
    }

    /// <summary>
    /// Asks whether the selected mails go into a saved filter by domainRelated or by subject, and shows the conditions
    /// (one per line, % = any text) so they can be edited before they are added.
    /// </summary>
    internal sealed class AddConditionsDialog : Form
    {
        // The choice is remembered per saved filter for the session: the same kind of mail tends to go into the same filter.
        private static readonly Dictionary<string, ConditionField> LastField = new Dictionary<string, ConditionField>(StringComparer.OrdinalIgnoreCase);

        private readonly string _filterName;

        private readonly RadioButton _byDomain;
        private readonly RadioButton _bySubject;
        private readonly TextBox _values;
        private readonly Dictionary<ConditionField, string> _texts = new Dictionary<ConditionField, string>();
        private ConditionField _shown;

        /// <param name="defaultField">The field chosen the first time for this filter in a session.</param>
        public AddConditionsDialog(string filterName, IList<string> domains, IList<string> subjectPatterns, ConditionField defaultField)
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
            ClientSize = new Size(520, 300);

            _texts[ConditionField.DomainRelated] = string.Join(Environment.NewLine, domains);
            _texts[ConditionField.Subject] = string.Join(Environment.NewLine, subjectPatterns);

            var intro = new Label
            {
                Text = "Add the selected mails to the saved filter \"" + filterName + "\" by:",
                Location = new Point(12, 12),
                AutoSize = true,
            };
            _byDomain = new RadioButton
            {
                Text = "domainRelated  (mail from these domains)",
                Location = new Point(24, 38),
                AutoSize = true,
                Enabled = domains.Count > 0,
            };
            _bySubject = new RadioButton
            {
                Text = "Subject  (numbers, dates, month and weekday names become %)",
                Location = new Point(24, 62),
                AutoSize = true,
                Enabled = subjectPatterns.Count > 0,
            };
            var hint = new Label
            {
                Text = "One condition per line; edit as needed.  % = any text.",
                Location = new Point(12, 94),
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
            };
            _values = new TextBox
            {
                Location = new Point(12, 116),
                Size = new Size(496, 136),
                Multiline = true,
                ScrollBars = ScrollBars.Both,
                WordWrap = false,
                AcceptsReturn = true,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
            };
            var ok = new Button { Text = "Add", DialogResult = DialogResult.OK, Location = new Point(352, 264), Size = new Size(75, 26) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, Location = new Point(433, 264), Size = new Size(75, 26) };
            AcceptButton = ok;
            CancelButton = cancel;
            Controls.AddRange(new Control[] { intro, _byDomain, _bySubject, hint, _values, ok, cancel });

            ConditionField start;
            if (!LastField.TryGetValue(filterName, out start))
                start = defaultField;
            if (start == ConditionField.Subject && subjectPatterns.Count == 0) start = ConditionField.DomainRelated;
            if (start == ConditionField.DomainRelated && domains.Count == 0) start = ConditionField.Subject;
            _shown = start;
            _values.Text = _texts[start];
            (start == ConditionField.Subject ? _bySubject : _byDomain).Checked = true;
            _byDomain.CheckedChanged += (s, e) => Switch();
            _bySubject.CheckedChanged += (s, e) => Switch();
        }

        /// <summary>The field chosen.</summary>
        public ConditionField Field
        {
            get { return _bySubject.Checked ? ConditionField.Subject : ConditionField.DomainRelated; }
        }

        /// <summary>The non-empty, distinct lines of the box.</summary>
        public List<string> Values
        {
            get
            {
                return _values.Lines.Select(l => l.Trim()).Where(l => l.Length > 0)
                    .Distinct(StringComparer.Ordinal).ToList();
            }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
                LastField[_filterName] = Field;
            base.OnFormClosed(e);
        }

        /// <summary>Keeps each field's (possibly edited) text when switching between them.</summary>
        private void Switch()
        {
            var now = Field;
            if (now == _shown)
                return;
            _texts[_shown] = _values.Text;
            _shown = now;
            _values.Text = _texts[now];
        }
    }
}
