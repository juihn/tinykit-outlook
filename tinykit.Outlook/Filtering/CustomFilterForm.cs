using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>
    /// Custom Filter window (resizable, stays open): the text to find, a check box per field of this kind of folder,
    /// the filter it makes (updated as you type or tick), then Clear All Conditions and Apply.
    /// </summary>
    internal sealed class CustomFilterForm : Form
    {
        // The last text and ticked fields per kind of folder, for the next time the window opens (this session).
        private static readonly Dictionary<CustomFilterKind, string> LastText = new Dictionary<CustomFilterKind, string>();
        private static readonly Dictionary<CustomFilterKind, HashSet<string>> LastChecked = new Dictionary<CustomFilterKind, HashSet<string>>();

        private readonly CustomFilterKind _kind;
        private readonly Func<string, string> _apply; // filter ("" = none) → status text
        private readonly TextBox _text;
        private readonly List<KeyValuePair<CheckBox, CustomFilterField>> _boxes = new List<KeyValuePair<CheckBox, CustomFilterField>>();
        private readonly RichTextBox _sql;
        private readonly Label _status;

        public CustomFilterForm(CustomFilterKind kind, string folderName, Func<string, string> apply)
        {
            _kind = kind;
            _apply = apply;

            Text = "Custom Filter – " + folderName;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            ClientSize = new Size(720, 460);
            MinimumSize = new Size(420, 300);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // text
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // check boxes
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // SQL
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // status
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));    // buttons

            _text = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 6) };
            string last;
            _text.Text = LastText.TryGetValue(kind, out last) ? last : "";
            _text.TextChanged += (s, e) => Rebuild();
            layout.Controls.Add(_text, 0, 0);

            var checks = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true, Margin = new Padding(0, 0, 0, 6) };
            HashSet<string> ticked;
            LastChecked.TryGetValue(kind, out ticked);
            foreach (var field in CustomFilter.FieldsOf(kind))
            {
                var box = new CheckBox
                {
                    Text = field.Label,
                    AutoSize = true,
                    Checked = ticked == null ? field.CheckedByDefault : ticked.Contains(field.Label),
                    Margin = new Padding(0, 2, 14, 2),
                };
                box.CheckedChanged += (s, e) => Rebuild();
                checks.Controls.Add(box);
                _boxes.Add(new KeyValuePair<CheckBox, CustomFilterField>(box, field));
            }
            layout.Controls.Add(checks, 0, 1);

            _sql = new RichTextBox
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                DetectUrls = false, // the property names are URLs; no links
                Multiline = true,
                WordWrap = true,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                BackColor = SystemColors.Window,
            };
            layout.Controls.Add(_sql, 0, 2);

            _status = new Label { AutoSize = true, Dock = DockStyle.Fill, ForeColor = SystemColors.GrayText, Margin = new Padding(0, 4, 0, 4) };
            layout.Controls.Add(_status, 0, 3);

            var buttons = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3, RowCount = 1, Margin = new Padding(0) };
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            buttons.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            var clear = new Button { Text = "Clear All Conditions", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 2, 8, 2) };
            clear.Click += (s, e) => ClearAll();
            var applyButton = new Button { Text = "Apply", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(16, 2, 16, 2) };
            applyButton.Click += (s, e) => Apply();
            buttons.Controls.Add(clear, 0, 0);
            buttons.Controls.Add(applyButton, 2, 0);
            layout.Controls.Add(buttons, 0, 4);

            Controls.Add(layout);
            AcceptButton = applyButton;
            Rebuild();
            _status.Text = "Type the text to find, tick where to look, then Apply. Apply with no conditions shows all items again.";
        }

        private string Sql
        {
            get { return CustomFilter.Build(_text.Text, _boxes.Where(b => b.Key.Checked).Select(b => b.Value)); }
        }

        private void Rebuild()
        {
            var sql = Sql;
            _sql.Text = sql.Length > 0 ? sql : "(no conditions: all items)";
            _sql.ForeColor = sql.Length > 0 ? SystemColors.WindowText : SystemColors.GrayText;
            LastText[_kind] = _text.Text;
            LastChecked[_kind] = new HashSet<string>(_boxes.Where(b => b.Key.Checked).Select(b => b.Value.Label));
        }

        private void ClearAll()
        {
            _text.Text = "";
            foreach (var b in _boxes)
                b.Key.Checked = false;
            Rebuild();
            _text.Focus();
        }

        private void Apply()
        {
            try
            {
                _status.Text = _apply(Sql.Replace(Environment.NewLine, " "));
            }
            catch (Common.UserMessageException ex)
            {
                _status.Text = ex.Message;
            }
            catch (Exception ex)
            {
                Common.Log.Error("Custom Filter", ex);
                MessageBox.Show(this, ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _text.Focus();
            _text.SelectAll();
        }
    }
}
