using System;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Formatting
{
    /// <summary>Picks the font name and size for a whole table view.</summary>
    internal sealed class ViewFontDialog : Form
    {
        // Table views store whole point sizes only (9.5pt is saved as 9pt), so the choices are whole sizes.
        private static readonly int[] Sizes = { 9, 10, 11, 12 };

        private readonly ComboBox _font = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDown, AutoCompleteMode = AutoCompleteMode.SuggestAppend,
            AutoCompleteSource = AutoCompleteSource.ListItems, Width = 220,
        };
        private readonly RadioButton[] _sizes = Sizes.Select(s => new RadioButton
        {
            Text = s.ToString(CultureInfo.InvariantCulture), Tag = s, AutoSize = true, Margin = new Padding(0, 0, 14, 0),
        }).ToArray();
        private readonly CheckBox _headers = new CheckBox { Text = "Column headers", AutoSize = true, Checked = true };
        private readonly CheckBox _rules = new CheckBox { Text = "Conditional formatting rules (unread, etc.)", AutoSize = true, Checked = true };
        private readonly Label _preview = new Label
        {
            Text = "Sample mail subject  샘플 메일 제목  123", AutoSize = false, Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleLeft, BorderStyle = BorderStyle.FixedSingle,
            BackColor = SystemColors.Window, Padding = new Padding(6, 0, 0, 0),
        };
        private readonly Button _ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(80, 0) };

        public ViewFontDialog(string viewName, string fontName, int fontSize)
        {
            Text = "View Font" + (string.IsNullOrEmpty(viewName) ? "" : " – " + viewName);
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Font;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterParent;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            using (var installed = new InstalledFontCollection())
                _font.Items.AddRange(installed.Families.Select(f => f.Name).Where(n => !n.StartsWith("@")).Cast<object>().ToArray());
            _font.Text = fontName;
            // The view's current size, or the nearest choice when it is not one of them.
            var nearest = Sizes.OrderBy(s => Math.Abs(s - fontSize)).First();
            foreach (var radio in _sizes)
            {
                radio.Checked = (int)radio.Tag == nearest;
                radio.CheckedChanged += (s, e) => UpdatePreview();
            }

            _font.TextChanged += (s, e) => UpdatePreview();

            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(80, 0) };
            AcceptButton = _ok;
            CancelButton = cancel;

            var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            AddRow(grid, "Font", _font);
            var sizeRow = new FlowLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false, Margin = new Padding(0),
            };
            sizeRow.Controls.AddRange(_sizes);
            AddRow(grid, "Size", sizeRow);
            AddRow(grid, "Also apply to", Stack(_headers, _rules));
            AddRow(grid, "Preview", _preview);
            _preview.MinimumSize = new Size(0, Font.Height * 3);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false,
                FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0),
            };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_ok);
            grid.Controls.Add(buttons, 0, grid.RowCount);
            grid.SetColumnSpan(buttons, 2);
            grid.RowCount++;

            Controls.Add(grid);
            UpdatePreview();
        }

        public string FontName
        {
            get { return _font.Text.Trim(); }
        }

        public int FontSize
        {
            get
            {
                var chosen = _sizes.FirstOrDefault(r => r.Checked);
                return chosen == null ? 0 : (int)chosen.Tag;
            }
        }

        public bool ApplyToHeaders
        {
            get { return _headers.Checked; }
        }

        public bool ApplyToRules
        {
            get { return _rules.Checked; }
        }

        private bool IsValid
        {
            get { return FontSize >= 6 && FontSize <= 72 && _font.Items.Cast<string>().Any(n => string.Equals(n, FontName, StringComparison.OrdinalIgnoreCase)); }
        }

        private void UpdatePreview()
        {
            _ok.Enabled = IsValid;
            if (!IsValid)
                return;
            var old = _preview.Font;
            try
            {
                _preview.Font = new Font(FontName, FontSize);
            }
            catch (ArgumentException)
            {
                _ok.Enabled = false;
                return;
            }
            if (old != null && old != Font)
                old.Dispose();
        }

        private static void AddRow(TableLayoutPanel grid, string label, Control control)
        {
            var l = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left | AnchorStyles.Top, Margin = new Padding(0, 6, 12, 6) };
            control.Margin = new Padding(0, 3, 0, 3);
            if (!(control is Label))
                control.Anchor = AnchorStyles.Left;
            grid.Controls.Add(l, 0, grid.RowCount);
            grid.Controls.Add(control, 1, grid.RowCount);
            grid.RowCount++;
        }

        private static FlowLayoutPanel Stack(params Control[] controls)
        {
            var p = new FlowLayoutPanel
            {
                AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, FlowDirection = FlowDirection.TopDown,
                WrapContents = false, Margin = new Padding(0),
            };
            p.Controls.AddRange(controls);
            return p;
        }
    }
}
