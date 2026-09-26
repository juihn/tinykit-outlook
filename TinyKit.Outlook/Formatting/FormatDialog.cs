using System;
using System.Drawing;
using System.Windows.Forms;
using TinyKit.OutlookAddin.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TinyKit.OutlookAddin.Formatting
{
    /// <summary>
    /// Edits a saved filter's conditional format: font style, strikeout, underline and color
    /// (the parts Outlook's ViewFont supports; font name and size are left as they are).
    /// </summary>
    internal sealed class FormatDialog : Form
    {
        private static readonly string[] Styles = { "Regular", "Bold", "Italic", "Bold Italic" };

        private readonly FilterFormat _original;
        private readonly ComboBox _style = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 160 };
        private readonly ComboBox _color = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList, DrawMode = DrawMode.OwnerDrawFixed, Width = 160,
        };
        private readonly CheckBox _strikeout = new CheckBox { Text = "Strikeout", AutoSize = true };
        private readonly CheckBox _underline = new CheckBox { Text = "Underline", AutoSize = true };
        private readonly Label _preview = new Label
        {
            Text = "Sample mail subject", AutoSize = false, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft,
            BorderStyle = BorderStyle.FixedSingle, BackColor = SystemColors.Window, Padding = new Padding(6, 0, 0, 0),
        };

        public FormatDialog(string filterName, FilterFormat format)
        {
            _original = format ?? new FilterFormat();

            Text = "Format – " + filterName;
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

            _style.Items.AddRange(Styles);
            _style.SelectedIndex = (_original.Bold ? 1 : 0) + (_original.Italic ? 2 : 0);

            _color.Items.Add(Outlook.OlColor.olAutoColor);
            foreach (var c in OlColorMap.Palette)
                _color.Items.Add(c);
            _color.SelectedItem = _original.Color;
            if (_color.SelectedIndex < 0)
                _color.SelectedIndex = 0;
            _color.ItemHeight = Font.Height + 4;
            _color.DrawItem += DrawColorItem;

            _strikeout.Checked = _original.Strikeout;
            _underline.Checked = _original.Underline;

            _style.SelectedIndexChanged += (s, e) => UpdatePreview();
            _color.SelectedIndexChanged += (s, e) => UpdatePreview();
            _strikeout.CheckedChanged += (s, e) => UpdatePreview();
            _underline.CheckedChanged += (s, e) => UpdatePreview();

            var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(80, 0) };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(80, 0) };
            AcceptButton = ok;
            CancelButton = cancel;

            var grid = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Fill };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            AddRow(grid, "Font style", _style);
            AddRow(grid, "Color", _color);
            AddRow(grid, "Effects", Flow(_strikeout, _underline));
            AddRow(grid, "Preview", _preview);
            _preview.Height = Font.Height * 2;
            _preview.MinimumSize = new Size(0, Font.Height * 2);

            var buttons = Flow(cancel, ok); // right to left: [OK] [Cancel] aligned right
            buttons.FlowDirection = FlowDirection.RightToLeft;
            buttons.Dock = DockStyle.Fill;
            buttons.Margin = new Padding(0, 10, 0, 0);
            grid.Controls.Add(buttons, 0, grid.RowCount);
            grid.SetColumnSpan(buttons, 2);
            grid.RowCount++;

            Controls.Add(grid);
            UpdatePreview();
        }

        /// <summary>The edited format; font name and size are carried over unchanged.</summary>
        public FilterFormat Result
        {
            get
            {
                return new FilterFormat
                {
                    FontName = _original.FontName,
                    Size = _original.Size,
                    Bold = _style.SelectedIndex == 1 || _style.SelectedIndex == 3,
                    Italic = _style.SelectedIndex == 2 || _style.SelectedIndex == 3,
                    Strikeout = _strikeout.Checked,
                    Underline = _underline.Checked,
                    Color = (Outlook.OlColor)_color.SelectedItem,
                };
            }
        }

        private void UpdatePreview()
        {
            var r = Result;
            var style = (r.Bold ? FontStyle.Bold : 0) | (r.Italic ? FontStyle.Italic : 0)
                      | (r.Strikeout ? FontStyle.Strikeout : 0) | (r.Underline ? FontStyle.Underline : 0);
            var old = _preview.Font;
            _preview.Font = new Font(Font, style);
            if (old != null && old != Font)
                old.Dispose();
            _preview.ForeColor = OlColorMap.ToColor(r.Color) ?? SystemColors.WindowText;
        }

        private void DrawColorItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index < 0)
                return;
            var value = (Outlook.OlColor)_color.Items[e.Index];
            var swatch = new Rectangle(e.Bounds.Left + 3, e.Bounds.Top + 3, e.Bounds.Height * 2, e.Bounds.Height - 6);
            var rgb = OlColorMap.ToColor(value);
            if (rgb.HasValue)
            {
                using (var b = new SolidBrush(rgb.Value))
                    e.Graphics.FillRectangle(b, swatch);
            }
            e.Graphics.DrawRectangle(SystemPens.ControlDark, swatch);
            var textLeft = swatch.Right + 6;
            TextRenderer.DrawText(e.Graphics, OlColorMap.ToName(value), e.Font,
                new Rectangle(textLeft, e.Bounds.Top, e.Bounds.Right - textLeft, e.Bounds.Height),
                e.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            e.DrawFocusRectangle();
        }

        private static void AddRow(TableLayoutPanel grid, string label, Control control)
        {
            var l = new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 6, 12, 6) };
            control.Margin = new Padding(0, 3, 0, 3);
            if (!(control is Label))
                control.Anchor = AnchorStyles.Left;
            grid.Controls.Add(l, 0, grid.RowCount);
            grid.Controls.Add(control, 1, grid.RowCount);
            grid.RowCount++;
        }

        private static FlowLayoutPanel Flow(params Control[] controls)
        {
            var p = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, Margin = new Padding(0) };
            p.Controls.AddRange(controls);
            return p;
        }
    }
}
