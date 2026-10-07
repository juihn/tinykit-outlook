using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>
    /// Column Symbols window: a drop-down per symbol (Contact: the mark before a name from Contacts in nameRelated; Sent,
    /// To, Cc: the me column's) showing each with its code point and name, and above them a font (the current table
    /// view's at first) in which the drop-downs show the symbols, to see how they will look in the list. The font is only
    /// for that preview.
    /// </summary>
    internal sealed class MeSymbolsForm : Form
    {
        private readonly ComboBox _contact, _sent, _to, _cc, _font;
        private readonly float _previewSize;
        private Font _previewFont;

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
            if (disposing && _previewFont != null)
                _previewFont.Dispose();
        }

        /// <param name="fontName">The current table view's font (null when unknown), shown first.</param>
        /// <param name="fontSize">Its size in points (0 when unknown).</param>
        public MeSymbolsForm(string contact, string sent, string to, string cc, string fontName, float fontSize)
        {
            Text = "Column Symbols";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;

            var grid = new TableLayoutPanel { AutoSize = true, ColumnCount = 2, Padding = new Padding(10), Dock = DockStyle.Fill };
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            grid.Controls.Add(new Label
            {
                AutoSize = true,
                MaximumSize = new Size(LogicalToDeviceUnits(520), 0),
                Text = "Contact: the mark before a name from Contacts in nameRelated. Sent, To, Cc: the me column for mail you "
                    + "sent, mail with you in To, and mail with you only in Cc (Sent also starts a sent mail's nameRelated). "
                    + "Mail that already has an old symbol, and saved filters that use it, are changed too.",
                Margin = new Padding(0, 0, 0, 10),
            }, 0, 0);
            grid.SetColumnSpan(grid.GetControlFromPosition(0, 0), 2);
            _previewSize = fontSize > 0 ? fontSize : Font.SizeInPoints; // the list's own size

            // Font: preview only (the view's font is set with View Font...).
            grid.Controls.Add(new Label { Text = "Font", AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 10, 0) }, 0, 1);
            _font = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = LogicalToDeviceUnits(440),
                MaxDropDownItems = 20,
                Margin = new Padding(0, 2, 0, 6),
            };
            foreach (var family in FontFamily.Families)
                _font.Items.Add(family.Name);
            var start = fontName != null && _font.Items.Contains(fontName) ? fontName : "Segoe UI Symbol";
            if (!_font.Items.Contains(start))
                _font.Items.Insert(0, start);
            grid.Controls.Add(_font, 1, 1);

            _contact = AddRow(grid, 2, "Contact", MeSymbols.ContactChoices, contact);
            _sent = AddRow(grid, 3, "Sent", MeSymbols.SentChoices, sent);
            _to = AddRow(grid, 4, "To", MeSymbols.ToChoices, to);
            _cc = AddRow(grid, 5, "Cc", MeSymbols.CcChoices, cc);
            _font.SelectedIndexChanged += (s, e) => Preview((string)_font.SelectedItem);
            _font.SelectedItem = start; // shows the symbol drop-downs in the view's font

            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            var ok = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            grid.Controls.Add(buttons, 0, 6);
            grid.SetColumnSpan(buttons, 2);
            AcceptButton = ok;
            CancelButton = cancel;
            Controls.Add(grid);
        }

        public string Contact { get { return ((MeSymbol)_contact.SelectedItem).Text; } }
        public string Sent { get { return ((MeSymbol)_sent.SelectedItem).Text; } }
        public string To { get { return ((MeSymbol)_to.SelectedItem).Text; } }
        public string Cc { get { return ((MeSymbol)_cc.SelectedItem).Text; } }

        // The symbol drop-downs in the chosen font, at the table view's size: as the symbols will look in the list.
        private void Preview(string fontName)
        {
            var font = new Font(fontName, _previewSize);
            foreach (var box in new[] { _contact, _sent, _to, _cc })
                box.Font = font;
            // Only the preview's own previous font is let go: the one the boxes had at first belongs to the window.
            if (_previewFont != null)
                _previewFont.Dispose();
            _previewFont = font;
        }

        private ComboBox AddRow(TableLayoutPanel grid, int row, string label, IList<MeSymbol> choices, string current)
        {
            grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 10, 0) }, 0, row);
            var box = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = LogicalToDeviceUnits(440),
                MaxDropDownItems = 12,
                Margin = new Padding(0, 2, 0, 2),
            };
            box.Items.AddRange(choices.Cast<object>().ToArray());
            var selected = choices.FirstOrDefault(c => c.Text == current);
            if (selected == null && !string.IsNullOrEmpty(current))
            {
                selected = new MeSymbol(current, "(from the settings file)");
                box.Items.Insert(0, selected);
            }
            box.SelectedItem = selected ?? choices[0];
            grid.Controls.Add(box, 1, row);
            return box;
        }
    }
}
