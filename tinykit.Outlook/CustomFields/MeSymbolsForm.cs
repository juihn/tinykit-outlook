using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>me Column Symbols window: a drop-down per case (Sent Mail, To, Cc) showing each symbol with its code point and name.</summary>
    internal sealed class MeSymbolsForm : Form
    {
        private readonly ComboBox _sent, _to, _cc;

        public MeSymbolsForm(string sent, string to, string cc)
        {
            Text = "me Column Symbols";
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
                Text = "The me column shows one of these for mail you sent, mail with you in To, and mail with you only in Cc. "
                    + "Mail that already has the old symbol, and saved filters that use it, are changed too.",
                Margin = new Padding(0, 0, 0, 10),
            }, 0, 0);
            grid.SetColumnSpan(grid.GetControlFromPosition(0, 0), 2);
            _sent = AddRow(grid, 1, "Sent Mail", MeSymbols.SentChoices, sent);
            _to = AddRow(grid, 2, "To", MeSymbols.ToChoices, to);
            _cc = AddRow(grid, 3, "Cc", MeSymbols.CcChoices, cc);

            var buttons = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, Margin = new Padding(0, 10, 0, 0) };
            var cancel = new Button { Text = "Cancel", AutoSize = true, DialogResult = DialogResult.Cancel };
            var ok = new Button { Text = "OK", AutoSize = true, DialogResult = DialogResult.OK };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            grid.Controls.Add(buttons, 0, 4);
            grid.SetColumnSpan(buttons, 2);
            AcceptButton = ok;
            CancelButton = cancel;
            Controls.Add(grid);
        }

        public string Sent { get { return ((MeSymbol)_sent.SelectedItem).Text; } }
        public string To { get { return ((MeSymbol)_to.SelectedItem).Text; } }
        public string Cc { get { return ((MeSymbol)_cc.SelectedItem).Text; } }

        private ComboBox AddRow(TableLayoutPanel grid, int row, string label, IList<MeSymbol> choices, string current)
        {
            grid.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(0, 0, 10, 0) }, 0, row);
            var box = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = LogicalToDeviceUnits(440),
                MaxDropDownItems = 12,
                Font = new Font("Segoe UI Symbol", Font.SizeInPoints + 1),
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
