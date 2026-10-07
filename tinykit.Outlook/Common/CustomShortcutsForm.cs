using System;
using System.Drawing;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>Custom Shortcuts window: the add-in's keyboard shortcuts with a check box each to turn it on or off (saved at once).</summary>
    internal sealed class CustomShortcutsForm : Form
    {
        public CustomShortcutsForm(CustomShortcuts shortcuts)
        {
            Text = "Custom Shortcuts";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = new Size(760, 260);
            MinimumSize = new Size(480, 200);

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "They work in the Outlook main window, but not while you type in a reply in the reading pane (there "
                    + "Outlook's own keys apply). Untick one to give the key back to Outlook.",
                MaximumSize = new Size(LogicalToDeviceUnits(720), 0),
                Margin = new Padding(0, 0, 0, 8),
            }, 0, 0);

            var list = new ListView { Dock = DockStyle.Fill, View = View.Details, CheckBoxes = true, FullRowSelect = true, HeaderStyle = ColumnHeaderStyle.Nonclickable };
            list.Columns.Add("Shortcut", LogicalToDeviceUnits(110));
            list.Columns.Add("Command", LogicalToDeviceUnits(190));
            list.Columns.Add("What it does", LogicalToDeviceUnits(420));
            foreach (var s in shortcuts.All)
            {
                var item = new ListViewItem(s.KeyText) { Checked = s.Enabled, Tag = s };
                item.SubItems.Add(s.Command);
                item.SubItems.Add(s.Description);
                list.Items.Add(item);
            }
            list.ItemChecked += (sender, e) =>
            {
                ((CustomShortcut)e.Item.Tag).Enabled = e.Item.Checked;
                shortcuts.Save();
            };
            layout.Controls.Add(list, 0, 1);

            var close = new Button { Text = "Close", AutoSize = true, Anchor = AnchorStyles.Right, DialogResult = DialogResult.OK, Margin = new Padding(0, 8, 0, 0) };
            layout.Controls.Add(close, 0, 2);
            AcceptButton = close;
            CancelButton = close;
            Controls.Add(layout);
        }
    }
}
