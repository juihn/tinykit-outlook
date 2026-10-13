using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>
    /// A check box per mail account for an option that applies account by account (e.g. Move Junk Mail to Inbox); each
    /// change is handed on at once with the ticked accounts.
    /// </summary>
    internal sealed class AccountsForm : Form
    {
        private readonly ListView _list;
        private readonly Action<List<string>> _apply;
        private bool _shown; // ItemChecked also fires while the list is first drawn: only clicks count

        public AccountsForm(string title, string intro, IEnumerable<KeyValuePair<string, bool>> accounts, Action<List<string>> apply)
        {
            _apply = apply;
            Text = title;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = LogicalToDeviceUnits(new Size(560, 320));
            MinimumSize = LogicalToDeviceUnits(new Size(380, 240));

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                AutoSize = true,
                Text = intro,
                UseMnemonic = false,
                MaximumSize = new Size(LogicalToDeviceUnits(520), 0),
                Margin = new Padding(0, 0, 0, 8),
            }, 0, 0);

            _list = new ListView { Dock = DockStyle.Fill, View = View.Details, CheckBoxes = true, FullRowSelect = true, HeaderStyle = ColumnHeaderStyle.Nonclickable };
            _list.Columns.Add("Account", LogicalToDeviceUnits(500));
            _list.ClientSizeChanged += (sender, e) => _list.Columns[0].Width = Math.Max(LogicalToDeviceUnits(200), _list.ClientSize.Width);
            foreach (var account in accounts)
                _list.Items.Add(new ListViewItem(account.Key) { Checked = account.Value });
            _list.ItemChecked += (sender, e) =>
            {
                if (!_shown)
                    return;
                try
                {
                    _apply(_list.Items.Cast<ListViewItem>().Where(i => i.Checked).Select(i => i.Text).ToList());
                }
                catch (Exception ex)
                {
                    Log.Error(title, ex);
                    MessageBox.Show(this, ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            };
            layout.Controls.Add(_list, 0, 1);

            var close = new Button { Text = "Close", AutoSize = true, Anchor = AnchorStyles.Right, DialogResult = DialogResult.OK, Margin = new Padding(0, 8, 0, 0) };
            layout.Controls.Add(close, 0, 2);
            AcceptButton = close;
            CancelButton = close;
            Controls.Add(layout);
            Shown += (sender, e) => _shown = true;
        }
    }
}
