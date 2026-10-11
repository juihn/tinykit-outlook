using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.CustomFields;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Contacts
{
    /// <summary>
    /// Email Address Books window: every Contacts folder (all accounts, with subfolders) with a check box for the folder's
    /// Properties &gt; Outlook Address Book &gt; "Show this folder as an e-mail Address Book" (ShowAsOutlookAB), set at once.
    /// </summary>
    internal sealed class AddressBooksForm : Form
    {
        private readonly ListView _list;
        private bool _loading;
        private bool _shown; // ItemChecked also fires (off, then on) while the list is first drawn: only clicks count

        public AddressBooksForm(Outlook.NameSpace session)
        {
            Text = "Email Address Books";
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = false;
            ClientSize = LogicalToDeviceUnits(new Size(900, 440));
            MinimumSize = LogicalToDeviceUnits(new Size(480, 300));

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label
            {
                AutoSize = true,
                Text = "A ticked Contacts folder shows in the Address Book (and is searched by Check Names) "
                    + "— the folder's Properties > Outlook Address Book > Show this folder as an e-mail Address Book. "
                    + "A change is made at once.",
                MaximumSize = new Size(LogicalToDeviceUnits(720), 0),
                Margin = new Padding(0, 0, 0, 8),
            }, 0, 0);

            _list = new ListView { Dock = DockStyle.Fill, View = View.Details, CheckBoxes = true, FullRowSelect = true, HeaderStyle = ColumnHeaderStyle.Nonclickable };
            _list.Columns.Add("Contacts folder", LogicalToDeviceUnits(600));
            _list.Columns.Add("Address Book name", LogicalToDeviceUnits(200));
            // The folder column takes the rest of the width (the paths are long), so there is no sideways scroll bar.
            _list.ClientSizeChanged += (sender, e) =>
                _list.Columns[0].Width = Math.Max(LogicalToDeviceUnits(200), _list.ClientSize.Width - _list.Columns[1].Width);
            Fill(session);
            _list.ItemChecked += (sender, e) => Apply(e.Item);
            layout.Controls.Add(_list, 0, 1);

            var close = new Button { Text = "Close", AutoSize = true, Anchor = AnchorStyles.Right, DialogResult = DialogResult.OK, Margin = new Padding(0, 8, 0, 0) };
            layout.Controls.Add(close, 0, 2);
            AcceptButton = close;
            CancelButton = close;
            Controls.Add(layout);
            Shown += (sender, e) => _shown = true;
            FormClosed += (sender, e) => Release();
        }

        private void Fill(Outlook.NameSpace session)
        {
            _loading = true;
            try
            {
                foreach (var folder in ContactIndex.ContactFolders(session))
                {
                    // FolderPath writes a "/" in a folder's name as %2F (e.g. the folder "olk/family")
                    var path = Safe(() => folder.FolderPath, folder.Name).Replace("%2F", "/");
                    var item = new ListViewItem(path) { Tag = folder };
                    item.SubItems.Add("");
                    ShowState(item, folder);
                    _list.Items.Add(item);
                }
            }
            finally
            {
                _loading = false;
            }
        }

        // Writes the tick to the folder, then shows what the folder now has (Outlook may refuse a change). Only when it
        // differs: writing the same value still makes Outlook redo the folder's address book.
        private void Apply(ListViewItem item)
        {
            if (_loading || !_shown)
                return;
            var folder = (Outlook.MAPIFolder)item.Tag;
            bool wanted = item.Checked;
            try
            {
                if (folder.ShowAsOutlookAB == wanted)
                    return;
                folder.ShowAsOutlookAB = wanted;
                Log.Info("Address book: " + folder.FolderPath + (wanted ? " shown" : " not shown"));
            }
            catch (COMException ex)
            {
                Log.Error("Address book: " + Safe(() => folder.FolderPath, "?"), ex);
                MessageBox.Show(this, "Outlook did not change \"" + Safe(() => folder.Name, "?") + "\":\n" + ex.Message,
                    ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            _loading = true;
            try
            {
                ShowState(item, folder);
            }
            finally
            {
                _loading = false;
            }
        }

        private static void ShowState(ListViewItem item, Outlook.MAPIFolder folder)
        {
            bool shown = Safe(() => folder.ShowAsOutlookAB, false);
            item.Checked = shown;
            item.SubItems[1].Text = shown ? Safe(() => folder.AddressBookName, "") : "";
        }

        private void Release()
        {
            foreach (ListViewItem item in _list.Items)
            {
                if (item.Tag != null && Marshal.IsComObject(item.Tag))
                    Marshal.ReleaseComObject(item.Tag);
                item.Tag = null;
            }
        }

        private static T Safe<T>(Func<T> get, T fallback)
        {
            try
            {
                return get();
            }
            catch (COMException)
            {
                return fallback;
            }
        }
    }
}
