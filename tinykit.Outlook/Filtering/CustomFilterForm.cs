using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using tinykit.OutlookAddin.Settings;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>
    /// Custom Filter window (resizable, stays open): the text to find with Clear All Conditions and Apply beside it,
    /// a check box per field of this kind of folder, then the filter it makes (updated as you type or tick), filling
    /// the rest. Enter applies, Esc empties the text; the window opens at its last size.
    /// </summary>
    internal sealed class CustomFilterForm : Form
    {
        // The last text and ticked fields per kind of folder, for the next time the window opens (this session).
        private static readonly Dictionary<CustomFilterKind, string> LastText = new Dictionary<CustomFilterKind, string>();
        private static readonly Dictionary<CustomFilterKind, HashSet<string>> LastChecked = new Dictionary<CustomFilterKind, HashSet<string>>();

        private readonly CustomFilterKind _kind;
        private readonly string _title;
        private readonly Func<string, string> _apply; // filter ("" = none) → status text
        private readonly TextBox _text;
        private readonly List<KeyValuePair<CheckBox, CustomFilterField>> _boxes = new List<KeyValuePair<CheckBox, CustomFilterField>>();
        private readonly RichTextBox _sql;

        public CustomFilterForm(CustomFilterKind kind, string folderName, Func<string, string> apply)
        {
            _kind = kind;
            _apply = apply;
            _title = "Custom Filter – " + folderName;

            Text = _title;
            Font = SystemFonts.MessageBoxFont;
            AutoScaleMode = AutoScaleMode.Dpi;
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            KeyPreview = true; // Enter / Esc wherever the focus is
            ClientSize = new Size(720, 460);
            MinimumSize = new Size(420, 240);
            var saved = LoadSize();
            if (saved.HasValue)
                Size = saved.Value;

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // text + buttons
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // check boxes
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // SQL, to the bottom

            // Text, then the two buttons on its right.
            var top = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 3, RowCount = 1, Margin = new Padding(0, 0, 0, 6) };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _text = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 1, 6, 0) };
            string last;
            _text.Text = LastText.TryGetValue(kind, out last) ? last : "";
            _text.TextChanged += (s, e) => Rebuild();
            var clear = new Button { Text = "Clear All Conditions", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0, 0, 6, 0) };
            clear.Click += (s, e) => ClearAll();
            var applyButton = new Button { Text = "Apply", AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Margin = new Padding(0) };
            applyButton.Click += (s, e) => Apply();
            top.Controls.Add(_text, 0, 0);
            top.Controls.Add(clear, 1, 0);
            top.Controls.Add(applyButton, 2, 0);
            layout.Controls.Add(top, 0, 0);

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
                    Margin = new Padding(0, 0, 7, 0),
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
                Margin = new Padding(0),
            };
            layout.Controls.Add(_sql, 0, 2);

            Controls.Add(layout);
            Rebuild();
        }

        private string Sql
        {
            get { return CustomFilter.Build(_text.Text, _boxes.Where(b => b.Key.Checked).Select(b => b.Value)); }
        }

        private void Rebuild()
        {
            var sql = Sql;
            _sql.Text = sql.Length > 0 ? sql : "(no conditions: Apply shows all items)";
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
                Text = _title + " — " + _apply(Sql.Replace(Environment.NewLine, " "));
            }
            catch (Common.UserMessageException ex)
            {
                MessageBox.Show(this, ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Common.Log.Error("Custom Filter", ex);
                MessageBox.Show(this, ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Apply();
                e.Handled = e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _text.Text = "";
                _text.Focus();
                e.Handled = e.SuppressKeyPress = true;
            }
            base.OnKeyDown(e);
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            _text.Focus();
            _text.SelectAll();
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            SaveSize(WindowState == FormWindowState.Normal ? Size : RestoreBounds.Size);
            base.OnFormClosed(e);
        }

        private static Size? LoadSize()
        {
            try
            {
                var path = SettingsPaths.CustomFilterSizeFile;
                if (!File.Exists(path))
                    return null;
                var parts = File.ReadAllText(path).Split((char[])null, StringSplitOptions.RemoveEmptyEntries);
                int w, h;
                if (parts.Length == 2 && int.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out w)
                    && int.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out h) && w > 0 && h > 0)
                {
                    var screen = Screen.PrimaryScreen.WorkingArea;
                    return new Size(Math.Min(w, screen.Width), Math.Min(h, screen.Height));
                }
            }
            catch (Exception ex)
            {
                Common.Log.Error("Custom Filter size", ex);
            }
            return null;
        }

        private static void SaveSize(Size size)
        {
            try
            {
                Directory.CreateDirectory(SettingsPaths.LocalFolder);
                File.WriteAllText(SettingsPaths.CustomFilterSizeFile,
                    size.Width.ToString(CultureInfo.InvariantCulture) + " " + size.Height.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                Common.Log.Error("Custom Filter size", ex);
            }
        }
    }
}
