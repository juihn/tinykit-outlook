using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using tinykit.OutlookAddin.Settings;

namespace tinykit.OutlookAddin.Common
{
    /// <summary>One of the add-in's keyboard shortcuts (see <see cref="KeyboardShortcuts"/>), which can be turned off.</summary>
    internal sealed class CustomShortcut
    {
        public string Id;
        public Keys Keys;
        public string Command;
        public string Description;
        public bool Enabled = true;

        /// <summary>"Ctrl+Alt+E".</summary>
        public string KeyText
        {
            get
            {
                var parts = new List<string>();
                if ((Keys & Keys.Control) != 0) parts.Add("Ctrl");
                if ((Keys & Keys.Shift) != 0) parts.Add("Shift");
                if ((Keys & Keys.Alt) != 0) parts.Add("Alt");
                var key = Keys & Keys.KeyCode;
                parts.Add(key >= Keys.D0 && key <= Keys.D9 ? ((char)('0' + (key - Keys.D0))).ToString() : key.ToString());
                return string.Join("+", parts);
            }
        }
    }

    /// <summary>
    /// The add-in's keyboard shortcuts and which are on, kept in Custom Shortcuts.txt (one "id=on|off" per line) in the
    /// settings folder; every shortcut is on unless the file turns it off.
    /// </summary>
    internal sealed class CustomShortcuts
    {
        private readonly List<CustomShortcut> _all = new List<CustomShortcut>();

        public IReadOnlyList<CustomShortcut> All
        {
            get { return _all; }
        }

        public CustomShortcut Add(string id, Keys keys, string command, string description)
        {
            var s = new CustomShortcut { Id = id, Keys = keys, Command = command, Description = description };
            _all.Add(s);
            return s;
        }

        /// <summary>Reads which shortcuts are off.</summary>
        public void Load()
        {
            try
            {
                var path = SettingsPaths.ShortcutsFile;
                if (!File.Exists(path))
                    return;
                foreach (var line in File.ReadAllLines(path))
                {
                    var t = line.Trim();
                    int eq = t.IndexOf('=');
                    if (t.StartsWith("#") || eq <= 0)
                        continue;
                    var s = _all.FirstOrDefault(x => string.Equals(x.Id, t.Substring(0, eq).Trim(), StringComparison.OrdinalIgnoreCase));
                    if (s != null)
                        s.Enabled = !string.Equals(t.Substring(eq + 1).Trim(), "off", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch (Exception ex)
            {
                Log.Error("Custom shortcuts", ex);
            }
        }

        public void Save()
        {
            try
            {
                var lines = new List<string> { "# tinykit keyboard shortcuts: id=on|off (Custom Shortcuts button in the TinyKit tab)" };
                lines.AddRange(_all.Select(s => s.Id + "=" + (s.Enabled ? "on" : "off") + "    # " + s.KeyText + " " + s.Command));
                Directory.CreateDirectory(Path.GetDirectoryName(SettingsPaths.ShortcutsFile));
                File.WriteAllLines(SettingsPaths.ShortcutsFile, lines);
            }
            catch (Exception ex)
            {
                Log.Error("Custom shortcuts", ex);
            }
        }
    }
}
