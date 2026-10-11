using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.Settings;

namespace tinykit.OutlookAddin.Ribbon
{
    /// <summary>
    /// Labels of the main tinykit tab shown or hidden per group (Shift+click the group's dialog launcher). The ribbon XML
    /// is written as usual and then gone through once (<see cref="Apply"/>): each control with an icon gets getShowLabel,
    /// each check box its label through getLabel (check boxes have no showLabel), and a group without a dialog launcher
    /// gets one. Controls without an icon (the user's own saved filters) and ones already icon-only stay as they are; a
    /// group with a short name shows it (getLabel) while its labels are hidden.
    /// Groups with hidden labels are kept in Ribbon Labels.txt in the settings folder; without that file, the groups in
    /// <see cref="HiddenByDefault"/>.
    /// </summary>
    internal sealed class RibbonLabels
    {
        private static readonly XNamespace Ns = "http://schemas.microsoft.com/office/2009/07/customui";

        /// <summary>Label that hides a check box's text (an empty one is not allowed).</summary>
        private const string Blank = " ";

        private readonly Dictionary<string, string> _groupOf = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _checkBoxLabels = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> _launcherGroup = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string[]> _groupLabels = new Dictionary<string, string[]>(StringComparer.Ordinal); // full, short
        private readonly HashSet<string> _hidden = Load();

        /// <summary>Groups whose labels are hidden until toggled (Built-in: Outlook's commands, icon only).</summary>
        public static readonly string[] HiddenByDefault = { "grpBuiltIn" };

        /// <summary>
        /// Adds the callbacks and launchers to <paramref name="xml"/> (customUI of the explorer). A new launcher calls
        /// <paramref name="onAction"/> with the group's key in its tag; a group in <paramref name="shortNames"/> gets
        /// <paramref name="getGroupLabel"/>. The groups in <paramref name="skip"/> are left as they are (no launcher).
        /// </summary>
        public string Apply(string xml, string onAction, string getShowLabel, string getCheckBoxLabel, string getGroupLabel,
            IDictionary<string, string> shortNames, ICollection<string> skip)
        {
            var doc = XDocument.Parse(xml);
            foreach (var group in doc.Descendants(Ns + "group"))
            {
                var key = KeyOf(group);
                if (key == null || skip.Contains(key))
                    continue;
                string shortName;
                var label = group.Attribute("label");
                if (label != null && shortNames.TryGetValue(key, out shortName))
                {
                    _groupLabels[key] = new[] { label.Value, shortName };
                    label.Remove();
                    group.Add(new XAttribute("getLabel", getGroupLabel));
                }
                Visit(group, key, getShowLabel, getCheckBoxLabel);
                var launcher = group.Element(Ns + "dialogBoxLauncher");
                if (launcher == null)
                {
                    var id = "lbl" + key;
                    group.Add(new XElement(Ns + "dialogBoxLauncher",
                        new XElement(Ns + "button",
                            new XAttribute("id", id), new XAttribute("tag", key), new XAttribute("onAction", onAction),
                            new XAttribute("screentip", "Show/Hide Labels"),
                            new XAttribute("supertip", "Click or Shift+click: show or hide the labels of this group's buttons."))));
                    _launcherGroup[id] = key;
                }
                else
                {
                    var button = launcher.Element(Ns + "button");
                    var id = button == null ? null : (string)button.Attribute("id");
                    if (id != null)
                        _launcherGroup[id] = key;
                }
            }
            return doc.ToString(SaveOptions.DisableFormatting);
        }

        // id, or the local part of idQ (tk:Items -> Items).
        private static string KeyOf(XElement group)
        {
            var id = (string)group.Attribute("id") ?? (string)group.Attribute("idQ");
            if (id == null)
                return null;
            int colon = id.IndexOf(':');
            return colon < 0 ? id : id.Substring(colon + 1);
        }

        private void Visit(XElement parent, string key, string getShowLabel, string getCheckBoxLabel)
        {
            foreach (var e in parent.Elements())
            {
                // Only the add-in's own controls: Outlook ignores label callbacks on its own (idMso), so the Built-in
                // group switches between two copies of them instead (OutlookRibbon.AppendBuiltInColumn).
                var id = (string)e.Attribute("id");
                switch (e.Name.LocalName)
                {
                    case "box":
                    case "buttonGroup":
                        Visit(e, key, getShowLabel, getCheckBoxLabel);
                        break;
                    case "button":
                    case "toggleButton":
                    case "dropDown":
                    case "menu":
                    case "dynamicMenu":
                    case "control":
                        // Not inside a menu (Visit does not go into one): menu items always show their labels.
                        if (id != null && HasImage(e) && e.Attribute("showLabel") == null && e.Attribute("getShowLabel") == null)
                        {
                            e.Add(new XAttribute("getShowLabel", getShowLabel));
                            _groupOf[id] = key;
                        }
                        break;
                    case "checkBox":
                        var label = (string)e.Attribute("label");
                        if (id != null && !string.IsNullOrWhiteSpace(label))
                        {
                            e.Attribute("label").Remove();
                            e.Add(new XAttribute("getLabel", getCheckBoxLabel));
                            _checkBoxLabels[id] = label;
                            _groupOf[id] = key;
                        }
                        break;
                }
            }
        }

        private static bool HasImage(XElement e)
        {
            return e.Attribute("imageMso") != null || e.Attribute("getImage") != null || e.Attribute("image") != null;
        }

        /// <summary>Whether the control's label shows (getShowLabel).</summary>
        public bool Shown(string controlId)
        {
            string key;
            return !(_groupOf.TryGetValue(controlId, out key) && _hidden.Contains(key));
        }

        /// <summary>Whether the group's labels are hidden.</summary>
        public bool Hidden(string groupKey)
        {
            return _hidden.Contains(groupKey);
        }

        /// <summary>The group's name: the short one while its labels are hidden (getLabel of a group).</summary>
        public string GroupLabel(string groupKey)
        {
            string[] names;
            if (!_groupLabels.TryGetValue(groupKey, out names))
                return groupKey;
            return _hidden.Contains(groupKey) ? names[1] : names[0];
        }

        /// <summary>A check box's label, or a blank one while its group's labels are hidden (getLabel).</summary>
        public string CheckBoxLabel(string controlId)
        {
            string label;
            if (!_checkBoxLabels.TryGetValue(controlId, out label))
                return Blank;
            return Shown(controlId) ? label : Blank;
        }

        /// <summary>The group whose dialog launcher this is, or null.</summary>
        public string GroupOfLauncher(string controlId)
        {
            string key;
            return _launcherGroup.TryGetValue(controlId, out key) ? key : null;
        }

        /// <summary>Shows the group's labels if hidden, hides them otherwise, and keeps that in the file.</summary>
        public void Toggle(string key)
        {
            if (!_hidden.Remove(key))
                _hidden.Add(key);
            Log.Info("Ribbon labels of " + key + (_hidden.Contains(key) ? " hidden" : " shown"));
            Save(_hidden);
        }

        private static HashSet<string> Load()
        {
            var hidden = new HashSet<string>(StringComparer.Ordinal);
            try
            {
                var path = SettingsPaths.RibbonLabelsFile;
                if (!File.Exists(path))
                    hidden.UnionWith(HiddenByDefault);
                else
                {
                    foreach (var line in File.ReadAllLines(path).Select(l => l.Trim()))
                    {
                        if (line.Length > 0 && !line.StartsWith("#", StringComparison.Ordinal))
                            hidden.Add(line);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Ribbon labels: load", ex);
            }
            return hidden;
        }

        private static void Save(HashSet<string> hidden)
        {
            try
            {
                Directory.CreateDirectory(SettingsPaths.SettingsFolder);
                var lines = new List<string> { "# tinykit ribbon: groups whose button labels are hidden (Shift+click a group's dialog launcher)" };
                lines.AddRange(hidden.OrderBy(k => k, StringComparer.Ordinal));
                File.WriteAllLines(SettingsPaths.RibbonLabelsFile, lines);
            }
            catch (Exception ex)
            {
                Log.Error("Ribbon labels: save", ex);
            }
        }
    }
}
