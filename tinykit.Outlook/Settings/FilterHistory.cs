using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.Filtering;

namespace tinykit.OutlookAddin.Settings
{
    internal sealed class HistoryEntry
    {
        public string Value;
        public DateTime Used;

        /// <summary>"[short date] value" as shown in the combo box.</summary>
        public string Label
        {
            get { return "[" + Used.ToString("d", CultureInfo.CurrentCulture) + "] " + Value; }
        }
    }

    /// <summary>Most-recently-used quick-filter values per kind (History.xml).</summary>
    internal sealed class FilterHistory
    {
        public const int MaxItems = 19;

        private readonly string _path;
        private readonly Dictionary<QuickKind, List<HistoryEntry>> _items = new Dictionary<QuickKind, List<HistoryEntry>>();

        private FilterHistory(string path)
        {
            _path = path;
            foreach (QuickKind k in Enum.GetValues(typeof(QuickKind)))
                _items[k] = new List<HistoryEntry>();
        }

        public IList<HistoryEntry> Get(QuickKind kind)
        {
            return _items[kind];
        }

        /// <summary>Moves (or adds) <paramref name="value"/> to the top with the current time.</summary>
        public void Touch(QuickKind kind, string value)
        {
            var list = _items[kind];
            list.RemoveAll(h => string.Equals(h.Value, value, StringComparison.OrdinalIgnoreCase));
            list.Insert(0, new HistoryEntry { Value = value, Used = DateTime.Now });
            if (list.Count > MaxItems)
                list.RemoveRange(MaxItems, list.Count - MaxItems);
            Save();
        }

        /// <summary>Removes <paramref name="value"/> from the list (Ctrl+click in the combo box).</summary>
        public void Remove(QuickKind kind, string value)
        {
            if (_items[kind].RemoveAll(h => string.Equals(h.Value, value, StringComparison.OrdinalIgnoreCase)) > 0)
                Save();
        }

        public static FilterHistory Load(string path)
        {
            var h = new FilterHistory(path);
            try
            {
                if (!File.Exists(path))
                    return h;
                var root = XDocument.Load(path).Root;
                foreach (QuickKind k in Enum.GetValues(typeof(QuickKind)))
                {
                    // "Sender" is the pre-rename name of From.
                    var ke = root.Element(k.ToString()) ?? (k == QuickKind.From ? root.Element("Sender") : null);
                    if (ke == null)
                        continue;
                    foreach (var ie in ke.Elements("Item").Take(MaxItems))
                    {
                        var value = (string)ie.Attribute("value");
                        DateTime used;
                        if (string.IsNullOrEmpty(value)
                            || !DateTime.TryParse((string)ie.Attribute("used"), CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out used))
                            continue;
                        h._items[k].Add(new HistoryEntry { Value = value, Used = used });
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("FilterHistory.Load", ex);
            }
            return h;
        }

        private void Save()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_path));
                new XDocument(new XElement("History",
                    _items.Select(p => new XElement(p.Key.ToString(),
                        p.Value.Select(e => new XElement("Item",
                            new XAttribute("used", e.Used.ToString("s", CultureInfo.InvariantCulture)),
                            new XAttribute("value", e.Value)))))))
                    .Save(_path);
            }
            catch (Exception ex)
            {
                Log.Error("FilterHistory.Save", ex);
            }
        }
    }
}
