using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using TinyKit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TinyKit.OutlookAddin.Filtering
{
    /// <summary>
    /// Applies DASL filters to the current table view and restores the view's own filter on Clear.
    /// The pre-existing filter of each (folder, view) is persisted so Clear also works after a restart.
    /// </summary>
    internal sealed class ViewFilterService
    {
        private readonly string _stateFile;
        private readonly Dictionary<string, string> _originals;
        private readonly Dictionary<string, Applied> _applied = new Dictionary<string, Applied>();

        private sealed class Applied
        {
            public string Source;
            public string ReadBack;
        }

        public ViewFilterService(string stateFile)
        {
            _stateFile = stateFile;
            _originals = LoadOriginals(stateFile);
        }

        /// <summary>The current view if it is a table view (the only kind with Filter + conditional formatting).</summary>
        public static Outlook.View GetTableView(Outlook.Explorer explorer)
        {
            Outlook.View view;
            try
            {
                view = explorer.CurrentView as Outlook.View;
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                return null;
            }
            return view != null && view.ViewType == Outlook.OlViewType.olTableView ? view : null;
        }

        public static Outlook.View RequireTableView(Outlook.Explorer explorer)
        {
            var view = GetTableView(explorer);
            if (view == null)
                throw new UserMessageException("The current folder is not shown in a table (list) view.\n"
                    + "Switch to a list view (View > Change View > Compact/Single/Preview) and try again.");
            return view;
        }

        /// <summary>Applies <paramref name="dasl"/>; <paramref name="source"/> identifies who applied it (for pressed states).</summary>
        public void Apply(Outlook.Explorer explorer, string dasl, string source)
        {
            var view = RequireTableView(explorer);
            var key = ViewKey(explorer, view);
            if (!_originals.ContainsKey(key))
            {
                var current = view.Filter ?? "";
                _originals[key] = IsAddinFilter(current) ? "" : current;
                SaveOriginals();
            }

            view.Filter = dasl;
            view.Apply();
            // Outlook may normalize the string; remember what it reads back so we can recognize it later.
            _applied[key] = new Applied { Source = source, ReadBack = view.Filter ?? "" };
        }

        /// <summary>Restores the view's own filter (empty if none was recorded). Returns false if not a table view.</summary>
        public bool Clear(Outlook.Explorer explorer)
        {
            var view = GetTableView(explorer);
            if (view == null)
                return false;
            var key = ViewKey(explorer, view);
            string original;
            if (!_originals.TryGetValue(key, out original) || IsAddinFilter(original))
                original = "";

            view.Filter = original;
            view.Apply();
            _applied.Remove(key);
            if (_originals.Remove(key))
                SaveOriginals();
            return true;
        }

        /// <summary>
        /// Recognizes filters that are the add-in's own (e.g. a saved filter left on the view from an earlier
        /// session, or set by hand from its SQL). Those are never recorded or restored as the view's own filter.
        /// </summary>
        public Func<string, bool> IsAddinFilter = f => false;

        /// <summary>
        /// The view's own filter while an add-in filter is shown (to persist instead of it), or null when the shown
        /// filter is the view's own.
        /// </summary>
        public string OwnFilter(Outlook.Explorer explorer)
        {
            var view = GetTableView(explorer);
            if (view == null)
                return null;
            string original;
            if (_originals.TryGetValue(ViewKey(explorer, view), out original))
                return IsAddinFilter(original) ? "" : original;
            return IsAddinFilter(view.Filter ?? "") ? "" : null;
        }

        /// <summary>Source tag of the filter this add-in applied to the current view, if it is still in effect.</summary>
        public string ActiveSource(Outlook.Explorer explorer)
        {
            var view = GetTableView(explorer);
            if (view == null)
                return null;
            Applied applied;
            if (!_applied.TryGetValue(ViewKey(explorer, view), out applied))
                return null;
            return string.Equals(applied.ReadBack, view.Filter ?? "", StringComparison.Ordinal) ? applied.Source : null;
        }

        public static string ViewKey(Outlook.Explorer explorer, Outlook.View view)
        {
            var folder = explorer.CurrentFolder;
            return (folder == null ? "" : folder.EntryID) + "|" + view.Name;
        }

        private static Dictionary<string, string> LoadOriginals(string path)
        {
            var map = new Dictionary<string, string>(StringComparer.Ordinal);
            try
            {
                if (File.Exists(path))
                {
                    foreach (var e in XDocument.Load(path).Root.Elements("View"))
                    {
                        var key = (string)e.Attribute("key");
                        if (!string.IsNullOrEmpty(key))
                            map[key] = (string)e.Attribute("filter") ?? "";
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("LoadOriginals", ex);
            }
            return map;
        }

        private void SaveOriginals()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_stateFile));
                new XDocument(new XElement("ViewState",
                    new XComment(" Filters the views had before an add-in filter was applied; restored by Clear. "),
                    _originals.Select(p => new XElement("View",
                        new XAttribute("key", p.Key), new XAttribute("filter", p.Value)))))
                    .Save(_stateFile);
            }
            catch (Exception ex)
            {
                Log.Error("SaveOriginals", ex);
            }
        }
    }
}
