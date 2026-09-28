using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using tinykit.OutlookAddin.CustomFields;
using tinykit.OutlookAddin.Filtering;
using tinykit.OutlookAddin.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Formatting
{
    /// <summary>Replaces the columns of the current table view with the ones from View Columns.txt.</summary>
    internal static class ViewColumnsService
    {
        /// <summary>The add-in's own fields; created as text fields in a folder that doesn't have them yet.</summary>
        private static readonly HashSet<string> OwnFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            CustomFieldNames.DomainRelated, CustomFieldNames.NameRelated, CustomFieldNames.Me, CustomFieldNames.Tos, CustomFieldNames.Ccs,
            CustomFieldNames.UnknownDomain,
        };

        /// <summary>
        /// Applies <paramref name="columns"/> to the current table view and saves it. Returns the problems (columns that
        /// could not be added, properties Outlook rejected); the other columns are applied anyway.
        /// </summary>
        public static List<string> Apply(Outlook.Explorer explorer, IList<ViewColumn> columns, string ownFilter)
        {
            var view = (Outlook.TableView)ViewFilterService.RequireTableView(explorer);
            var folder = explorer.CurrentFolder;
            var problems = new List<string>();
            // A view can show user-defined columns the folder does not define (e.g. a new folder whose view was copied
            // from the Inbox); Outlook refuses to remove those ("Field does not exist in this folder"), so define them first.
            DefineViewUserFields(view, folder);
            var fields = view.ViewFields;

            // A table view always keeps at least one column, so keep the first one until another has been added.
            while (fields.Count > 1)
                fields.Remove(2);
            bool keptPresent = true;

            foreach (var c in columns)
            {
                var added = TryAdd(fields, c.FieldName);
                if (added == null && keptPresent && fields.Count > 1)
                {
                    // Add also fails when the field is already there: it may be the kept column. Drop that and retry.
                    fields.Remove(1);
                    keptPresent = false;
                    added = TryAdd(fields, c.FieldName);
                }
                if (added == null && keptPresent && fields.Count == 1
                    && string.Equals(fields[1].ColumnFormat.Label, c.FieldName, StringComparison.OrdinalIgnoreCase))
                {
                    // Nothing added yet and the kept column is this field: keep it in place.
                    keptPresent = false;
                    added = fields[1];
                }
                if (added == null && CreateUserField(folder, c))
                    added = TryAdd(fields, c.FieldName);
                if (added == null)
                {
                    problems.Add("Line " + c.Line + ": \"" + c.FieldName + "\" is not a field of this folder"
                        + (c.Type == null && !OwnFields.Contains(c.FieldName) ? " (give it a Type to create it as a user-defined field)" : "")
                        + ".");
                    continue;
                }
                SetFormat(added.ColumnFormat, c, problems);
            }

            if (keptPresent)
            {
                if (fields.Count <= 1)
                    throw new Common.UserMessageException("None of the columns in View Columns.txt could be added:\n" + string.Join("\n", problems));
                fields.Remove(1);
            }

            ConditionalFormatService.SaveView(view, ownFilter);
            Reselect(explorer, view, ownFilter);
            return problems;
        }

        /// <summary>Replaces the current table view's sort with the given fields, all ascending, and saves the view.</summary>
        public static void SortBy(Outlook.Explorer explorer, string ownFilter, params string[] fieldNames)
        {
            var view = (Outlook.TableView)ViewFilterService.RequireTableView(explorer);
            var sort = view.SortFields;
            while (sort.Count > 0)
                sort.Remove(1);
            foreach (var name in fieldNames)
            {
                try
                {
                    sort.Add(name, false);
                }
                catch (COMException)
                {
                    throw new Common.UserMessageException("This view cannot be sorted by \"" + name + "\".");
                }
            }
            ConditionalFormatService.SaveView(view, ownFilter);
            Reselect(explorer, view, ownFilter);
        }

        /// <summary>
        /// In some folders (seen in contact folders of a cached Exchange mailbox) View.Apply leaves the list showing rows
        /// it did not read: dates "None", other cells empty, mail icons for contacts, until the list is sorted or the
        /// folder is opened again. Selecting the (just saved) view again shows it properly. Skipped while a quick or
        /// saved filter is shown, which re-selecting would drop (the saved view has the view's own filter).
        /// </summary>
        private static void Reselect(Outlook.Explorer explorer, Outlook.TableView view, string ownFilter)
        {
            if (ownFilter != null && ownFilter != (view.Filter ?? ""))
                return;
            try
            {
                explorer.CurrentView = view.Name;
            }
            catch (COMException ex)
            {
                Common.Log.Error("View Columns: reselect " + view.Name, ex);
            }
        }

        private const string UserPropertyPrefix = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";

        /// <summary>
        /// Defines in the folder every user-defined field that a column of the view shows but the folder lacks, with the
        /// column's type (string → text, datetime → date/time, i4 → integer, r8 → number, boolean → yes/no).
        /// </summary>
        private static void DefineViewUserFields(Outlook.TableView view, Outlook.MAPIFolder folder)
        {
            System.Xml.Linq.XDocument xml;
            try
            {
                xml = System.Xml.Linq.XDocument.Parse(view.XML);
            }
            catch (Exception ex) when (ex is COMException || ex is System.Xml.XmlException)
            {
                return;
            }
            foreach (var column in xml.Descendants("column"))
            {
                var prop = (string)column.Element("prop") ?? "";
                if (!prop.StartsWith(UserPropertyPrefix, StringComparison.OrdinalIgnoreCase))
                    continue;
                var name = prop.Substring(UserPropertyPrefix.Length);
                try
                {
                    if (name.Length == 0 || folder.UserDefinedProperties.Find(name) != null)
                        continue;
                    folder.UserDefinedProperties.Add(name, TypeOf((string)column.Element("type")));
                }
                catch (COMException ex)
                {
                    Common.Log.Error("Define user field " + name + " in " + folder.Name, ex);
                }
            }
        }

        private static Outlook.OlUserPropertyType TypeOf(string viewType)
        {
            switch ((viewType ?? "").ToLowerInvariant())
            {
                case "datetime": return Outlook.OlUserPropertyType.olDateTime;
                case "i4": return Outlook.OlUserPropertyType.olInteger;
                case "r8": return Outlook.OlUserPropertyType.olNumber;
                case "boolean": return Outlook.OlUserPropertyType.olYesNo;
                default: return Outlook.OlUserPropertyType.olText;
            }
        }

        private static Outlook.ViewField TryAdd(Outlook.ViewFields fields, string name)
        {
            try
            {
                return fields.Add(name);
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>Creates a missing user-defined field in the folder when its type is known. Returns true if created.</summary>
        private static bool CreateUserField(Outlook.MAPIFolder folder, ViewColumn c)
        {
            Outlook.OlUserPropertyType type;
            if (c.Type.HasValue)
                type = c.Type.Value;
            else if (OwnFields.Contains(c.FieldName))
                type = Outlook.OlUserPropertyType.olText;
            else
                return false;
            try
            {
                if (folder.UserDefinedProperties.Find(c.FieldName) != null)
                    return false;
                folder.UserDefinedProperties.Add(c.FieldName, type);
                return true;
            }
            catch (COMException)
            {
                return false;
            }
        }

        private static void SetFormat(Outlook.ColumnFormat format, ViewColumn c, List<string> problems)
        {
            Set(() => { if (c.Width.HasValue) format.Width = c.Width.Value; }, c, "Width " + c.Width, problems);
            Set(() => { if (c.Format.HasValue) format.FieldFormat = c.Format.Value; }, c,
                "Format " + c.Format + " (not in this field's Format list)", problems);
            Set(() => { if (c.Align.HasValue) format.Align = c.Align.Value; }, c, "Alignment", problems);
            Set(() => { if (!string.IsNullOrEmpty(c.Alias)) format.Label = c.Alias; }, c, "Alias \"" + c.Alias + "\"", problems);
        }

        private static void Set(Action set, ViewColumn c, string what, List<string> problems)
        {
            try
            {
                set();
            }
            catch (Exception ex) when (ex is COMException || ex is ArgumentException)
            {
                problems.Add("Line " + c.Line + " (" + c.FieldName + "): Outlook rejected " + what + ".");
            }
        }
    }
}
