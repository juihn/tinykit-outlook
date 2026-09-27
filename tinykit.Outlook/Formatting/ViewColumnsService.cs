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
            return problems;
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
