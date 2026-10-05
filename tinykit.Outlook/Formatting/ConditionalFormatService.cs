using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.Filtering;
using tinykit.OutlookAddin.Settings;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Formatting
{
    /// <summary>
    /// Mirrors the enabled saved filters into the table view's conditional formatting
    /// (View Settings &gt; Conditional Formatting) as rules named "[TK] &lt;filter name&gt; #&lt;sql hash&gt;".
    /// Rules without that prefix (Outlook's own and the user's) are left alone.
    /// The add-in's rules always use the view's row font name and size, so only style and color come from the filter.
    /// </summary>
    /// <remarks>
    /// AutoFormatRule.Filter is write-only in practice (it reads back empty, even for rules made in the UI),
    /// so the SQL is identified by a short hash in the rule name.
    /// </remarks>
    internal static class ConditionalFormatService
    {
        public const string RulePrefix = "[TK] ";

        /// <summary>Prefix used before the add-in was named tinykit; such rules are still ours and get replaced on sync.</summary>
        private const string LegacyRulePrefix = "[OA] ";

        /// <summary>
        /// Brings the current table view's add-in rules in line with <paramref name="filters"/>.
        /// Returns false when the current view is not a table view; unchanged views are not rewritten unless forced.
        /// <paramref name="ownFilter"/> is the view's own filter to persist (null: the shown one is the view's own).
        /// </summary>
        public static bool Sync(Outlook.Explorer explorer, IEnumerable<SavedFilter> filters, bool force, string ownFilter,
            string reason = "")
        {
            var view = ViewFilterService.GetTableView(explorer) as Outlook.TableView;
            if (view == null)
                return false;

            var wanted = filters
                .Where(f => f.FormatEnabled && f.Format != null && !string.IsNullOrWhiteSpace(f.Sql))
                .ToList();
            var rules = view.AutoFormatRules;

            var rowFont = view.RowFont;
            if (!force && InSync(rules, wanted, rowFont))
                return true;
            if (wanted.Count == 0 && !HasOurs(rules))
                return true; // nothing to write or remove: don't touch the view

            // Logged with what the view had, so a rule that went wrong can be traced to when it was (re)written.
            var where = "view \"" + view.Name + "\" (" + SaveOptionText(view) + ") of " + FolderPathOf(explorer);
            var shownFilter = view.Filter ?? "";
            Log.Info("Formats: " + (force ? "rewriting" : "fixing") + " " + wanted.Count + " rule(s) in " + where + " [" + reason + "]"
                + "; before: " + DescribeOurs(rules) + "; shown filter " + shownFilter.Length + " chars, own filter "
                + (ownFilter == null ? "same" : ownFilter.Length + " chars"));

            var othersBefore = OtherCustomRules(rules);
            RemoveOurs(rules);
            var total = 0;
            foreach (var f in wanted)
            {
                var rule = rules.Add(RuleName(f));
                var sql = Dasl.StripSqlPrefix(f.Sql);
                rule.Filter = sql;
                rule.Enabled = true;
                ApplyFont(rule.Font, f.Format, rowFont);
                total += sql.Length;
            }
            try
            {
                Save(view, rules, othersBefore, ownFilter);
            }
            catch (Exception ex)
            {
                Log.Info("Formats: saving " + where + " failed: " + ex.Message);
                throw;
            }
            Log.Info("Formats: wrote " + wanted.Count + " rule(s), " + total + " chars of conditions, in " + where
                + "; after: " + DescribeOurs(view.AutoFormatRules));
            return true;
        }

        // The add-in's rules of a view: "name (on|off)", or "none".
        private static string DescribeOurs(Outlook.AutoFormatRules rules)
        {
            var names = new List<string>();
            for (int i = 1; i <= rules.Count; i++)
            {
                var r = rules[i];
                if (IsOurs(r))
                    names.Add(r.Name + (r.Enabled ? "" : " (off)"));
            }
            return names.Count == 0 ? "none" : string.Join(", ", names);
        }

        private static string SaveOptionText(Outlook.TableView view)
        {
            try
            {
                // as Outlook names it, e.g. olViewSaveOptionAllFoldersOfType (a view shared by all folders of the type)
                return view.SaveOption.ToString();
            }
            catch (Exception)
            {
                return "?";
            }
        }

        private static string FolderPathOf(Outlook.Explorer explorer)
        {
            try
            {
                return explorer.CurrentFolder.FolderPath;
            }
            catch (Exception)
            {
                return "?";
            }
        }

        /// <summary>The current table view's row font (name, size), or null when not a table view.</summary>
        public static Tuple<string, int> GetViewFont(Outlook.Explorer explorer)
        {
            var view = ViewFilterService.GetTableView(explorer) as Outlook.TableView;
            if (view == null)
                return null;
            var font = view.RowFont;
            return Tuple.Create(font.Name, font.Size);
        }

        /// <summary>
        /// Sets the font name and size of the whole table view: rows, optionally column headers, and optionally every
        /// conditional formatting rule (Outlook's and the user's keep their style and color; only name and size change).
        /// The add-in's own rules always follow. Returns false when the current view is not a table view.
        /// </summary>
        public static bool SetViewFont(Outlook.Explorer explorer, string name, int size, bool headers, bool allRules, string ownFilter)
        {
            var view = ViewFilterService.GetTableView(explorer) as Outlook.TableView;
            if (view == null)
                return false;

            var row = view.RowFont;
            row.Name = name;
            row.Size = size;
            if (headers)
            {
                var column = view.ColumnFont;
                column.Name = name;
                column.Size = size;
            }

            var rules = view.AutoFormatRules;
            var othersBefore = OtherCustomRules(rules);
            for (int i = 1; i <= rules.Count; i++)
            {
                var r = rules[i];
                if (allRules || IsOurs(r))
                {
                    r.Font.Name = name;
                    r.Font.Size = size;
                }
            }
            Save(view, rules, othersBefore, ownFilter);
            return true;
        }

        /// <summary>Removes the add-in's rules from the current table view. Returns false when not a table view.</summary>
        public static bool Remove(Outlook.Explorer explorer, string ownFilter)
        {
            var view = ViewFilterService.GetTableView(explorer) as Outlook.TableView;
            if (view == null)
                return false;
            var rules = view.AutoFormatRules;
            var othersBefore = OtherCustomRules(rules);
            if (RemoveOurs(rules) > 0)
                Save(view, rules, othersBefore, ownFilter);
            return true;
        }

        /// <summary>Saves a table view changed elsewhere (e.g. its columns) with the same filter swap and rule guard.</summary>
        public static void SaveView(Outlook.TableView view, string ownFilter)
        {
            var rules = view.AutoFormatRules;
            Save(view, rules, OtherCustomRules(rules), ownFilter);
        }

        private static void Save(Outlook.TableView view, Outlook.AutoFormatRules rules, IList<string> othersBefore, string ownFilter)
        {
            rules.Save();
            // View.Save persists the filter shown now; while an add-in filter (quick/saved) is applied, persist the
            // view's own filter instead so the temporary one doesn't become permanent, then show it again.
            var shown = view.Filter ?? "";
            var swap = ownFilter != null && ownFilter != shown;
            if (swap)
                view.Filter = ownFilter;
            view.Save();
            if (swap)
                view.Filter = shown;
            view.Apply();

            // Guard: the user's own rules must survive. Report loudly if Outlook dropped any.
            var othersAfter = OtherCustomRules(view.AutoFormatRules);
            var lost = othersBefore.Except(othersAfter).ToList();
            if (lost.Count > 0)
            {
                var msg = "Outlook dropped " + lost.Count + " of your conditional formatting rules in view \""
                    + view.Name + "\" while saving: " + string.Join(", ", lost);
                Log.Info(msg);
                throw new UserMessageException(msg);
            }
        }

        private static string RuleName(SavedFilter f)
        {
            using (var sha = SHA1.Create())
            {
                var hash = sha.ComputeHash(Encoding.UTF8.GetBytes(Dasl.StripSqlPrefix(f.Sql)));
                return RulePrefix + f.Name + " #" + BitConverter.ToString(hash, 0, 3).Replace("-", "").ToLowerInvariant();
            }
        }

        private static bool IsOurs(Outlook.AutoFormatRule r)
        {
            return !r.Standard && (r.Name.StartsWith(RulePrefix, StringComparison.Ordinal)
                                || r.Name.StartsWith(LegacyRulePrefix, StringComparison.Ordinal));
        }

        private static bool HasOurs(Outlook.AutoFormatRules rules)
        {
            for (int i = 1; i <= rules.Count; i++)
            {
                if (IsOurs(rules[i]))
                    return true;
            }
            return false;
        }

        private static List<string> OtherCustomRules(Outlook.AutoFormatRules rules)
        {
            var names = new List<string>();
            for (int i = 1; i <= rules.Count; i++)
            {
                var r = rules[i];
                if (!r.Standard && !IsOurs(r))
                    names.Add(r.Name);
            }
            return names;
        }

        private static int RemoveOurs(Outlook.AutoFormatRules rules)
        {
            int removed = 0;
            for (int i = rules.Count; i >= 1; i--)
            {
                if (IsOurs(rules[i]))
                {
                    rules.Remove(i);
                    removed++;
                }
            }
            return removed;
        }

        private static bool InSync(Outlook.AutoFormatRules rules, IList<SavedFilter> wanted, Outlook.ViewFont rowFont)
        {
            var ours = new List<Outlook.AutoFormatRule>();
            for (int i = 1; i <= rules.Count; i++)
            {
                var r = rules[i];
                if (IsOurs(r))
                    ours.Add(r);
            }
            if (ours.Count != wanted.Count)
                return false;

            for (int i = 0; i < wanted.Count; i++)
            {
                var r = ours[i];
                if (r.Name != RuleName(wanted[i]) || !r.Enabled || !FontMatches(r.Font, wanted[i].Format, rowFont))
                    return false;
            }
            return true;
        }

        private static void ApplyFont(Outlook.ViewFont font, FilterFormat fmt, Outlook.ViewFont rowFont)
        {
            font.Name = rowFont.Name;
            font.Size = rowFont.Size;
            font.Bold = fmt.Bold;
            font.Italic = fmt.Italic;
            font.Strikethrough = fmt.Strikeout;
            font.Underline = fmt.Underline;
            font.Color = fmt.Color;
        }

        private static bool FontMatches(Outlook.ViewFont font, FilterFormat fmt, Outlook.ViewFont rowFont)
        {
            return string.Equals(font.Name, rowFont.Name, StringComparison.OrdinalIgnoreCase)
                && font.Size == rowFont.Size
                && font.Bold == fmt.Bold
                && font.Italic == fmt.Italic
                && font.Strikethrough == fmt.Strikeout
                && font.Underline == fmt.Underline
                && font.Color == fmt.Color;
        }
    }
}
