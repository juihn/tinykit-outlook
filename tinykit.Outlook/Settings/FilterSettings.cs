using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using tinykit.OutlookAddin.Formatting;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Settings
{
    /// <summary>Font format for a saved filter's conditional-formatting rule.</summary>
    internal sealed class FilterFormat
    {
        public string FontName;      // null = keep the view's font
        public int? Size;            // null = keep the view's size
        public bool Bold;
        public bool Italic;
        public bool Strikeout;
        public bool Underline;
        public Outlook.OlColor Color = Outlook.OlColor.olAutoColor;

        public string Style
        {
            get
            {
                if (Bold && Italic) return "Bold Italic";
                if (Bold) return "Bold";
                if (Italic) return "Italic";
                return "Regular";
            }
        }

        public string Describe()
        {
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(FontName)) parts.Add(FontName);
            if (Size.HasValue) parts.Add(Size.Value + "pt");
            parts.Add(Style);
            if (Strikeout) parts.Add("Strikeout");
            if (Underline) parts.Add("Underline");
            parts.Add(OlColorMap.ToName(Color));
            return string.Join(", ", parts);
        }
    }

    /// <summary>A saved DASL filter, shown as a button in the ribbon's Saved Filters group.</summary>
    internal sealed class SavedFilter
    {
        public string Name;
        public string Sql;

        // Whitespace around the SQL in the file (e.g. a line break before "]]>"), written back unchanged.
        public string SqlLeading = "";
        public string SqlTrailing = "";
        public bool FormatEnabled;
        public FilterFormat Format;
    }

    /// <summary>
    /// Saved Filters - Mail.xml / - Contacts.xml / - Tasks.xml — hand-editable. SQL is Outlook DASL (the "SQL" tab of View Settings &gt; Filter),
    /// with or without the "@SQL=" prefix.
    /// </summary>
    internal sealed class FilterSettings
    {
        /// <summary>The kind of folder these saved filters are for (one file per kind).</summary>
        public ItemKind Kind = ItemKind.Mail;

        public bool AutoApplyFormats = true;

        /// <summary>Mail only: fill the custom mail fields of arriving mail.</summary>
        public bool AutoFillFields = true;

        /// <summary>Global switch for the saved filters' formats: All Formats Off clears it, Refresh Formats sets it.
        /// Each filter's own formatEnabled is kept, so turning formats back on restores the same set.</summary>
        public bool FormatsOn = true;
        public List<SavedFilter> Filters = new List<SavedFilter>();

        private const string Schema = @"
  tinykit Outlook saved filters. One file per kind of folder: Saved Filters - Mail.xml, - Contacts.xml and
  - Tasks.xml; the ribbon shows the ones of the folder you are in. Saved changes are picked up automatically.

  <SavedFilters autoApplyFormats=""true|false""         autoApplyFormats: sync formatting rules into
                                                       each table view when you switch folders/views
                autoFillFields=""true|false""           autoFillFields (Mail only): fill domainRelated/nameRelated/
                                                       me/tos/ccs of mail arriving in Inbox / Sent Items
                formatsOn=""true|false"">               formatsOn: all formats on/off (All Formats Off /
                                                       Refresh Formats); formatEnabled per filter is kept
    <Filter name=""Ribbon button label"" formatEnabled=""true|false"">
      <Sql><![CDATA[ DASL filter, e.g. ""urn:schemas:httpmail:read"" = 0 ]]></Sql>
      <Format font=""Segoe UI"" size=""9"" style=""Regular|Bold|Italic|Bold Italic""
              strikeout=""false"" underline=""false"" color=""Auto|Black|Maroon|Green|Olive|Navy|Purple|
              Teal|Gray|Silver|Red|Lime|Yellow|Blue|Fuchsia|Aqua|White|#RRGGBB"" />
    </Filter>
  </SavedFilters>

  font/size are optional (omit to keep the view's font). #RRGGBB maps to the nearest of the 16 colors
  Outlook conditional formatting supports. Tip: build a filter in View Settings > Filter, then use
  Manage > Save View Filter on the ribbon, or copy the text of its SQL tab here.
  Up to 20 saved filters are shown on the ribbon.

  Location: %OneDriveConsumer%\.config\tinykit\Outlook\ when that folder exists (shared by every PC
  signed in to the same personal OneDrive; create it and restart Outlook to move there), otherwise
  %APPDATA%\tinykit\Outlook\. History.xml and View Columns - *.txt follow the same rule; ViewState.xml and the
  log stay local. Mail: ""Add to Delete"" appends the selected mails' subjects to the saved filter named ""Delete""
  (created if missing), ""Add to Issue"" their domainRelated values to ""Issue"".
";

        public static FilterSettings Load(string path, ItemKind kind)
        {
            if (!File.Exists(path))
            {
                var created = CreateDefault(kind);
                created.Save(path);
                return created;
            }

            var root = XDocument.Load(path).Root;
            // <OutlookFilters> is the pre-"Saved Filters" root name.
            if (root == null || (root.Name.LocalName != "SavedFilters" && root.Name.LocalName != "OutlookFilters"))
                throw new FormatException("Root element must be <SavedFilters>.");

            var settings = new FilterSettings
            {
                Kind = kind,
                AutoApplyFormats = ParseBool(root.Attribute("autoApplyFormats"), true),
                AutoFillFields = ParseBool(root.Attribute("autoFillFields"), true),
                FormatsOn = ParseBool(root.Attribute("formatsOn"), true),
            };
            foreach (var e in root.Elements("Filter"))
            {
                var name = ((string)e.Attribute("name") ?? "").Trim();
                var rawSql = (string)e.Element("Sql") ?? "";
                var sql = rawSql.Trim();
                if (name.Length == 0 || sql.Length == 0)
                    continue;
                var fe = e.Element("Format");
                settings.Filters.Add(new SavedFilter
                {
                    Name = name,
                    Sql = sql,
                    SqlLeading = rawSql.Substring(0, rawSql.Length - rawSql.TrimStart().Length),
                    SqlTrailing = rawSql.Substring(rawSql.TrimEnd().Length),
                    FormatEnabled = ParseBool(e.Attribute("formatEnabled"), fe != null),
                    Format = fe == null ? null : ParseFormat(fe),
                });
            }
            return settings;
        }

        public void Save(string path)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var doc = new XDocument(
                new XDeclaration("1.0", "utf-8", null),
                new XComment(Schema),
                new XElement("SavedFilters",
                    new XAttribute("autoApplyFormats", AutoApplyFormats ? "true" : "false"),
                    Kind == ItemKind.Mail ? new XAttribute("autoFillFields", AutoFillFields ? "true" : "false") : null,
                    new XAttribute("formatsOn", FormatsOn ? "true" : "false"),
                    Filters.Select(ToElement)));
            doc.Save(path);
        }

        public SavedFilter Find(string name)
        {
            return Filters.FirstOrDefault(f => string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase));
        }

        private static XElement ToElement(SavedFilter f)
        {
            var e = new XElement("Filter",
                new XAttribute("name", f.Name),
                new XAttribute("formatEnabled", f.FormatEnabled ? "true" : "false"),
                new XElement("Sql", new XCData(f.SqlLeading + f.Sql + f.SqlTrailing)));
            if (f.Format != null)
            {
                var fmt = f.Format;
                var fe = new XElement("Format");
                if (!string.IsNullOrEmpty(fmt.FontName)) fe.Add(new XAttribute("font", fmt.FontName));
                if (fmt.Size.HasValue) fe.Add(new XAttribute("size", fmt.Size.Value));
                fe.Add(new XAttribute("style", fmt.Style),
                       new XAttribute("strikeout", fmt.Strikeout ? "true" : "false"),
                       new XAttribute("underline", fmt.Underline ? "true" : "false"),
                       new XAttribute("color", OlColorMap.ToName(fmt.Color)));
                e.Add(fe);
            }
            return e;
        }

        private static FilterFormat ParseFormat(XElement fe)
        {
            var style = ((string)fe.Attribute("style") ?? "").ToLowerInvariant();
            int size;
            var sizeText = (string)fe.Attribute("size");
            var font = ((string)fe.Attribute("font") ?? "").Trim();
            return new FilterFormat
            {
                FontName = font.Length == 0 ? null : font,
                Size = int.TryParse(sizeText, NumberStyles.Integer, CultureInfo.InvariantCulture, out size) && size > 0
                    ? size : (int?)null,
                Bold = style.Contains("bold") || ParseBool(fe.Attribute("bold"), false),
                Italic = style.Contains("italic") || ParseBool(fe.Attribute("italic"), false),
                Strikeout = ParseBool(fe.Attribute("strikeout"), false),
                Underline = ParseBool(fe.Attribute("underline"), false),
                Color = OlColorMap.Parse((string)fe.Attribute("color")),
            };
        }

        private static bool ParseBool(XAttribute a, bool fallback)
        {
            if (a == null)
                return fallback;
            var v = a.Value.Trim();
            if (v == "1" || v.Equals("true", StringComparison.OrdinalIgnoreCase) || v.Equals("yes", StringComparison.OrdinalIgnoreCase))
                return true;
            if (v == "0" || v.Equals("false", StringComparison.OrdinalIgnoreCase) || v.Equals("no", StringComparison.OrdinalIgnoreCase))
                return false;
            return fallback;
        }

        private const string FlagStatus = "\"http://schemas.microsoft.com/mapi/proptag/0x10900003\" IS NOT NULL";

        /// <summary>The saved filters of a first install (no saved filters file of that kind yet).</summary>
        private static FilterSettings CreateDefault(ItemKind kind)
        {
            switch (kind)
            {
                case ItemKind.Contact: return CreateContactDefault();
                case ItemKind.Task: return CreateTaskDefault();
                default: return CreateMailDefault();
            }
        }

        private static FilterSettings CreateContactDefault()
        {
            const string email = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/";
            var s = new FilterSettings { Kind = ItemKind.Contact };
            // No e-mail address at all (Email1..3 = PidLidEmail1/2/3EmailAddress).
            s.Filters.Add(new SavedFilter
            {
                Name = "No Email",
                Sql = "\"" + email + "8083001f\" IS NULL AND \"" + email + "8093001f\" IS NULL AND \"" + email + "80a3001f\" IS NULL",
                FormatEnabled = true,
                Format = new FilterFormat { Color = Outlook.OlColor.olColorGray },
            });
            s.Filters.Add(new SavedFilter
            {
                Name = "Flagged",
                Sql = FlagStatus,
                FormatEnabled = false,
                Format = new FilterFormat { Color = Outlook.OlColor.olColorRed },
            });
            return s;
        }

        private static FilterSettings CreateTaskDefault()
        {
            const string complete = "\"http://schemas.microsoft.com/mapi/id/{00062003-0000-0000-C000-000000000046}/811c000b\"";
            var s = new FilterSettings { Kind = ItemKind.Task };
            s.Filters.Add(new SavedFilter
            {
                Name = "Active",
                Sql = complete + " = 0",
                FormatEnabled = false,
                Format = new FilterFormat { Bold = true },
            });
            s.Filters.Add(new SavedFilter
            {
                Name = "Completed",
                Sql = complete + " = 1",
                FormatEnabled = true,
                Format = new FilterFormat { Strikeout = true, Color = Outlook.OlColor.olColorGray },
            });
            s.Filters.Add(new SavedFilter
            {
                Name = "High",
                Sql = "\"urn:schemas:httpmail:importance\" = 2 AND " + complete + " = 0",
                FormatEnabled = true,
                Format = new FilterFormat { Color = Outlook.OlColor.olColorRed },
            });
            return s;
        }

        private static FilterSettings CreateMailDefault()
        {
            const string userProp = "http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/";
            var s = new FilterSettings { Kind = ItemKind.Mail };
            // Flagged or completed (PR_FLAG_STATUS is set).
            s.Filters.Add(new SavedFilter
            {
                Name = "Flagged",
                Sql = FlagStatus,
                FormatEnabled = false,
                Format = new FilterFormat { Color = Outlook.OlColor.olColorRed },
            });
            // Mail I sent (me = ▶, written by Custom Fields).
            s.Filters.Add(new SavedFilter
            {
                Name = "Sent",
                Sql = "\"" + userProp + "me\" = '▶'",
                FormatEnabled = true,
                Format = new FilterFormat { Underline = true, Color = Outlook.OlColor.olColorTeal },
            });
            // Received mail whose sender is not in Contacts (nameRelated is "[contact name]" only for contacts).
            s.Filters.Add(new SavedFilter
            {
                Name = "Unknown",
                Sql = "\"" + userProp + "nameRelated\" IS NOT NULL AND NOT (\"" + userProp + "nameRelated\" LIKE '[%')"
                    + " AND \"" + userProp + "me\" <> '▶'",
                FormatEnabled = true,
                Format = new FilterFormat { Color = Outlook.OlColor.olColorGray },
            });
            return s;
        }
    }
}
