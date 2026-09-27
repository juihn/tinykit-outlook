using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Settings
{
    /// <summary>One column of a table view, from one line of View Columns.txt.</summary>
    internal sealed class ViewColumn
    {
        public int Line;
        public string FieldName;
        public int? Width;
        /// <summary>Type of a user-defined field, used only when the field has to be created in the folder.</summary>
        public Outlook.OlUserPropertyType? Type;
        /// <summary>Position (1-based) in the column's Format drop-down of View Settings.</summary>
        public int? Format;
        public Outlook.OlAlign? Align;
        public string Alias;
    }

    /// <summary>
    /// View Columns.txt: the columns of a table view, one per line. The fields of a line are aligned with tabs
    /// (tab width 4): FieldName at column 0, Width at 16, Type at 20, Format at 32, Alignment at 36, Alias at 44.
    /// An empty field is just more tabs. Lines starting with # are comments (e.g. a column that is switched off).
    /// </summary>
    internal static class ViewColumns
    {
        private const int TabWidth = 4;
        // Start column of each field: FieldName, Width, Type, Format, Alignment, Alias.
        private static readonly int[] FieldStops = { 0, 16, 20, 32, 36, 44 };

        private const string Header =
@"# tinykit Outlook: the columns of the current table view. One file per kind of folder:
#   View Columns - Mail.txt, View Columns - Contacts.txt, View Columns - Tasks.txt.
#   View Columns (ribbon, View group) replaces the view's columns with the lines below, in this order.
#   Ctrl+click View Columns to edit the file of the folder you are in. Lines starting with # are skipped.
#
# Fields, aligned with tabs (tab width 4); leave a field empty with more tabs:
#   FieldName   column 0    field name as in View Settings > Columns (e.g. Received, Flag Status) or a user-defined field
#   Width       column 16   width in characters
#   Type        column 20   type of a user-defined field, used only to create it in a folder that lacks it:
#                           olText, olDateTime, olInteger, olNumber, olYesNo (OlUserPropertyType)
#                           https://learn.microsoft.com/dotnet/api/microsoft.office.interop.outlook.oluserpropertytype
#   Format      column 32   position (1, 2, ...) in the column's Format drop-down of View Settings
#   Alignment   column 36   Left, Center or Right
#   Alias       column 44   column heading, when it should differ from the field name
#
#FieldName		Width
#					Type
#								Format
#									Alignment
#											Alias
";

        private const string MailColumns =
@"Icon			1
Reminder		1				4
Importance		1
Received		18	olDateTime	3
nameRelated		18					Right
domainRelated	18
me				4					Center
tos				5					Right
ccs				5					Right
Size			9				3	Right
Categories		3
Attachment		12	olYesNo		4
Subject			60
Flag Status		4
";

        private const string ContactColumns =
@"Icon			1
File As			25
Company			20
Department		15
Job Title		15
Email			25
Business Phone	15
Mobile Phone	15
Categories		10
Flag Status		4
";

        private const string TaskColumns =
@"Icon			1
Complete		3
Priority		3
Subject			50
Status			12
Due Date		15
Start Date		15
% Complete		8
Categories		10
In Folder		12
";

        /// <summary>The View Columns file written on first use for <paramref name="kind"/>.</summary>
        public static string DefaultContent(ItemKind kind)
        {
            switch (kind)
            {
                case ItemKind.Contact: return Header + ContactColumns;
                case ItemKind.Task: return Header + TaskColumns;
                default: return Header + MailColumns;
            }
        }

        public static List<ViewColumn> Load(string path)
        {
            return Parse(File.ReadAllLines(path, Encoding.UTF8));
        }

        public static void CreateDefault(string path, ItemKind kind)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            File.WriteAllText(path, DefaultContent(kind).Replace("\r\n", "\n").Replace("\n", "\r\n"), new UTF8Encoding(true));
        }

        public static List<ViewColumn> Parse(IEnumerable<string> lines)
        {
            var columns = new List<ViewColumn>();
            int lineNo = 0;
            foreach (var raw in lines)
            {
                lineNo++;
                var line = raw.TrimEnd();
                if (line.Trim().Length == 0 || line.TrimStart().StartsWith("#", StringComparison.Ordinal))
                    continue;

                var fields = SplitAligned(line);
                var c = new ViewColumn { Line = lineNo, FieldName = fields[0] };
                if (string.IsNullOrEmpty(c.FieldName))
                    throw new FormatException("Line " + lineNo + ": the field name must start at the beginning of the line.");
                c.Width = ParseInt(fields[1], lineNo, "Width");
                c.Format = ParseInt(fields[3], lineNo, "Format");
                if (fields[2] != null)
                {
                    Outlook.OlUserPropertyType type;
                    // Types of built-in fields (e.g. olSize) are not OlUserPropertyType values; they are not needed.
                    if (Enum.TryParse(fields[2], true, out type))
                        c.Type = type;
                }
                if (fields[4] != null)
                {
                    switch (fields[4].ToLowerInvariant())
                    {
                        case "left": c.Align = Outlook.OlAlign.olAlignLeft; break;
                        case "center": c.Align = Outlook.OlAlign.olAlignCenter; break;
                        case "right": c.Align = Outlook.OlAlign.olAlignRight; break;
                        default: throw new FormatException("Line " + lineNo + ": Alignment must be Left, Center or Right, not \"" + fields[4] + "\".");
                    }
                }
                c.Alias = fields[5];
                columns.Add(c);
            }
            return columns;
        }

        /// <summary>
        /// Splits a line at its tabs and assigns each piece to the field whose tab stop it starts at (a piece that
        /// starts before its natural stop, e.g. after a long field name, goes to the next field).
        /// </summary>
        private static string[] SplitAligned(string line)
        {
            var fields = new string[FieldStops.Length];
            int col = 0, field = -1, i = 0;
            while (i < line.Length)
            {
                if (line[i] == '\t')
                {
                    col = (col / TabWidth + 1) * TabWidth;
                    i++;
                    continue;
                }
                int start = i;
                while (i < line.Length && line[i] != '\t')
                    i++;
                var text = line.Substring(start, i - start).Trim();
                int byColumn = 0;
                for (int f = FieldStops.Length - 1; f >= 0; f--)
                {
                    if (col >= FieldStops[f]) { byColumn = f; break; }
                }
                field = Math.Max(field + 1, byColumn);
                if (field < fields.Length && text.Length > 0)
                    fields[field] = text;
                col += i - start;
            }
            return fields;
        }

        private static int? ParseInt(string text, int lineNo, string what)
        {
            if (text == null)
                return null;
            int value;
            if (!int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) || value <= 0)
                throw new FormatException("Line " + lineNo + ": " + what + " must be a positive number, not \"" + text + "\".");
            return value;
        }
    }
}
