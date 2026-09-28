using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
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
    /// View Columns - &lt;kind&gt;.txt: the columns of a table view, one per line. The values of a line are separated by
    /// tabs (several in a row count as one) and matched to FieldName, Width, Type, Format, Alignment, Alias in that
    /// order, each by its format, so fields can be left out. Lines starting with # are comments.
    /// </summary>
    internal static class ViewColumns
    {
        private const string Header =
@"# tinykit Outlook: the columns of the current table view. One file per kind of folder:
#   View Columns - Mail.txt, View Columns - Contacts.txt, View Columns - Tasks.txt.
#   View Columns (ribbon, View group) replaces the view's columns with the lines below, in this order.
#   Ctrl+click View Columns to edit the file of the folder you are in. Lines starting with # are skipped.
#
# Fields, in this order, separated by tabs (several tabs in a row count as one, so they can be lined up). A field can
# be left out: each value goes to the next field whose format it fits.
#   FieldName   field name as in View Settings > Columns (e.g. Received, Flag Status) or a user-defined field
#   Width       width in characters (a number)
#   Type        type of a user-defined field, used only to create it in a folder that lacks it:
#               olText, olDateTime, olInteger, olNumber, olYesNo (OlUserPropertyType)
#               https://learn.microsoft.com/dotnet/api/microsoft.office.interop.outlook.oluserpropertytype
#   Format      position (1, 2, ...) in the column's Format drop-down of View Settings (a number after Width)
#   Alignment   Left, Center or Right
#   Alias       column heading, when it should differ from the field name (any other text)
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

                var fields = SplitFields(line, lineNo);
                var c = new ViewColumn { Line = lineNo, FieldName = fields[FieldName] };
                c.Width = PositiveInt(fields[Width]);
                c.Format = PositiveInt(fields[Format]);
                Outlook.OlUserPropertyType type;
                // Types of built-in fields (e.g. olSize) are not OlUserPropertyType values; they are not needed.
                if (fields[Type] != null && Enum.TryParse(fields[Type], true, out type))
                    c.Type = type;
                if (fields[Alignment] != null)
                    c.Align = AlignOf(fields[Alignment]);
                c.Alias = fields[Alias];
                columns.Add(c);
            }
            return columns;
        }

        // Field positions, in the order the values of a line are matched to them.
        private const int FieldName = 0, Width = 1, Type = 2, Format = 3, Alignment = 4, Alias = 5;
        private static readonly string[] FieldNames = { "FieldName", "Width", "Type", "Format", "Alignment", "Alias" };

        /// <summary>
        /// Splits a line at its tabs (several tabs in a row count as one) and assigns the values to the fields in order:
        /// the first value is the field name; each further value goes to the next field whose format it fits, skipping
        /// the fields it does not fit (Width and Format: a positive number; Type: ol + a capital letter, e.g. olDateTime;
        /// Alignment: Left, Center or Right; Alias: any text).
        /// </summary>
        private static string[] SplitFields(string line, int lineNo)
        {
            var values = line.Split('\t').Select(v => v.Trim()).Where(v => v.Length > 0).ToList();
            var fields = new string[FieldNames.Length];
            fields[FieldName] = values[0];
            int next = Width;
            foreach (var value in values.Skip(1))
            {
                int f = next;
                while (f < fields.Length && !Fits(f, value))
                    f++;
                if (f == fields.Length)
                    throw new FormatException("Line " + lineNo + ": \"" + value + "\" does not fit "
                        + (next < fields.Length ? string.Join(", ", FieldNames.Skip(next)) : "any field (all six are used)") + ".");
                fields[f] = value;
                next = f + 1;
            }
            return fields;
        }

        private static bool Fits(int field, string value)
        {
            switch (field)
            {
                case Width:
                case Format:
                    return PositiveInt(value) != null;
                case Type:
                    return value.Length > 2 && value.StartsWith("ol", StringComparison.OrdinalIgnoreCase) && char.IsUpper(value[2]);
                case Alignment:
                    return AlignOf(value) != null;
                default:
                    return true; // Alias: any text
            }
        }

        private static int? PositiveInt(string text)
        {
            int value;
            return text != null && int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) && value > 0
                ? value : (int?)null;
        }

        private static Outlook.OlAlign? AlignOf(string text)
        {
            switch (text.ToLowerInvariant())
            {
                case "left": return Outlook.OlAlign.olAlignLeft;
                case "center": return Outlook.OlAlign.olAlignCenter;
                case "right": return Outlook.OlAlign.olAlignRight;
                default: return null;
            }
        }
    }
}
