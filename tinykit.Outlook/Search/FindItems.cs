using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.Filtering;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Search
{
    /// <summary>The kinds of items Find Items looks for, each in the folders of that kind.</summary>
    internal enum FindKind
    {
        Mail,
        Appointment,
        Contact,
        Task,
    }

    /// <summary>A found item: what the result list shows, and how to open it again.</summary>
    internal sealed class FindResult
    {
        public FindKind Kind;
        public string Account;
        public string Folder;
        public DateTime? Date;
        public bool DateHasTime;
        public string Subject;
        public string Field1;
        public string Field2;
        public string Field3;
        public string EntryId;
        public string StoreId;
    }

    /// <summary>A folder to search, with what the result list needs to show its items.</summary>
    internal sealed class FindFolder
    {
        public Outlook.MAPIFolder Folder;
        public FindKind Kind;
        public string Account;
        public string Path;
        public string StoreId;
    }

    /// <summary>
    /// Find Items: the text anywhere (LIKE '%text%') in the subject and the fields of each kind of item, read folder by
    /// folder with Folder.GetTable (fast, without opening the items), in the chosen accounts (stores).
    /// </summary>
    internal static class FindItems
    {
        private const string PropTag = "http://schemas.microsoft.com/mapi/proptag/";
        private const string ContactId = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/";
        private const string Body = "urn:schemas:httpmail:textdescription";                                              // PR_BODY
        private const string FromName = "urn:schemas:httpmail:fromname";                                                  // PR_SENT_REPRESENTING_NAME
        private const string SenderName = PropTag + "0x0C1A001F";                                                         // PR_SENDER_NAME
        private const string Location = "urn:schemas:calendar:location";
        private const string TimeZoneDescription = "http://schemas.microsoft.com/mapi/id/{00062002-0000-0000-C000-000000000046}/8234001F";
        private const string Hidden = PropTag + "0x10F4000B";                                                             // PR_ATTR_HIDDEN
        private const string StringType = "/0x0000001F";

        private static readonly CultureInfo Korean = CultureInfo.GetCultureInfo("ko-KR");

        /// <summary>The fields searched besides the subject, per kind (the body or notes only when asked).</summary>
        private static readonly Dictionary<FindKind, string[]> Fields = new Dictionary<FindKind, string[]>
        {
            { FindKind.Mail, new[] { FromName, SenderName } },
            { FindKind.Appointment, new[] { Location } },
            {
                FindKind.Contact, new[]
                {
                    Dasl.FileAs, Dasl.Company, Dasl.Department, "urn:schemas:contacts:nickname",
                    ContactId + "8083001f", ContactId + "8093001f", ContactId + "80a3001f", // e-mail 1/2/3 addresses
                    ContactId + "8080001f", ContactId + "8090001f", ContactId + "80a0001f", // e-mail 1/2/3 display names
                    PropTag + "0x3A08001F", PropTag + "0x3A1B001F", // business, business 2
                    PropTag + "0x3A1C001F", PropTag + "0x3A21001F", // mobile, pager
                    PropTag + "0x3A09001F", PropTag + "0x3A2F001F", // home, home 2
                    PropTag + "0x3A1F001F", PropTag + "0x3A57001F", // other, company main
                    PropTag + "0x3A1A001F", PropTag + "0x3A1E001F", // primary, car
                    PropTag + "0x3A24001F",                         // business fax
                }
            },
            { FindKind.Task, new string[0] },
        };

        // The table columns per kind: date, subject, field 1, field 2 (field 3 is the body).
        private static readonly Dictionary<FindKind, string[]> Columns = new Dictionary<FindKind, string[]>
        {
            // A named property as a table column needs its type (string: /0x0000001F); a filter takes it without.
            { FindKind.Mail, new[] { "ReceivedTime", "Subject", Dasl.NameRelated + StringType, Dasl.DomainRelated + StringType } },
            { FindKind.Appointment, new[] { "Start", "Subject", "End", TimeZoneDescription } },
            { FindKind.Contact, new[] { "CreationTime", "FullName", "CompanyName", "Department" } }, // + NickName (cell 6)
            { FindKind.Task, new[] { "DueDate", "Subject", "Status", "DateCompleted" } },
        };

        public static string LabelOf(FindKind kind)
        {
            switch (kind)
            {
                case FindKind.Mail: return "Mail";
                case FindKind.Appointment: return "Calendar";
                case FindKind.Contact: return "Contacts";
                default: return "Tasks";
            }
        }

        /// <summary>The Office image of the kind (its "New ..." command).</summary>
        public static string ImageMsoOf(FindKind kind)
        {
            switch (kind)
            {
                case FindKind.Mail: return "NewMailMessage";
                case FindKind.Appointment: return "NewAppointment";
                case FindKind.Contact: return "NewContact";
                default: return "NewTask";
            }
        }

        public static FindKind? KindOf(Outlook.OlItemType type)
        {
            switch (type)
            {
                case Outlook.OlItemType.olMailItem: return FindKind.Mail;
                case Outlook.OlItemType.olAppointmentItem: return FindKind.Appointment;
                case Outlook.OlItemType.olContactItem: return FindKind.Contact;
                case Outlook.OlItemType.olTaskItem: return FindKind.Task;
                default: return null;
            }
        }

        /// <summary>The filter for one kind of folder (without "@SQL="), or "" when there is no text.</summary>
        public static string FilterFor(FindKind kind, string text, bool body)
        {
            text = (text ?? "").Trim();
            if (text.Length == 0)
                return "";
            var pattern = "%" + text + "%";
            var properties = new List<string> { Dasl.Subject };
            properties.AddRange(Fields[kind]);
            if (body)
                properties.Add(Body);
            return "(" + string.Join(" OR ", properties.Select(p => Dasl.Like(p, pattern))) + ")";
        }

        /// <summary>
        /// The folders of the chosen kinds in a store, all levels (hidden folders and Sync Issues left out).
        /// </summary>
        public static List<FindFolder> FoldersOf(Outlook.Store store, ICollection<FindKind> kinds)
        {
            var found = new List<FindFolder>();
            Outlook.MAPIFolder root;
            try
            {
                root = store.GetRootFolder();
            }
            catch (COMException ex)
            {
                Log.Info("Find Items: store " + store.DisplayName + " skipped: " + ex.Message);
                return found;
            }
            var skip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var id in new[]
            {
                Outlook.OlDefaultFolders.olFolderSyncIssues, Outlook.OlDefaultFolders.olFolderConflicts,
                Outlook.OlDefaultFolders.olFolderLocalFailures, Outlook.OlDefaultFolders.olFolderServerFailures,
            })
            {
                try
                {
                    skip.Add(store.GetDefaultFolder(id).EntryID);
                }
                catch (COMException)
                {
                    // not in this store
                }
            }
            var prefix = root.FolderPath.Replace("%2F", "/").TrimEnd('\\') + "\\";
            var account = store.DisplayName;
            var storeId = store.StoreID;
            Collect(root, kinds, skip, prefix, account, storeId, found);
            return found;
        }

        private static void Collect(Outlook.MAPIFolder parent, ICollection<FindKind> kinds, HashSet<string> skip, string prefix,
            string account, string storeId, List<FindFolder> found)
        {
            foreach (Outlook.MAPIFolder folder in parent.Folders)
            {
                try
                {
                    if (skip.Contains(folder.EntryID) || IsHidden(folder))
                        continue;
                    var kind = KindOf(folder.DefaultItemType);
                    if (kind != null && kinds.Contains(kind.Value))
                    {
                        var path = folder.FolderPath.Replace("%2F", "/"); // FolderPath escapes "/" in folder names
                        found.Add(new FindFolder
                        {
                            Folder = folder,
                            Kind = kind.Value,
                            Account = account,
                            Path = path.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) ? path.Substring(prefix.Length) : folder.Name,
                            StoreId = storeId,
                        });
                    }
                    Collect(folder, kinds, skip, prefix, account, storeId, found);
                }
                catch (COMException ex)
                {
                    Log.Info("Find Items: folder skipped: " + ex.Message);
                }
            }
        }

        private static bool IsHidden(Outlook.MAPIFolder folder)
        {
            try
            {
                return folder.PropertyAccessor.GetProperty(Hidden) is bool b && b;
            }
            catch (COMException)
            {
                return false; // not set
            }
        }

        /// <summary>
        /// The items of one folder that match, up to <paramref name="max"/>; <paramref name="keepGoing"/> is asked between
        /// chunks of rows (it may pump messages) and stops the reading when it returns false.
        /// </summary>
        public static List<FindResult> Search(FindFolder folder, string filter, int max, Func<bool> keepGoing)
        {
            var results = new List<FindResult>();
            var table = folder.Folder.GetTable("@SQL=" + filter, Outlook.OlTableContents.olUserItems);
            table.Columns.RemoveAll();
            var columns = new List<string> { "EntryID" };
            columns.AddRange(Columns[folder.Kind]);
            columns.Add(Body);
            if (folder.Kind == FindKind.Contact)
                columns.Add("NickName"); // the subject of a contact: full name / nickname
            var present = new bool[columns.Count];
            for (int i = 0; i < columns.Count; i++)
            {
                try
                {
                    table.Columns.Add(columns[i]);
                    present[i] = true;
                }
                catch (Exception ex) // COMException, or ArgumentException for a named property this store does not have
                {
                    Log.Info("Find Items: column " + columns[i] + " not available in " + folder.Path + ": " + ex.Message);
                }
            }
            // GetArray returns only the columns that were added, in order.
            var index = new int[columns.Count];
            for (int i = 0, k = 0; i < columns.Count; i++)
                index[i] = present[i] ? k++ : -1;

            while (!table.EndOfTable && results.Count < max)
            {
                var rows = (object[,])table.GetArray(200);
                if (rows == null)
                    break;
                int r0 = rows.GetLowerBound(0), c0 = rows.GetLowerBound(1);
                for (int r = r0; r <= rows.GetUpperBound(0) && results.Count < max; r++)
                {
                    Func<int, object> cell = i => index[i] < 0 ? null : rows[r, c0 + index[i]];
                    results.Add(ToResult(folder, cell));
                }
                if (!keepGoing())
                    break;
            }
            return results;
        }

        private static FindResult ToResult(FindFolder folder, Func<int, object> cell)
        {
            // cell: 0 EntryID, 1 date, 2 subject, 3 field 1, 4 field 2, 5 body (contacts: 2 full name, 6 nickname)
            var result = new FindResult
            {
                Kind = folder.Kind,
                Account = folder.Account,
                Folder = folder.Path,
                StoreId = folder.StoreId,
                EntryId = cell(0) as string,
                Date = DateOf(cell(1)),
                DateHasTime = folder.Kind != FindKind.Task, // a due date has no time
                Subject = folder.Kind == FindKind.Contact ? JoinPresent(" / ", Text(cell(2)), Text(cell(6))) : Text(cell(2)),
                Field3 = OneLine(Text(cell(5))),
            };
            switch (folder.Kind)
            {
                case FindKind.Appointment:
                    result.Field1 = ListDateText(DateOf(cell(3)), true);
                    result.Field2 = Text(cell(4));
                    break;
                case FindKind.Task:
                    result.Field1 = StatusText(cell(3));
                    result.Field2 = ListDateText(DateOf(cell(4)), false);
                    break;
                default:
                    result.Field1 = Text(cell(3));
                    result.Field2 = Text(cell(4));
                    break;
            }
            return result;
        }

        /// <summary>
        /// The text shown under the results for an item, read from the item itself. A field without a value is left out
        /// with its label, and a line with nothing in it is left out.
        /// </summary>
        public static string Details(Outlook.NameSpace session, FindResult result)
        {
            var item = session.GetItemFromID(result.EntryId, result.StoreId);
            try
            {
                var sb = new StringBuilder();
                if (item is Outlook.ContactItem contact)
                {
                    Line(sb, contact.FileAs);
                    Line(sb, Part("Created", ListDateText(DateOf(contact.CreationTime), false), " "),
                        Part("Modified", ListDateText(DateOf(contact.LastModificationTime), false), " "),
                        Part("Anniversary", ListDateText(DateOf(contact.Anniversary), false), " "));
                    Line(sb, Part("Full Name", contact.FullName), Part("NickName", contact.NickName));
                    Line(sb, contact.CompanyName, contact.Department, contact.JobTitle);
                    Line(sb, Address(contact.Email1Address, contact.Email1DisplayName), Address(contact.Email2Address, contact.Email2DisplayName),
                        Address(contact.Email3Address, contact.Email3DisplayName));
                    Line(sb, Part("mobile", contact.MobileTelephoneNumber), Part("pager", contact.PagerNumber),
                        Part("work", contact.BusinessTelephoneNumber), Part("work 2", contact.Business2TelephoneNumber),
                        Part("company", contact.CompanyMainTelephoneNumber), Part("home", contact.HomeTelephoneNumber),
                        Part("home 2", contact.Home2TelephoneNumber), Part("other", contact.OtherTelephoneNumber),
                        Part("primary", contact.PrimaryTelephoneNumber), Part("car", contact.CarTelephoneNumber),
                        Part("fax", contact.BusinessFaxNumber));
                    Line(sb, Part("Business Address", OneLineAddress(contact.BusinessAddress)));
                    Line(sb, Part("Free/Busy Address", contact.InternetFreeBusyAddress));
                    Line(sb, Part("IM Address", contact.IMAddress));
                    AppendBody(sb, contact.Body);
                }
                else if (item is Outlook.AppointmentItem appointment)
                {
                    Line(sb, Part("Subject", appointment.Subject));
                    Line(sb, AppointmentTime(appointment));
                    Line(sb, Part("Location", appointment.Location));
                    AppendBody(sb, appointment.Body);
                }
                else if (item is Outlook.TaskItem task)
                {
                    Line(sb, Part("Subject", task.Subject));
                    Line(sb, Part("DueDate", ListDateText(DateOf(task.DueDate), false)), Part("Status", StatusText((int)task.Status), " "),
                        Part("Completed", ListDateText(DateOf(task.DateCompleted), false), " "));
                    if (task.ReminderSet)
                        Line(sb, Part("Reminder Time", ListDateText(DateOf(task.ReminderTime), true)));
                    if (task.IsRecurring)
                        Line(sb, Part("Recurrence", RecurrenceText(task.GetRecurrencePattern())));
                    AppendBody(sb, task.Body);
                }
                else
                {
                    // Mail, meeting requests and responses, reports: read late-bound.
                    dynamic d = item;
                    Line(sb, Part("Subject", (string)d.Subject));
                    DateTime received = d.ReceivedTime;
                    Line(sb, Part("Received", FullTimeText(received)));
                    string name = null, address = null, to = null, cc = null;
                    try
                    {
                        name = d.SenderName;
                        address = item is Outlook.MailItem mail ? MailInfo.SenderSmtpOf(mail) : (string)d.SenderEmailAddress;
                    }
                    catch (Exception)
                    {
                        // no sender (e.g. a report)
                    }
                    try
                    {
                        to = d.To;
                        cc = d.CC;
                    }
                    catch (Exception)
                    {
                        // no recipients (e.g. a meeting response)
                    }
                    var sender = string.IsNullOrEmpty(address) ? name : (string.IsNullOrEmpty(name) ? "" : name + " ") + "<" + address + ">";
                    Line(sb, Part("Sender", sender));
                    Line(sb, Part("Recipients", to));
                    Line(sb, Part("Carbon Copy", cc));
                    AppendBody(sb, (string)d.Body);
                }
                return sb.ToString();
            }
            finally
            {
                Marshal.ReleaseComObject(item);
            }
        }

        /// <summary>"label: value", or null when there is no value (so the label is left out too).</summary>
        private static string Part(string label, string value, string separator = ": ")
        {
            return string.IsNullOrWhiteSpace(value) ? null : label + separator + value.Trim();
        }

        /// <summary>The parts that have a value, joined with " / "; nothing when none has.</summary>
        private static void Line(StringBuilder sb, params string[] parts)
        {
            var present = parts.Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p.Trim()).ToList();
            if (present.Count > 0)
                sb.AppendLine(string.Join(" / ", present));
        }

        private static string OneLineAddress(string address)
        {
            if (string.IsNullOrWhiteSpace(address))
                return null;
            return string.Join(", ", address.Replace("\r\n", "\n").Split('\n').Select(l => l.Trim()).Where(l => l.Length > 0));
        }

        private static readonly string[] WeekdayNames = { "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat" };

        /// <summary>
        /// A recurrence in words: "every 2 weeks on Mon, Thu", "monthly on day 15", "yearly on the 2nd Tue of March", then
        /// "from 'yy.MM.ddW" and the end (until a date, N times, or no end).
        /// </summary>
        private static string RecurrenceText(Outlook.RecurrencePattern p)
        {
            var interval = p.Interval;
            string text;
            switch (p.RecurrenceType)
            {
                case Outlook.OlRecurrenceType.olRecursDaily:
                    text = interval > 1 ? "every " + interval + " days" : "daily";
                    break;
                case Outlook.OlRecurrenceType.olRecursWeekly:
                    text = (interval > 1 ? "every " + interval + " weeks" : "weekly") + " on " + Weekdays(p.DayOfWeekMask);
                    break;
                case Outlook.OlRecurrenceType.olRecursMonthly:
                    text = (interval > 1 ? "every " + interval + " months" : "monthly") + " on day " + p.DayOfMonth;
                    break;
                case Outlook.OlRecurrenceType.olRecursMonthNth:
                    text = (interval > 1 ? "every " + interval + " months" : "monthly") + " on the " + Nth(p.Instance) + " " + Weekdays(p.DayOfWeekMask);
                    break;
                case Outlook.OlRecurrenceType.olRecursYearly:
                    text = "yearly on " + CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(p.MonthOfYear) + " " + p.DayOfMonth;
                    break;
                case Outlook.OlRecurrenceType.olRecursYearNth:
                    text = "yearly on the " + Nth(p.Instance) + " " + Weekdays(p.DayOfWeekMask) + " of "
                        + CultureInfo.InvariantCulture.DateTimeFormat.GetMonthName(p.MonthOfYear);
                    break;
                default:
                    text = p.RecurrenceType.ToString();
                    break;
            }
            text += ", from " + ListDateText(DateOf(p.PatternStartDate), false);
            if (p.NoEndDate)
                text += ", no end";
            else if (p.Occurrences > 0 && DateOf(p.PatternEndDate) == null)
                text += ", " + p.Occurrences + " times";
            else
                text += " until " + ListDateText(DateOf(p.PatternEndDate), false);
            return text;
        }

        private static string Weekdays(Outlook.OlDaysOfWeek mask)
        {
            var days = new List<string>();
            for (int i = 0; i < 7; i++)
            {
                if (((int)mask & (1 << i)) != 0)
                    days.Add(WeekdayNames[i]);
            }
            return string.Join(", ", days);
        }

        private static string Nth(int instance)
        {
            switch (instance)
            {
                case 1: return "1st";
                case 2: return "2nd";
                case 3: return "3rd";
                case 5: return "last";
                default: return instance + "th";
            }
        }

        /// <summary>Opens the item in its own window.</summary>
        public static void Open(Outlook.NameSpace session, FindResult result)
        {
            dynamic item = session.GetItemFromID(result.EntryId, result.StoreId);
            item.Display();
        }

        private static string AppointmentTime(Outlook.AppointmentItem appointment)
        {
            string zone = null;
            try
            {
                zone = appointment.StartTimeZone.Name;
            }
            catch (COMException)
            {
            }
            if (appointment.AllDayEvent)
                return DateText(appointment.Start) + " ~ " + DateText(appointment.End.AddDays(-1)) + " (all day)";
            return DateTimeText(appointment.Start) + " ~ " + DateTimeText(appointment.End) + (string.IsNullOrEmpty(zone) ? "" : " " + zone);
        }

        /// <summary>The body right after the fields, without its lines that are empty or only white space (e.g. &amp;nbsp;).</summary>
        private static void AppendBody(StringBuilder sb, string body)
        {
            if (string.IsNullOrEmpty(body))
                return;
            var lines = body.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n').Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
            if (lines.Count == 0)
                return;
            sb.Append(string.Join(Environment.NewLine, lines.Select(l => l.TrimEnd())));
        }

        private static string Address(string address, string displayName)
        {
            if (string.IsNullOrEmpty(address))
                return null;
            return string.IsNullOrEmpty(displayName) ? address : address + " " + displayName;
        }

        private static string StatusText(object status)
        {
            if (!(status is int))
                return null;
            switch ((Outlook.OlTaskStatus)(int)status)
            {
                case Outlook.OlTaskStatus.olTaskNotStarted: return "Not Started";
                case Outlook.OlTaskStatus.olTaskInProgress: return "In Progress";
                case Outlook.OlTaskStatus.olTaskComplete: return "Completed";
                case Outlook.OlTaskStatus.olTaskWaiting: return "Waiting on someone else";
                case Outlook.OlTaskStatus.olTaskDeferred: return "Deferred";
                default: return status.ToString();
            }
        }

        /// <summary>Outlook's "None" date (4501-01-01) and missing values as null.</summary>
        private static DateTime? DateOf(object value)
        {
            if (!(value is DateTime))
                return null;
            var date = (DateTime)value;
            return date.Year >= 4500 ? (DateTime?)null : date;
        }

        /// <summary>'yyMMdd and the Korean weekday: '261005일.</summary>
        public static string DateText(DateTime? date)
        {
            if (date == null)
                return null;
            return "'" + date.Value.ToString("yyMMdd", CultureInfo.InvariantCulture) + date.Value.ToString("ddd", Korean);
        }

        /// <summary>'261005일 14:30.</summary>
        /// <summary>
        /// The result list's dates: 'yy.MM.dd, the Korean weekday, and HH:mm when <paramref name="withTime"/>: '26.10.05일 14:30.
        /// Local time, as the table gives built-in date properties (ReceivedTime, Start, ...) in local time.
        /// </summary>
        public static string ListDateText(DateTime? date, bool withTime)
        {
            if (date == null)
                return null;
            return "'" + date.Value.ToString("yy.MM.dd", CultureInfo.InvariantCulture) + date.Value.ToString("ddd", Korean)
                + (withTime ? " " + date.Value.ToString("HH:mm", CultureInfo.InvariantCulture) : "");
        }

        public static string DateTimeText(DateTime? date)
        {
            return date == null ? null : DateText(date) + " " + date.Value.ToString("HH:mm", CultureInfo.InvariantCulture);
        }

        /// <summary>'261005일 14:30:15 +09:00 (the local offset on that day).</summary>
        private static string FullTimeText(DateTime date)
        {
            var offset = TimeZoneInfo.Local.GetUtcOffset(date);
            return DateText(date) + " " + date.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + " "
                + (offset < TimeSpan.Zero ? "-" : "+") + offset.ToString("hh\\:mm", CultureInfo.InvariantCulture);
        }

        private static string Text(object value)
        {
            return value == null ? null : Convert.ToString(value, CultureInfo.CurrentCulture);
        }

        /// <summary>The values that are there, joined with <paramref name="separator"/> ("" when none is).</summary>
        private static string JoinPresent(string separator, params string[] values)
        {
            return string.Join(separator, values.Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim()));
        }

        private static string OneLine(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            return System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();
        }
    }
}
