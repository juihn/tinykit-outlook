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
            { FindKind.Contact, new[] { "CreationTime", "FileAs", "CompanyName", "Department" } },
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
            // cell: 0 EntryID, 1 date, 2 subject, 3 field 1, 4 field 2, 5 body
            var result = new FindResult
            {
                Kind = folder.Kind,
                Account = folder.Account,
                Folder = folder.Path,
                StoreId = folder.StoreId,
                EntryId = cell(0) as string,
                Date = DateOf(cell(1)),
                DateHasTime = folder.Kind != FindKind.Task, // a due date has no time
                Subject = Text(cell(2)),
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

        /// <summary>The text shown under the results for an item, read from the item itself.</summary>
        public static string Details(Outlook.NameSpace session, FindResult result)
        {
            var item = session.GetItemFromID(result.EntryId, result.StoreId);
            try
            {
                var sb = new StringBuilder();
                if (item is Outlook.ContactItem contact)
                {
                    sb.AppendLine(contact.FileAs);
                    sb.AppendLine(DateText(contact.CreationTime) + " / " + DateText(contact.LastModificationTime));
                    var mails = new[]
                    {
                        Address(contact.Email1Address, contact.Email1DisplayName),
                        Address(contact.Email2Address, contact.Email2DisplayName),
                        Address(contact.Email3Address, contact.Email3DisplayName),
                    }.Where(s => s != null).ToList();
                    if (mails.Count > 0)
                        sb.AppendLine(string.Join(" / ", mails));
                    var phones = new[]
                    {
                        Phone("mobile", contact.MobileTelephoneNumber), Phone("pager", contact.PagerNumber),
                        Phone("business", contact.BusinessTelephoneNumber), Phone("business 2", contact.Business2TelephoneNumber),
                        Phone("company", contact.CompanyMainTelephoneNumber), Phone("home", contact.HomeTelephoneNumber),
                        Phone("home 2", contact.Home2TelephoneNumber), Phone("other", contact.OtherTelephoneNumber),
                        Phone("primary", contact.PrimaryTelephoneNumber), Phone("car", contact.CarTelephoneNumber),
                        Phone("fax", contact.BusinessFaxNumber),
                    }.Where(s => s != null).ToList();
                    if (phones.Count > 0)
                        sb.AppendLine(string.Join(" / ", phones));
                    AppendBody(sb, contact.Body);
                }
                else if (item is Outlook.AppointmentItem appointment)
                {
                    sb.AppendLine("Subject: " + appointment.Subject);
                    sb.AppendLine(AppointmentTime(appointment));
                    if (!string.IsNullOrEmpty(appointment.Location))
                        sb.AppendLine("Location: " + appointment.Location);
                    AppendBody(sb, appointment.Body);
                }
                else if (item is Outlook.TaskItem task)
                {
                    sb.AppendLine("Subject: " + task.Subject);
                    sb.AppendLine("DueDate: " + (DateText(DateOf(task.DueDate)) ?? "None") + " / Status " + StatusText((int)task.Status)
                        + " / Completed " + (DateText(DateOf(task.DateCompleted)) ?? "None"));
                    AppendBody(sb, task.Body);
                }
                else
                {
                    // Mail, meeting requests and responses, reports: read late-bound.
                    dynamic d = item;
                    sb.AppendLine("Subject: " + (string)d.Subject);
                    DateTime received = d.ReceivedTime;
                    sb.AppendLine("Received: " + FullTimeText(received));
                    string name = null, address = null;
                    try
                    {
                        name = d.SenderName;
                        address = item is Outlook.MailItem mail ? Compose.RecipientCommands.SenderSmtp(mail) : (string)d.SenderEmailAddress;
                    }
                    catch (COMException)
                    {
                        // no sender (e.g. a report)
                    }
                    sb.AppendLine("Sender: " + name + (string.IsNullOrEmpty(address) ? "" : " <" + address + ">"));
                    AppendBody(sb, (string)d.Body);
                }
                return sb.ToString();
            }
            finally
            {
                Marshal.ReleaseComObject(item);
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

        private static void AppendBody(StringBuilder sb, string body)
        {
            if (string.IsNullOrEmpty(body))
                return;
            sb.AppendLine();
            sb.Append(body.Replace("\r\n", "\n").Replace("\r", "\n").Replace("\n", Environment.NewLine).TrimEnd());
        }

        private static string Address(string address, string displayName)
        {
            if (string.IsNullOrEmpty(address))
                return null;
            return string.IsNullOrEmpty(displayName) ? address : address + " (" + displayName + ")";
        }

        private static string Phone(string label, string number)
        {
            return string.IsNullOrEmpty(number) ? null : label + ": " + number;
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

        private static string OneLine(string text)
        {
            if (string.IsNullOrEmpty(text))
                return text;
            return System.Text.RegularExpressions.Regex.Replace(text, @"\s+", " ").Trim();
        }
    }
}
