using System;
using System.Collections.Generic;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>The kinds of folders Custom Filter works in; each has its own set of fields.</summary>
    internal enum CustomFilterKind
    {
        Mail,
        Calendar,
        Contact,
        Task,
    }

    /// <summary>A check box of Custom Filter: a label and the properties it searches (any of them).</summary>
    internal sealed class CustomFilterField
    {
        public string Label;
        public string[] Properties;
        public bool CheckedByDefault = true;

        public CustomFilterField(string label, bool checkedByDefault, params string[] properties)
        {
            Label = label;
            CheckedByDefault = checkedByDefault;
            Properties = properties;
        }
    }

    /// <summary>
    /// Custom Filter: the items where the text appears (LIKE '%text%') in any of the checked fields, the fields joined
    /// with OR; a field with several properties (e.g. a contact's names) matches if any of them does.
    /// </summary>
    internal static class CustomFilter
    {
        private const string PropTag = "http://schemas.microsoft.com/mapi/proptag/";
        private const string ContactId = "http://schemas.microsoft.com/mapi/id/{00062004-0000-0000-C000-000000000046}/";

        private const string Body = "urn:schemas:httpmail:textdescription";                 // PR_BODY
        private const string FromName = "urn:schemas:httpmail:fromname";                     // PR_SENT_REPRESENTING_NAME
        private const string SenderName = PropTag + "0x0C1A001F";                            // PR_SENDER_NAME
        private const string RepresentingSmtp = PropTag + "0x5D02001F";                      // PR_SENT_REPRESENTING_SMTP_ADDRESS
        private const string RepresentingEmail = PropTag + "0x0065001F";                     // PR_SENT_REPRESENTING_EMAIL_ADDRESS

        private static readonly Dictionary<CustomFilterKind, CustomFilterField[]> Fields = new Dictionary<CustomFilterKind, CustomFilterField[]>
        {
            {
                CustomFilterKind.Mail, new[]
                {
                    new CustomFilterField("Sender Name", true, FromName, SenderName),
                    new CustomFilterField("From Address", true, Dasl.SenderSmtp, Dasl.SenderEmail),
                    new CustomFilterField("nameRelated", true, Dasl.NameRelated),
                    new CustomFilterField("domainRelated", true, Dasl.DomainRelated),
                    new CustomFilterField("Recipient Name", true,
                        PropTag + "0x0E04001F", PropTag + "0x0E03001F", PropTag + "0x0E02001F"), // PR_DISPLAY_TO / CC / BCC
                    new CustomFilterField("Subject", true, Dasl.Subject),
                    new CustomFilterField("Message Body", false, Body),
                }
            },
            {
                CustomFilterKind.Calendar, new[]
                {
                    new CustomFilterField("Sender Name", true, FromName, SenderName),                     // the organizer
                    new CustomFilterField("From Address", true, RepresentingSmtp, RepresentingEmail, Dasl.SenderSmtp, Dasl.SenderEmail),
                    new CustomFilterField("Subject", true, Dasl.Subject),
                    new CustomFilterField("Message Body", false, Body),
                }
            },
            {
                CustomFilterKind.Contact, new[]
                {
                    new CustomFilterField("Company", true, Dasl.Company),
                    new CustomFilterField("Department", true, Dasl.Department),
                    new CustomFilterField("Name", true,
                        "urn:schemas:contacts:givenName", "urn:schemas:contacts:sn", "urn:schemas:contacts:middlename",
                        "urn:schemas:contacts:nickname", ContactId + "8080001f", ContactId + "8090001f"), // + e-mail 1/2 display names
                    new CustomFilterField("Email Address", true, ContactId + "8083001f", ContactId + "8093001f"), // e-mail 1/2
                    new CustomFilterField("Phone Number", true,
                        PropTag + "0x3A08001F", PropTag + "0x3A1B001F", // business, business 2
                        PropTag + "0x3A1C001F", PropTag + "0x3A21001F", // mobile, pager
                        PropTag + "0x3A09001F", PropTag + "0x3A2F001F", // home, home 2
                        PropTag + "0x3A1F001F", PropTag + "0x3A57001F", // other, company main
                        PropTag + "0x3A1A001F"),                        // primary
                    new CustomFilterField("Notes", false, Body),
                }
            },
            {
                CustomFilterKind.Task, new[]
                {
                    new CustomFilterField("Subject", true, Dasl.Subject),
                    new CustomFilterField("Message Body", false, Body),
                }
            },
        };

        public static CustomFilterField[] FieldsOf(CustomFilterKind kind)
        {
            return Fields[kind];
        }

        /// <summary>The kind of folder the explorer shows, or null when Custom Filter does not apply there.</summary>
        public static CustomFilterKind? KindOf(Outlook.MAPIFolder folder)
        {
            if (folder == null)
                return null;
            switch (folder.DefaultItemType)
            {
                case Outlook.OlItemType.olMailItem: return CustomFilterKind.Mail;
                case Outlook.OlItemType.olAppointmentItem: return CustomFilterKind.Calendar;
                case Outlook.OlItemType.olContactItem: return CustomFilterKind.Contact;
                case Outlook.OlItemType.olTaskItem: return CustomFilterKind.Task;
                default: return null;
            }
        }

        /// <summary>
        /// The filter, one checked field per line ("OR" at the start of the next ones), or "" when there is no text or no
        /// field checked. View.Filter takes it as is (without "@SQL=").
        /// </summary>
        public static string Build(string text, IEnumerable<CustomFilterField> checkedFields)
        {
            text = (text ?? "").Trim();
            if (text.Length == 0)
                return "";
            var pattern = "%" + text + "%";
            var lines = checkedFields
                .Select(f => f.Properties.Length == 1
                    ? Dasl.Like(f.Properties[0], pattern)
                    : "(" + string.Join(" OR ", f.Properties.Select(p => Dasl.Like(p, pattern))) + ")")
                .ToList();
            if (lines.Count == 0)
                return "";
            return lines.Count == 1 ? lines[0] : "(" + lines[0] + Environment.NewLine + string.Join(Environment.NewLine, lines.Skip(1).Select(l => "OR " + l)) + ")";
        }
    }
}
