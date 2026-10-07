using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Filtering
{
    /// <summary>The filterable attributes of one mail/meeting item.</summary>
    internal sealed class MailInfo
    {
        private const string NormalizedSubjectTag = "http://schemas.microsoft.com/mapi/proptag/0x0E1D001F";
        private const string SenderSmtpTag = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F";

        // Second-level labels that are registry categories, not organizations (co.kr, com.au, ...).
        private static readonly HashSet<string> GenericSecondLevel = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "co", "or", "ac", "go", "ne", "re", "pe", "hs", "ms", "es", "sc", "kg",
            "com", "net", "org", "gov", "edu", "mil", "biz", "info",
        };

        public string Subject { get; private set; }
        public string NormalizedSubject { get; private set; }
        public string SenderSmtp { get; private set; }
        public string SenderAddress { get; private set; }
        public string SenderName { get; private set; }
        public DateTime Time { get; private set; }

        /// <summary>The item's domainRelated (sender's, or first recipient's for my own mail); may be null.</summary>
        public string DomainRelated { get; private set; }

        /// <summary>The item's nameRelated (e.g. "▶ [Kim (contoso)] (+)"); may be null.</summary>
        public string NameRelated { get; private set; }

        public string ValueFor(QuickKind kind)
        {
            switch (kind)
            {
                case QuickKind.Subject:
                    return string.IsNullOrEmpty(NormalizedSubject) ? Subject : NormalizedSubject;
                case QuickKind.From:
                    return SenderSmtp;
                case QuickKind.Domain:
                    return DomainRelated ?? DomainPath(SenderSmtp);
                case QuickKind.Name:
                    return CoreName(NameRelated) ?? SenderName;
            }
            return null;
        }

        /// <summary>
        /// The first mail/meeting item of the explorer selection, or null.
        /// <paramref name="domainRelatedOf"/> and <paramref name="nameRelatedOf"/> supply the item's custom columns (optional).
        /// </summary>
        public static MailInfo FromFirstSelected(Outlook.Explorer explorer, Func<object, string> domainRelatedOf,
            Func<object, string> nameRelatedOf)
        {
            Outlook.Selection selection;
            try
            {
                selection = explorer.Selection;
            }
            catch (COMException)
            {
                return null; // e.g. no item list in the current pane
            }

            int count = selection.Count;
            for (int i = 1; i <= count; i++)
            {
                object item = selection[i];
                try
                {
                    var info = From(item);
                    if (info == null)
                        continue;
                    if (domainRelatedOf != null)
                        info.DomainRelated = domainRelatedOf(item);
                    if (nameRelatedOf != null)
                        info.NameRelated = nameRelatedOf(item);
                    return info;
                }
                finally
                {
                    Marshal.ReleaseComObject(item);
                }
            }
            return null;
        }

        /// <summary>
        /// The name inside a nameRelated value, so received and sent mail of the same person match:
        /// "▶ [Kim (contoso)] (+)" → "Kim (contoso)", "(unialerts)" → "unialerts".
        /// </summary>
        public static string CoreName(string nameRelated)
        {
            if (string.IsNullOrWhiteSpace(nameRelated))
                return null;
            var s = nameRelated.Trim();
            if (s.StartsWith("▶ ", StringComparison.Ordinal))
                s = s.Substring(2).Trim();
            if (s.EndsWith(" (+)", StringComparison.Ordinal))
                s = s.Substring(0, s.Length - 4).Trim();
            if (s.Length > 2 && ((s[0] == '[' && s[s.Length - 1] == ']') || (s[0] == '(' && s[s.Length - 1] == ')')))
                s = s.Substring(1, s.Length - 2).Trim();
            return s.Length == 0 ? null : s;
        }

        /// <summary>
        /// The organization label of an address's host: a@contoso.com → contoso,
        /// a@mail.contoso.co.kr → contoso.
        /// </summary>
        public static string DomainLabel(string smtp)
        {
            int i;
            var labels = HostLabels(smtp, out i);
            return labels == null ? null : labels[i];
        }

        /// <summary>
        /// The domain label followed by its subdomains, nearest first, joined with "/":
        /// a@billing.fabrikam.com → fabrikam/billing, a@x.mail.contoso.co.kr → contoso/mail/x.
        /// </summary>
        public static string DomainPath(string smtp)
        {
            int i;
            var labels = HostLabels(smtp, out i);
            if (labels == null)
                return null;
            var parts = new List<string>();
            for (int j = i; j >= 0; j--)
                parts.Add(labels[j]);
            return string.Join("/", parts);
        }

        /// <summary>
        /// The base (registrable) domain of an address's host: a@billing.fabrikam.com → fabrikam.com,
        /// a@x.mail.contoso.co.kr → contoso.co.kr.
        /// </summary>
        public static string BaseDomain(string smtp)
        {
            int i;
            var labels = HostLabels(smtp, out i);
            return labels == null ? null : string.Join(".", labels, i, labels.Length - i);
        }

        /// <summary>Lower-cased host labels of an address and the index of its organization label; null if no host.</summary>
        private static string[] HostLabels(string smtp, out int labelIndex)
        {
            labelIndex = -1;
            if (string.IsNullOrEmpty(smtp))
                return null;
            int at = smtp.LastIndexOf('@');
            if (at < 0)
                return null;
            var labels = smtp.Substring(at + 1).ToLowerInvariant().Split(new[] { '.' }, StringSplitOptions.RemoveEmptyEntries);
            if (labels.Length == 0)
                return null;
            if (labels.Length == 1)
            {
                labelIndex = 0;
                return labels;
            }

            int i = labels.Length - 2;
            if (labels.Length >= 3 && labels[labels.Length - 1].Length == 2 && GenericSecondLevel.Contains(labels[i]))
                i--;
            labelIndex = i;
            return labels;
        }

        private static MailInfo From(object item)
        {
            var mail = item as Outlook.MailItem;
            if (mail != null)
            {
                var pa = mail.PropertyAccessor;
                return new MailInfo
                {
                    Subject = mail.Subject,
                    NormalizedSubject = GetString(pa, NormalizedSubjectTag) ?? mail.ConversationTopic,
                    SenderAddress = mail.SenderEmailAddress,
                    SenderName = mail.SenderName,
                    SenderSmtp = ResolveSmtp(mail.SenderEmailType, mail.SenderEmailAddress, pa, () => mail.Sender),
                    Time = mail.ReceivedTime,
                };
            }

            var meeting = item as Outlook.MeetingItem;
            if (meeting != null)
            {
                var pa = meeting.PropertyAccessor;
                return new MailInfo
                {
                    Subject = meeting.Subject,
                    NormalizedSubject = GetString(pa, NormalizedSubjectTag) ?? meeting.ConversationTopic,
                    SenderAddress = meeting.SenderEmailAddress,
                    SenderName = meeting.SenderName,
                    SenderSmtp = ResolveSmtp(meeting.SenderEmailType, meeting.SenderEmailAddress, pa, null),
                    Time = meeting.ReceivedTime,
                };
            }
            return null;
        }

        /// <summary>
        /// The SMTP address of the person a mail is from (an Exchange sender resolved), or null. Not the account a mail
        /// being written is sent from: that is Compose.RecipientCommands.SenderSmtp.
        /// </summary>
        public static string SenderSmtpOf(Outlook.MailItem mail)
        {
            return ResolveSmtp(mail.SenderEmailType, mail.SenderEmailAddress, mail.PropertyAccessor, () => mail.Sender);
        }

        private static string ResolveSmtp(string type, string address, Outlook.PropertyAccessor pa, Func<Outlook.AddressEntry> sender)
        {
            if (!string.Equals(type, "EX", StringComparison.OrdinalIgnoreCase))
                return address;

            var smtp = GetString(pa, SenderSmtpTag);
            if (!string.IsNullOrEmpty(smtp))
                return smtp;

            if (sender != null)
            {
                try
                {
                    var entry = sender();
                    var user = entry == null ? null : entry.GetExchangeUser();
                    if (user != null && !string.IsNullOrEmpty(user.PrimarySmtpAddress))
                        return user.PrimarySmtpAddress;
                }
                catch (COMException)
                {
                    // offline / GAL unavailable
                }
            }
            return address;
        }

        private static string GetString(Outlook.PropertyAccessor pa, string schema)
        {
            try
            {
                var s = pa.GetProperty(schema) as string;
                return string.IsNullOrEmpty(s) ? null : s;
            }
            catch (COMException)
            {
                return null; // property not present on this item
            }
        }
    }
}
