using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using tinykit.OutlookAddin.Filtering;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>
    /// Computes domainRelated / nameRelated / me / tos / ccs for one item.
    /// <list type="bullet">
    /// <item>domainRelated — domain label then subdomains, e.g. gojek/invoicing.
    ///   Received: the sender's (on behalf of: the principal's);
    ///   sent by me: first To recipient's (no To: first Cc's).</item>
    /// <item>nameRelated — received: [contact name] if the sender is in Contacts, else the sender display name
    ///   ("(local part)" when the display name is just the address);
    ///   sent by me: "▶ " + first recipient's name (+ " (+)" with several recipients), where a recipient
    ///   picked from an address book keeps its shown name, and a typed address becomes [contact name] or the address.</item>
    /// <item>me — ▶ sent by me, else ● I am in To, ○ I am in Cc, else "-".  tos / ccs — number of To / Cc recipients, "-" for 0.</item>
    /// </list>
    /// </summary>
    internal sealed class CustomFieldCalculator
    {
        private const string PropTag = "http://schemas.microsoft.com/mapi/proptag/";
        private const string SenderName = PropTag + "0x0C1A001F";
        private const string SenderAddrType = PropTag + "0x0C1E001F";
        private const string SenderEmail = PropTag + "0x0C1F001F";
        private const string SenderSmtp = PropTag + "0x5D01001F";
        private const string RepName = PropTag + "0x0042001F";
        private const string RepAddrType = PropTag + "0x0064001F";
        private const string RepEmail = PropTag + "0x0065001F";
        private const string RepSmtp = PropTag + "0x5D02001F";
        private const string RecipientSmtpTag = PropTag + "0x39FE001F";

        private const string SentPrefix = "▶ ";
        private const string MultipleSuffix = " (+)";
        private const string MeInTo = "●";
        private const string MeInCc = "○";
        private const string MeSender = "▶";

        private readonly Outlook.NameSpace _session;
        private readonly Dictionary<string, string> _exToSmtp = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public MyIdentity Me { get; private set; }
        public ContactIndex Contacts { get; private set; }

        public CustomFieldCalculator(Outlook.NameSpace session)
        {
            _session = session;
            Me = new MyIdentity(session);
            Contacts = new ContactIndex(session);
        }

        public CustomFieldValues Compute(ItemView v)
        {
            var values = new CustomFieldValues();

            // Sender and, for "on behalf of", the principal (PR_SENT_REPRESENTING_*).
            var senderAddress = Str(v.Props, SenderEmail);
            var senderSmtp = ToSmtp(Str(v.Props, SenderAddrType), senderAddress, Str(v.Props, SenderSmtp), v.Sender);
            var senderName = Str(v.Props, SenderName) ?? v.SenderName;

            var repAddress = Str(v.Props, RepEmail);
            var onBehalf = repAddress != null && !SameAddress(repAddress, senderAddress);
            var repSmtp = onBehalf ? ToSmtp(Str(v.Props, RepAddrType), repAddress, Str(v.Props, RepSmtp), null) : senderSmtp;
            if (onBehalf && SameAddress(repSmtp, senderSmtp))
                onBehalf = false;

            var fromSmtp = onBehalf ? repSmtp : senderSmtp;
            var fromName = onBehalf ? (Str(v.Props, RepName) ?? senderName) : senderName;
            var sentByMe = Me.IsMe(senderSmtp) || Me.IsMe(senderAddress) || Me.IsMe(repSmtp) || Me.IsMe(repAddress);

            // Recipients.
            var to = new List<Outlook.Recipient>();
            var cc = new List<Outlook.Recipient>();
            int total = v.Recipients.Count;
            for (int i = 1; i <= total; i++)
            {
                var r = v.Recipients[i];
                if (r.Type == (int)Outlook.OlMailRecipientType.olTo)
                    to.Add(r);
                else if (r.Type == (int)Outlook.OlMailRecipientType.olCC)
                    cc.Add(r);
            }
            values.Tos = Count(to.Count);
            values.Ccs = Count(cc.Count);
            values.Me = sentByMe ? MeSender : to.Any(IsMe) ? MeInTo : cc.Any(IsMe) ? MeInCc : CustomFieldValues.None;

            if (sentByMe)
            {
                var first = to.FirstOrDefault() ?? cc.FirstOrDefault();
                if (first != null)
                {
                    var smtp = RecipientSmtp(first);
                    values.DomainRelated = MailInfo.DomainPath(smtp) ?? CustomFieldValues.None;
                    values.NameRelated = SentPrefix + RecipientName(first, smtp) + (total > 1 ? MultipleSuffix : "");
                }
            }
            else
            {
                values.DomainRelated = MailInfo.DomainPath(fromSmtp) ?? CustomFieldValues.None;
                var contact = Contacts.Find(fromSmtp);
                values.NameRelated = contact != null ? "[" + contact + "]" : SenderDisplay(fromName, fromSmtp);
            }
            return values;
        }

        /// <summary>The sender's display name; "(local part)" when it is empty or just the address.</summary>
        private static string SenderDisplay(string name, string smtp)
        {
            name = (name ?? "").Trim();
            if (name.Length > 0 && !SameAddress(name, smtp))
                return name;
            if (string.IsNullOrEmpty(smtp))
                return name.Length > 0 ? name : CustomFieldValues.None;
            int at = smtp.IndexOf('@');
            return "(" + (at > 0 ? smtp.Substring(0, at) : smtp) + ")";
        }

        /// <summary>
        /// Name shown for a recipient of my mail: the resolved name when it was picked from an address book,
        /// otherwise (typed address, "Name &lt;address&gt;") [contact name] or the address itself.
        /// </summary>
        private string RecipientName(Outlook.Recipient r, string smtp)
        {
            if (!IsOneOff(r) && !string.IsNullOrWhiteSpace(r.Name))
                return r.Name.Trim();
            var contact = Contacts.Find(smtp);
            if (contact != null)
                return "[" + contact + "]";
            return smtp ?? r.Name ?? r.Address ?? CustomFieldValues.None;
        }

        private static bool IsOneOff(Outlook.Recipient r)
        {
            try
            {
                var entry = r.AddressEntry;
                if (entry == null)
                    return true;
                var type = entry.AddressEntryUserType;
                return type == Outlook.OlAddressEntryUserType.olSmtpAddressEntry
                    || type == Outlook.OlAddressEntryUserType.olOtherAddressEntry;
            }
            catch (COMException)
            {
                return true;
            }
        }

        private bool IsMe(Outlook.Recipient r)
        {
            var address = r.Address;
            if (Me.IsMe(address))
                return true;
            return address != null && address.IndexOf('@') < 0 && Me.IsMe(RecipientSmtp(r));
        }

        private string RecipientSmtp(Outlook.Recipient r)
        {
            var address = r.Address;
            if (address != null && address.IndexOf('@') >= 0)
                return address;
            var smtp = Str(r.PropertyAccessor, RecipientSmtpTag);
            if (smtp != null)
                return smtp;
            return ToSmtp("EX", address, null, () => r.AddressEntry);
        }

        /// <summary>SMTP address for a sender/principal/recipient; Exchange DNs are resolved (and cached).</summary>
        private string ToSmtp(string addrType, string address, string smtpProp, Func<Outlook.AddressEntry> entry)
        {
            if (!string.IsNullOrEmpty(smtpProp) && smtpProp.IndexOf('@') >= 0)
                return smtpProp;
            if (string.IsNullOrEmpty(address))
                return null;
            if (address.IndexOf('@') >= 0 && !string.Equals(addrType, "EX", StringComparison.OrdinalIgnoreCase))
                return address;

            string smtp;
            if (_exToSmtp.TryGetValue(address, out smtp))
                return smtp;

            smtp = null;
            try
            {
                var ae = entry == null ? null : entry();
                if (ae == null)
                {
                    var recipient = _session.CreateRecipient(address);
                    if (recipient.Resolve())
                        ae = recipient.AddressEntry;
                }
                if (ae != null)
                {
                    var user = ae.GetExchangeUser();
                    if (user != null)
                        smtp = user.PrimarySmtpAddress;
                    else
                    {
                        var list = ae.GetExchangeDistributionList();
                        if (list != null)
                            smtp = list.PrimarySmtpAddress;
                    }
                }
            }
            catch (COMException)
            {
                // offline / not in the GAL
            }
            if (string.IsNullOrEmpty(smtp))
                smtp = address;
            _exToSmtp[address] = smtp;
            return smtp;
        }

        private static bool SameAddress(string a, string b)
        {
            return string.Equals((a ?? "").Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static string Count(int n)
        {
            return n == 0 ? CustomFieldValues.None : n.ToString();
        }

        private static string Str(Outlook.PropertyAccessor pa, string schema)
        {
            try
            {
                var s = pa.GetProperty(schema) as string;
                return string.IsNullOrWhiteSpace(s) ? null : s;
            }
            catch (COMException)
            {
                return null;
            }
        }
    }
}
