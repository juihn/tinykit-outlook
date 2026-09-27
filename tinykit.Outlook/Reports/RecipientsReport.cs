using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using tinykit.OutlookAddin.CustomFields;
using tinykit.OutlookAddin.Filtering;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Reports
{
    /// <summary>A sender or recipient in the report.</summary>
    internal sealed class ReportPerson
    {
        /// <summary>Main text: the contact's name, else the recipient's display name, else the address.</summary>
        public string Description;

        /// <summary>The display name on the mail when it differs from <see cref="Description"/> and the address.</summary>
        public string Name;

        public string Smtp;
        public string UserName;
        public bool IsTo;
        public string Type = "";
        public ContactEntry Contact;
    }

    internal sealed class ReportDepartment
    {
        /// <summary>The contacts' department; null for recipients without one.</summary>
        public string Name;
        public List<ReportPerson> People = new List<ReportPerson>();
    }

    internal sealed class ReportDomain
    {
        public string Name;
        public List<ReportDepartment> Departments = new List<ReportDepartment>();

        public int Count
        {
            get { return Departments.Sum(d => d.People.Count); }
        }
    }

    /// <summary>
    /// The recipients of a mail grouped by domain, then by the contacts' department: who (else) got it,
    /// with To before Cc and contacts marked.
    /// </summary>
    internal sealed class RecipientsReport
    {
        public object Item;
        public string Subject;
        public ReportPerson Sender;
        public List<ReportDomain> Domains = new List<ReportDomain>();
        public int RecipientCount;

        /// <summary>True when the sender is also among the recipients (left out of the list).</summary>
        public bool IncludesSender;

        public static RecipientsReport Build(object item, CustomFieldCalculator calc)
        {
            var view = ItemView.From(item);
            if (view == null)
                return null;
            var report = new RecipientsReport { Item = item, Subject = ((dynamic)item).Subject as string ?? "" };

            var senderSmtp = calc.FromSmtpOf(view);
            var senderContact = calc.Contacts.Lookup(senderSmtp);
            var senderName = TrimQuotes(view.SenderName);
            report.Sender = Person(senderSmtp, senderName, senderContact);

            var domains = new Dictionary<string, ReportDomain>(StringComparer.OrdinalIgnoreCase);
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var recipients = view.Recipients;
            for (int i = 1; i <= recipients.Count; i++)
            {
                Outlook.Recipient r = recipients[i];
                string smtp;
                try
                {
                    smtp = calc.RecipientSmtp(r);
                }
                catch (COMException)
                {
                    continue;
                }
                if (string.IsNullOrWhiteSpace(smtp) || smtp.IndexOf('@') <= 0)
                    continue;
                smtp = smtp.Trim();
                if (string.Equals(smtp, senderSmtp, StringComparison.OrdinalIgnoreCase))
                {
                    report.IncludesSender = true;
                    continue;
                }
                if (!seen.Add(smtp))
                    continue;

                var person = Person(smtp, TrimQuotes(r.Name), calc.Contacts.Lookup(smtp));
                person.IsTo = r.Type == (int)Outlook.OlMailRecipientType.olTo;
                person.Type = r.Type == (int)Outlook.OlMailRecipientType.olCC ? "Cc"
                            : r.Type == (int)Outlook.OlMailRecipientType.olBCC ? "Bcc" : "To";

                var domainName = smtp.Substring(smtp.LastIndexOf('@') + 1).ToLowerInvariant();
                ReportDomain domain;
                if (!domains.TryGetValue(domainName, out domain))
                    domains[domainName] = domain = new ReportDomain { Name = domainName };
                var departmentName = person.Contact == null ? null : person.Contact.Department;
                var department = domain.Departments.FirstOrDefault(d => d.Name == departmentName);
                if (department == null)
                    domain.Departments.Add(department = new ReportDepartment { Name = departmentName });
                department.People.Add(person);
                report.RecipientCount++;
            }

            // Domains by organization, then subdomain (fabrikam, fabrikam/billing, ...); recipients without a department
            // first; To before Cc; then by user name.
            report.Domains = domains.Values.OrderBy(d => MailInfo.DomainPath("x@" + d.Name) ?? d.Name, StringComparer.Ordinal).ToList();
            foreach (var domain in report.Domains)
            {
                domain.Departments = domain.Departments.OrderBy(d => d.Name == null ? 0 : 1)
                    .ThenBy(d => d.Name, StringComparer.CurrentCultureIgnoreCase).ToList();
                foreach (var department in domain.Departments)
                    department.People = department.People.OrderByDescending(p => p.IsTo)
                        .ThenBy(p => p.UserName, StringComparer.OrdinalIgnoreCase).ToList();
            }
            return report;
        }

        private static ReportPerson Person(string smtp, string displayName, ContactEntry contact)
        {
            var description = contact != null ? contact.Name
                            : !string.IsNullOrEmpty(displayName) ? displayName
                            : smtp ?? "";
            int at = smtp == null ? -1 : smtp.LastIndexOf('@');
            return new ReportPerson
            {
                Description = description,
                Name = !string.IsNullOrEmpty(displayName)
                       && !string.Equals(displayName, description, StringComparison.OrdinalIgnoreCase)
                       && !string.Equals(displayName, smtp, StringComparison.OrdinalIgnoreCase) ? displayName : null,
                Smtp = smtp,
                UserName = at > 0 ? smtp.Substring(0, at) : smtp,
                Contact = contact,
            };
        }

        private static string TrimQuotes(string s)
        {
            return s == null ? null : s.Trim().Trim('\'', '"').Trim();
        }
    }
}
