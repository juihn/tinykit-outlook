using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.CustomFields;
using tinykit.OutlookAddin.Filtering;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Compose
{
    /// <summary>The Recipients group of the TinyKit tab in a mail being written.</summary>
    internal sealed class RecipientCommands
    {
        private readonly CustomFieldCalculator _calc;

        public RecipientCommands(CustomFieldCalculator calc)
        {
            _calc = calc;
        }

        /// <summary>The address the mail is sent from (its account, else the first account); null if unknown.</summary>
        public static string SenderSmtp(Outlook.MailItem mail)
        {
            try
            {
                var account = mail.SendUsingAccount;
                if (account == null)
                {
                    var accounts = mail.Session.Accounts;
                    account = accounts.Count > 0 ? accounts[1] : null;
                }
                return account == null ? null : account.SmtpAddress;
            }
            catch (COMException)
            {
                return null;
            }
        }

        /// <summary>Removes the sending account's own address from To, Cc and Bcc (e.g. after Reply All). Returns how many.</summary>
        public int RemoveSender(Outlook.MailItem mail)
        {
            var sender = SenderSmtp(mail);
            if (string.IsNullOrEmpty(sender))
                throw new UserMessageException("The mail has no sending account.");
            int removed = 0;
            var recipients = mail.Recipients;
            for (int i = recipients.Count; i >= 1; i--)
            {
                if (SameAddress(_calc.RecipientSmtp(recipients[i]), sender))
                {
                    recipients.Remove(i);
                    removed++;
                }
            }
            return removed;
        }

        /// <summary>
        /// Restate Recipients: each recipient found in Contacts (by address, else by the contact's e-mail display name)
        /// is replaced by that contact entry, so it shows with the name set in Contacts; the type (To/Cc/Bcc) is kept.
        /// The replaced ones are added back after the others, those in the sender's own domain last.
        /// Returns how many were replaced.
        /// </summary>
        public int Restate(Outlook.MailItem mail)
        {
            var senderDomain = MailInfo.BaseDomain(SenderSmtp(mail) ?? "");
            var recipients = mail.Recipients;
            var external = new List<Replacement>();
            var internalOnes = new List<Replacement>();
            var remove = new List<int>();

            for (int i = 1; i <= recipients.Count; i++)
            {
                var r = recipients[i];
                var smtp = _calc.RecipientSmtp(r);
                string address = null;
                if (!string.IsNullOrEmpty(smtp) && _calc.Contacts.Lookup(smtp) != null)
                    address = smtp;
                else
                    address = _calc.Contacts.AddressNamed(Unquote(r.Name));
                if (address == null)
                    continue;

                var replacement = new Replacement
                {
                    Name = _calc.Contacts.Find(address),
                    Address = address,
                    Type = r.Type,
                };
                var domain = MailInfo.BaseDomain(address);
                if (!string.IsNullOrEmpty(senderDomain) && string.Equals(domain, senderDomain, StringComparison.OrdinalIgnoreCase))
                    internalOnes.Add(replacement);
                else
                    external.Add(replacement);
                remove.Add(i);
            }

            for (int k = remove.Count - 1; k >= 0; k--)
                recipients.Remove(remove[k]);
            foreach (var x in external)
                Add(recipients, x);
            foreach (var x in internalOnes)
                Add(recipients, x);
            return remove.Count;
        }

        private sealed class Replacement
        {
            public string Name;
            public string Address;
            public int Type;
        }

        // By the contact's display name, which resolves to the contact entry; if that name does not resolve to the same
        // address (ambiguous, or another entry), by the address instead.
        private void Add(Outlook.Recipients recipients, Replacement x)
        {
            var recipient = recipients.Add(x.Name ?? x.Address);
            recipient.Type = x.Type;
            bool ok;
            try
            {
                ok = recipient.Resolve() && SameAddress(_calc.RecipientSmtp(recipient), x.Address);
            }
            catch (COMException)
            {
                ok = false;
            }
            if (ok)
                return;
            recipients.Remove(recipient.Index);
            recipient = recipients.Add(x.Address);
            recipient.Type = x.Type;
            try
            {
                recipient.Resolve();
            }
            catch (COMException)
            {
                // left for Check Names
            }
        }

        private static string Unquote(string name)
        {
            name = (name ?? "").Trim();
            if (name.Length >= 2 && ((name[0] == '\'' && name[name.Length - 1] == '\'') || (name[0] == '"' && name[name.Length - 1] == '"')))
                name = name.Substring(1, name.Length - 2);
            return name;
        }

        private static bool SameAddress(string a, string b)
        {
            return !string.IsNullOrEmpty(a) && string.Equals(a.Trim(), (b ?? "").Trim(), StringComparison.OrdinalIgnoreCase);
        }
    }
}
