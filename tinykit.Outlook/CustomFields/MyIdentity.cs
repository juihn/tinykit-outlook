using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.CustomFields
{
    /// <summary>All addresses that mean "me": every account's SMTP address and Exchange DN.</summary>
    internal sealed class MyIdentity
    {
        private readonly HashSet<string> _addresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public MyIdentity(Outlook.NameSpace session)
        {
            Add(() => session.CurrentUser.Address);
            foreach (Outlook.Account account in session.Accounts)
            {
                var a = account;
                Add(() => a.SmtpAddress);
                Add(() => a.CurrentUser.Address);
                Add(() =>
                {
                    var user = a.CurrentUser.AddressEntry.GetExchangeUser();
                    return user == null ? null : user.PrimarySmtpAddress;
                });
            }
        }

        public bool IsMe(string address)
        {
            return !string.IsNullOrEmpty(address) && _addresses.Contains(address.Trim());
        }

        public override string ToString()
        {
            return string.Join(", ", _addresses);
        }

        private void Add(Func<string> get)
        {
            try
            {
                var s = get();
                if (!string.IsNullOrWhiteSpace(s))
                    _addresses.Add(s.Trim());
            }
            catch (COMException)
            {
                // account without that information (e.g. offline Exchange)
            }
        }
    }
}
