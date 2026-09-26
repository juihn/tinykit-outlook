using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace TinyKit.OutlookAddin.CustomFields
{
    /// <summary>The members custom-field stamping needs, over MailItem and MeetingItem.</summary>
    internal sealed class ItemView
    {
        public object Item { get; private set; }
        public Outlook.Recipients Recipients { get; private set; }
        public Outlook.PropertyAccessor Props { get; private set; }
        public Outlook.UserProperties UserProperties { get; private set; }
        public string SenderName { get; private set; }

        /// <summary>The sender AddressEntry (MailItem only; null otherwise).</summary>
        public Func<Outlook.AddressEntry> Sender { get; private set; }

        public Action Save { get; private set; }

        /// <summary>Returns null for item types that have no sender/recipients to stamp.</summary>
        public static ItemView From(object item)
        {
            var mail = item as Outlook.MailItem;
            if (mail != null)
            {
                return new ItemView
                {
                    Item = mail,
                    Recipients = mail.Recipients,
                    Props = mail.PropertyAccessor,
                    UserProperties = mail.UserProperties,
                    SenderName = mail.SenderName,
                    Sender = () => mail.Sender,
                    Save = mail.Save,
                };
            }

            var meeting = item as Outlook.MeetingItem;
            if (meeting != null)
            {
                return new ItemView
                {
                    Item = meeting,
                    Recipients = meeting.Recipients,
                    Props = meeting.PropertyAccessor,
                    UserProperties = meeting.UserProperties,
                    SenderName = meeting.SenderName,
                    Sender = () => null,
                    Save = meeting.Save,
                };
            }
            return null;
        }
    }
}
