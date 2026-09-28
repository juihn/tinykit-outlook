using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Contacts
{
    /// <summary>The commands of the TinyKit tab in a contact's window; each works on the window's contact.</summary>
    internal static class ContactCommands
    {
        private const string AttachDataProp = "http://schemas.microsoft.com/mapi/proptag/0x37010102"; // PR_ATTACH_DATA_BIN
        private const string PictureFile = "ContactPicture.jpg";

        public static Outlook.ContactItem ContactOf(Outlook.Inspector inspector)
        {
            var contact = inspector == null ? null : inspector.CurrentItem as Outlook.ContactItem;
            if (contact == null)
                throw new UserMessageException("Open a contact first.");
            return contact;
        }

        /// <summary>Opens the contact's business address (or else home, other) in Google Maps.</summary>
        public static void OpenInGoogleMaps(Outlook.ContactItem contact)
        {
            var address = FirstNonEmpty(contact.BusinessAddress, contact.HomeAddress, contact.OtherAddress);
            if (address == null)
                throw new UserMessageException("This contact has no address.");
            address = string.Join(" ", address.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries));
            Process.Start("https://www.google.com/maps/search/" + Uri.EscapeDataString(address));
        }

        private static string FirstNonEmpty(params string[] values)
        {
            foreach (var v in values)
                if (!string.IsNullOrWhiteSpace(v))
                    return v;
            return null;
        }

        /// <summary>The contact's picture, cropped to a square; null when it has none.</summary>
        public static Bitmap PictureOf(Outlook.ContactItem contact)
        {
            if (contact == null || !contact.HasPicture)
                return null;
            foreach (Outlook.Attachment attachment in contact.Attachments)
            {
                if (!string.Equals(attachment.FileName, PictureFile, StringComparison.OrdinalIgnoreCase))
                    continue;
                var data = attachment.PropertyAccessor.GetProperty(AttachDataProp) as byte[];
                if (data == null)
                    return null;
                using (var stream = new MemoryStream(data))
                using (var image = new Bitmap(stream))
                {
                    int side = Math.Min(image.Width, image.Height);
                    return image.Clone(new Rectangle((image.Width - side) / 2, (image.Height - side) / 2, side, side), image.PixelFormat);
                }
            }
            return null;
        }

        /// <summary>
        /// Contact Picture: adds a picture when there is none, else changes it; Shift+click removes it. Outlook's own
        /// commands of the window do the work (they show the file dialog), then the contact is saved.
        /// </summary>
        public static void EditPicture(Outlook.Inspector inspector, Outlook.ContactItem contact, bool remove)
        {
            string command = !contact.HasPicture ? "ContactAddPicture" : remove ? "ContactRemovePicture" : "ContactChangePicture";
            inspector.CommandBars.ExecuteMso(command);
            contact.Save();
        }

        /// <summary>Sets the contact's message class and saves it; it opens with that form from the next time.</summary>
        public static string SetForm(Outlook.ContactItem contact, string messageClass)
        {
            if (string.Equals(contact.MessageClass, messageClass, StringComparison.OrdinalIgnoreCase))
                return "This contact already uses the " + FilterController.FormName(messageClass) + " form.";
            contact.MessageClass = messageClass;
            contact.Save();
            return "This contact now uses the " + FilterController.FormName(messageClass) + " form. "
                + "Close it and open it again to see it (restart Outlook if it still shows the old form).";
        }

        /// <summary>The custom contact form of the contact's folder, or else of the default Contacts folder.</summary>
        public static string CustomForm(Outlook.ContactItem contact)
        {
            Outlook.MAPIFolder folder = null;
            try
            {
                folder = contact == null ? null : contact.Parent as Outlook.MAPIFolder;
            }
            catch (COMException)
            {
            }
            return FilterController.CustomFormOf(folder)
                ?? FilterController.CustomFormOf(Globals.ThisAddIn.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts));
        }

        /// <summary>
        /// Copy to Clipboard: <c>company / department / name (job title) e-mail T.phone M.mobile</c>, tab-separated;
        /// Shift+click keeps the clipboard's text after it.
        /// </summary>
        public static void CopyToClipboard(Outlook.ContactItem contact, bool append)
        {
            var parts = new List<string>();
            if (!string.IsNullOrEmpty(contact.CompanyName))
                parts.Add(contact.CompanyName + " /");
            if (!string.IsNullOrEmpty(contact.Department))
                parts.Add(contact.Department + " /");
            var name = contact.Subject;
            if (!string.IsNullOrEmpty(contact.JobTitle))
                name += " (" + contact.JobTitle + ")";
            parts.Add(name);
            parts.Add(contact.Email1Address ?? "");
            if (!string.IsNullOrEmpty(contact.BusinessTelephoneNumber))
                parts.Add("T." + contact.BusinessTelephoneNumber);
            if (!string.IsNullOrEmpty(contact.MobileTelephoneNumber))
                parts.Add("M." + contact.MobileTelephoneNumber);
            var text = string.Join("\t", parts).Trim() + Environment.NewLine;
            if (append && Clipboard.ContainsText())
                text += Clipboard.GetText() + Environment.NewLine;
            Clipboard.SetText(text);
        }
    }
}
