using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Windows.Forms;
using tinykit.OutlookAddin.Common;
using tinykit.OutlookAddin.Filtering;
using tinykit.OutlookAddin.Settings;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace tinykit.OutlookAddin.Ribbon
{
    /// <summary>
    /// Ribbon XML for the Outlook explorer: a "TinyKit" tab with Custom Mail Fields, Quick Filter, Clear, Saved Filters, Items
    /// (a toggle per saved filter, formats and the file/format management buttons) and View groups.
    /// </summary>
    [ComVisible(true)]
    public class OutlookRibbon : Office.IRibbonExtensibility
    {
        private const string ExplorerRibbonId = "Microsoft.Outlook.Explorer";
        private const string SlotSeparator = "_";

        // Quick filter slots: two rows of two; which field a slot filters by depends on the folder
        // (FilterController.QuickKindsFor: mail F S / N D, contacts F E / C D).
        private const int QuickSlots = 4;

        /// <summary>
        /// The one-letter labels padded with narrow Unicode spaces to the widest one in the ribbon font, so the four
        /// buttons come out the same width. A zero-width non-joiner at the end keeps Office from trimming the padding.
        /// </summary>
        private static readonly Dictionary<QuickKind, string> PaddedLabels = QuickLabels();

        private static Dictionary<QuickKind, string> QuickLabels()
        {
            var kinds = (QuickKind[])Enum.GetValues(typeof(QuickKind));
            char[] pads = { '\u2007', '\u2009', '\u200A' }; // figure, thin, hair space: widest first
            using (var font = new Font("Segoe UI", 9f))
            {
                Func<string, int> width = t => TextRenderer.MeasureText(t, font, Size.Empty, TextFormatFlags.NoPadding).Width;
                int bare = width("xx");
                var padWidth = pads.Select(c => Math.Max(1, width("x" + c + "x") - bare)).ToArray();
                int target = kinds.Max(k => width(QuickLabel(k)));
                return kinds.ToDictionary(k => k, k =>
                {
                    var text = new StringBuilder(QuickLabel(k));
                    int missing = target - width(QuickLabel(k));
                    for (int p = 0; p < pads.Length; p++)
                    {
                        while (missing >= padWidth[p])
                        {
                            text.Append(pads[p]);
                            missing -= padWidth[p];
                        }
                    }
                    return text.Append('\u200C').ToString();
                });
            }
        }

        /// <summary>One-letter button label: F(rom), N(ame), S(ubject), D(omain); F(ile As), E(mail), C(ompany), D(epartment).</summary>
        private static string QuickLabel(QuickKind kind)
        {
            return kind.ToString().Substring(0, 1);
        }

        private static string QuickTip(QuickKind kind)
        {
            switch (kind)
            {
                case QuickKind.From:
                    return "Uses the From e-mail address: shows mail whose From address contains the value.";
                case QuickKind.Name:
                    return "Uses the nameRelated column: shows mail whose nameRelated contains the value "
                         + "(the sender's contact or display name; the first recipient for mail you sent).";
                case QuickKind.Subject:
                    return "Uses the subject: shows mail whose subject contains the value (RE:/FW: ignored when taken from a mail).";
                case QuickKind.Domain:
                    return "Uses the domainRelated column: shows mail whose domainRelated contains the value "
                         + "(the sender's domain label, e.g. fabrikam/billing; the first recipient's for mail you sent).";
                case QuickKind.FileAs:
                    return "Uses File As: shows contacts whose File As contains the value.";
                case QuickKind.Email:
                    return "Uses the e-mail addresses: shows contacts whose E-mail, E-mail 2 or E-mail 3 contains the value.";
                case QuickKind.Company:
                    return "Uses Company: shows contacts whose company contains the value.";
                case QuickKind.Department:
                    return "Uses Department: shows contacts whose department contains the value.";
            }
            return "";
        }

        private static string QuickName(QuickKind kind)
        {
            return kind == QuickKind.FileAs ? "File As" : kind.ToString();
        }

        // Outlook's own commands shown in the Built-in group: one array per column, null = separator line.
        // Only one of MarkAsRead / MarkAsUnread is visible at a time (Outlook shows the one that applies).
        private static readonly string[][] BuiltInColumns =
        {
            new[] { "CategorizeMenu", "AllCategories", "MoveToOneNote" },
            new[] { "AddressBook", "ShowRemindersWindow" },
            null,
            new[] { "FlagNoDate", "ClearFlag", "MoveToFolderGallery" },
            new[] { "MarkAsRead", "MarkAsUnread", "ShowInConversations", "FindRelatedMessages" },
        };

        // Built-in commands that have no icon of their own (they live in menus), with the image to show instead.
        private static readonly Dictionary<string, string> BuiltInImages = new Dictionary<string, string>
        {
            { "ClearFlag", "Delete" },
            { "FindRelatedMessages", "GroupConversations" },
        };

        // Built-in check boxes: showLabel does not apply to them, so their label is replaced by a blank one.
        private static readonly HashSet<string> BuiltInCheckBoxes = new HashSet<string> { "ShowInConversations" };

        private readonly FilterController _controller;
        private Office.IRibbonUI _ui;

        internal OutlookRibbon(FilterController controller)
        {
            _controller = controller;
            _controller.Invalidate = () => { if (_ui != null) _ui.Invalidate(); };
        }

        public string GetCustomUI(string ribbonID)
        {
            return ribbonID == ExplorerRibbonId ? BuildXml() : null;
        }

        private static string BuildXml()
        {
            var sb = new StringBuilder();
            sb.Append("<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\" onLoad=\"OnLoad\">");
            sb.Append("<ribbon><tabs><tab id=\"tabOAFilter\" label=\"TinyKit\" insertBeforeMso=\"TabMail\">");

            // Built-in: frequently used Outlook commands, icon only, in columns.
            sb.Append("<group id=\"grpBuiltIn\" label=\"Built-in\">");
            int separators = 0;
            foreach (var column in BuiltInColumns)
            {
                if (column == null)
                {
                    sb.Append("<separator id=\"biSep").Append(separators++).Append("\"/>");
                    continue;
                }
                sb.Append("<box id=\"bi").Append(column[0]).Append("\" boxStyle=\"vertical\">");
                foreach (var idMso in column)
                {
                    string image;
                    if (BuiltInCheckBoxes.Contains(idMso))
                        sb.Append("<checkBox idMso=\"").Append(idMso).Append("\" label=\" \"/>");
                    else if (BuiltInImages.TryGetValue(idMso, out image))
                        sb.Append("<button idMso=\"").Append(idMso).Append("\" imageMso=\"").Append(image).Append("\" showLabel=\"false\"/>");
                    else
                        sb.Append("<control idMso=\"").Append(idMso).Append("\" showLabel=\"false\"/>");
                }
                sb.Append("</box>");
            }
            // Contact folders: switch the selected contacts between Outlook's form and the folder's custom form.
            sb.Append("<box id=\"biForms\" boxStyle=\"vertical\">");
            sb.Append("<button id=\"cfFormDefault\" label=\"Default Contact Form\" showLabel=\"false\" imageMso=\"NewContact\"")
              .Append(" getVisible=\"GetContactVisible\" onAction=\"OnSetDefaultContactForm\" screentip=\"Default Contact Form\"")
              .Append(" supertip=\"Set the selected contacts to open with Outlook's own contact form (message class IPM.Contact).\"/>");
            sb.Append("<button id=\"cfFormCustom\" getLabel=\"GetCustomFormLabel\" showLabel=\"false\" imageMso=\"ChooseForm\"")
              .Append(" getVisible=\"GetCustomFormVisible\" onAction=\"OnSetCustomContactForm\" getScreentip=\"GetCustomFormLabel\"")
              .Append(" getSupertip=\"GetCustomFormSupertip\"/>");
            sb.Append("</box>");
            sb.Append("</group>");

            // Custom mail fields (domainRelated, nameRelated, me, tos, ccs, unknownDomain): mail folders only.
            sb.Append("<group id=\"grpFields\" label=\"Custom Mail Fields\" getVisible=\"GetMailVisible\">");
            sb.Append("<button id=\"cfFill\" label=\"Fill Fields\" imageMso=\"PropertySheet\" onAction=\"OnFillFields\"")
              .Append(" screentip=\"Fill Fields\" supertip=\"Fill domainRelated, nameRelated, me, tos, ccs and unknownDomain. ")
              .Append("Click: the items of the current folder that do not have them yet (e.g. mail received while Outlook was closed, ")
              .Append("or in folders other than Inbox and Sent Items). Shift+click: recompute the selected items, even if they already have ")
              .Append("values (e.g. after editing Known Domains.txt or adding a contact).\"/>");
            sb.Append("<button id=\"cfAddKnown\" label=\"Add Known Domain\" imageMso=\"AddToFavorites\" onAction=\"OnAddKnownDomain\"")
              .Append(" screentip=\"Add Known Domain\" getSupertip=\"GetAddKnownSupertip\"/>");
            sb.Append("<checkBox id=\"cfAutoFill\" label=\"Auto-fill new mail\" getPressed=\"GetAutoFillPressed\" onAction=\"OnAutoFillToggle\"")
              .Append(" screentip=\"Auto-fill new mail\" supertip=\"Fill the fields of mail arriving in each account's Inbox and Sent Items.\"/>");
            sb.Append("</group>");

            // Quick Filter (mail and contact folders): one input box, then four buttons, each followed by its own
            // history drop-down; what they filter by depends on the folder.
            sb.Append("<group id=\"grpQuick\" label=\"Quick Filter\" getVisible=\"GetQuickVisible\">");
            sb.Append("<editBox id=\"qInput\" label=\"Value\" showLabel=\"false\" sizeString=\"WWWWWWWWWWm\"")
              .Append(" getText=\"GetInputText\" onChange=\"OnInputChange\" screentip=\"Quick filter value\"")
              .Append(" supertip=\"Type a value, then press a button below to show the items whose field contains it. ")
              .Append("Leave it empty to use the value of the first selected mail or contact.\"/>");
            for (int row = 0; row < QuickSlots / 2; row++)
            {
                sb.Append("<box id=\"qRow").Append(row).Append("\" boxStyle=\"horizontal\">");
                AppendQuick(sb, row * 2);
                AppendQuick(sb, row * 2 + 1);
                sb.Append("</box>");
            }
            sb.Append("</group>");

            // Clear.
            sb.Append("<group id=\"grpClear\" label=\"Clear\">");
            sb.Append("<toggleButton id=\"qClearOnExit\" label=\"Clear Inboxes on Exit\" size=\"large\" imageMso=\"FilterClearAllFilters\"")
              .Append(" getPressed=\"GetClearOnExitPressed\" onAction=\"OnClearOnExitToggle\"")
              .Append(" screentip=\"Clear Inboxes on exit\" supertip=\"When Outlook closes, clear the quick, saved and Others filters in the ")
              .Append("Inbox of every account (as Clear Filter does; the views' own filters stay), so Outlook opens with full Inboxes.\"/>");
            sb.Append("<button id=\"qClear\" label=\"Clear Filter\" size=\"large\" imageMso=\"FilterClearAllFilters\" onAction=\"OnClear\"")
              .Append(" screentip=\"Clear Filter\" supertip=\"Remove the quick or saved filter and restore the view's own filter.\"/>");
            sb.Append("</group>");

            // Saved filters: a toggle per saved filter (fixed slots shown/hidden by callbacks, so edits need no
            // restart), formats of the applied one, the Add to menu, then the file and format management.
            sb.Append("<group id=\"grpSaved\" label=\"Saved Filters\" getVisible=\"GetSavedGroupVisible\">");
            for (int i = 0; i < FilterController.MaxSavedFilters; i++)
            {
                sb.Append("<toggleButton id=\"sf").Append(SlotSeparator).Append(i).Append("\" tag=\"").Append(i)
                  .Append("\" getLabel=\"GetSavedLabel\" getVisible=\"GetSavedVisible\" getPressed=\"GetSavedPressed\" onAction=\"OnSavedToggle\"")
                  .Append(" getScreentip=\"GetSavedLabel\" getSupertip=\"GetSavedSupertip\"/>");
            }
            sb.Append("<toggleButton id=\"sfOthers\" label=\"Others\" getVisible=\"GetOthersVisible\" getPressed=\"GetOthersPressed\"")
              .Append(" onAction=\"OnOthersToggle\" screentip=\"Others\" supertip=\"Show only the items that none of the saved filters on the left ")
              .Append("match (NOT (filter 1 OR filter 2 ...)). Press again to restore the view's own filter.\"/>");
            sb.Append("<separator id=\"sepSaved1\"/>");
            sb.Append("<toggleButton id=\"sfFormat\" label=\"Format\" imageMso=\"ConditionalFormattingMenu\"")
              .Append(" getPressed=\"GetApplyFormatPressed\" onAction=\"OnFormatToggle\" screentip=\"Format\" getSupertip=\"GetFormatSupertip\"/>");
            sb.Append("<dynamicMenu id=\"sfAddTo\" label=\"Add to\" imageMso=\"").Append(AddIcon)
              .Append("\" getVisible=\"GetMailVisible\" getContent=\"GetAddToContent\" invalidateContentOnDrop=\"true\" screentip=\"Add to\"")
              .Append(" supertip=\"Add the selected mails to a saved filter (all except Flagged, Sent and Unknown). A small window asks ")
              .Append("whether by DOMAINRELATED, by SUBJECT, or by both together (domainRelated = ... AND subject); subjects become ")
              .Append("patterns in which numbers, dates and month/weekday names are % (e.g. Your trip with Gojek on %), which you can ")
              .Append("edit before adding.\"/>");
            sb.Append("<separator id=\"sepSaved2\"/>");
            sb.Append("<button id=\"mSettings\" label=\"Edit Saved Filters\" imageMso=\"").Append(XmlIcon)
              .Append("\" onAction=\"OnOpenSettings\" getScreentip=\"GetSettingsScreentip\" getSupertip=\"GetSettingsSupertip\"/>");
            sb.Append("<button id=\"mSaveView\" label=\"Save View as Filter...\" imageMso=\"FileSaveAs\" onAction=\"OnSaveViewFilter\"")
              .Append(" screentip=\"Save View as Filter\" supertip=\"Save the current view's filter (e.g. from View Settings &gt; Filter) as a saved filter.\"/>");

            // Manage: reloading the file and the conditional-format housekeeping, out of the way in one menu.
            sb.Append("<menu id=\"mManage\" label=\"Manage\" imageMso=\"").Append(ManageIcon).Append("\" screentip=\"Manage\"")
              .Append(" supertip=\"Reload the saved filters file and manage the saved filters' conditional formatting.\">");
            sb.Append("<button id=\"mReload\" label=\"Reload\" imageMso=\"Refresh\" onAction=\"OnReload\" getScreentip=\"GetReloadScreentip\"")
              .Append(" supertip=\"Read the saved filters file again now (changes saved in an editor are also picked up automatically).\"/>");
            sb.Append("<menuSeparator id=\"mManageSep\"/>");
            sb.Append("<button id=\"mApplyFormats\" label=\"Refresh Formats\" imageMso=\"Refresh\" onAction=\"OnApplyFormats\"")
              .Append(" screentip=\"Refresh Formats\" supertip=\"Turn the saved filters' formats on (after All Formats Off) and rewrite the formats of all saved filters whose Format is on into this view's conditional formatting.\"/>");
            sb.Append("<button id=\"mFormatsOff\" label=\"All Formats Off\" imageMso=\"").Append(FormatsOffIcon).Append("\" onAction=\"OnAllFormatsOff\"")
              .Append(" screentip=\"All Formats Off\" supertip=\"Take all saved filters' conditional formatting out of the views. Each filter's Format on/off setting is kept, so Refresh Formats turns the same formats back on.\"/>");
            sb.Append("<button id=\"mRemoveFormats\" label=\"Remove Formats\" imageMso=\"ClearFormatting\" onAction=\"OnRemoveFormats\"")
              .Append(" screentip=\"Remove Formats\" supertip=\"Delete this add-in's conditional formatting rules ([TK] ...) from the current view (asks first). Settings are unchanged.\"/>");
            sb.Append("<checkBox id=\"mAutoApply\" label=\"Auto-apply formats\" getPressed=\"GetAutoApplyPressed\" onAction=\"OnAutoApplyToggle\"")
              .Append(" screentip=\"Auto-apply formats\" supertip=\"Sync the formats into each table view when switching folders or views.\"/>");
            sb.Append("</menu>");
            sb.Append("</group>");

            // Items: tools for the selected items, in mail, contact and task folders alike.
            sb.Append("<group id=\"grpItems\" label=\"Items\">");
            sb.Append("<button id=\"cfRecipients\" label=\"Recipients Report\" imageMso=\"ContactProperties\" onAction=\"OnRecipientsReport\"")
              .Append(" screentip=\"Recipients Report\" supertip=\"Show the selected mail's sender and recipients grouped by domain and ")
              .Append("department (from Contacts). Blue: To, gray: Cc/Bcc; green: in Contacts. In the window, click a person to open the ")
              .Append("contact (or search LinkedIn), Ctrl+click for a new contact, Shift+click to add the address to the clipboard.\"/>");
            sb.Append("<button id=\"cfCopyItems\" label=\"Copy Items\" imageMso=\"GroupClipboard\" onAction=\"OnCopyItems\"")
              .Append(" screentip=\"Copy Items\" supertip=\"Copy one line per selected mail to the clipboard: ")
              .Append("'yy.MM.dd요일 HH:mm &lt;sender&gt; subject. Shift+click: put the new lines before the clipboard's current text.\"/>");
            sb.Append("</group>");

            // View.
            sb.Append("<group id=\"grpView\" label=\"View\">");
            sb.Append("<button id=\"mViewColumns\" label=\"View Columns\" imageMso=\"TableInsert\" onAction=\"OnViewColumns\"")
              .Append(" screentip=\"View Columns\" getSupertip=\"GetViewColumnsSupertip\"/>");
            sb.Append("<button id=\"mViewFont\" label=\"View Font...\" imageMso=\"FontDialog\" onAction=\"OnViewFont\"")
              .Append(" screentip=\"View Font\" supertip=\"Choose the font and size of the whole table view: rows, column headers and conditional formatting.\"/>");
            sb.Append("</group>");

            sb.Append("</tab></tabs></ribbon></customUI>");
            return sb.ToString();
        }

        private const string AddIcon = "OutlineExpand";
        private const string XmlIcon = "EditItem";
        private const string FormatsOffIcon = "ConditionalFormattingClearMenu";
        private const string ManageIcon = "AlignJustify";

        /// <summary>
        /// A filter button followed by an arrow-only gallery of that field's recent values (a gallery placed directly
        /// in the group opens in one click, and without item images it has no icon column). Labels and tips come from
        /// callbacks, because the field of a slot depends on the folder.
        /// </summary>
        private static void AppendQuick(StringBuilder sb, int slot)
        {
            sb.Append("<button id=\"q").Append(slot).Append("\" tag=\"").Append(slot)
              .Append("\" getLabel=\"GetQuickLabel\" imageMso=\"Filter\" onAction=\"OnQuick\"")
              .Append(" getScreentip=\"GetQuickScreentip\" getSupertip=\"GetQuickSupertip\"/>");
            sb.Append("<gallery id=\"qh").Append(slot).Append("\" tag=\"").Append(slot).Append("\" label=\"Recent values\"")
              .Append(" showLabel=\"false\" columns=\"1\" showItemImage=\"false\" invalidateContentOnDrop=\"true\"")
              .Append(" getItemCount=\"GetItemCount\" getItemLabel=\"GetItemLabel\" getItemID=\"GetItemID\" onAction=\"OnHistoryPick\"")
              .Append(" getScreentip=\"GetHistoryScreentip\" supertip=\"Click a value to filter by it (it also goes into the box). Ctrl+click removes it from this list.\"/>");
        }

        // ---------- Callbacks ----------

        public void OnLoad(Office.IRibbonUI ribbonUI)
        {
            _ui = ribbonUI;
        }

        public bool GetMailVisible(Office.IRibbonControl control)
        {
            return Safe(() => Focus(control) == ItemKind.Mail, true);
        }

        public bool GetContactVisible(Office.IRibbonControl control)
        {
            return Safe(() => Focus(control) == ItemKind.Contact, false);
        }

        public bool GetCustomFormVisible(Office.IRibbonControl control)
        {
            return Safe(() => Focus(control) == ItemKind.Contact && CustomForm(control) != null, false);
        }

        public string GetCustomFormLabel(Office.IRibbonControl control)
        {
            return Safe(() => FilterController.FormName(CustomForm(control) ?? "IPM.Contact"), "Custom Contact Form");
        }

        public string GetCustomFormSupertip(Office.IRibbonControl control)
        {
            return Safe(() => "Set the selected contacts to open with the Contacts folder's form (message class "
                + (CustomForm(control) ?? "?") + ").", "");
        }

        private string CustomForm(Office.IRibbonControl control)
        {
            var explorer = control.Context as Outlook.Explorer ?? Globals.ThisAddIn.Application.ActiveExplorer();
            return explorer == null ? null : _controller.CustomContactForm(explorer);
        }

        public void OnSetDefaultContactForm(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.SetContactForm(ex, FilterController.DefaultContactForm));
        }

        public void OnSetCustomContactForm(Office.IRibbonControl control)
        {
            Run(control, ex =>
            {
                var form = _controller.CustomContactForm(ex);
                if (form == null)
                    throw new UserMessageException("The Contacts folder has no custom form (its default form is Outlook's own).");
                _controller.SetContactForm(ex, form);
            });
        }

        public bool GetQuickVisible(Office.IRibbonControl control)
        {
            return Safe(() => FilterController.QuickKindsFor(Focus(control)).Length > 0, true);
        }

        public bool GetSavedGroupVisible(Office.IRibbonControl control)
        {
            return Safe(() => Focus(control) != null, true);
        }

        public string GetQuickLabel(Office.IRibbonControl control)
        {
            return Safe(() => { var k = QuickKindOf(control); return k == null ? "" : PaddedLabels[k.Value]; }, "");
        }

        public string GetQuickScreentip(Office.IRibbonControl control)
        {
            return Safe(() => { var k = QuickKindOf(control); return k == null ? "" : "Filter by " + QuickName(k.Value); }, "");
        }

        public string GetQuickSupertip(Office.IRibbonControl control)
        {
            return Safe(() =>
            {
                var k = QuickKindOf(control);
                return k == null ? "" : QuickTip(k.Value) + " Uses the text in the box; with the box empty, the value of the first selected "
                    + (_controller.Kind == ItemKind.Contact ? "contact" : "mail") + ". The arrow next to it lists recent "
                    + QuickName(k.Value) + " values.";
            }, "");
        }

        public string GetHistoryScreentip(Office.IRibbonControl control)
        {
            return Safe(() => { var k = QuickKindOf(control); return k == null ? "" : "Recent " + QuickName(k.Value) + " values"; }, "");
        }

        public void OnQuick(Office.IRibbonControl control)
        {
            Run(control, ex =>
            {
                var k = QuickKindOf(control);
                if (k != null)
                    _controller.QuickButton(ex, k.Value);
            });
        }

        public string GetInputText(Office.IRibbonControl control)
        {
            return _controller.InputText;
        }

        public void OnInputChange(Office.IRibbonControl control, string text)
        {
            _controller.SetInputText(text);
        }

        public int GetItemCount(Office.IRibbonControl control)
        {
            return Safe(() => { var k = QuickKindOf(control); return k == null ? 0 : _controller.History.Get(k.Value).Count; }, 0);
        }

        public string GetItemLabel(Office.IRibbonControl control, int index)
        {
            return Safe(() =>
            {
                var k = QuickKindOf(control);
                if (k == null)
                    return "";
                var list = _controller.History.Get(k.Value);
                return index < list.Count ? list[index].Label : "";
            }, "");
        }

        public string GetItemID(Office.IRibbonControl control, int index)
        {
            return control.Id + SlotSeparator + index;
        }

        public void OnHistoryPick(Office.IRibbonControl control, string selectedId, int selectedIndex)
        {
            Run(control, ex =>
            {
                var k = QuickKindOf(control);
                if (k != null)
                    _controller.QuickFromHistory(ex, k.Value, selectedIndex);
            });
        }

        public void OnClear(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.Clear(ex));
        }

        public string GetSavedLabel(Office.IRibbonControl control)
        {
            return Safe(() => { Focus(control); var f = _controller.FilterAt(SlotOf(control)); return f == null ? "" : f.Name; }, "");
        }

        public bool GetSavedVisible(Office.IRibbonControl control)
        {
            return Safe(() => { Focus(control); return _controller.FilterAt(SlotOf(control)) != null; }, false);
        }

        public string GetSavedSupertip(Office.IRibbonControl control)
        {
            return Safe(() => { Focus(control); var f = _controller.FilterAt(SlotOf(control)); return f == null ? "" : FilterController.Truncate(f.Sql, 1000); }, "");
        }

        public bool GetSavedPressed(Office.IRibbonControl control)
        {
            var ex = control.Context as Outlook.Explorer;
            return ex != null && Safe(() => { Focus(control); return _controller.IsSavedActive(ex, SlotOf(control)); }, false);
        }

        public bool GetOthersVisible(Office.IRibbonControl control)
        {
            return Safe(() => Focus(control) != null && _controller.IsOthersVisible, false);
        }

        public bool GetOthersPressed(Office.IRibbonControl control)
        {
            var ex = control.Context as Outlook.Explorer;
            return ex != null && Safe(() => { Focus(control); return _controller.IsOthersActive(ex); }, false);
        }

        public void OnOthersToggle(Office.IRibbonControl control, bool pressed)
        {
            Run(control, ex => _controller.ToggleOthers(ex, pressed));
        }

        public void OnSavedToggle(Office.IRibbonControl control, bool pressed)
        {
            Run(control, ex => _controller.ToggleSaved(ex, SlotOf(control), pressed));
        }

        public string GetFormatSupertip(Office.IRibbonControl control)
        {
            var ex = control.Context as Outlook.Explorer;
            var f = ex == null ? null : Safe(() => { Focus(control); return _controller.ActiveSavedFilter(ex); }, null);
            const string how = " Click: turn the format on/off in table views. Ctrl+click: edit it (style, strikeout, underline, color).";
            if (f == null)
                return "Conditional formatting of the applied saved filter. Apply a saved filter first." + how;
            return f.Format == null
                ? "\"" + f.Name + "\" has no format yet: clicking opens the Format dialog and turns the new format on."
                : "\"" + f.Name + "\": " + f.Format.Describe() + "." + how;
        }


        public bool GetApplyFormatPressed(Office.IRibbonControl control)
        {
            var ex = control.Context as Outlook.Explorer;
            return ex != null && Safe(() => { Focus(control); return _controller.IsActiveFormatEnabled(ex); }, false);
        }

        public void OnFormatToggle(Office.IRibbonControl control, bool pressed)
        {
            Run(control, ex => _controller.FormatButton(ex, pressed));
        }

        /// <summary>
        /// Items of the Add to menu: one per saved filter it lists (see FilterController.AddToTargets), built each time
        /// the menu opens, so filters added to the file later show up too. The tag carries the filter name.
        /// </summary>
        public string GetAddToContent(Office.IRibbonControl control)
        {
            return Safe(() =>
            {
                Focus(control);
                var sb = new StringBuilder("<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">");
                var targets = _controller.AddToTargets;
                for (int i = 0; i < targets.Count; i++)
                {
                    var name = SecurityElement.Escape(targets[i]);
                    sb.Append("<button id=\"addTo").Append(SlotSeparator).Append(i).Append("\" label=\"").Append(name)
                      .Append("\" tag=\"").Append(name).Append("\" imageMso=\"").Append(AddIcon).Append("\" onAction=\"OnAddToFilter\"/>");
                }
                if (targets.Count == 0)
                    sb.Append("<button id=\"addToNone\" label=\"(no saved filters to add to)\" enabled=\"false\"/>");
                return sb.Append("</menu>").ToString();
            }, "<menu xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\"/>");
        }

        public void OnAddToFilter(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.AddSelectionToFilter(ex, control.Tag));
        }

        public string GetSettingsScreentip(Office.IRibbonControl control)
        {
            return Safe(() => { Focus(control); return "Edit " + SettingsPaths.SavedFiltersName(_controller.Kind); }, "Edit Saved Filters");
        }

        public string GetReloadScreentip(Office.IRibbonControl control)
        {
            return Safe(() => { Focus(control); return "Reload " + SettingsPaths.SavedFiltersName(_controller.Kind); }, "Reload");
        }

        public string GetSettingsSupertip(Office.IRibbonControl control)
        {
            return Safe(() =>
            {
                Focus(control);
                return "Open the saved filters of " + FolderWord(_controller.Kind) + " folders ("
                    + SettingsPaths.SavedFiltersName(_controller.Kind) + ") in VS Code (Notepad if VS Code is not installed). "
                    + "Mail, contacts and tasks each have their own file. Saved changes are picked up automatically. "
                    + SettingsPaths.Describe();
            }, "");
        }

        private static string FolderWord(ItemKind kind)
        {
            return kind == ItemKind.Contact ? "contact" : kind == ItemKind.Task ? "task" : "mail";
        }

        public void OnOpenSettings(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.OpenSettingsFile());
        }

        public void OnReload(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.ReloadSettings(ex));
        }

        public void OnSaveViewFilter(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.SaveCurrentViewFilter(ex));
        }

        public void OnApplyFormats(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.RefreshFormats(ex));
        }

        public void OnAllFormatsOff(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.AllFormatsOff(ex));
        }

        public void OnRemoveFormats(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.RemoveFormats(ex));
        }

        public void OnViewColumns(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.ViewColumnsButton(ex));
        }

        public string GetViewColumnsSupertip(Office.IRibbonControl control)
        {
            return Safe(() =>
            {
                var kind = Focus(control) ?? ItemKind.Mail;
                return "Replace the columns of the current table view with the ones defined for " + FolderWord(kind)
                    + " folders (field, width, format, alignment, heading). Mail, contacts and tasks each have their own file. "
                    + "Ctrl+click: edit the file. File: " + SettingsPaths.ViewColumnsFile(kind);
            }, "");
        }

        public void OnViewFont(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.EditViewFont(ex));
        }

        public void OnFillFields(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.FillFieldsButton(ex));
        }

        public void OnRecipientsReport(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.ShowRecipientsReport(ex));
        }

        public void OnCopyItems(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.CopySelectedItems(ex));
        }

        public void OnAddKnownDomain(Office.IRibbonControl control)
        {
            Run(control, ex => _controller.AddKnownDomainButton(ex));
        }

        public string GetAddKnownSupertip(Office.IRibbonControl control)
        {
            return Safe(() => "Add the base domain of each selected mail's sender (billing.fabrikam.com → fabrikam.com) to the known domains, "
                + "then refill unknownDomain of the selected mails and of this folder's mail from those domains. "
                + "unknownDomain: * unknown sender domain, + known sender but an unknown recipient domain, - all known. "
                + "Ctrl+click: edit the list. File: " + SettingsPaths.KnownDomainsFile, "");
        }

        public bool GetClearOnExitPressed(Office.IRibbonControl control)
        {
            return Safe(() => _controller.ClearInboxFiltersOnExit, true);
        }

        public void OnClearOnExitToggle(Office.IRibbonControl control, bool pressed)
        {
            Run(control, ex => _controller.SetClearInboxFiltersOnExit(pressed));
        }

        public bool GetAutoFillPressed(Office.IRibbonControl control)
        {
            return _controller.AutoFillFields;
        }

        public void OnAutoFillToggle(Office.IRibbonControl control, bool pressed)
        {
            Run(control, ex => _controller.SetAutoFillFields(pressed));
        }

        public bool GetAutoApplyPressed(Office.IRibbonControl control)
        {
            return _controller.AutoApplyFormats;
        }

        public void OnAutoApplyToggle(Office.IRibbonControl control, bool pressed)
        {
            Run(control, ex => _controller.SetAutoApplyFormats(ex, pressed));
        }

        // ---------- Helpers ----------

        /// <summary>Makes the kind of the control's explorer folder current in the controller; null for other folders.</summary>
        private ItemKind? Focus(Office.IRibbonControl control)
        {
            return _controller.Focus(control.Context as Outlook.Explorer ?? Globals.ThisAddIn.Application.ActiveExplorer());
        }

        /// <summary>The field a quick filter slot filters by in the control's folder, or null if none.</summary>
        private QuickKind? QuickKindOf(Office.IRibbonControl control)
        {
            var kinds = FilterController.QuickKindsFor(Focus(control));
            int slot = SlotOf(control);
            return slot >= 0 && slot < kinds.Length ? kinds[slot] : (QuickKind?)null;
        }

        private static int SlotOf(Office.IRibbonControl control)
        {
            int i;
            return int.TryParse(control.Tag, out i) ? i : -1;
        }

        /// <summary>Runs an action for the explorer the control lives in; exceptions become message boxes
        /// (Office silently swallows exceptions thrown from ribbon callbacks).</summary>
        private void Run(Office.IRibbonControl control, Action<Outlook.Explorer> action)
        {
            var explorer = control.Context as Outlook.Explorer ?? Globals.ThisAddIn.Application.ActiveExplorer();
            try
            {
                _controller.Focus(explorer);
                action(explorer);
            }
            catch (UserMessageException ex)
            {
                Notifier.Info(explorer, ex.Message); // guidance, not an error: a notification instead of a message box
                _controller.Invalidate();
            }
            catch (Exception ex)
            {
                Log.Error(control.Id, ex);
                MessageBox.Show(WindowOwner.From(explorer), ex.Message, ThisAddIn.Title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                _controller.Invalidate();
            }
        }

        private static T Safe<T>(Func<T> f, T fallback)
        {
            try
            {
                return f();
            }
            catch (Exception ex)
            {
                Log.Error("ribbon getter", ex);
                return fallback;
            }
        }
    }
}
