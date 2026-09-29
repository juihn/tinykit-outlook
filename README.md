# tinykit for Outlook

A small VSTO add-in (C#, .NET Framework 4.8) for **Outlook Classic** on Windows. It adds a **TinyKit** ribbon tab with:

- **Quick filters** by From, Name, Subject and Domain, with recent values.
- **Saved filters** (DASL/SQL) kept in an XML file, each with an optional conditional format (style, strikeout, underline, color).
- **Custom mail fields** (`domainRelated`, `nameRelated`, `me`, `tos`, `ccs`) to sort, group and filter by.
- **View Font** for the whole table view.

tinykit is a set of small Office tools; add-ins for other hosts (e.g. `tinykit.Word`) live in their own repositories.

## Build / install

Requirements: Windows, Outlook Classic (Microsoft 365 / 2016 or later), Visual Studio with the *Office/SharePoint development* workload.

```powershell
.\New-SigningCert.ps1        # once per PC: self-signed cert for the VSTO manifests + Signing.props (not committed)
& "$env:ProgramFiles\Microsoft Visual Studio\18\Enterprise\MSBuild\Current\Bin\MSBuild.exe" tinykit.Outlook.sln -p:Configuration=Debug
.\Register-Addin.ps1         # -Unregister to remove, -ShareSettings to share settings through OneDrive
```

Building in Visual Studio registers the add-in automatically; a command-line build does not, hence `Register-Addin.ps1`.
Close Outlook before building (it locks the add-in DLL) and start it afterwards. The first time, VSTO asks you to trust the
self-signed publisher; choose **Install**.

The namespace is `tinykit.OutlookAddin` (not `tinykit.Outlook`) so the usual
`using Outlook = Microsoft.Office.Interop.Outlook;` alias doesn't collide.

## Contact form (`Forms\`)

`Forms\myContactForm.fdm` (and the same form as `Forms\myContactForm.oft`) is a custom contact form,
message class `IPM.Contact.myContactForm`. To use it:

1. **Install:** *File > Options > Advanced > Custom Forms??> Manage Forms??, choose *Personal Forms* on the right,
   **Install??*, pick `myContactForm.fdm`, OK.
2. **Make it the Contacts folder's form:** right-click *Contacts* > *Properties* > *When posting to this folder, use:*
   **myContactForm**. New contacts then open with it, and the Built-in group shows a **myContactForm** button that
   switches selected contacts to it (**Default Contact Form** switches them back).

The files had their author details (name, e-mail, Exchange address, message IDs) replaced with neutral values of the
same length; the form itself is unchanged.

## Settings location (sharing across PCs)

Shared settings files are kept in the personal OneDrive when this folder exists:

```
%OneDriveConsumer%\.config\            (hidden; one folder per program)
    tinykit\
        Outlook\   Saved Filters - Mail.xml, Saved Filters - Contacts.xml, Saved Filters - Tasks.xml,
                   View Columns - Mail.txt, View Columns - Contacts.txt, View Columns - Tasks.txt,
                   Known Domains.txt, History.xml
        Word\      (later)
```

- **Folder exists:** every PC signed in to the same personal OneDrive uses the same saved filters and recent values.
  Changes are picked up automatically when OneDrive syncs the file.
- **Folder missing:** the files stay in `%APPDATA%\tinykit\Outlook\` (this PC only).
- **Always local:** `ViewState.xml` (each view's own filter) and `OutlookAddin.log` stay in `%APPDATA%\tinykit\Outlook\`.
- `%OneDriveConsumer%` is the *personal* OneDrive (`%OneDrive%` can be a work account). If the variable is missing, the
  OneDrive client's registry entry is used.
- The **Edit Saved Filters** tip and the header comment of the saved filters files show where the settings are.
- Files from before the per-kind split (`Saved Filters.xml`, `View Columns.txt`) are renamed to the `- Mail` files at startup.

**To start sharing:** create the folder (or run `Register-Addin.ps1 -ShareSettings`) and restart Outlook.
The current local files are copied there; the local copies stay as a backup.

**On another PC:**
1. Clone this repository, build, and run `.\Register-Addin.ps1` (see *Build / install*).
   The script also sets `.config` hidden, because OneDrive syncs the folder but not its Hidden attribute.
2. Make sure OneDrive has synced `.config\tinykit\Outlook\`, then start Outlook and accept the VSTO trust prompt once.

## **TinyKit** tab

Groups, left to right: Built-in 쨌 Custom Mail Fields 쨌 Table View 쨌 Quick Filter 쨌 Clear 쨌 Saved Filters 쨌 Items.
The tab follows the kind of folder you are in:

| Folder | Custom Mail Fields | Quick Filter | Saved Filters, View Columns |
|---|---|---|---|
| Mail | shown | F S / N D | `- Mail` files |
| Contacts | hidden | F E / C D | `- Contacts` files |
| Tasks | hidden | hidden | `- Tasks` files |
| Other (calendar, notes, ...) | hidden | hidden | Saved Filters hidden |

In calendar folders Clear and Table View are hidden too. Items shows in mail, calendar and contact folders, each button
only where it applies.

Informational messages (results such as *Custom fields: 3 updated*, and hints such as *Select the mails first*) appear as
Windows notifications under Outlook (classic), using the *urgent* scenario so they also show in Do Not Disturb. Errors,
warnings and questions stay message boxes. If Outlook's notifications are turned off in Windows Settings, message boxes
are used instead.

### Built-in group
Frequently used Outlook commands, icons only, in columns (the list is `BuiltInColumns` in `OutlookRibbon.cs`):
- Categorize, All Categories, Send to OneNote
- Address Book, Reminders Window
- *separator*
- Flag (no date), Clear Flag, Move to Folder
- Mark as Read / Mark as Unread (only the one that applies is shown), Show as Conversations (check box, no label),
  Messages in this Conversation

The rightmost column has Outlook's **Month**, **Week** and **Go To Date**, with labels.

Outlook enables and disables them as on its own tabs. Clear Flag and Messages in this Conversation have no icon of
their own (they live in menus), so they get Delete and GroupConversations.

### Quick Filter group
One input box, then two rows of two buttons, each followed by a ??drop-down.
The filter shows the items whose field *contains* the value (`LIKE '%value%'`):

| Folder | Button | Column used |
|---|---|---|
| Mail | **F** | the From (sender) e-mail address |
| Mail | **S** | the subject (RE:/FW: ignored when taken from a mail) |
| Mail | **N** | `nameRelated` (taken from a mail without `??`, ` (+)`, `[ ]`/`( )`, so received and sent mail of the same person match) |
| Mail | **D** | `domainRelated` |
| Contacts | **F** | File As |
| Contacts | **E** | E-mail, E-mail 2 or E-mail 3 |
| Contacts | **C** | Company |
| Contacts | **D** | Department |

- **Button:** uses the text in the input box. If the box is empty, it uses the value of the first selected mail or contact.
- **??** opens a one-level list of that field's last 19 values as `[short date] value`. Clicking one filters by it and puts it in the box.
  **Ctrl+click** removes it from the list. Every field keeps its own list, so mail and contact values never mix.
- The box keeps its text after filtering, so the same word can be tried with another field. **Clear Filter** empties it.

### Clear group
- **Clear Inboxes on Exit** (toggle button, on by default): when Outlook closes (its last window), the quick, saved and Others filters are
  cleared in every view of every account's Inbox, the same way as Clear Filter, so Outlook opens with full Inboxes.
- **Clear Filter** removes the quick or saved filter and restores the view's own filter (recorded in `ViewState.xml`, so it survives a restart).
  Filters that are part of a view's own definition (View Settings > Filter) stay. Stored as `clearInboxFiltersOnExit` in
  `Saved Filters - Mail.xml`.

### Custom Mail Fields group (mail folders)
The add-in writes these text columns (user properties) on mail and meeting items:

| Column | Received mail | Mail I sent |
|---|---|---|
| `domainRelated` | sender's domain label, then its subdomains nearest first: `a@billing.fabrikam.com` ??`fabrikam/billing` (for *on behalf of*, the principal's) | same, for the first To recipient (no To: first Cc) |
| `nameRelated` | `[contact name]` if the sender is in Contacts, else the sender display name; `(local part)` when that is just the address | `??` + first recipient + ` (+)` when there are several recipients |
| `me` | `?? I am in To 쨌 `?? I am in Cc 쨌 `-` | `?? |
| `tos` / `ccs` | number of To / Cc recipients (`-` for 0) | same |
| `unknownDomain` | `*` the sender's domain is not in *Known Domains.txt* 쨌 `+` it is, but a recipient's is not 쨌 `-` all known | `+` / `-` for the recipients (I am a known sender) |

- For mail I sent, the first recipient is shown as follows. If it was picked from an address book (contact or GAL), it shows the name Outlook displayed. If an address was typed or pasted (`Name <address>`), it shows `[contact name]` when the address is in Contacts, otherwise the address itself.
- Contact name means the contact's e-mail display name for that address, or File As if that is empty. Contacts are looked up in every store's Contacts folder and its subfolders (e.g. `olk/family`), skipping system folders such as Recipient Cache and GAL Contacts.
- A mail counts as "mine" when its sender (or principal) is one of my accounts' addresses.
- **Auto-fill new mail** fills mail as it arrives in each account's Inbox or Sent Items.
- **Fill Fields** fills the current folder's items that have no `domainRelated` or `unknownDomain` yet, e.g. mail that arrived
  while Outlook was closed or in folders other than Inbox and Sent Items.
  **Shift+click** recomputes the selected items instead, even if they already have values (e.g. after editing Known Domains.txt).
- **Add Known Domain** adds the base domain of each selected mail's sender (`a@billing.fabrikam.com` ??`fabrikam.com`) to
  `Known Domains.txt`, then refills the selected mails and this folder's mail from those domains that is still marked `*` or `+`.
  **Ctrl+click** opens `Known Domains.txt` in VS Code (Notepad if not installed).

`Known Domains.txt` (in the settings folder) has one domain per line after the date it was added and a tab, e.g.
`'26.09.07??15:42:39 +08<Tab>fabrikam.com` (date, Korean day of week, time, UTC offset); a line may also be just a domain, and `#` lines are comments. A domain covers
its subdomains. My own addresses among the recipients are ignored. Edits apply to the next mail filled; Shift+click
Fill Fields to recompute existing mail.

### Saved Filters group (from `Saved Filters - Mail.xml`, `- Contacts.xml`, `- Tasks.xml`)
Mail, contact and task folders each have their own saved filters file; the group shows the ones of the folder you are in.
Each file starts with these saved filters, which you can edit or delete:

| File | Filter | Shows | Format |
|---|---|---|---|
| Mail | **Flagged** | flagged (or completed) mail | red, off |
| Mail | **Sent** | mail I sent (`me` = `??) | teal, underlined |
| Mail | **Unknown** | received mail whose sender is not in Contacts (`nameRelated` not `[??`) | gray |
| Contacts | **No Email** | contacts without any e-mail address | gray |
| Contacts | **Flagged** | flagged contacts | red, off |
| Tasks | **Active** | tasks not completed | bold, off |
| Tasks | **Completed** | completed tasks | gray, strikeout |
| Tasks | **High** | active tasks of high importance | red |

Sent and Unknown use the Custom Mail Fields, so older mail needs **Fill Fields** once.

- **One toggle button per saved filter** (up to 20). Pressing it applies the filter's SQL as the view filter. Pressing it again restores the view's own filter.
- **Others** (toggle, after the saved filters) shows only the items that **none** of the saved filters match:
  `NOT ((filter 1) OR (filter 2) ...)`. Filters without SQL are left out. Pressing it again restores the view's own filter.
- **Format** (toggle) applies to the saved filter that is currently applied:
  - **Click:** turns its format on or off as conditional formatting.
  - **No format defined yet:** opens the Format dialog first, and the new format is turned on.
  - **Ctrl+click:** opens the Format dialog to edit it (font style, strikeout, underline, color). Font name and size follow the view (View Font).
- **Add to** (menu, mail) lists the saved filters except `Flagged`, `Sent` and `Unknown` (the first-install filters that
  are computed from the mail itself), including ones added to the file later and ones that have a name and format but no
  SQL yet (they are filled in place). Picking one opens a small window with three check boxes:
  - **domainRelated**: `domainRelated = '...'` for each selected mail's domain.
  - **From address**: the sender's address (`PR_SENDER_SMTP_ADDRESS` or the raw sender address, as the **F** quick filter).
  - **Subject**: a subject pattern for each selected mail.
  - **several**: all of them together, e.g. `(From = '...' AND subject pattern)`, one line per mail with the values in that
    order separated by tabs (`domainRelated <Tab> from <Tab> subject`); e.g. only the invoices of one company sent through
    a shared billing service.

  The first time in a session `Issue` starts with domainRelated and every other filter with Subject; afterwards the
  last choice for that filter. Subjects become patterns in which numbers (dates, times, amounts, ids, `9??27??,
  `1,234??) and month/weekday names are `%` (`Your trip with Gojek on Friday, 26 September` ??`Your trip with Gojek on %`);
  a subject without them is matched exactly. The list can be edited before adding.
  DASL `LIKE` only honours `%` at the start or end, so a pattern with `%` in the middle is added as prefix/middle/suffix
  conditions joined with `AND`.
  Each value is added as one condition per line (`... OR` + new line); values the filter already covers are skipped.
- **Edit Saved Filters** opens the current folder kind's saved filters file in VS Code (Notepad if not installed). Saved changes are picked up automatically; **Manage > Reload** forces it.
- **Save View as Filter...** saves the current view filter (e.g. one built in View Settings > Filter) as a saved filter.
- **Manage** (menu):
  - **Reload** reads the saved filters file again now.
  - **Refresh Formats** turns the formats back on (after All Formats Off) and rewrites the rules of all filters whose Format is on into the current view.
  - **All Formats Off** takes all saved filters' formats out of the views (`formatsOn="false"`). Each filter's Format on/off setting is kept, so Refresh Formats brings the same set back. Turning a filter's Format on also turns formats back on.
  - **Remove Formats** deletes the add-in's rules from the current view after a warning. Settings are unchanged, so the rules come back with Refresh Formats or the next auto-apply.
  - **Auto-apply formats**: when on, syncs the rules into each table view as you switch folders and views.

The rules are written to the view's *View Settings > Conditional Formatting* and named `[TK] <filter name> #<SQL hash>`.
Rules without that prefix are never changed. Color is limited to the 16 colors Outlook's conditional formatting supports.

### Items group (mail, calendar and contact folders)
Tools for the selected items; each button shows only in the folders where it applies.
- **Recipients Report** (mail and calendar folders) opens a window for the selected mail: the sender, then the
  recipients grouped by **domain** and by the contacts' **department**. Blue bullet: To, gray: Cc/Bcc; green background:
  in Contacts. The check boxes add each person's display name and user name. Click a person to open the contact (or
  search LinkedIn when there is none), **Ctrl+click** for a new contact with the name and address filled in,
  **Shift+click** to add the address to the addresses on the clipboard. **Search** finds people by name or user name and
  lists their addresses; **Copy Contents** copies the report as text; **Refresh** reads the mail and Contacts again. The
  subject opens the mail. For a calendar item: the organizer, then the attendees (required as To, optional ones and
  resources as Cc).
- **Copy Items** (mail folders) copies one line per selected mail (or meeting request) to the clipboard, e.g.
  `'26.09.28??17:01 <The Mulia Bali> Ultimate Getaway`, in the order the view shows them. **Shift+click** puts the new
  lines before the clipboard's current text, to collect mails from several folders.
- **Default Contact Form** (contact folders) sets the selected contacts' message class to `IPM.Contact`, so they open
  with Outlook's form.
- **<custom form>** (contact folders, e.g. *myContactForm*) sets it to the custom form that is the default form of this
  contact folder (or else of the default Contacts folder), e.g. `IPM.Contact.myContactForm`. Hidden when no such form is
  set. Outlook keeps using the old form for an item it has in memory, so after a change both buttons clear the
  selection, briefly switch to the Inbox (with the window's painting suspended, so it does not flicker), come back and
  select the same contacts again. If a contact still opens with the old form, restart Outlook.

### Table View group
- **View Columns** replaces the columns of the current table view with the ones in the View Columns file of the folder's
  kind (`View Columns - Mail.txt`, `- Contacts.txt`, `- Tasks.txt`), in that order. **Ctrl+click** opens that file in
  VS Code (Notepad if not installed). Each file is created with default columns on first use.
- **View Font...** sets the font and size (9, 10, 11 or 12) for the whole current table view: rows, and optionally column
  headers and all conditional formatting rules, which keep their own style and color. Table views store whole point sizes
  only (9.5pt is saved as 9pt), so there are no half sizes.
- **Sort by Company/Dept** (contact folders) sorts the view by Company, and within a company by Department, both A to Z;
  the view keeps this sort (its grouping, if any, stays).

### Saved Filters - Mail.xml / - Contacts.xml / - Tasks.xml
```xml
<SavedFilters autoApplyFormats="true" autoFillFields="true" formatsOn="true">
  <Filter name="Unread" formatEnabled="true">
    <Sql><![CDATA["urn:schemas:httpmail:read" = 0]]></Sql>
    <Format font="Segoe UI" size="9" style="Bold" strikeout="false" underline="false" color="Navy" />
  </Filter>
</SavedFilters>
```
`Sql` is DASL, the same text as the SQL tab of View Settings > Filter. The `@SQL=` prefix is optional.
`font` and `size` are optional. `color` takes `Auto`, a palette name, or `#RRGGBB` (mapped to the nearest palette color).
The file is rewritten when you change settings from the ribbon, so comments you add will not be kept.

### View Columns - Mail.txt / - Contacts.txt / - Tasks.txt
One column per line; `#` lines are skipped. The values are separated by tabs, and several tabs in a row count as one,
so the lines can be lined up freely. After the field name, each value goes to the next field whose format it fits,
so fields can be left out (`me?쩉enter` is a centered column with no width; `unknownDomain???쩉enter?쪀d` skips Type and
Format):

| Field (in order) | Format | Meaning |
|---|---|---|
| FieldName | first value | a field as named in View Settings > Columns (`Received`, `Flag Status`, ...) or a user-defined field |
| Width | a number | width in characters |
| Type | `ol` + capital letter | `olText`, `olDateTime`, `olInteger`, `olNumber`, `olYesNo`: creates a missing user-defined field in the folder (others, e.g. `olSize`, are allowed and ignored) |
| Format | a number (after Width) | position (1, 2, ...) in the column's Format drop-down of View Settings |
| Alignment | `Left`, `Center`, `Right` | column alignment |
| Alias | any other text | column heading, when it should differ from the field name |

```
#FieldName		Width
#					Type
#								Format
#									Alignment
#											Alias
Received		18	olDateTime	3
nameRelated		18					Right
unknownDomain	4					Center	ud
Size			9				3	Right
```

The add-in's own fields (`domainRelated`, `nameRelated`, `me`, `tos`, `ccs`) are created in the folder automatically when missing.
Columns that can't be added, or a Format that is not in a field's list, are reported; the other columns are still applied.

Files: `Saved Filters - <kind>.xml`, `View Columns - <kind>.txt` and `History.xml` in the settings folder (see *Settings location*;
older `Filters.xml`, `Saved Filters.xml` and `View Columns.txt` are renamed automatically); `ViewState.xml` and `OutlookAddin.log` (errors) in `%APPDATA%\tinykit\Outlook\`.

## **TinyKit** tab in a received mail's window
Before the window's Message tab, with one **Message** group:
- Delete, Archive, Send to OneNote 쨌 Follow Up, Flag (no date), Clear Flag 쨌 Translate, Show Original, translation
  preferences 쨌 Approve / Reject (approval requests only) 쨌 Find 쨌 Edit Message: Outlook's own commands.
- **Recipients Report** of the open mail, as in the Items group.

## **TinyKit** tab in a mail being written
Before the window's Message tab, with a **Recipients** group:
- **Remove Sender from Recipients** removes the sending account's own address from To, Cc and Bcc (e.g. after Reply All).
- **Restate Recipients** replaces each recipient found in Contacts (by address, or by the contact's e-mail display name)
  with that contact entry, so it shows with the name set in Contacts; its type (To/Cc/Bcc) is kept. The replaced
  recipients move after the others, those in the sender's own domain (base domain, e.g. `contoso.com` for
  `a@mail.contoso.com`) last. Then Outlook's Check Names runs.
- **Recipients Report** of this mail, as in the Items group.

and a **Compose** group of Outlook's own commands: theme Fonts, Ruler, Bcc.

The tab and group have qualified ids (`tk:ComposeTab`, `tk:ComposeRecipients`, before `tk:AfterRecipients`) so other
add-ins can add groups, as in the contact window.

## **TinyKit** tab in a contact's window
Before the window's Contact tab, with one **Built-in** group:
- General, Details, All Fields (the window's pages).
- **Open in Google Map** (the business address, or else home or other), Delete, Save & Close.
- **Contact Picture** (large, shows the contact's picture): adds a picture, or changes it; **Shift+click** removes it.
- **Default Message Class** / **Custom Message Class**: make this contact open with Outlook's form (`IPM.Contact`) or with
  the Contacts folder's custom form (e.g. `IPM.Contact.myContactForm`); it takes effect when the contact is opened again.
- **Copy to Clipboard**: `company / department / name (job title)`, e-mail, `T.`phone, `M.`mobile, tab-separated;
  **Shift+click** keeps the clipboard's text after it.

Other add-ins can add their own groups to this tab: declare `xmlns:tk="tinykit"` in their ribbon XML and use
`<tab idQ="tk:ContactTab" label="TinyKit" insertBeforeMso="TabContact">`. To put a group right after Built-in, give it
`idQ="tk:AfterBuiltIn"` and `insertAfterQ="tk:ContactBuiltIn"`: Outlook merges groups in load order and ignores a
reference to a group not loaded yet, so Built-in also names `tk:AfterBuiltIn` in its `insertBeforeQ`.

## License

[MIT](LICENSE)
