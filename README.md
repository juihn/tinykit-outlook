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
.\build.ps1                  # -Configuration Release; runs New-SigningCert.ps1 first if Signing.props is missing
.\Register-Addin.ps1         # only if the add-in is not registered yet; -Unregister to remove, -ShareSettings to share settings through OneDrive
```

`build.ps1` builds with the x64 MSBuild of the latest Visual Studio, which also registers the add-in for the current user
(as building in Visual Studio does). `New-SigningCert.ps1` makes a self-signed cert for the VSTO manifests and
`Signing.props` (per PC, not committed).
Close Outlook before building (it locks the add-in DLL; `build.ps1` stops if it is running) and start it afterwards. The
first time, VSTO asks you to trust the self-signed publisher; choose **Install**.

**Windows on ARM** (e.g. Parallels on Apple Silicon): Visual Studio for ARM64 has no *Office/SharePoint development*
workload. Install it with *.NET desktop development* (and the .NET Framework 4.8 targeting pack), and copy from a PC with
x64 Visual Studio and that workload:

| To `%LOCALAPPDATA%\tinykit\vsto-build-kit\` | From the x64 PC |
|---|---|
| `OfficeTools\` | `<VS>\MSBuild\Microsoft\VisualStudio\v18.0\OfficeTools\*` |
| `VSTO40\` | `Microsoft.Office.Tools.Common.v4.0.Utilities.dll`, `Microsoft.Office.Tools.Outlook.v4.0.Utilities.dll` (any VSTO add-in's output folder has them) |

`build.ps1` uses this folder (`-Kit` for another one) when Visual Studio has no OfficeTools. The Office interop
assemblies come from the GAC (installed with Office), and the x64 MSBuild runs under emulation: the add-in registration
step loads the x64 `vstoee.dll`, which the ARM64 MSBuild cannot.

The namespace is `tinykit.OutlookAddin` (not `tinykit.Outlook`) so the usual
`using Outlook = Microsoft.Office.Interop.Outlook;` alias doesn't collide.

## Contact form (`Forms\`)

`Forms\myContactForm.fdm` (and the same form as `Forms\myContactForm.oft`) is a custom contact form,
message class `IPM.Contact.myContactForm`. To use it:

1. **Install:** *File > Options > Advanced > Custom Forms… > Manage Forms…*, choose *Personal Forms* on the right,
   **Install…**, pick `myContactForm.fdm`, OK.
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

Groups, left to right: Built-in · Custom Mail Fields · Table View · Quick Filter · Clear Filter · Saved Filters · Items.
The tab follows the kind of folder you are in:

| Folder | Custom Mail Fields | Quick Filter | Saved Filters, View Columns |
|---|---|---|---|
| Mail | shown | F S / N D | `- Mail` files |
| Contacts | hidden | F E / C D | `- Contacts` files |
| Tasks | hidden | hidden | `- Tasks` files |
| Other (calendar, notes, ...) | hidden | hidden | Saved Filters hidden |

In calendar folders Clear Filter and Table View are hidden too. Items shows in mail, calendar, contact and task folders, each
button only where it applies.

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
One input box, then two rows of two buttons, each followed by a ▼ drop-down.
The filter shows the items whose field *contains* the value (`LIKE '%value%'`):

| Folder | Button | Column used |
|---|---|---|
| Mail | **F** | the From (sender) e-mail address |
| Mail | **S** | the subject (RE:/FW: ignored when taken from a mail) |
| Mail | **N** | `nameRelated` (taken from a mail without the sent symbol, the contact mark, ` (+)`, `[ ]`/`( )`, so received and sent mail of the same person match) |
| Mail | **D** | `domainRelated` |
| Contacts | **F** | File As |
| Contacts | **E** | E-mail, E-mail 2 or E-mail 3 |
| Contacts | **C** | Company |
| Contacts | **D** | Department |

- **Button:** uses the text in the input box. If the box is empty, it uses the value of the first selected mail or contact.
- **Alt + the button's letter** in the box (e.g. **Alt+S** filter by Subject, **Alt+D** by `domainRelated` in mail
  folders) does what the button does, with the text as typed, without leaving the box.
- **▼:** opens a one-level list of that field's last 19 values as `[short date] value`. Clicking one filters by it and puts it in the box.
  **Ctrl+click** removes it from the list. Every field keeps its own list, so mail and contact values never mix.
- **Tab / Shift+Tab** go from the box through the buttons only, skipping the ▼ lists (the ribbon has no tab-stop
  setting: the add-in follows the focus and sends another Tab when it lands on a list).
- The box keeps its text after filtering, so the same word can be tried with another field. **Clear Filter** empties it.

### Clear Filter group
- **Clear Filter** (**Ctrl+Alt+C**) removes the quick or saved filter and restores the view's own filter (recorded in `ViewState.xml`, so it survives a restart).
- **Clear Inboxes on Exit** (check box, on by default): when Outlook closes (its last window), the quick, saved and Others filters are
  cleared in every view of every account's Inbox, the same way as Clear Filter, so Outlook opens with full Inboxes.
  Filters that are part of a view's own definition (View Settings > Filter) stay. Stored as `clearInboxFiltersOnExit` in
  `Saved Filters - Mail.xml`.

### Custom Mail Fields group (mail folders)
The add-in writes these text columns (user properties) on mail and meeting items:

| Column | Received mail | Mail I sent |
|---|---|---|
| `domainRelated` | sender's domain label, then its subdomains nearest first: `a@billing.fabrikam.com` → `fabrikam/billing` (for *on behalf of*, the principal's) | same, for the first To recipient (no To: first Cc) |
| `nameRelated` | `☺ contact name` if the sender is in Contacts (mail filled before 2026-10: `[contact name]`), else the sender display name; `(local part)` when that is just the address | the `me` column's sent symbol (`▶` by default) + a space + first recipient (`tos` / `ccs` count the recipients; mail filled before 2026-10 may still end in ` (+)`) |
| `me` | `●` I am in To · `○` I am only in Cc · `-` neither (e.g. Bcc, a list address) | `▶` |
| `tos` / `ccs` | number of To / Cc recipients (`-` for 0) | same |
| `domainMark` | `🅄` the sender's domain is not in *Known Domains.txt* · `+` it is, but a recipient's is not · `-` all known · `🅙` moved from Junk Email by Move Junk Mail to Inbox (kept when the fields are filled again) | `+` / `-` for the recipients (I am a known sender) |

- For mail I sent, the first recipient is shown as `☺ contact name` when the address is in Contacts, however it was entered. Otherwise it shows the name Outlook displayed if it was picked from an address book (e.g. the GAL), else the address itself.
- Contact name means the contact's e-mail display name for that address, or File As if that is empty. A nickname in the display name is left out when the contact also has a first, middle or last name (it is used only when it is the only name). Contacts are looked up in every store's Contacts folder and its subfolders (e.g. `olk/family`), skipping system folders such as Recipient Cache and GAL Contacts.
- `domainMark` was called `unknownDomain` before 2026-10, with `*` for an unknown sender. A while after Outlook starts,
  mail that still has `unknownDomain` gets its value in `domainMark` (`*` → `🅄`), a few at a time; saved filters and the
  View Columns file are read with the new name too.
- A mail counts as "mine" when its sender (or principal) is one of my accounts' addresses.
- **Auto-fill new mail** fills mail as it arrives in each account's Inbox or Sent Items.
- **Fill Fields** fills the current folder's items that have no `domainRelated` or `domainMark` yet, e.g. mail that arrived
  while Outlook was closed or in folders other than Inbox and Sent Items.
  **Shift+click** recomputes the selected items instead, even if they already have values (e.g. after editing Known Domains.txt).
  **Ctrl+click** clears the fields of the selected items (asked first); they stay empty until filled again.
- **Add Known Domain** adds the base domain of each selected mail's sender (`a@billing.fabrikam.com` → `fabrikam.com`) to
  `Known Domains.txt`, then refills the selected mails and this folder's mail from those domains that are still marked `🅄` or `+`.
  **Ctrl+click** opens `Known Domains.txt` in VS Code (Notepad if not installed).
- The Custom Mail Fields group's **dialog button** (corner arrow) opens **Column Symbols**: a drop-down for the **Contact** mark before a
  name from Contacts in `nameRelated` (☺ ☻ ✆ ☎ and Ⓒ in a circle or square, plain or negative; kept as `contactMark`), and one each for mail I **Sent**
  (→ ⇒ ⇥ ⇨ ▶ ▷ ⟶ ⟹), me in To and me only in Cc (■ □ ▢ ◆ ○ ● ✓ and Ⓣ/Ⓒ in a circle or square, plain or negative),
  each shown with its code point and Unicode name. A changed symbol is also changed in the saved filters that compare
  `me` with it, and, if you agree, in every mail that has it (all mail folders). Sent mail's `nameRelated` starts with
  the same sent symbol and one space: on OK, mail whose `nameRelated` has another sent symbol, no space after it, or
  another contact mark is set right too (asked first, with the counts). A Font drop-down at the top (the view's font and size at first) shows the
  symbol drop-downs in that font, as a preview of how they will look in the list; it does not change the view's font
  (use View Font... for that). Kept as `meSent` / `meTo` / `meCc` in `Saved Filters - Mail.xml`.

`Known Domains.txt` (in the settings folder) has one domain per line after the date it was added and a tab, e.g.
`'26.09.07월 15:42:39 +08<Tab>fabrikam.com` (date, Korean day of week, time, UTC offset); a line may also be just a domain, and `#` lines are comments. A domain covers
its subdomains. My own addresses among the recipients are ignored. Edits apply to the next mail filled; Shift+click
Fill Fields to recompute existing mail.

### Saved Filters group (from `Saved Filters - Mail.xml`, `- Contacts.xml`, `- Tasks.xml`)
Mail, contact and task folders each have their own saved filters file; the group shows the ones of the folder you are in.
Each file starts with these saved filters, which you can edit or delete:

| File | Filter | Shows | Format |
|---|---|---|---|
| Mail | **Flagged** | flagged (or completed) mail | red, off |
| Mail | **Sent** | mail I sent (`me` = `▶`) | teal, underlined |
| Mail | **Unknown** | received mail whose sender is not in Contacts (`nameRelated` not `[…]`) | gray |
| Contacts | **No Email** | contacts without any e-mail address | gray |
| Contacts | **Flagged** | flagged contacts | red, off |
| Tasks | **Active** | tasks not completed | bold, off |
| Tasks | **Completed** | completed tasks | gray, strikeout |
| Tasks | **High** | active tasks of high importance | red |

Sent and Unknown use the Custom Mail Fields, so older mail needs **Fill Fields** once.

- **One toggle button per saved filter** (up to 20). Pressing it applies the filter's SQL as the view filter. Pressing it again restores the view's own filter.
  **Shift+click** on one of your own filters (mail; not Flagged, Sent, Unknown or Others) adds the selected mails to it,
  the same as picking it in **Add/New...**.
  Filters with an `icon` in the file (the first-install Flagged, Sent and Unknown: the mail list's own kinds of items)
  come first, with their icon, then a separator and the other filters as text.
- **Others** (toggle with a funnel icon, the last filter button) shows only the items that **none** of the saved filters
  match: `NOT ((filter 1) OR (filter 2) ...)`. Filters without SQL are left out. Pressing it again restores the view's own filter.
- **Add/New...** (menu, mail) lists the saved filters except `Flagged`, `Sent` and `Unknown` (the first-install filters that
  are computed from the mail itself), including ones added to the file later and ones that have a name and format but no
  SQL yet (they are filled in place), and last **New...**, which asks for the name of a new saved filter and then builds
  it from the selected mails in the same way (the filter is created when its conditions are added).
  Picking a filter opens a small window with three check boxes:
  - **domainRelated**: `domainRelated = '...'` for each selected mail's domain.
  - **From address**: the sender's address (`PR_SENDER_SMTP_ADDRESS` or the raw sender address, as the **F** quick filter).
  - **Subject**: a subject pattern for each selected mail.
  - **several**: all of them together, e.g. `(From = '...' AND subject pattern)`, one line per mail with the values in that
    order separated by tabs (`domainRelated <Tab> from <Tab> subject`); e.g. only the invoices of one company sent through
    a shared billing service.

  The first time in a session `Issue` starts with domainRelated and every other filter with Subject; afterwards the
  last choice for that filter. Subjects become patterns in which numbers (dates, times, amounts, ids, `9월 27일`,
  `1,234원`) and month/weekday names are `%` (`Your trip with Gojek on Friday, 26 September` → `Your trip with Gojek on %`);
  a subject without them is matched exactly. The list can be edited before adding.
  DASL `LIKE` only honours `%` at the start or end, so a pattern with `%` in the middle is added as prefix/middle/suffix
  conditions joined with `AND`.
  Each value is added as one condition per line (`... OR` + new line); values the filter already covers are skipped.
- After a separator, one column of **Format**, **Refresh Formats** and **Manage**:
- **Format** (toggle) applies to the saved filter that is currently applied:
  - **Click:** turns its format on or off as conditional formatting.
  - **No format defined yet:** opens the Format dialog first, and the new format is turned on.
  - **Ctrl+click:** opens the Format dialog to edit it (font style, strikeout, underline, color). Font name and size follow the view (View Font).
- **Refresh Formats** turns the formats back on (after All Formats Off) and rewrites the rules of all filters whose Format is on into the current view.
- **Manage** (menu):
  - **Edit Saved Filters** opens the current folder kind's saved filters file in VS Code (Notepad if not installed). Saved changes are picked up automatically; **Reload** forces it.
  - **Save View as Filter...** saves the current view filter (e.g. one built in View Settings > Filter) as a saved filter.
  - **Reload** reads the saved filters file again now.
  - **All Formats Off** takes all saved filters' formats out of the views (`formatsOn="false"`). Each filter's Format on/off setting is kept, so Refresh Formats brings the same set back. Turning a filter's Format on also turns formats back on.
  - **Remove Formats** deletes the add-in's rules from the current view after a warning. Settings are unchanged, so the rules come back with Refresh Formats or the next auto-apply.
  - **Auto-apply formats**: when on, syncs the rules into each table view as you switch folders and views.

The rules are written to the view's *View Settings > Conditional Formatting* and named `[TK] <filter name> #<SQL hash>`.
Rules without that prefix are never changed. Color is limited to the 16 colors Outlook's conditional formatting supports.

### Tools group (mail, calendar, contact and task folders)
Tools for the selected items; each button shows only in the folders where it applies.
- **Custom Filter** (mail, calendar, contact and task folders; **Ctrl+Alt+F** in the main window) opens a resizable window: the text to find with
  **Clear All Conditions** / **Apply** beside it, a check box per field, and the filter it makes (shown as you type or
  tick). The item shows when the text appears in any ticked field (`LIKE '%text%'`, fields joined with OR). **Enter**
  applies, **Esc** empties the text (or, when it is empty, closes the window); the title bar shows when it was applied. **Apply** with no text or no field ticked
  shows all items again (so does Clear Filter in table views). The window stays open and filters the folder it was
  opened in; it remembers the text and ticks per kind of folder until Outlook closes, and opens at its last size
  (`Custom Filter Size.txt` in `%APPDATA%\tinykit\Outlook\`). Fields:

  | Folder | Check boxes (a field with several properties matches if any of them does) |
  |---|---|
  | Mail | Sender Name, From Address, nameRelated, domainRelated, Recipient Name (To, Cc, Bcc), Subject, Message Body |
  | Calendar | Sender Name and From Address (the organizer), Subject, Message Body |
  | Contacts | Company, Department, Name (first, last, middle, nickname, e-mail 1/2 display names), Email Address (1, 2), Phone Number (business, mobile, home, other, pager, ...), Notes |
  | Tasks | Subject, Message Body |

  Message Body and Notes start unticked (searching bodies is slower). Calendar, card and other views can be filtered
  too, not only table views.
- **Find Items** (**Ctrl+Alt+E** in the main window) opens a resizable window that finds text in mail, calendar items, contacts and tasks of all folders
  at once (each folder read with `Folder.GetTable`, `LIKE '%text%'`, up to 5000 results). It searches the subject and

  | Items | Also searched |
  |---|---|
  | Mail | sender name |
  | Calendar | location |
  | Contacts | File As, company, department, nickname, e-mail 1-3 addresses and display names, phone numbers |
  | Tasks | – |

  and, with **Message Body / Notes** ticked, the body. Under the text, check boxes choose the item types (with their
  icons) and the accounts (stores; public folders start unticked); hidden folders and Sync Issues are left out.
  **Enter** searches. **Esc** stops a search; otherwise it puts the focus in the text, or, when the focus is there,
  selects the text, and closes the window when the text is empty. The results:

  | Type (icon) | Date | Subject | Field 1 | Field 2 | Field 3 |
  |---|---|---|---|---|---|
  | Mail | received | subject | nameRelated | domainRelated | body |
  | Calendar | start | subject | end | time zone | body |
  | Contacts | created | full name / nickname | company | department | notes |
  | Tasks | due | subject | status | completed | body |

  each with its account and folder; dates are local time, `'yy.MM.dd요일 HH:mm` (a due date without the time); click a
  header to sort. Under them, the selected item's details (a field without a value is left out with its label):
  - mail: subject, received `'yyMMdd요일 HH:mm:ss +09:00`, sender name and address, recipients, carbon copy;
  - calendar: subject, start ~ end and time zone, location;
  - contacts: File As; created / modified / anniversary; full name / nickname; company / department / job title;
    e-mail addresses with display names; phone numbers (mobile, pager, work, ...); business address; free/busy
    address; IM address;
  - tasks: subject; due date / status / completed; reminder time; recurrence;

  then the body or notes, without its blank (or white-space only) lines. **Enter** searches wherever the focus is,
  except in the results, where it (or a double-click) opens the item; after a search the first result is selected and
  has the focus, so Enter opens it (Esc goes back to the text). The window keeps its size,
  splitter, Message Body, item types and accounts (`Find Items.txt` in `%APPDATA%\tinykit\Outlook\`).
- **Send Delay** (mail folders): Off, 5, 10, 15, 30 or 60 seconds (default 60). Mail you send waits that long in the
  Outbox (*Do not deliver before*): open it there to change it and send it again, or delete it to cancel. A later delivery
  time you set yourself is kept. Exchange accounts are held by the server, so Outlook may be closed meanwhile; IMAP accounts
  go only while Outlook runs. With *Send immediately when connected* off (File > Options > Advanced), mail goes at the
  first Send/Receive after the delay. (`sendDelaySeconds` in `Saved Filters - Mail.xml`.)
- **Custom Shortcuts** shows the add-in's keyboard shortcuts with a check box each; untick one to give the key back to
  Outlook (`Custom Shortcuts.txt` in the settings folder). They work in the main window, but not while typing in a reply
  in the reading pane, where Outlook's own keys apply:

  | Key | Command | Outlook's own use of the key in the main window |
  |---|---|---|
  | Ctrl+Alt+Q | Quick Filter box (shows the TinyKit tab, cursor in the box) | – |
  | Ctrl+Alt+W | Open in New Window (the current folder) | – |
  | Ctrl+Alt+E | Find Items | – |
  | Ctrl+Alt+R | Reading pane Right → Bottom → Off → Right | Reply with Meeting |
  | Ctrl+Alt+F | Custom Filter | Forward as attachment |
  | Ctrl+Alt+C | Clear Filter | – |

- **Open Contact of MailItem** (mail folders) opens the contact of the person in the selected mail's
  `nameRelated`: the first recipient of mail I sent, otherwise the sender (any Contacts folder). Without one, a new
  contact with that name and address opens in the Contacts folder of the mail's account (else of the default account),
  to be saved there.
- **Move Sent Mail to Inbox** (mail folders) moves mail in each account's Sent Items to the same account's Inbox, so a conversation reads
  in one place: a few seconds after Outlook starts and after mail arrives in Sent Items, each time all of Sent Items (so
  mail sent from another mail client and synchronized is moved too). Meeting requests stay. Gmail accounts too:
  there Sent Mail is a label, and moving out of it over IMAP takes that label off.
  (`moveSentToInbox` in `Saved Filters - Mail.xml`; on by default.)
- **Move Junk Mail to Inbox** (mail folders) moves mail arriving in each account's Junk Email to the same account's Inbox,
  a few seconds after it arrives, so nothing new is lost to a wrong junk verdict. Mail already in Junk Email when Outlook
  starts or the option is turned on stays there. For Gmail, moving out of Spam marks the mail not spam. The moved
  mail gets `🅙` in `domainMark`.
  (`moveJunkToInbox` in `Saved Filters - Mail.xml`; off by default.)
- **Recipients Report** (mail and calendar folders) opens a window for the selected mail: the sender, then the
  recipients grouped by **domain** and by the contacts' **department**. Blue bullet: To, gray: Cc/Bcc; green background:
  in Contacts. The check boxes add each person's display name and user name. Click a person to open the contact (or
  search LinkedIn when there is none), **Ctrl+click** for a new contact with the name and address filled in,
  **Shift+click** to add the address to the addresses on the clipboard. **Search** finds people by name or user name and
  lists their addresses; **Copy Contents** copies the report as text; **Refresh** reads the mail and Contacts again. The
  subject opens the mail. For a calendar item: the organizer, then the attendees (required as To, optional ones and
  resources as Cc).
- **Copy Items Text** (mail folders) copies one line per selected mail (or meeting request) to the clipboard, e.g.
  `'26.09.28월 17:01 <The Mulia Bali> Ultimate Getaway`, in the order the view shows them. **Shift+click** puts the new
  lines before the clipboard's current text, to collect mails from several folders.
- **Default Contact Form** (contact folders) sets the selected contacts' message class to `IPM.Contact`, so they open
  with Outlook's form.
- **<custom form>** (contact folders, e.g. *myContactForm*) sets it to the custom form that is the default form of this
  contact folder (or else of the default Contacts folder), e.g. `IPM.Contact.myContactForm`. Hidden when no such form is
  set. Outlook keeps using the old form for an item it has in memory, so after a change both buttons clear the
  selection, briefly switch to the Inbox (with the window's painting suspended, so it does not flicker), come back and
  select the same contacts again. If a contact still opens with the old form, restart Outlook.

### Table View group
- **Apply Predefined Columns** switches the folder to the **TinyKit** view and gives it the columns in the View Columns
  file of the folder's kind (`View Columns - Mail.txt`, `- Contacts.txt`, `- Tasks.txt`), in that order. The TinyKit view
  is made on first use as a copy of the current table view (its font, sort, filter and conditional formatting), for all
  folders of that kind. When the folder is shown otherwise (e.g. contacts as business cards), it is a copy of the
  folder's first table view (e.g. List or Phone), else a new table view. Compact and Outlook's other views stay as they are. A quick or saved filter shown is cleared
  first. **Ctrl+click** opens the file in VS Code (Notepad if not installed). Each file is created with default columns
  on first use.
- **Automatic column sizing** turns the current view's View Settings > Other Settings > Automatic column sizing on or off
  (on: the columns fill the width of the list; off: they keep their widths and the list scrolls sideways). Turned off in
  the TinyKit view, the predefined columns are applied again, so the columns always get the file's widths.
- **View Font...** sets the font and size (9, 10, 11 or 12) for the whole current table view: rows, and optionally column
  headers and all conditional formatting rules, which keep their own style and color. Table views store whole point sizes
  only (9.5pt is saved as 9pt), so there are no half sizes.
- **Sort by Company/Dept** (contact folders) sorts the view by Company, and within a company by Department, both A to Z;
  the view keeps this sort (its grouping, if any, stays).

### Saved Filters - Mail.xml / - Contacts.xml / - Tasks.xml
```xml
<SavedFilters autoApplyFormats="true" autoFillFields="true" moveSentToInbox="true" sendDelaySeconds="60" formatsOn="true">
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
so fields can be left out (`me⇥Center` is a centered column with no width; `domainMark⇥4⇥Center⇥ud` skips Type and
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
domainMark	4					Center	ud
Size			9				3	Right
```

The add-in's own fields (`domainRelated`, `nameRelated`, `me`, `tos`, `ccs`) are created in the folder automatically when missing.
Columns that can't be added, or a Format that is not in a field's list, are reported; the other columns are still applied.

Files: `Saved Filters - <kind>.xml`, `View Columns - <kind>.txt` and `History.xml` in the settings folder (see *Settings location*;
older `Filters.xml`, `Saved Filters.xml` and `View Columns.txt` are renamed automatically); `ViewState.xml` and `OutlookAddin.log` (errors) in `%APPDATA%\tinykit\Outlook\`.

## **TinyKit** tab in a received mail's window
Before the window's Message tab, with one **Message** group:
- Delete, Archive, Send to OneNote · Follow Up, Flag (no date), Clear Flag · Translate, Show Original, translation
  preferences · Approve / Reject (approval requests only) · Find · Edit Message: Outlook's own commands.
- **Recipients Report** of the open mail, as in the Tools group.

## **TinyKit** tab in a mail being written
Before the window's Message tab, with a **Recipients** group:
- **Remove Sender from Recipients** removes the sending account's own address from To, Cc and Bcc (e.g. after Reply All).
- **Restate Recipients** replaces each recipient found in Contacts (by address, or by the contact's e-mail display name)
  with that contact entry, so it shows with the name set in Contacts; its type (To/Cc/Bcc) is kept. The replaced
  recipients move after the others, those in the sender's own domain (base domain, e.g. `contoso.com` for
  `a@mail.contoso.com`) last. Then Outlook's Check Names runs.
- **Recipients Report** of this mail, as in the Tools group.

and a **Compose** group of Outlook's own commands: theme Fonts, Ruler, Bcc.

The tab and group have qualified ids (`tk:ComposeTab`, `tk:ComposeRecipients`, before `tk:AfterRecipients`) so other
add-ins can add groups, as in the contact window.

## **TinyKit** tab in a contact's window
Before the window's Contact tab, with two groups:
- **Built-in**: General, Details, All Fields (the window's pages), Delete, Save & Close.
- **TinyKit**:
  - **Contact Picture** (large, shows the contact's picture): adds a picture, or changes it; **Shift+click** removes it.
  - **Open Address in Google Map** (the business address, or else home or other).
  - **Copy to Clipboard**: `company / department / name (job title)`, e-mail, `T.`phone, `M.`mobile, tab-separated;
    **Shift+click** keeps the clipboard's text after it.

To switch contacts between Outlook's form and the folder's custom form, use the Tools group of the main window (the
contact window keeps the form it opened with).

Other add-ins can add their own groups to this tab: declare `xmlns:tk="tinykit"` in their ribbon XML and use
`<tab idQ="tk:ContactTab" label="TinyKit" insertBeforeMso="TabContact">`. To put a group right after TinyKit's groups,
give it `idQ="tk:AfterBuiltIn"` and `insertAfterQ="tk:ContactTools"`: Outlook merges groups in load order and ignores a
reference to a group not loaded yet, so Built-in and TinyKit also name `tk:AfterBuiltIn` in their `insertBeforeQ`.

## License

[MIT](LICENSE)
