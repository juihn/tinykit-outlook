# tinykit for Outlook

A small VSTO add-in (C#, .NET Framework 4.8) for **Outlook Classic** on Windows. It adds a **Tools** ribbon tab with:

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

## Settings location (sharing across PCs)

Shared settings files are kept in the personal OneDrive when this folder exists:

```
%OneDriveConsumer%\.config\            (hidden; one folder per program)
    tinykit\
        Outlook\   Saved Filters.xml, History.xml
        Word\      (later)
```

- **Folder exists:** every PC signed in to the same personal OneDrive uses the same saved filters and recent values.
  Changes are picked up automatically when OneDrive syncs the file.
- **Folder missing:** the files stay in `%APPDATA%\tinykit\Outlook\` (this PC only).
- **Always local:** `ViewState.xml` (each view's own filter) and `OutlookAddin.log` stay in `%APPDATA%\tinykit\Outlook\`.
- `%OneDriveConsumer%` is the *personal* OneDrive (`%OneDrive%` can be a work account). If the variable is missing, the
  OneDrive client's registry entry is used.
- The **Edit Saved Filters** tip and the header comment of `Saved Filters.xml` show where the settings are.

**To start sharing:** create the folder (or run `Register-Addin.ps1 -ShareSettings`) and restart Outlook.
The current local files are copied there; the local copies stay as a backup.

**On another PC:**
1. Clone this repository, build, and run `.\Register-Addin.ps1` (see *Build / install*).
   The script also sets `.config` hidden, because OneDrive syncs the folder but not its Hidden attribute.
2. Make sure OneDrive has synced `.config\tinykit\Outlook\`, then start Outlook and accept the VSTO trust prompt once.

## **Tools** tab

Groups, left to right: Custom Fields · Quick Filter · Clear · Saved Filters · View.

### Quick Filter group
One input box, then **F** (From) and **S** (Subject) on the second row and **N** (Name) and **D** (Domain) on the third, each followed by a ▼ drop-down.
The filter shows mail whose field *contains* the value (`LIKE '%value%'`):

| Button | Column used |
|---|---|
| **F** | the From (sender) e-mail address |
| **N** | `nameRelated` (taken from a mail without `▶ `, ` (+)`, `[ ]`/`( )`, so received and sent mail of the same person match) |
| **S** | the subject (RE:/FW: ignored when taken from a mail) |
| **D** | `domainRelated` |

- **Button:** uses the text in the input box. If the box is empty, it uses the value of the first selected mail.
- **▼:** opens a one-level list of that field's last 19 values as `[short date] value`. Clicking one filters by it and puts it in the box.
  **Ctrl+click** removes it from the list.
- The box keeps its text after filtering, so the same word can be tried with another field. **Clear Filter** empties it.

### Clear group
- **Clear Filter** removes the quick or saved filter and restores the view's own filter (recorded in `ViewState.xml`, so it survives a restart).

### Custom Fields group
The add-in writes these text columns (user properties) on mail and meeting items:

| Column | Received mail | Mail I sent |
|---|---|---|
| `domainRelated` | sender's domain label, then its subdomains nearest first: `a@billing.fabrikam.com` → `fabrikam/billing` (for *on behalf of*, the principal's) | same, for the first To recipient (no To: first Cc) |
| `nameRelated` | `[contact name]` if the sender is in Contacts, else the sender display name; `(local part)` when that is just the address | `▶ ` + first recipient + ` (+)` when there are several recipients |
| `me` | `●` I am in To · `○` I am in Cc · `-` | `▶` |
| `tos` / `ccs` | number of To / Cc recipients (`-` for 0) | same |

- For mail I sent, the first recipient is shown as follows. If it was picked from an address book (contact or GAL), it shows the name Outlook displayed. If an address was typed or pasted (`Name <address>`), it shows `[contact name]` when the address is in Contacts, otherwise the address itself.
- Contact name means the contact's e-mail display name for that address, or File As if that is empty. Contacts are looked up in every store's Contacts folder and its subfolders (e.g. `olk/family`), skipping system folders such as Recipient Cache and GAL Contacts.
- A mail counts as "mine" when its sender (or principal) is one of my accounts' addresses.
- **Auto-fill new mail** fills mail as it arrives in each account's Inbox or Sent Items.
- **Fill Fields** recomputes the selected items.
- **Fill Missing** fills the current folder's items that have no `domainRelated` yet, e.g. mail that arrived while Outlook was closed.

### Saved Filters group (from `Saved Filters.xml`)
A first install starts with three saved filters, which you can edit or delete:

| Filter | Shows | Format |
|---|---|---|
| **Flagged** | flagged (or completed) mail | red, off |
| **Sent** | mail I sent (`me` = `▶`) | teal, underlined |
| **Unknown** | received mail whose sender is not in Contacts (`nameRelated` not `[…]`) | gray |

Sent and Unknown use the Custom Fields, so older mail needs **Fill Missing** once.

- **One toggle button per saved filter** (up to 20). Pressing it applies the filter's SQL as the view filter. Pressing it again restores the view's own filter.
- **Format** (toggle) applies to the saved filter that is currently applied:
  - **Click:** turns its format on or off as conditional formatting.
  - **No format defined yet:** opens the Format dialog first, and the new format is turned on.
  - **Ctrl+click:** opens the Format dialog to edit it (font style, strikeout, underline, color). Font name and size follow the view (View Font).
- **Add to Delete** adds the **subject** of each selected mail to the saved filter `Delete` (`subject = ...`).
- **Add to Issue** adds the **domainRelated** of each selected mail to the saved filter `Issue` (`domainRelated = ...`). `Issue` is created right after `Delete` on first use.
  Both add one condition per line (`... OR` + new line) and skip values the filter already covers.
- **Edit Saved Filters** opens `Saved Filters.xml` in VS Code (Notepad if not installed). Saved changes are picked up automatically; **Reload** forces it.
- **Save View as Filter...** saves the current view filter (e.g. one built in View Settings > Filter) as a saved filter.
- **Refresh Formats** turns the formats back on (after All Formats Off) and rewrites the rules of all filters whose Format is on into the current view.
- **All Formats Off** takes all saved filters' formats out of the views (`formatsOn="false"`). Each filter's Format on/off setting is kept, so Refresh Formats brings the same set back. Turning a filter's Format on also turns formats back on.
- **Remove Formats** deletes the add-in's rules from the current view after a warning. Settings are unchanged, so the rules come back with Refresh Formats or the next auto-apply.
- **Auto-apply formats**: when on, syncs the rules into each table view as you switch folders and views.

The rules are written to the view's *View Settings > Conditional Formatting* and named `[TK] <filter name> #<SQL hash>`.
Rules without that prefix are never changed. Color is limited to the 16 colors Outlook's conditional formatting supports.

### View group
- **View Font...** sets the font and size for the whole current table view: rows, and optionally column headers and all conditional
  formatting rules, which keep their own style and color.

### Saved Filters.xml
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

Files: `Saved Filters.xml` and `History.xml` in the settings folder (see *Settings location*; an older `Filters.xml` is renamed
automatically); `ViewState.xml` and `OutlookAddin.log` (errors) in `%APPDATA%\tinykit\Outlook\`.

## License

[MIT](LICENSE)
