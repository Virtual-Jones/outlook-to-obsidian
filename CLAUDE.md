# Claude Instructions — Outlook to Obsidian (Office Add-in)

## Project Overview
An **Office Add-in** (previously a Chrome extension) that extracts a **calendar appointment or an email message** from Outlook using the Office.js API and creates a formatted Obsidian note via the `obsidian://` URI scheme. The task pane auto-detects which item type is open and adapts.

Works in: Outlook on the Web (Microsoft 365), Outlook Desktop (Windows & Mac), New Outlook for Windows.

**Architecture:**
- **Manifest** ([manifest.xml](manifest.xml)) — Office Add-in manifest; declares the ribbon buttons and task pane URL
- **Task Pane UI** ([taskpane.html](taskpane.html)) — the add-in interface, includes both main view and settings panel
- **Task Pane Logic** ([taskpane.js](taskpane.js)) — all extraction and UI logic using `Office.context.mailbox.item`
- **Function File** ([commands.html](commands.html)) — required by the manifest; currently a no-op placeholder

**Runtime dependencies** (loaded from CDN in [taskpane.html](taskpane.html)):
- **Office.js** — Outlook API
- **Turndown** + **@joplin/turndown-plugin-gfm** — HTML → Markdown conversion of item bodies

## Key Files & Responsibilities

### [taskpane.js](taskpane.js) — Core Logic
- `Office.onReady()` bootstraps the add-in
- `getCurrentItemType()` / `applyContextLabels()` — detect appointment vs message and relabel the UI
- `extractMeetingDetails()` reads all meeting data via Office.js API
- `extractEmailDetails()` reads message data (from / to / cc / sent / body)
- `getItemProperty(propOrValue)` — handles both read-mode (direct value) and compose-mode (`getAsync()`) property access
- `getBodyMarkdown(item)` — pulls the body as HTML and converts to Markdown via Turndown
- `cleanOutlookHtml()`, `promoteOutlookTableHeaders()`, `flattenTableCellBlocks()` — pre-Turndown HTML normalisation
- `getAttendeesViaEws(itemId)` — fetches RSVP status via EWS `GetItem` SOAP call; requires `ReadWriteMailbox` permission
- `parseEwsAttendees(xmlString)` — parses the EWS XML response into the `attendeesByStatus` map
- `getAttendeesViaOfficeJs(item)` — fallback when EWS is unavailable; buckets by `appointmentResponse`
- Settings stored in `Office.context.roamingSettings` (roam with the user's Exchange account)

### [taskpane.html](taskpane.html) — UI
- Main panel: config display, Extract button, status, details summary, Open in Obsidian + Copy buttons
- Settings panel: vault name, meeting folder, email folder, internal domain, options checkboxes
- Hidden `<a id="obsidianLink">` anchor — clicked programmatically to trigger the `obsidian://` protocol handler (most reliable cross-platform approach)

### [manifest.xml](manifest.xml) — Add-in Manifest
- Add-in ID: `55267e54-b786-4113-a715-5b864213fa0b` (do NOT change after deployment)
- Version: 2.0.0.0
- Requires Mailbox 1.5+
- Permission: `ReadWriteMailbox` (originally for `makeEwsRequestAsync`; `ReadItem` would now suffice)
- Extension points: `AppointmentAttendeeCommandSurface`, `AppointmentOrganizerCommandSurface`, `MessageReadCommandSurface`
- Hosted at `https://virtualjones.github.io/outlook-to-obsidian` — forks must replace every occurrence with their own Pages URL and regenerate `<Id>`
- Validate with `npx --yes office-addin-manifest validate manifest.xml`

## Important Patterns & Conventions

### Read Mode vs Compose Mode
Office.js exposes properties differently depending on context:

| Property | Read mode | Compose mode |
|---|---|---|
| `item.subject` | `string` (direct) | `Subject` object → `.getAsync()` |
| `item.start` | `Date` (direct) | `Time` object → `.getAsync()` |
| `item.location` | `string` (direct) | `Location` object → `.getAsync()` |
| `item.requiredAttendees` | `EmailAddressDetails[]` | `Recipients` object → `.getAsync()` |
| `item.organizer` | `EmailAddressDetails` | not available |
| `item.itemId` | defined (string) | `undefined` |

`getItemProperty()` in [taskpane.js](taskpane.js) transparently handles both modes.

### EWS RSVP Status
EWS `GetItem` with `calendar:RequiredAttendees` and `calendar:OptionalAttendees` additional properties returns attendees with their `ResponseType`:

| EWS ResponseType | Bucket |
|---|---|
| `Accept` | Accepted |
| `Tentative` | Tentative |
| `Decline` | Declined |
| `None` / `NoResponseReceived` | No Response |
| `Organizer` | skipped (captured via `item.organizer`) |

EWS is only available in read mode (requires `item.itemId`). If EWS fails, the code falls back to `getAttendeesViaOfficeJs()`, which buckets attendees by `EmailAddressDetails.appointmentResponse` — this only populates when the current user **organized** the meeting; otherwise everyone lands in "No Response".

**EWS is being retired for Exchange Online (October 2026)** and most tenants already block it, so the fallback path is the common case on modern M365. `getAttendeesViaEws()` is still attempted first in read mode and typically fails with `GenericResponseError`.

### Settings
Stored in `Office.context.roamingSettings` (max 32 KB, roams with Exchange account):

| Key | Default | Description |
|---|---|---|
| `vaultName` | `Notes` | Obsidian vault name |
| `folderPath` | `02. Meeting Notes` | Base folder for meeting notes |
| `emailFolderPath` | `04. Email Notes` | Base folder for email notes |
| `internalDomain` | `example.com` | Email domain treated as internal |
| `copyToClipboard` | `true` | Auto-copy note after extracting |
| `includeExternal` | `false` | Prefix title with `[EXTERNAL]` |

Code reads the fallback from `DEFAULT_SETTINGS`, never a duplicated literal.

### Data Formatting
- **File path**: `folderPath|emailFolderPath/YYYY/MM/YYYY-MM-DD Title.md`
- **Safe filenames**: `replace(/[<>:"/\\|?*]/g, '-')`
- **Line endings**: CRLF (`\r\n`)
- **Obsidian URI**: `obsidian://new?vault=NAME&file=PATH&content=ENCODED`

### Note Structure
Generated in [taskpane.js](taskpane.js):
```markdown
---
tags: meeting
date: YYYY-MM-DD
type: outlook-meeting
external: true/false
attendees: ["Name 1", "Name 2"]
summary:
---

# Meeting Title

## Meeting Details
**Date/Time:** ...
**Location:** ...
**Total Attendees:** N

### Attendees by RSVP Status
**Organizer:** ...
**Accepted (N):** ...
**Tentative (N):** ...
**No Response (N):** ...

### Meeting Description
...

## Notes


## Action Items


## Follow-up

```

## Deployment

The manifest already points at `https://virtualjones.github.io/outlook-to-obsidian` (GitHub Pages, `main` branch, root).

### Option A — GitHub Pages (current setup; also the path for forks)
1. Enable Pages on the repo (Settings → Pages → Deploy from `main` / root)
2. If forking: replace every `https://virtualjones.github.io/outlook-to-obsidian` in [manifest.xml](manifest.xml)
   with `https://<your-username>.github.io/<repo-name>`, and generate a fresh `<Id>` (`[guid]::NewGuid()`)
3. Sideload via Outlook → Get Add-ins → My Add-ins → Add a custom add-in → Add from file

### Option B — Microsoft 365 Admin Centre (organisation-wide)
1. Host files on any HTTPS server (update the URLs in manifest.xml)
2. admin.microsoft.com → Settings → Integrated apps → Upload custom app → Upload manifest XML

### Option C — Local development
```bash
# Install dev certificate (once)
npx office-addin-dev-certs install

# Serve the folder
npx http-server . -S -C ~/.office-addin-dev-certs/localhost.crt -K ~/.office-addin-dev-certs/localhost.key -p 3000

# In manifest.xml: point the URLs at https://localhost:3000
# Sideload via Outlook ribbon → Get Add-ins → My Add-ins → Add from file
```

## Common Tasks

### Changing the Internal Email Domain
Update `internalDomain` default in `DEFAULT_SETTINGS` in [taskpane.js](taskpane.js) — or simply configure it in the Settings panel at runtime.

### Modifying the Note Template
Edit the `noteLines` array in `extractMeetingDetails()` in [taskpane.js](taskpane.js). Maintain YAML frontmatter format and CRLF line endings.

### Adding New Settings
1. Add to `DEFAULT_SETTINGS` in [taskpane.js](taskpane.js)
2. Add a form field to the settings panel in [taskpane.html](taskpane.html)
3. Read/write it in `loadSettings()` and `onSaveSettings()` in [taskpane.js](taskpane.js)

### Debugging
- Open the add-in task pane → right-click → Inspect (opens DevTools attached to the WebView2)
- On Outlook Web: F12 in the browser, then find the task pane iframe
- Check `Office.context.mailbox.item` in the console to inspect the live meeting object
- Check `result.error` in EWS callback for SOAP-level errors

## Critical Integration Points

### Office.js API Surface Used
```javascript
Office.context.mailbox.item.subject           // string | Subject
Office.context.mailbox.item.start             // Date | Time
Office.context.mailbox.item.end               // Date | Time
Office.context.mailbox.item.location          // string | Location
Office.context.mailbox.item.organizer         // EmailAddressDetails (read only)
Office.context.mailbox.item.requiredAttendees // EmailAddressDetails[] | Recipients
Office.context.mailbox.item.optionalAttendees // EmailAddressDetails[] | Recipients
Office.context.mailbox.item.from || .sender   // message mode
Office.context.mailbox.item.to / .cc          // EmailAddressDetails[]
Office.context.mailbox.item.dateTimeCreated   // sent time (falls back to dateTimeModified)
Office.context.mailbox.item.body.getAsync()   // Office.CoercionType.Html (not text)
Office.context.mailbox.item.itemId            // string (read mode only)
Office.context.mailbox.item.itemType          // Appointment | Message
Office.context.mailbox.makeEwsRequestAsync()  // EWS SOAP for RSVP status
Office.context.roamingSettings                // persistent settings storage
```

### Obsidian URI Construction
- Pattern: `obsidian://new?vault=NAME&file=PATH/FILE&content=ENCODED_NOTE`
- Triggered by programmatic `.click()` on a hidden `<a>` element
- Works in Outlook Desktop (WebView2) and Outlook Web (browser tab)

## Code Style & Principles
- **No DOM scraping** — use only Office.js API
- **Graceful EWS fallback** — if EWS fails, show attendees without RSVP status (do not fail)
- **Read/compose transparency** — `getItemProperty()` handles both modes; extraction code should not branch on mode unnecessarily
- **Roaming settings** — always use `Office.context.roamingSettings`, never `localStorage` (localStorage doesn't roam across devices)
- **Protocol handler** — use the hidden anchor click pattern for `obsidian://`, not `window.location.href` (sandboxed iframes block navigation)
