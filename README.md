# Send to Obsidian — Outlook Add-in

An Outlook Add-in that extracts the current calendar appointment or email message into a formatted Markdown note and opens it directly in Obsidian via the `obsidian://` URI scheme.

Originally a Chrome extension, this is now a cross-platform **Office Add-in** that runs anywhere Outlook does: Outlook on the Web, New Outlook for Windows, Classic Outlook Desktop (Windows & Mac).

## Features

- **One-click extraction** for either calendar appointments or email messages
- **Auto-detects context** — the same task pane adapts based on whether you opened a meeting or an email
- **Markdown body conversion** — email/meeting bodies converted from HTML to clean Markdown (tables, lists, blockquotes, links, formatting preserved)
- **Folder routing** — meetings go to one folder, emails to another; year/month subfolders created automatically
- **RSVP status** for meetings you organize (Accepted / Tentative / Declined / No Response)
- **External attendee/sender detection** based on a configurable internal email domain
- **Roaming settings** — preferences sync with your Microsoft 365 account across devices
- **Direct Obsidian launch** via the `obsidian://new` URI handler

## Architecture

| File | Role |
|---|---|
| [manifest.xml](manifest.xml) | Office Add-in manifest — declares ribbon buttons and rules |
| [taskpane.html](taskpane.html) | Task pane UI (main panel + settings panel) |
| [taskpane.js](taskpane.js) | Extraction logic, settings, Markdown conversion, URI building |
| [commands.html](commands.html) | Placeholder function file (required by manifest) |
| `icon{16,48,128}.png` | Ribbon icons |

External dependencies (loaded at runtime via CDN):

- **[Office.js](https://appsforoffice.microsoft.com/lib/1/hosted/office.js)** — Outlook API
- **[Turndown 7.2.0](https://github.com/mixmark-io/turndown)** — HTML → Markdown converter
- **[@joplin/turndown-plugin-gfm](https://github.com/laurent22/joplin-turndown-plugin-gfm)** — GitHub-Flavored Markdown extensions (tables, strikethrough, task lists)

## Installation

The add-in is hosted via **GitHub Pages** at `https://virtualjones.github.io/outlook-to-obsidian/` and sideloaded into Outlook by uploading the `manifest.xml` file.

### Per-user sideload

1. Download [manifest.xml](manifest.xml) from this repo (or clone the repo).
2. Open Outlook → **Get Add-ins** → **My add-ins** → **Custom Addins** → **Add a custom add-in → Add from file**.
3. Select `manifest.xml` and accept the warning.

Alternatively use [aka.ms/olksideload](https://aka.ms/olksideload) which routes you to the same dialog.

### Org-wide deployment (admin)

Upload `manifest.xml` at https://admin.microsoft.com → Settings → Integrated apps → Upload custom app.

### Self-hosting

If you fork this repo:

1. Enable GitHub Pages on your fork (Settings → Pages → Deploy from `main` / root).
2. Replace all `https://virtualjones.github.io/outlook-to-obsidian` URLs in `manifest.xml` with your own Pages URL.
3. Generate a new GUID for the `<Id>` element (PowerShell: `[guid]::NewGuid()`).
4. Sideload the updated manifest as above.

## Usage

### Capturing a meeting

1. Open a calendar appointment.
2. Click **Obsidian → Create Note** in the ribbon.
3. Click **Extract Meeting Details** in the task pane.
4. Click **Open in Obsidian** (or **Copy** to paste manually).

Output location: `<Vault>/<MeetingNotesFolder>/YYYY/MM/YYYY-MM-DD <Title>.md`

### Capturing an email

1. Open an email message.
2. Click **Obsidian → Create Note** in the ribbon.
3. Click **Extract Email Details** in the task pane.
4. Click **Open in Obsidian**.

Output location: `<Vault>/<EmailNotesFolder>/YYYY/MM/YYYY-MM-DD <Subject>.md`

## Note templates

### Meeting note

```markdown
---
tags: meeting
date: 2026-06-15
type: outlook-meeting
external: false
attendees: ["Alice Brown", "Bob Johnson"]
summary:
---

# Weekly Team Sync

## Meeting Details
**Date/Time:** 6/15/2026 14:00 – 15:00
**Location:** Conference Room A
**Total Attendees:** 8

### Attendees by RSVP Status
**Organizer:** Mike Jones
**Accepted (5):** Alice Brown, Bob Johnson, Carol White, David Lee, Emma Davis
**Tentative (2):** Frank Miller, Grace Wilson
**No Response (1):** Henry Taylor

### Meeting Description
Discuss Q2 deliverables…

## Notes

## Action Items

## Follow-up
```

### Email note

```markdown
---
tags: email
date: 2026-06-15
type: outlook-email
from: "Luke Kremer <luke.kremer@syntax.com>"
external: false
summary:
---

# Re: Project Status

## Email Details
**From:** Luke Kremer <luke.kremer@syntax.com>
**Sent:** 6/15/2026 09:42
**To (3):** Leonardo De Araujo, Rebecca Murray, Sarah Abikhzir
**Cc (1):** Antonella Costanzo

### Body
Body content converted from HTML to Markdown, with
**bold**, *italic*, [links](https://example.com), tables,
and reply chains as > blockquotes.

## Notes

## Action Items

## Follow-up
```

## Settings

Open the task pane → ⚙️ in the top right. Settings roam with your Microsoft 365 account.

| Setting | Default | Description |
|---|---|---|
| Vault Name | `Syntax` | Case-sensitive Obsidian vault name |
| Meeting Notes Folder | `02. Meeting Notes` | Base folder for meeting notes; `YYYY/MM` subfolders are added automatically |
| Email Notes Folder | `04. Email Notes` | Base folder for email notes |
| Internal Email Domain | `syntax.com` | Anyone with another domain is flagged as external |
| Copy to clipboard | ✓ | Auto-copy the note after extracting |
| Prefix `[EXTERNAL]` | ✗ | Prepend `[EXTERNAL]` to titles when external participants are present |

## Key implementation notes

### RSVP detection

Meeting RSVPs come from `EmailAddressDetails.appointmentResponse` on each attendee. This only populates when you're the **organizer** of the meeting. Attendees viewing someone else's meeting cannot see other attendees' responses.

(Historically the add-in used the EWS `GetItem` SOAP call for RSVPs, but Microsoft is retiring EWS for Exchange Online in October 2026 and most tenants already block it. The EWS code path remains as a fallback but is expected to fail with `GenericResponseError` on modern M365 tenants.)

### HTML → Markdown body conversion

Bodies are pulled as HTML (plain-text coercion collapses line breaks in New Outlook) and run through Turndown plus the GFM plugin. Outlook-specific cruft (`<o:p>`, MSO conditional comments, inline `<style>`, embedded `cid:` images) is stripped before conversion.

### Permissions

`ReadWriteMailbox` is requested in the manifest. The high-permission level was originally for EWS access; with EWS retired, `ReadItem` would now suffice, but the higher permission is harmless.

### Read vs Compose mode

Office.js exposes properties differently depending on whether you're reading a saved item or composing a new one. The `getItemProperty()` helper transparently handles both — see [CLAUDE.md](CLAUDE.md) for details.

## Troubleshooting

### Sideload fails

- Confirm the Pages URLs in `manifest.xml` are reachable in a browser (`/taskpane.html`, `/commands.html`, `/icon128.png`).
- Remove any prior partial install under **My add-ins** before retrying.
- Restart Outlook after removing.

### "RSVPs all show as No Response"

You're either:
- Viewing a meeting you didn't organize (Microsoft platform limitation), or
- The meeting genuinely has no responses yet.

### "Obsidian doesn't open"

- Obsidian desktop must be installed and the vault name must match exactly (case-sensitive).
- Verify the vault is open at least once so it's registered with Obsidian's URI handler.

### Manifest validation

```powershell
npx --yes office-addin-manifest validate manifest.xml
```

### Task pane DevTools

- **Outlook on the Web / Classic Desktop**: right-click in the task pane → Inspect.
- **New Outlook for Windows**: open Edge → `edge://inspect/#devices` → find the task pane and click **inspect**.

## Limitations

- **Right-click context menu**: Office Add-ins cannot add to Outlook's right-click menu. The add-in only appears via the ribbon. ("Send to OneNote" is a built-in Outlook feature, not an add-in.)
- **RSVPs for attendee-side views**: see RSVP note above.
- **Mobile**: not supported. Office.js Mailbox 1.5 features and the `obsidian://` URI handler aren't available in iOS/Android Outlook.
- **Body conversion accuracy**: Turndown is excellent but Outlook produces some unusual HTML (especially in reply chains). Edge cases may occasionally render imperfectly.

## License

MIT.
