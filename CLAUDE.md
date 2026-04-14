# Claude Instructions — Outlook to Obsidian Extension

## Project Overview
This is a Manifest V3 Chrome extension that extracts meeting details from Outlook Web and creates formatted Obsidian notes via the `obsidian://` URI scheme.

**Architecture:**
- **Popup UI** ([popup.html](popup.html) / [popup.js](popup.js)) — main interface for triggering extraction
- **Options Page** ([options.html](options.html) / [options.js](options.js)) — user settings configuration
- **Content Script** ([content.js](content.js)) — DOM scraper running on Outlook pages
- **Manifest** ([manifest.json](manifest.json)) — extension config, permissions, content script registration

## Key Files & Responsibilities

### [content.js](content.js) - Core Scraper
- Contains `extractMeetingDetails()` function with all scraping logic
- Uses fallback selector arrays for robustness (e.g., `titleSelectors`, `timeSelectors`)
- Implements scrolling extraction for attendee lists
- Returns structured JSON payload with meeting data
- Message listener responds to `ping` and `extractMeeting` actions

### [popup.js](popup.js) - Orchestration
- Injects content script dynamically if needed via `chrome.scripting.executeScript`
- Sends messages to content script and handles responses
- Constructs Obsidian URI: `obsidian://new?vault=...&file=...&content=...`
- Applies `[EXTERNAL]` prefix to title if configured
- Copies note to clipboard based on settings

### [options.js](options.js) - Settings Management
- Manages `chrome.storage.sync` with keys: `vaultName`, `folderPath`, `autoClose`, `copyToClipboard`, `includeExternal`
- Defines `DEFAULT_SETTINGS` — update here when adding new settings

## Important Patterns & Conventions

### Scraping Patterns
- **Always preserve fallback selector arrays** — order matters, first match wins
- **Scrolling logic** — attendees extracted via scroll loop in `content.js:179-266`
- **External detection** — any email NOT ending with `@syntax.com` triggers `hasExternalAttendees`
- **RSVP grouping** — attendees stored in `attendeesByStatus` Map with statuses: Accepted, Tentative, Declined, No Response

### Data Formatting
- **File paths**: `folderPath/YYYY/MM/YYYY-MM-DD Title.md`
- **Safe filenames**: `replace(/[<>:"/\\|?*]/g, '-')`
- **Line endings**: CRLF (`\r\n`) used in note content
- **Date format**: ISO-like `YYYY-MM-DD`

### Note Structure
Generated in [content.js:343-376](content.js#L343-L376):
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

### Meeting Description
...

## Notes


## Action Items


## Follow-up

```

## Common Tasks

### Changing Outlook Selectors
1. Inspect Outlook DOM to identify new selectors
2. Update appropriate selector array in [content.js](content.js) (`titleSelectors`, `timeSelectors`, etc.)
3. Preserve fallback order — add new selectors but keep existing ones for compatibility
4. Test across different Outlook page types (new meeting, existing meeting, invite)

### Modifying Note Template
- Edit the `noteLines` array in [content.js:343-376](content.js#L343-L376)
- Maintain YAML frontmatter format
- Use CRLF (`\r\n`) for line endings
- Test with `[EXTERNAL]` prefix feature

### Adding New Settings
1. Add to `DEFAULT_SETTINGS` in [options.js](options.js)
2. Update UI in [options.html](options.html)
3. Add save/load logic in [options.js](options.js)
4. Use setting in [popup.js](popup.js) or [content.js](content.js) as needed
5. Consider backward compatibility with existing users

### Debugging Extraction Issues
1. Open Chrome DevTools on Outlook page (for content script logs)
2. Right-click extension icon → Inspect popup (for popup logs)
3. Check `chrome.runtime.lastError` in message passing
4. Verify content script injection in [popup.js:38-66](popup.js#L38-L66)

## Critical Integration Points

### Message Passing
```javascript
// popup.js → content.js
chrome.tabs.sendMessage(tabId, {action: 'extractMeeting'}, callback)

// content.js response format
{
  success: true/false,
  data: {
    title, time, location, totalAttendees,
    hasExternalAttendees, meetingDate, safeTitle, note
  },
  error: "..." // if success: false
}
```

### Obsidian URI Construction
- Pattern: `obsidian://new?vault=NAME&file=PATH/FILE&content=ENCODED_NOTE`
- Vault and folder must match user's Obsidian setup
- All parameters must be URI-encoded
- Folder structure: `folderPath/YYYY/MM`
- Generated in [popup.js:177-197](popup.js#L177-L197)

### Content Script Injection
- Declared in [manifest.json](manifest.json) for automatic injection
- Manual injection via [popup.js:38-66](popup.js#L38-L66) if not loaded
- Uses ping/pong pattern to check if already injected

## Development Workflow

1. **Edit** files in this directory
2. **Reload extension** in `chrome://extensions` (Developer mode must be enabled)
3. **Refresh** Outlook page to reload content script
4. **Test** via extension popup
5. **Debug** using DevTools on both Outlook page and popup

**No build system** — changes to JS/HTML files take effect after reload.

## Safety & Permissions

- **Host permissions**: Only Outlook domains (`outlook.office.com`, `outlook.office365.com`, `outlook.live.com`)
- **Storage**: Uses `chrome.storage.sync` — preserve existing key names for compatibility
- **Clipboard**: `clipboardWrite` permission used for auto-copy feature
- **Scripting**: `scripting` permission required for dynamic content script injection

## Code Style & Principles

- **No over-engineering** — keep changes minimal and focused
- **Preserve robustness** — maintain selector fallbacks and error handling
- **Test manually** — no automated tests exist, validate all changes in browser
- **Backward compatible** — don't break existing user settings or workflows
- **Security aware** — sanitize all user-controlled strings in filenames and URIs

## Common Gotchas

1. **Outlook DOM changes frequently** — selectors may break with UI updates
2. **Scrolling required** — attendee list is virtualized, must scroll to extract all
3. **Content script timing** — may not be loaded when popup opens (hence injection logic)
4. **CRLF line endings** — Windows-style required for note formatting
5. **Email domain hardcoded** — `@syntax.com` is treated as internal, update in [content.js:237](content.js#L237) and [content.js:299](content.js#L299) if needed
6. **URI encoding** — must encode Obsidian URI parameters or vault creation will fail
7. **Year/month folders** — created automatically in Obsidian if they don't exist

## When Making Changes

**Before modifying:**
- Read the relevant file(s) completely
- Understand the existing pattern
- Check if similar logic exists elsewhere

**When testing:**
- Test with internal-only meetings
- Test with external attendees
- Test with missing fields (no location, no description)
- Test with very long titles and attendee lists
- Verify Obsidian URI opens correctly and creates note in right folder

**After changes:**
- Update this file if architectural changes made
- Consider edge cases
- Manually test the full flow: extract → copy → create in Obsidian
