# Copilot instructions — outlook-to-obsidian

Purpose: help AI coding agents immediately understand and safely modify this Chrome extension that extracts Outlook Web meeting data and creates Obsidian notes.

- **Big picture:** This is a Manifest V3 Chrome extension. The UI consists of a popup (`popup.html` / `popup.js`) and an options page (`options.html` / `options.js`). A content script (`content.js`) runs on Outlook Web domains defined in `manifest.json` and is responsible for scraping meeting details. The popup injects/communicates with the content script and then opens Obsidian via an `obsidian://new` URI.

- **Key files:**
  - [manifest.json](manifest.json) — extension metadata, permissions and content script registration.
  - [content.js](content.js) — core scraper. Look for the `extractMeetingDetails()` flow and the `chrome.runtime.onMessage` listener.
  - [popup.js](popup.js) — orchestrates injection, messaging, settings, clipboard copy, and constructs the Obsidian URI.
  - [options.js](options.js) — default settings and storage keys (`vaultName`, `folderPath`, `autoClose`, `copyToClipboard`, `includeExternal`).

- **Data flow & service boundaries:**
  - The popup calls `chrome.tabs.sendMessage` to request extraction. If the content script is not present, `popup.js` uses `chrome.scripting.executeScript` to inject `content.js` at runtime.
  - `content.js` returns a single JSON payload with `data` (title, note, meetingDate, totalAttendees, hasExternalAttendees, safeTitle, etc.). The popup then optionally modifies title based on `includeExternal` and opens Obsidian with an encoded URI.

- **Important implementation patterns to preserve:**
  - Robust scrapers: `content.js` uses arrays of selector fallbacks (`titleSelectors`, `bodySelectors`, etc.). Preserve the fallback order when changing selectors.
  - Scrolling extraction: attendees are collected by scrolling a scrollable container; see the `scrollContainer` / `maxScrollAttempts` logic. Avoid breaking this when refactoring DOM traversal.
  - Attendee grouping: attendees are stored in a Map per RSVP status (`attendeesByStatus`) and later formatted into the note.
  - Settings come from `chrome.storage.sync` and defaults are defined in `options.js` (`DEFAULT_SETTINGS`).
  - Filenames and note text use `safeTitle` sanitation (`replace(/[<>:\"/\\|?*]/g, '-')`) and note lines use CRLF (`\r\n`).

- **Conventions & project-specific quirks:**
  - External attendee detection treats `@syntax.com` as internal; `hasExternalAttendees` flips if other domains are present.
  - Date format in notes is ISO-like (`YYYY-MM-DD`) and folder structure uses `folderPath/YYYY/MM` — `popup.js` constructs this.
  - Note content and metadata (YAML-like block) are produced inside `content.js`; editing that output is the main place to change note templates.

- **Dev workflow (no build system):**
  1. Edit JS/HTML files in this folder.
  2. Open `chrome://extensions`, enable Developer mode, click Reload for the unpacked extension (or press the reload icon).
  3. Refresh the Outlook Web page and use the popup. For rapid iteration, `popup.js` tries to inject `content.js` if needed.
 4. Debugging: inspect the Outlook tab (DevTools) for `content.js` logs. Inspect the popup via `chrome://extensions` → Inspect views → `popup.html` for popup console logs.

- **Safety & permissions:**
  - Manifest lists `host_permissions` for Outlook domains; add any new host requiring scraping to `manifest.json` and re-load the extension.
  - `clipboardWrite` and `storage` are used; changes to storage keys must preserve existing key names to remain backward-compatible.

- **When modifying behavior, check these spots:**
  - Change scraping selectors in `content.js` only after confirming the Outlook DOM variation under test.
  - If changing message shapes, update handling in `popup.js` (the code that expects `response.success` and `response.data`).
  - If you add new settings, update `DEFAULT_SETTINGS` in `options.js` and the UI in `options.html`.

- **Integration points to be careful with:**
  - `obsidian://new` URI generation in `popup.js` (encoding and folder paths). Mistakes here can generate malformed URIs.
  - `chrome.scripting.executeScript` injection: this is used to load `content.js` at runtime and requires correct path and `tabId`.

- **No tests present:** there are no automated tests or build steps. Keep changes small and validate manually in Chrome/Edge using the flow above.

If you want, I can expand any section (examples of selectors, sample note output, or suggested unit-test strategies). Which section should I expand first?
