# Outlook to Obsidian Meeting Notes

A Chrome extension that extracts meeting details from Outlook Web and creates beautifully formatted notes in Obsidian with a single click.

## Features

- **One-Click Extraction** - Extract meeting title, date/time, location, attendees, and description
- **Automatic Attendee Collection** - Scrolls through attendee lists to capture all participants with RSVP status
- **Smart Organization** - Automatically organizes notes in `Year/Month` folder structure
- **External Meeting Detection** - Identifies and flags meetings with external attendees
- **RSVP Status Tracking** - Groups attendees by Accepted, Tentative, Declined, and No Response
- **Clipboard Integration** - Optionally copies note to clipboard for quick pasting
- **Direct Obsidian Integration** - Opens note directly in Obsidian via URI scheme

## Installation

### From Source

1. Clone or download this repository
2. Open Chrome/Edge and navigate to `chrome://extensions/`
3. Enable **Developer mode** (toggle in top-right)
4. Click **Load unpacked**
5. Select the extension directory

### First-Time Setup

1. Click the extension icon in your browser toolbar
2. Click the settings (⚙️) button
3. Configure:
   - **Vault Name**: Your Obsidian vault name (case-sensitive)
   - **Folder Path**: Base folder for meeting notes (e.g., `Meeting Notes` or `02. Meeting Notes`)
   - **Options**: Choose auto-close, clipboard copy, and external highlighting preferences
4. Click **Save Settings**

## Usage

### Basic Workflow

1. Navigate to a meeting invitation in Outlook Web (outlook.office.com)
2. Open the meeting to view its details
3. Click the extension icon
4. Click **Extract Meeting Details**
5. Review the extracted information
6. Click **Create in Obsidian**

The note will be created in: `VaultName/FolderPath/YYYY/MM/YYYY-MM-DD Title.md`

### Example Note Structure

```markdown
---
tags: meeting
date: 2025-01-21
type: outlook-meeting
external: false
attendees: ["John Doe", "Jane Smith", "Bob Johnson"]
summary:
---

# Weekly Team Sync

## Meeting Details
**Date/Time:** Tuesday, January 21, 2025, 2:00 PM - 3:00 PM
**Location:** Conference Room A
**Total Attendees:** 8
**External Meeting:** No

### Attendees by RSVP Status

**Organizer:** John Doe

**Accepted (5):** Alice Brown, Bob Johnson, Carol White, David Lee, Emma Davis

**Tentative (2):** Frank Miller, Grace Wilson

**No Response (1):** Henry Taylor

### Meeting Description
Discuss project milestones and upcoming deliverables for Q1.

## Notes


## Action Items


## Follow-up

```

## Configuration Options

### Vault Settings

| Setting | Description | Example |
|---------|-------------|---------|
| **Vault Name** | Exact name of your Obsidian vault | `Syntax`, `Work`, `Personal` |
| **Folder Path** | Base directory for meetings | `Meeting Notes`, `02. Meeting Notes`, `Work/Meetings` |

### Behavior Options

| Option | Default | Description |
|--------|---------|-------------|
| **Auto Close** | ✓ | Automatically closes popup after creating note |
| **Copy to Clipboard** | ✓ | Always copies note content to clipboard |
| **Highlight External** | ✗ | Adds `[EXTERNAL]` prefix to meetings with non-company attendees |

## How It Works

### Architecture

```
┌─────────────┐          ┌──────────────┐          ┌──────────────┐
│  Popup UI   │ ────────>│Content Script│ ────────>│ Outlook DOM  │
│ (popup.js)  │ inject   │(content.js)  │ scrape   │              │
└─────────────┘          └──────────────┘          └──────────────┘
       │                         │
       │                         │
       ▼                         ▼
┌─────────────┐          ┌──────────────┐
│  Settings   │          │  Extraction  │
│ (options.js)│          │   Engine     │
└─────────────┘          └──────────────┘
       │                         │
       │                         │
       ▼                         ▼
┌─────────────────────────────────────────┐
│         Obsidian URI Handler            │
│  obsidian://new?vault=...&file=...      │
└─────────────────────────────────────────┘
```

### Data Flow

1. **User Action**: User clicks "Extract Meeting Details" in popup
2. **Script Injection**: Popup checks if content script is loaded, injects if needed
3. **Message Passing**: Popup sends `extractMeeting` message to content script
4. **DOM Scraping**: Content script:
   - Extracts meeting metadata (title, date, location)
   - Scrolls through attendee list to capture all participants
   - Determines RSVP statuses and external attendees
   - Formats data into markdown note structure
5. **Response**: Content script returns JSON payload with extracted data
6. **URI Construction**: Popup builds Obsidian URI with encoded note content
7. **Obsidian Launch**: Browser opens `obsidian://` URI, creating note

### Extraction Process

#### Meeting Metadata
- **Title**: Uses fallback selectors for Outlook's dynamic UI
- **Date/Time**: Extracts from multiple possible locations, parses into ISO format
- **Location**: Searches input fields and button labels
- **Organizer**: Finds persona element with name and email

#### Attendee Collection
The extension employs a sophisticated scrolling mechanism:

1. Locates scrollable container in Outlook's attendee list
2. Scrolls incrementally (80% of viewport height per iteration)
3. Processes visible elements after each scroll
4. Tracks RSVP status headers (Accepted, Tentative, etc.)
5. Extracts name and email for each attendee
6. Deduplicates entries
7. Stops when no new attendees found (max 50 scroll attempts)

#### External Detection
Attendees and organizers are checked against `@syntax.com` domain. Any other domain triggers the `external: true` flag.

## File Structure

```
outlook-to-obsidian/
├── manifest.json          # Extension configuration and permissions
├── popup.html             # Extension popup interface
├── popup.js               # Popup logic and orchestration
├── options.html           # Settings page UI
├── options.js             # Settings management
├── content.js             # Outlook DOM scraper
├── icon16.png             # Extension icon (16x16)
├── icon48.png             # Extension icon (48x48)
├── icon128.png            # Extension icon (128x128)
├── README.md              # This file
└── CLAUDE.md              # AI assistant instructions
```

## Code Overview

### [manifest.json](manifest.json)
Defines extension metadata, permissions, and content script registration.

**Key Permissions:**
- `activeTab` - Access current tab
- `clipboardWrite` - Copy to clipboard
- `storage` - Save user settings
- `scripting` - Inject content script dynamically
- `host_permissions` - Access Outlook domains

### [popup.js](popup.js) (206 lines)
Main orchestration logic for the extension.

**Key Functions:**
- `loadAndDisplaySettings()` - Loads saved configuration
- `ensureContentScript(tabId)` - Injects content script if not present
- `extractBtn.addEventListener()` - Handles extraction workflow
- `createObsidianBtn.addEventListener()` - Builds and opens Obsidian URI
- `copyToClipboard(text)` - Copies note to clipboard

### [content.js](content.js) (402 lines)
Core scraping engine that runs on Outlook Web pages.

**Main Function:** `extractMeetingDetails()`

**Extraction Steps:**
1. Title extraction (lines 17-40)
2. Date/time parsing (lines 42-87)
3. Location detection (lines 89-108)
4. Organizer identification (lines 110-127)
5. Meeting body/description (lines 129-145)
6. Attendee collection with scrolling (lines 147-309)
7. External attendee detection (lines 311-313)
8. Note formatting (lines 315-378)

**Selector Strategy:**
Uses arrays of fallback selectors to handle Outlook UI variations:
```javascript
let titleSelectors = [
  'span.CWGkB',
  'div.FZzLA span',
  'span[class*="CWGkB"]',
  // ... more fallbacks
];
```

### [options.js](options.js) (109 lines)
Settings page logic with storage management.

**Key Features:**
- Default settings definition
- Preset buttons for common configurations
- Live folder structure preview
- Settings validation and save/reset

**Default Settings:**
```javascript
{
  vaultName: 'Syntax',
  folderPath: '02. Meeting Notes',
  autoClose: true,
  copyToClipboard: true,
  includeExternal: false
}
```

## Development

### Prerequisites
- Chrome or Edge browser
- Obsidian desktop app installed
- Node.js (optional, for future build tooling)

### Development Workflow

1. Make changes to source files
2. Go to `chrome://extensions`
3. Click reload icon for this extension
4. Refresh Outlook Web page
5. Test via popup interface

### Debugging

**Content Script:**
1. Open DevTools on Outlook Web page
2. Check Console tab for content script logs
3. Use Sources tab to set breakpoints in `content.js`

**Popup:**
1. Right-click extension icon → Inspect popup
2. View console logs and debug `popup.js`

**Extension Background:**
1. Go to `chrome://extensions`
2. Click "Inspect views" under extension

### Testing Checklist

- [ ] Extract meeting with all fields populated
- [ ] Extract meeting with missing fields (no location, no body)
- [ ] Extract meeting with 50+ attendees (scrolling test)
- [ ] Extract meeting with external attendees
- [ ] Extract meeting with various RSVP statuses
- [ ] Verify Obsidian URI opens correctly
- [ ] Test clipboard copy functionality
- [ ] Verify folder structure (Year/Month)
- [ ] Test settings save/load
- [ ] Test on different Outlook UI variations

## Troubleshooting

### Extension Not Working

**Issue**: "Please navigate to an Outlook Web meeting invitation first"
- **Solution**: Ensure you're on `outlook.office.com`, `outlook.office365.com`, or `outlook.live.com`
- **Solution**: Make sure you've opened a specific meeting (not just calendar view)

**Issue**: "Could not access the page. Please refresh and try again."
- **Solution**: Refresh the Outlook page and try again
- **Solution**: Reload the extension in `chrome://extensions`

**Issue**: "Failed to extract meeting details"
- **Solution**: Outlook's UI may have changed - selectors might need updating
- **Solution**: Check console logs in DevTools for specific errors

### Obsidian Not Opening

**Issue**: Clicking "Create in Obsidian" does nothing
- **Solution**: Ensure Obsidian desktop app is installed
- **Solution**: Verify vault name exactly matches (case-sensitive)
- **Solution**: Check that folder path doesn't contain invalid characters

**Issue**: Note created in wrong location
- **Solution**: Verify folder path in settings
- **Solution**: Check that year/month folders exist or can be created

### Missing Attendees

**Issue**: Not all attendees extracted
- **Solution**: The scrolling logic should handle this automatically
- **Solution**: If persistent, check `maxScrollAttempts` in [content.js:187](content.js#L187)
- **Solution**: Manually scroll down before extraction as a workaround

### External Detection Not Working

**Issue**: External attendees not flagged
- **Solution**: Check if company domain matches hardcoded `@syntax.com` in [content.js:237](content.js#L237) and [content.js:299](content.js#L299)
- **Solution**: Update domain check if using different email domain

## Customization

### Changing Note Template

Edit the `noteLines` array in [content.js:343-376](content.js#L343-L376):

```javascript
let noteLines = [
  '---',
  'tags: meeting',
  `date: ${meetingDate.full}`,
  'type: outlook-meeting',
  // ... customize frontmatter
  '---',
  '',
  `# ${title}`,
  // ... customize body
];
```

### Adding New Settings

1. Add to `DEFAULT_SETTINGS` in [options.js:2-8](options.js#L2-L8)
2. Add UI element in [options.html](options.html)
3. Add save/load logic in [options.js](options.js)
4. Use setting in [popup.js](popup.js) or [content.js](content.js)

### Changing Company Domain

Update email checks in [content.js](content.js):
- Line 237: `if (email && !email.endsWith('@syntax.com'))`
- Line 299: `if (!email.endsWith('@syntax.com'))`

Replace `@syntax.com` with your company domain.

### Modifying Folder Structure

Change folder path construction in [popup.js:181](popup.js#L181):

```javascript
// Current: folderPath/YYYY/MM
const folderPath = `${currentSettings.folderPath}/${extractedData.meetingDate.year}/${extractedData.meetingDate.month}`;

// Custom: folderPath/YYYY-QQ
const quarter = Math.ceil(parseInt(extractedData.meetingDate.month) / 3);
const folderPath = `${currentSettings.folderPath}/${extractedData.meetingDate.year}/Q${quarter}`;
```

## Privacy & Security

- **Local Only**: All processing happens locally in your browser
- **No Data Sent**: No meeting data is sent to external servers
- **Permissions**: Only accesses Outlook domains with explicit user action
- **Storage**: Settings stored in Chrome's sync storage (encrypted by Chrome)
- **Clipboard**: Optional - can be disabled in settings

## Browser Compatibility

- ✅ Chrome (Manifest V3)
- ✅ Edge (Chromium-based)
- ✅ Brave (Chromium-based)
- ❌ Firefox (uses different manifest format)
- ❌ Safari (uses different extension system)

## Known Limitations

- Only works with Outlook Web (not desktop Outlook)
- Requires Obsidian desktop app (not mobile)
- Scraping may break if Outlook significantly changes UI
- Maximum 50 scroll attempts for attendee collection
- Company domain hardcoded to `@syntax.com`

## Future Enhancements

Potential improvements for future versions:

- [ ] Support for multiple company domains
- [ ] Configurable note templates
- [ ] Meeting recurrence pattern extraction
- [ ] Attachment detection and links
- [ ] Teams meeting link extraction
- [ ] Custom tag support
- [ ] Batch extraction for multiple meetings
- [ ] Export to other formats (CSV, JSON)

## Contributing

Contributions welcome! To contribute:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

Please maintain existing code style and add comments for complex logic.

## License

This project is open source and available under the MIT License.

## Support

For issues, questions, or feature requests:
- Check the [Troubleshooting](#troubleshooting) section
- Review existing code comments
- Open an issue in the repository

## Changelog

### Version 1.1.0
- Added external meeting detection
- Improved attendee scrolling logic
- Added settings page with presets
- Enhanced error handling

### Version 1.0.0
- Initial release
- Basic meeting extraction
- Obsidian integration
- RSVP status tracking

## Acknowledgments

Built for seamless integration between Outlook Web and Obsidian note-taking workflows.
