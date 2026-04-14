# API Documentation

Complete code reference for the Outlook to Obsidian extension.

## Table of Contents

- [Content Script API](#content-script-api)
- [Popup API](#popup-api)
- [Options API](#options-api)
- [Message Protocol](#message-protocol)
- [Data Structures](#data-structures)
- [Obsidian URI Schema](#obsidian-uri-schema)

---

## Content Script API

**File**: [content.js](content.js)
**Context**: Runs on Outlook Web pages
**Manifest Registration**: Lines 32-42 in [manifest.json](manifest.json)

### Message Listener

```javascript
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'ping') {
    sendResponse({status: 'ready'});
    return true;
  }

  if (request.action === 'extractMeeting') {
    extractMeetingDetails().then(sendResponse);
    return true;
  }
});
```

**Supported Actions:**
- `ping` - Health check to verify content script is loaded
- `extractMeeting` - Triggers meeting detail extraction

### extractMeetingDetails()

**Signature**: `async function extractMeetingDetails(): Promise<ExtractionResult>`

**Returns**: Promise resolving to extraction result object

**Process Flow**:

1. **Title Extraction** (lines 17-40)
   ```javascript
   let titleSelectors = [
     'span.CWGkB',
     'div.FZzLA span',
     'span[class*="CWGkB"]',
     'div[class*="FZzLA"] span',
     'input[aria-label*="subject" i]',
     'input[placeholder*="subject" i]'
   ];
   ```
   - Iterates through selector fallbacks
   - Returns first non-empty match
   - Defaults to `'Untitled Meeting'` if not found

2. **Date/Time Extraction** (lines 42-87)
   ```javascript
   let timeSelectors = [
     'div.y79CD',
     'div[class*="y79CD"]',
     'button[aria-label*="Start time" i]',
     'input[aria-label*="start" i][aria-label*="time" i]'
   ];
   ```
   - Extracts text from date/time elements
   - Parses date pattern: `MM/DD/YYYY`
   - Constructs `meetingDate` object with year, month, day, full
   - Defaults to current date if not found

3. **Location Extraction** (lines 89-108)
   ```javascript
   let locationSelectors = [
     'input[aria-label*="location" i]',
     'button[aria-label*="location" i]',
     'div[aria-label*="location" i]:not([class*="calendar"])',
     'span[class*="location" i]:not([class*="calendar"])',
     'input[placeholder*="location" i]'
   ];
   ```
   - Strips `"Location"` prefix if present

4. **Organizer Extraction** (lines 110-127)
   ```javascript
   let organizerElement = document.querySelector(
     'div.lpxUr.EKtKK span.fui-Persona__primaryText'
   );
   ```
   - Extracts name from persona element
   - Parses email from `aria-label` attribute
   - Email pattern: `/[\w\.-]+@[\w\.-]+\.\w+/`

5. **Body Extraction** (lines 129-145)
   ```javascript
   let bodySelectors = [
     'div[role="textbox"][aria-label*="message" i]',
     'div[role="textbox"][contenteditable="true"]',
     'div[class*="body"][class*="content"]',
     'div[aria-label*="description" i]',
     'div[aria-label*="Message body"]'
   ];
   ```
   - Uses `innerText` for better formatting

6. **Attendee Extraction with Scrolling** (lines 147-309)

   **Scroll Container Detection** (lines 158-174):
   ```javascript
   let possibleContainers = [
     document.querySelector('div[role="tree"]'),
     document.querySelector('div[class*="scrollable"]'),
     // ... dynamic overflow detection
   ];
   ```

   **Scrolling Algorithm** (lines 179-266):
   ```javascript
   const maxScrollAttempts = 50;
   while (scrollAttempts < maxScrollAttempts) {
     // Process visible elements
     // Scroll by 80% of viewport
     scrollContainer.scrollTop += scrollContainer.clientHeight * 0.8;
     await new Promise(resolve => setTimeout(resolve, 300));

     // Break if no new attendees found
     if (noChangeCount > 3) break;
   }
   ```

   **RSVP Status Tracking**:
   ```javascript
   let attendeesByStatus = {
     'Accepted': new Map(),
     'Tentative': new Map(),
     'Declined': new Map(),
     'No Response': new Map()
   };
   ```

   **Element Types**:
   - Header elements: `div.ELl7R.TqkD6` (RSVP status sections)
   - Attendee elements: `div.lpxUr.VCV1f.sKKBX` (individual attendees)

7. **External Detection** (lines 311-313)
   ```javascript
   if (organizerEmail && !organizerEmail.endsWith('@syntax.com')) {
     hasExternalAttendees = true;
   }
   ```

8. **Note Formatting** (lines 315-378)
   - Builds YAML frontmatter
   - Formats attendee lists by RSVP status
   - Uses CRLF line endings (`\r\n`)
   - Sanitizes filename: `replace(/[<>:"/\\|?*]/g, '-')`

**Return Type**:
```typescript
interface ExtractionResult {
  success: boolean;
  data?: {
    title: string;
    time: string;
    location: string;
    totalAttendees: number;
    hasExternalAttendees: boolean;
    meetingDate: {
      year: string;
      month: string;
      day: string;
      full: string;  // YYYY-MM-DD
    };
    safeTitle: string;
    note: string;
  };
  error?: string;
}
```

**Error Handling**:
```javascript
try {
  // ... extraction logic
} catch (error) {
  return {
    success: false,
    error: error.message
  };
}
```

---

## Popup API

**File**: [popup.js](popup.js)
**Context**: Extension popup window
**UI**: [popup.html](popup.html)

### Global State

```javascript
let extractedData = null;      // Cached extraction result
let currentSettings = {};      // Loaded user settings
```

### loadAndDisplaySettings()

**Signature**: `function loadAndDisplaySettings(): void`

**Purpose**: Loads settings from Chrome storage and updates UI

```javascript
chrome.storage.sync.get({
  vaultName: '',
  folderPath: '',
  autoClose: true,
  copyToClipboard: true,
  includeExternal: false
}, function(items) {
  currentSettings = items;
  // Update UI...
});
```

**UI Updates**:
- Displays vault name and folder path
- Shows/hides configuration warning
- Enables/disables extract button

### ensureContentScript(tabId)

**Signature**: `async function ensureContentScript(tabId: number): Promise<boolean>`

**Purpose**: Ensures content script is loaded, injects if necessary

**Algorithm**:
1. Send ping message to tab
2. If no response (script not loaded):
   - Inject script via `chrome.scripting.executeScript`
   - Wait 100ms for initialization
3. Return success/failure

```javascript
chrome.tabs.sendMessage(tabId, {action: 'ping'}, function(response) {
  if (chrome.runtime.lastError) {
    // Inject content script
    chrome.scripting.executeScript({
      target: {tabId: tabId},
      files: ['content.js']
    }, callback);
  } else {
    // Already loaded
    resolve(true);
  }
});
```

### Extract Button Handler

**Event**: `click` on `#extractBtn`

**Flow**:
1. Validate current tab is Outlook domain
2. Ensure content script loaded
3. Send `extractMeeting` message
4. Process response
5. Apply `[EXTERNAL]` prefix if configured
6. Update UI with details
7. Copy to clipboard if configured
8. Show "Create in Obsidian" button

**URL Validation**:
```javascript
if (!tab.url.match(/https:\/\/(outlook\.office\.com|outlook\.office365\.com|outlook\.live\.com)/)) {
  // Show error
}
```

**External Prefix Application** (lines 118-129):
```javascript
if (currentSettings.includeExternal && response.data.hasExternalAttendees) {
  extractedData.title = '[EXTERNAL] ' + extractedData.title;
  // Update note content...
}
```

### Create Obsidian Button Handler

**Event**: `click` on `#createObsidianBtn`

**Flow**:
1. Build folder path with year/month
2. Sanitize filename
3. Construct Obsidian URI
4. Open URI in new tab
5. Close popup if configured

**Folder Path Construction** (line 181):
```javascript
const folderPath = `${currentSettings.folderPath}/${extractedData.meetingDate.year}/${extractedData.meetingDate.month}`;
```

**Filename Construction** (lines 184-185):
```javascript
const safeTitle = extractedData.title.replace(/[<>:"/\\|?*]/g, '-');
const filename = `${extractedData.meetingDate.full} ${safeTitle}`;
```

**URI Construction** (line 188):
```javascript
const obsidianUri = `obsidian://new?vault=${encodeURIComponent(vaultName)}&file=${encodeURIComponent(folderPath + '/' + filename)}&content=${encodeURIComponent(note)}`;
```

### copyToClipboard(text)

**Signature**: `function copyToClipboard(text: string): void`

**Implementation**:
```javascript
function copyToClipboard(text) {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand('copy');
  document.body.removeChild(textarea);
}
```

**Note**: Uses deprecated `document.execCommand('copy')` but widely supported

---

## Options API

**File**: [options.js](options.js)
**Context**: Extension options page
**UI**: [options.html](options.html)

### Constants

```javascript
const DEFAULT_SETTINGS = {
  vaultName: 'Syntax',
  folderPath: '02. Meeting Notes',
  autoClose: true,
  copyToClipboard: true,
  includeExternal: false
};
```

### Event Listeners

**Page Load**:
```javascript
document.addEventListener('DOMContentLoaded', loadSettings);
```

**Input Changes**:
```javascript
document.getElementById('folderPath').addEventListener('input', updatePreview);
document.getElementById('vaultName').addEventListener('input', updatePreview);
```

**Preset Buttons**:
```javascript
document.querySelectorAll('[data-vault]').forEach(btn => {
  btn.addEventListener('click', function() {
    document.getElementById('vaultName').value = this.dataset.vault;
    updatePreview();
  });
});

document.querySelectorAll('[data-path]').forEach(btn => {
  btn.addEventListener('click', function() {
    document.getElementById('folderPath').value = this.dataset.path;
    updatePreview();
  });
});
```

### loadSettings()

**Signature**: `function loadSettings(): void`

**Purpose**: Loads settings from storage and populates form

```javascript
chrome.storage.sync.get(DEFAULT_SETTINGS, function(items) {
  document.getElementById('vaultName').value = items.vaultName;
  document.getElementById('folderPath').value = items.folderPath;
  document.getElementById('autoClose').checked = items.autoClose;
  document.getElementById('copyToClipboard').checked = items.copyToClipboard;
  document.getElementById('includeExternal').checked = items.includeExternal;
  updatePreview();
});
```

### saveSettings()

**Signature**: `function saveSettings(): void`

**Validation**:
- Vault name must not be empty
- Folder path must not be empty

**Process**:
1. Collect form values
2. Trim text inputs
3. Validate required fields
4. Save to `chrome.storage.sync`
5. Show success/error status

```javascript
const settings = {
  vaultName: document.getElementById('vaultName').value.trim(),
  folderPath: document.getElementById('folderPath').value.trim(),
  autoClose: document.getElementById('autoClose').checked,
  copyToClipboard: document.getElementById('copyToClipboard').checked,
  includeExternal: document.getElementById('includeExternal').checked
};

chrome.storage.sync.set(settings, function() {
  showStatus('Settings saved successfully!', 'success');
});
```

### updatePreview()

**Signature**: `function updatePreview(): void`

**Purpose**: Shows real-time preview of folder structure

**Output Format**:
```
📁 VaultName/
  📁 FolderPath/
    📁 YYYY/
      📁 MM/
        📄 YYYY-MM-15 Example Meeting.md
        📄 YYYY-MM-20 Another Meeting.md
```

### showStatus(message, type)

**Signature**: `function showStatus(message: string, type: 'success' | 'error'): void`

**Purpose**: Displays temporary status message

**Auto-dismiss**: 3 seconds

```javascript
function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = `status ${type}`;

  setTimeout(() => {
    statusDiv.className = 'status';
  }, 3000);
}
```

---

## Message Protocol

### Chrome Extension Messages

**Popup → Content Script**

**Ping Message**:
```javascript
// Request
{
  action: 'ping'
}

// Response
{
  status: 'ready'
}
```

**Extract Meeting Message**:
```javascript
// Request
{
  action: 'extractMeeting'
}

// Response
{
  success: true,
  data: {
    title: string,
    time: string,
    location: string,
    totalAttendees: number,
    hasExternalAttendees: boolean,
    meetingDate: {
      year: string,
      month: string,
      day: string,
      full: string
    },
    safeTitle: string,
    note: string
  }
}

// Or on error
{
  success: false,
  error: string
}
```

### Message Passing Pattern

```javascript
// Sender (popup.js)
chrome.tabs.sendMessage(tabId, {action: 'extractMeeting'}, function(response) {
  if (chrome.runtime.lastError) {
    console.error('Runtime error:', chrome.runtime.lastError);
    return;
  }

  if (response && response.success) {
    // Handle success
  } else {
    // Handle error
  }
});

// Receiver (content.js)
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'extractMeeting') {
    extractMeetingDetails().then(sendResponse);
    return true; // Keep channel open for async response
  }
});
```

---

## Data Structures

### Settings Object

```typescript
interface Settings {
  vaultName: string;        // Obsidian vault name
  folderPath: string;       // Base folder for meetings
  autoClose: boolean;       // Close popup after creation
  copyToClipboard: boolean; // Auto-copy to clipboard
  includeExternal: boolean; // Add [EXTERNAL] prefix
}
```

**Storage Key**: `chrome.storage.sync`
**Scope**: Synced across user's Chrome instances

### Meeting Date Object

```typescript
interface MeetingDate {
  year: string;   // "2025"
  month: string;  // "01" (zero-padded)
  day: string;    // "21" (zero-padded)
  full: string;   // "2025-01-21" (ISO format)
}
```

### Attendees Data Structure

```typescript
type AttendeesByStatus = {
  'Accepted': Map<string, string>;    // name → email
  'Tentative': Map<string, string>;
  'Declined': Map<string, string>;
  'No Response': Map<string, string>;
};
```

**Why Map?**
- Prevents duplicate names
- Fast lookup
- Preserves insertion order for sorting

### Note Content Format

```markdown
---
tags: meeting
date: YYYY-MM-DD
type: outlook-meeting
external: boolean
attendees: ["Name 1", "Name 2", ...]
summary:
---

# Title

## Meeting Details
**Date/Time:** ...
**Location:** ...
**Total Attendees:** N
**External Meeting:** Yes/No

### Attendees by RSVP Status

**Organizer:** Name

**Accepted (N):** Name1, Name2, ...

**Tentative (N):** Name1, Name2, ...

**Declined (N):** Name1, Name2, ...

**No Response (N):** Name1, Name2, ...

### Meeting Description
Body text...

## Notes


## Action Items


## Follow-up

```

---

## Obsidian URI Schema

### obsidian://new

**Purpose**: Creates a new note in Obsidian

**Format**:
```
obsidian://new?vault=<vault>&file=<path>&content=<content>
```

**Parameters**:
- `vault` - Vault name (URI encoded)
- `file` - File path without `.md` extension (URI encoded)
- `content` - Note content (URI encoded)

**Example**:
```
obsidian://new?vault=Syntax&file=02.%20Meeting%20Notes%2F2025%2F01%2F2025-01-21%20Team%20Sync&content=%23%20Team%20Sync...
```

**Encoding Rules**:
- Use `encodeURIComponent()` for all parameters
- Spaces → `%20`
- Forward slashes → `%2F`
- Special chars properly escaped

**Path Construction**:
```javascript
const folderPath = `${folderPath}/${year}/${month}`;
const filename = `${date} ${safeTitle}`;
const fullPath = `${folderPath}/${filename}`;
```

**Note**: Obsidian automatically:
- Creates missing folders
- Adds `.md` extension
- Opens the note after creation

### File Naming Rules

**Safe Characters**: Letters, numbers, spaces, hyphens, underscores

**Unsafe Characters** (replaced with `-`):
```javascript
/[<>:"/\\|?*]/g
```

Specifically:
- `<` → `-`
- `>` → `-`
- `:` → `-`
- `"` → `-`
- `/` → `-`
- `\` → `-`
- `|` → `-`
- `?` → `-`
- `*` → `-`

**Example Sanitization**:
```javascript
"Q1 Planning: Review & Approve" → "Q1 Planning- Review - Approve"
```

---

## Selector Reference

### Common Outlook Selectors

**Meeting Title**:
- `span.CWGkB`
- `div.FZzLA span`
- `span[class*="CWGkB"]`
- `input[aria-label*="subject" i]`

**Date/Time**:
- `div.y79CD`
- `button[aria-label*="Start time" i]`
- `input[aria-label*="start" i][aria-label*="time" i]`

**Location**:
- `input[aria-label*="location" i]`
- `button[aria-label*="location" i]`

**Organizer**:
- `div.lpxUr.EKtKK span.fui-Persona__primaryText`

**Body**:
- `div[role="textbox"][aria-label*="message" i]`
- `div[role="textbox"][contenteditable="true"]`

**Attendee List**:
- Container: `div[role="tree"]`
- Section headers: `div.ELl7R.TqkD6`
- Attendee items: `div.lpxUr.VCV1f.sKKBX`
- Persona name: `span.fui-Persona__primaryText`

**Email Extraction**:
- From `aria-label` of parent `span`
- Pattern: `/[\w\.-]+@[\w\.-]+\.\w+/`

---

## Error Handling Patterns

### Content Script Errors

```javascript
try {
  // Extraction logic
  return {
    success: true,
    data: { /* ... */ }
  };
} catch (error) {
  return {
    success: false,
    error: error.message
  };
}
```

### Message Passing Errors

```javascript
chrome.tabs.sendMessage(tabId, message, function(response) {
  if (chrome.runtime.lastError) {
    // Content script not loaded or tab closed
    console.error('Runtime error:', chrome.runtime.lastError);
    return;
  }

  // Process response
});
```

### Script Injection Errors

```javascript
chrome.scripting.executeScript({
  target: {tabId: tabId},
  files: ['content.js']
}, () => {
  if (chrome.runtime.lastError) {
    // Injection failed (permissions, invalid tab, etc.)
    console.error('Injection failed:', chrome.runtime.lastError);
  }
});
```

### Validation Errors

```javascript
if (!settings.vaultName) {
  showStatus('Please enter a vault name', 'error');
  return;
}
```

---

## Performance Considerations

### Attendee Scrolling

**Optimization Techniques**:
1. **Incremental scrolling**: 80% viewport height prevents missing elements
2. **Throttling**: 300ms delay between scrolls allows DOM updates
3. **Early exit**: Stops after 3 iterations with no new attendees
4. **Max attempts**: Hard limit of 50 scrolls prevents infinite loops
5. **Deduplication**: Uses `Set` to track processed names

**Time Complexity**:
- Best case: O(n) where n = visible attendees (no scrolling needed)
- Worst case: O(50 × m) where m = attendees per viewport
- Typical: 5-10 seconds for 100+ attendees

### Selector Performance

**Strategy**: Fail-fast fallback chain
```javascript
for (let selector of selectors) {
  try {
    let element = document.querySelector(selector);
    if (element && element.textContent.trim()) {
      return element.textContent.trim();
    }
  } catch(e) {
    // Invalid selector, try next
  }
}
```

### Storage Access

**Async Pattern**:
```javascript
chrome.storage.sync.get(defaults, function(items) {
  // Use items
});
```

**Sync vs Local**:
- Uses `chrome.storage.sync` for cross-device settings
- Limit: 100KB total, 8KB per item
- Alternative: `chrome.storage.local` for larger data

---

## Security Considerations

### Content Security Policy

**Manifest V3 Requirements**:
- No inline scripts in HTML
- No `eval()` or `Function()` constructors
- External scripts must be bundled

### XSS Prevention

**User Input Sanitization**:
```javascript
// Filename sanitization
const safeTitle = title.replace(/[<>:"/\\|?*]/g, '-');

// URI encoding
encodeURIComponent(content)
```

**DOM Insertion**:
```javascript
// Safe: textContent
element.textContent = userInput;

// Unsafe: innerHTML (not used in this extension)
element.innerHTML = userInput; // DON'T DO THIS
```

### Permission Scope

**Minimal Permissions**:
- Only Outlook domains in `host_permissions`
- No broad `<all_urls>` access
- `activeTab` requires user interaction

**Storage Isolation**:
- Settings stored per-extension
- No cross-extension access
- Synced securely by Chrome

---

## Testing Utilities

### Manual Test Scenarios

**Content Script Loading**:
```javascript
// In browser console on Outlook page
chrome.runtime.sendMessage({action: 'ping'}, response => {
  console.log('Content script status:', response);
});
```

**Storage Inspection**:
```javascript
// In popup console
chrome.storage.sync.get(null, items => {
  console.log('All settings:', items);
});
```

**URI Validation**:
```javascript
// Test URI construction
const testUri = `obsidian://new?vault=${encodeURIComponent('Test Vault')}&file=${encodeURIComponent('Test/Note')}&content=${encodeURIComponent('# Test')}`;
console.log('Test URI:', testUri);
```

### Debug Logging

**Add to content.js**:
```javascript
console.log('Extraction started');
console.log('Found title:', title);
console.log('Found attendees:', attendeesByStatus);
console.log('Final note length:', note.length);
```

**Add to popup.js**:
```javascript
console.log('Settings loaded:', currentSettings);
console.log('Extraction response:', response);
console.log('Obsidian URI:', obsidianUri);
```

---

## Version History

### Manifest Version

```json
{
  "manifest_version": 3,
  "version": "1.1.0"
}
```

**Breaking Changes from V2 → V3**:
- `chrome.scripting.executeScript` instead of `tabs.executeScript`
- `host_permissions` separate from `permissions`
- Service workers instead of background pages (not used here)

---

This documentation provides complete API reference for developers working with the Outlook to Obsidian extension codebase.
