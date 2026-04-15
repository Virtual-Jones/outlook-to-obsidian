'use strict';

/* ═══════════════════════════════════════════════════════════════════════════
   Outlook to Obsidian — Office Add-in task pane
   ═══════════════════════════════════════════════════════════════════════════
   Uses Office.js (Office.context.mailbox.item) to read meeting data directly
   from the Outlook object model — no DOM scraping required.

   RSVP status is fetched via EWS (makeEwsRequestAsync) when the item is in
   read mode and the mailbox is Exchange / Microsoft 365.  If EWS is
   unavailable the add-in gracefully falls back to the basic Office.js
   attendee list (no RSVP status, all shown as "No Response").
   ═══════════════════════════════════════════════════════════════════════════ */

/* ── Defaults ────────────────────────────────────────────────────────────── */
const DEFAULT_SETTINGS = {
  vaultName:       'Syntax',
  folderPath:      '02. Meeting Notes',
  internalDomain:  'syntax.com',
  copyToClipboard: true,
  includeExternal: false,
};

let settings = { ...DEFAULT_SETTINGS };
let extractedData = null;

/* ── Bootstrap ───────────────────────────────────────────────────────────── */
Office.onReady(() => {
  loadSettings();
  bindUI();
});

/* ── UI binding ──────────────────────────────────────────────────────────── */
function bindUI() {
  // Main panel
  document.getElementById('extractBtn').addEventListener('click', onExtract);
  document.getElementById('openObsidianBtn').addEventListener('click', onOpenObsidian);
  document.getElementById('copyBtn').addEventListener('click', onCopy);
  document.getElementById('settingsBtn').addEventListener('click', () => showPanel('settings'));

  // Settings panel
  document.getElementById('backBtn').addEventListener('click', () => showPanel('main'));
  document.getElementById('saveSettingsBtn').addEventListener('click', onSaveSettings);

  // Preset buttons (vault & folder path quick-fills)
  document.querySelectorAll('.preset-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const target = btn.dataset.target;
      const value  = btn.dataset.value;
      if (target && value) {
        document.getElementById(target).value = value;
        if (target === 'vaultName' || target === 'folderPath') updateFolderPreview();
      }
    });
  });

  // Live preview in settings
  document.getElementById('vaultName').addEventListener('input', updateFolderPreview);
  document.getElementById('folderPath').addEventListener('input', updateFolderPreview);
}

/* ── Panel switching ─────────────────────────────────────────────────────── */
function showPanel(name) {
  const isMain = (name === 'main');
  document.getElementById('mainPanel').style.display     = isMain ? 'flex' : 'none';
  document.getElementById('settingsPanel').style.display = isMain ? 'none' : 'flex';
  if (!isMain) {
    populateSettingsForm();
    updateFolderPreview();
  }
}

/* ── Settings — load ─────────────────────────────────────────────────────── */
function loadSettings() {
  try {
    const rs = Office.context.roamingSettings;
    for (const key of Object.keys(DEFAULT_SETTINGS)) {
      const v = rs.get(key);
      if (v !== undefined && v !== null) settings[key] = v;
    }
  } catch (_) {
    // roamingSettings not available (e.g. unit test harness) — use defaults
  }
  updateConfigDisplay();
}

/* ── Settings — save ─────────────────────────────────────────────────────── */
function onSaveSettings() {
  const vaultName      = document.getElementById('vaultName').value.trim();
  const folderPath     = document.getElementById('folderPath').value.trim();
  const internalDomain = document.getElementById('internalDomain').value.trim();

  if (!vaultName)  { showStatus('Please enter a vault name.', 'error'); return; }
  if (!folderPath) { showStatus('Please enter a folder path.', 'error'); return; }

  settings.vaultName       = vaultName;
  settings.folderPath      = folderPath;
  settings.internalDomain  = internalDomain || DEFAULT_SETTINGS.internalDomain;
  settings.copyToClipboard = document.getElementById('copyToClipboard').checked;
  settings.includeExternal = document.getElementById('includeExternal').checked;

  try {
    const rs = Office.context.roamingSettings;
    for (const [k, v] of Object.entries(settings)) rs.set(k, v);
    rs.saveAsync(result => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        updateConfigDisplay();
        showPanel('main');
        showStatus('Settings saved.', 'success');
      } else {
        showStatus('Could not save settings: ' + (result.error?.message || 'unknown error'), 'error');
      }
    });
  } catch (e) {
    showStatus('Error saving settings: ' + e.message, 'error');
  }
}

function populateSettingsForm() {
  document.getElementById('vaultName').value       = settings.vaultName;
  document.getElementById('folderPath').value      = settings.folderPath;
  document.getElementById('internalDomain').value  = settings.internalDomain;
  document.getElementById('copyToClipboard').checked = settings.copyToClipboard;
  document.getElementById('includeExternal').checked = settings.includeExternal;
}

function updateConfigDisplay() {
  const ok = !!(settings.vaultName && settings.folderPath);
  document.getElementById('configWarning').style.display = ok ? 'none' : 'block';
  document.getElementById('vaultDisplay').textContent    = settings.vaultName  || 'Not set';
  document.getElementById('pathDisplay').textContent     = settings.folderPath || 'Not set';
  document.getElementById('extractBtn').disabled = !ok;
}

function updateFolderPreview() {
  const vault  = document.getElementById('vaultName').value  || 'YourVault';
  const folder = document.getElementById('folderPath').value || 'Meeting Notes';
  const now    = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  document.getElementById('folderPreview').textContent =
    `📁 ${vault}/\n  📁 ${folder}/\n    📁 ${y}/\n      📁 ${m}/\n        📄 ${y}-${m}-15 Example Meeting.md`;
}

/* ── Extract button handler ──────────────────────────────────────────────── */
async function onExtract() {
  const btn = document.getElementById('extractBtn');
  btn.disabled = true;
  document.getElementById('details').style.display = 'none';
  document.getElementById('actions').style.display = 'none';
  showStatus('Extracting meeting details…', 'info');

  try {
    extractedData = await extractMeetingDetails();

    // Apply [EXTERNAL] prefix when configured
    if (settings.includeExternal && extractedData.hasExternalAttendees) {
      extractedData.title = '[EXTERNAL] ' + extractedData.title;
      extractedData.note  = extractedData.note.replace(/^(# ).+$/m, `$1${extractedData.title}`);
    }

    renderDetails(extractedData);
    showStatus('Meeting details extracted successfully.', 'success');

    if (settings.copyToClipboard) await copyText(extractedData.note);

    document.getElementById('details').style.display = 'block';
    document.getElementById('actions').style.display = 'flex';
  } catch (err) {
    showStatus('Error: ' + err.message, 'error');
  } finally {
    btn.disabled = !(settings.vaultName && settings.folderPath);
  }
}

function renderDetails(data) {
  const truncate = (s, n) => s.length > n ? s.substring(0, n) + '…' : s;
  document.getElementById('details').innerHTML = `
    <div class="detail-row">
      <span class="detail-label">Title</span>
      <span class="detail-value">${esc(truncate(data.title, 60))}</span>
    </div>
    <div class="detail-row">
      <span class="detail-label">Date/Time</span>
      <span class="detail-value">${esc(data.time || 'Not specified')}</span>
    </div>
    ${data.location ? `
    <div class="detail-row">
      <span class="detail-label">Location</span>
      <span class="detail-value">${esc(truncate(data.location, 50))}</span>
    </div>` : ''}
    <div class="detail-row">
      <span class="detail-label">Attendees</span>
      <span class="detail-value">${data.totalAttendees}</span>
    </div>
    ${data.hasExternalAttendees ? `
    <div class="detail-row">
      <span class="detail-label">External</span>
      <span class="detail-value external">Yes</span>
    </div>` : ''}
  `;
}

/* ── Open in Obsidian ────────────────────────────────────────────────────── */
function onOpenObsidian() {
  if (!extractedData) return;
  const uri  = buildObsidianUri(extractedData);
  const link = document.getElementById('obsidianLink');
  link.href  = uri;
  // Clicking a hidden anchor is the most reliable way to trigger a custom
  // protocol handler across Outlook Desktop (WebView2) and Outlook Web.
  link.click();
}

function buildObsidianUri(data) {
  const folder   = `${settings.folderPath}/${data.meetingDate.year}/${data.meetingDate.month}`;
  const safeTitle = data.title.replace(/[<>:"/\\|?*]/g, '-');
  const file     = `${folder}/${data.meetingDate.full} ${safeTitle}`;
  return `obsidian://new?vault=${encodeURIComponent(settings.vaultName)}&file=${encodeURIComponent(file)}&content=${encodeURIComponent(data.note)}`;
}

/* ── Copy to clipboard ───────────────────────────────────────────────────── */
async function onCopy() {
  if (!extractedData) return;
  await copyText(extractedData.note);
  showStatus('Note copied to clipboard.', 'success');
}

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch (_) {
    // Fallback for older WebView2 builds
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity  = '0';
    document.body.appendChild(ta);
    ta.focus();
    ta.select();
    document.execCommand('copy');
    document.body.removeChild(ta);
  }
}

/* ── UI helpers ──────────────────────────────────────────────────────────── */
function showStatus(msg, type) {
  const el = document.getElementById('status');
  el.textContent  = msg;
  el.className    = `status ${type}`;
  el.style.display = 'block';
}

function esc(str) {
  return (str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;');
}

/* ═══════════════════════════════════════════════════════════════════════════
   CORE EXTRACTION — Office.js API
   ═══════════════════════════════════════════════════════════════════════════ */

async function extractMeetingDetails() {
  const item = Office.context.mailbox.item;

  // In read mode the item has a persisted itemId; compose mode does not.
  const isReadMode = !!(item.itemId);

  /* ── Subject ──────────────────────────────────────────────────────────── */
  const title = (await getItemProperty(item.subject)) || 'Untitled Meeting';

  /* ── Start / End ──────────────────────────────────────────────────────── */
  const startDate = await getItemProperty(item.start);
  const endDate   = await getItemProperty(item.end);

  let time        = '';
  let meetingDate = null;

  if (startDate instanceof Date && !isNaN(startDate)) {
    const y  = startDate.getFullYear();
    const mo = String(startDate.getMonth() + 1).padStart(2, '0');
    const d  = String(startDate.getDate()).padStart(2, '0');
    meetingDate = { year: String(y), month: mo, day: d, full: `${y}-${mo}-${d}` };

    const fmt = { hour: '2-digit', minute: '2-digit' };
    const startStr = startDate.toLocaleTimeString([], fmt);
    const endStr   = (endDate instanceof Date && !isNaN(endDate))
      ? endDate.toLocaleTimeString([], fmt)
      : '';
    time = `${startDate.toLocaleDateString()} ${startStr}${endStr ? ' – ' + endStr : ''}`;
  }

  if (!meetingDate) {
    const today = new Date();
    meetingDate = {
      year:  String(today.getFullYear()),
      month: String(today.getMonth() + 1).padStart(2, '0'),
      day:   String(today.getDate()).padStart(2, '0'),
      full:  today.toISOString().split('T')[0],
    };
  }

  /* ── Location ─────────────────────────────────────────────────────────── */
  const location = (await getItemProperty(item.location)) || '';

  /* ── Organizer (read mode only) ───────────────────────────────────────── */
  let organizer      = '';
  let organizerEmail = '';

  if (isReadMode && item.organizer) {
    organizer      = item.organizer.displayName || '';
    organizerEmail = (item.organizer.emailAddress || '').toLowerCase();
  }

  /* ── Body / Description ───────────────────────────────────────────────── */
  const body = await getBodyText(item);

  /* ── Attendees with RSVP ──────────────────────────────────────────────── */
  let attendeesByStatus = {
    'Accepted':    new Map(),
    'Tentative':   new Map(),
    'Declined':    new Map(),
    'No Response': new Map(),
  };

  if (isReadMode && item.itemId) {
    try {
      attendeesByStatus = await getAttendeesViaEws(item.itemId);
    } catch (ewsErr) {
      // EWS unavailable or failed — fall back to Office.js (no RSVP info)
      console.warn('EWS attendee fetch failed, using Office.js fallback:', ewsErr.message);
      attendeesByStatus = await getAttendeesViaOfficeJs(item);
    }
  } else {
    // Compose mode: no itemId yet
    attendeesByStatus = await getAttendeesViaOfficeJs(item);
  }

  /* ── External attendee detection ──────────────────────────────────────── */
  const internalSuffix = '@' + (settings.internalDomain || 'syntax.com')
    .toLowerCase()
    .replace(/^@/, '');

  let hasExternalAttendees = organizerEmail
    ? !organizerEmail.endsWith(internalSuffix)
    : false;

  if (!hasExternalAttendees) {
    for (const attendees of Object.values(attendeesByStatus)) {
      for (const email of attendees.values()) {
        if (email && !email.endsWith(internalSuffix)) {
          hasExternalAttendees = true;
          break;
        }
      }
      if (hasExternalAttendees) break;
    }
  }

  /* ── Format note ──────────────────────────────────────────────────────── */
  const acceptedList = Array.from(attendeesByStatus['Accepted'].keys()).sort();
  if (organizer) acceptedList.unshift(organizer);

  const attendeesFormatted = [];
  let totalAttendees = organizer ? 1 : 0;

  if (organizer) attendeesFormatted.push(`**Organizer:** ${organizer}`);

  for (const status of ['Accepted', 'Tentative', 'No Response', 'Declined']) {
    const list = Array.from(attendeesByStatus[status].keys()).sort();
    if (list.length > 0) {
      totalAttendees += list.length;
      attendeesFormatted.push(`**${status} (${list.length}):** ${list.join(', ')}`);
    }
  }

  const attendeesStr = attendeesFormatted.join('\n\n') || 'No attendees found';

  const noteLines = [
    '---',
    'tags: meeting',
    `date: ${meetingDate.full}`,
    'type: outlook-meeting',
    `external: ${hasExternalAttendees}`,
    `attendees: [${acceptedList.map(n => `"${n}"`).join(', ')}]`,
    'summary: ',
    '---',
    '',
    `# ${title}`,
    '',
    '## Meeting Details',
    `**Date/Time:** ${time || 'Not specified'}`,
    location ? `**Location:** ${location}` : '',
    `**Total Attendees:** ${totalAttendees}`,
    hasExternalAttendees ? '**External Meeting:** Yes' : '',
    '',
    '### Attendees by RSVP Status',
    attendeesStr,
    '',
    '### Meeting Description',
    body.trim() || 'No description available',
    '',
    '## Notes',
    '',
    '',
    '## Action Items',
    '',
    '',
    '## Follow-up',
    '',
    '',
  ];

  const note      = noteLines.filter(l => l !== null).join('\r\n');
  const safeTitle = title.replace(/[<>:"/\\|?*]/g, '-');

  return { title, time, location, totalAttendees, hasExternalAttendees, meetingDate, safeTitle, note };
}

/* ═══════════════════════════════════════════════════════════════════════════
   OFFICE.JS HELPERS
   ═══════════════════════════════════════════════════════════════════════════ */

/**
 * Reads a property that is either a direct value (read mode) or exposes
 * getAsync() (compose mode), and returns a Promise for the resolved value.
 */
function getItemProperty(propOrValue) {
  return new Promise(resolve => {
    if (propOrValue === undefined || propOrValue === null) {
      resolve('');
      return;
    }
    // Direct value: string, Date, number, boolean
    if (typeof propOrValue !== 'object' || propOrValue instanceof Date) {
      resolve(propOrValue);
      return;
    }
    // Object with getAsync (compose mode wrappers)
    if (typeof propOrValue.getAsync === 'function') {
      propOrValue.getAsync(result => {
        resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : '');
      });
    } else {
      resolve('');
    }
  });
}

/** Get the appointment body as plain text. */
function getBodyText(item) {
  return new Promise(resolve => {
    if (!item.body) { resolve(''); return; }
    item.body.getAsync(Office.CoercionType.Text, result => {
      resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value || '' : '');
    });
  });
}

/**
 * Fallback attendee extraction using the Office.js object model.
 * Provides name and email but NOT RSVP status (all marked "No Response").
 */
function getAttendeesViaOfficeJs(item) {
  return new Promise(async resolve => {
    const result = {
      'Accepted':    new Map(),
      'Tentative':   new Map(),
      'Declined':    new Map(),
      'No Response': new Map(),
    };

    async function fetchList(prop) {
      const value = item[prop];
      if (!value) return [];
      if (Array.isArray(value)) return value;                   // read mode: direct array
      if (typeof value.getAsync === 'function') {               // compose mode: async wrapper
        return await new Promise(res => {
          value.getAsync(r =>
            res(r.status === Office.AsyncResultStatus.Succeeded ? r.value || [] : [])
          );
        });
      }
      return [];
    }

    const required = await fetchList('requiredAttendees');
    const optional = await fetchList('optionalAttendees');

    for (const a of [...required, ...optional]) {
      const name  = a.displayName || a.emailAddress || '';
      const email = (a.emailAddress || '').toLowerCase();
      if (name) result['No Response'].set(name, email);
    }

    resolve(result);
  });
}

/* ═══════════════════════════════════════════════════════════════════════════
   EWS — fetch attendees with RSVP status
   ═══════════════════════════════════════════════════════════════════════════
   Requires ReadWriteMailbox permission in the manifest.
   Available in: Outlook desktop (Windows & Mac), Outlook on the web (M365).
   NOT available on iOS/Android Outlook.
   ═══════════════════════════════════════════════════════════════════════════ */

function getAttendeesViaEws(itemId) {
  return new Promise((resolve, reject) => {

    const soap = `<?xml version="1.0" encoding="UTF-8"?>
<soap:Envelope
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:xsd="http://www.w3.org/2001/XMLSchema"
  xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"
  xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"
  xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages">
  <soap:Header>
    <t:RequestServerVersion Version="Exchange2013"/>
  </soap:Header>
  <soap:Body>
    <m:GetItem>
      <m:ItemShape>
        <t:BaseShape>IdOnly</t:BaseShape>
        <t:AdditionalProperties>
          <t:FieldURI FieldURI="calendar:RequiredAttendees"/>
          <t:FieldURI FieldURI="calendar:OptionalAttendees"/>
        </t:AdditionalProperties>
      </m:ItemShape>
      <m:ItemIds>
        <t:ItemId Id="${itemId}"/>
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(soap, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        reject(new Error(result.error?.message || 'EWS request failed'));
        return;
      }
      try {
        resolve(parseEwsAttendees(result.value));
      } catch (e) {
        reject(e);
      }
    });
  });
}

/**
 * Parse the EWS GetItem XML response and return an attendeesByStatus object.
 *
 * EWS ResponseType values:
 *   None | Organizer | Tentative | Accept | Decline | NoResponseReceived
 */
function parseEwsAttendees(xmlString) {
  const parser = new DOMParser();
  const doc    = parser.parseFromString(xmlString, 'text/xml');
  const T      = 'http://schemas.microsoft.com/exchange/services/2006/types';
  const M      = 'http://schemas.microsoft.com/exchange/services/2006/messages';

  // Surface any EWS-level errors
  const responseMsgs = doc.getElementsByTagNameNS(M, 'GetItemResponseMessage');
  if (responseMsgs.length > 0) {
    const cls = responseMsgs[0].getAttribute('ResponseClass');
    if (cls === 'Error') {
      const code = responseMsgs[0].getElementsByTagNameNS(M, 'ResponseCode')[0]?.textContent || '';
      throw new Error(`EWS error: ${code}`);
    }
  }

  const result = {
    'Accepted':    new Map(),
    'Tentative':   new Map(),
    'Declined':    new Map(),
    'No Response': new Map(),
  };

  // Map EWS ResponseType → our bucket names
  const RSVP_MAP = {
    'accept':              'Accepted',
    'tentative':           'Tentative',
    'decline':             'Declined',
    'none':                'No Response',
    'noresponsereceived':  'No Response',
    'organizer':           null,   // skip — captured via item.organizer
  };

  function processAttendeeGroup(tagName) {
    for (const group of doc.getElementsByTagNameNS(T, tagName)) {
      for (const attendeeEl of group.getElementsByTagNameNS(T, 'Attendee')) {
        const name  = attendeeEl.getElementsByTagNameNS(T, 'Name')[0]?.textContent?.trim() || '';
        const email = (attendeeEl.getElementsByTagNameNS(T, 'EmailAddress')[0]?.textContent?.trim() || '').toLowerCase();
        const rt    = (attendeeEl.getElementsByTagNameNS(T, 'ResponseType')[0]?.textContent?.trim() || 'none').toLowerCase();

        if (!name) continue;

        const bucket = RSVP_MAP.hasOwnProperty(rt) ? RSVP_MAP[rt] : 'No Response';
        if (bucket === null) continue;  // organizer — skip
        result[bucket].set(name, email);
      }
    }
  }

  processAttendeeGroup('RequiredAttendees');
  processAttendeeGroup('OptionalAttendees');

  return result;
}
