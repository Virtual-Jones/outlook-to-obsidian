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
  emailFolderPath: '04. Email Notes',
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
  applyContextLabels();
});

/* Returns 'message' or 'meeting' based on the current Outlook item. */
function getCurrentItemType() {
  try {
    const t = Office.context.mailbox.item?.itemType;
    return t === Office.MailboxEnums.ItemType.Message ? 'message' : 'meeting';
  } catch (_) {
    return 'meeting';
  }
}

/* Update button label and header to match the current context (message vs meeting). */
function applyContextLabels() {
  const isMessage = getCurrentItemType() === 'message';
  const btn = document.getElementById('extractBtn');
  if (btn) {
    btn.innerHTML = isMessage
      ? '<span>📧</span><span>Extract Email Details</span>'
      : '<span>📅</span><span>Extract Meeting Details</span>';
  }
  const headerIcon = document.getElementById('headerIcon');
  if (headerIcon) headerIcon.textContent = isMessage ? '📧' : '📅';
  const headerTitle = document.getElementById('headerTitle');
  if (headerTitle) headerTitle.textContent = isMessage ? 'Send Email to Obsidian' : 'Send Meeting to Obsidian';
}

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
  const vaultName       = document.getElementById('vaultName').value.trim();
  const folderPath      = document.getElementById('folderPath').value.trim();
  const emailFolderPath = document.getElementById('emailFolderPath').value.trim();
  const internalDomain  = document.getElementById('internalDomain').value.trim();

  if (!vaultName)       { showStatus('Please enter a vault name.', 'error'); return; }
  if (!folderPath)      { showStatus('Please enter a meeting notes folder.', 'error'); return; }
  if (!emailFolderPath) { showStatus('Please enter an email notes folder.', 'error'); return; }

  settings.vaultName       = vaultName;
  settings.folderPath      = folderPath;
  settings.emailFolderPath = emailFolderPath;
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
  document.getElementById('vaultName').value         = settings.vaultName;
  document.getElementById('folderPath').value        = settings.folderPath;
  document.getElementById('emailFolderPath').value   = settings.emailFolderPath;
  document.getElementById('internalDomain').value    = settings.internalDomain;
  document.getElementById('copyToClipboard').checked = settings.copyToClipboard;
  document.getElementById('includeExternal').checked = settings.includeExternal;
}

function updateConfigDisplay() {
  const isMessage = getCurrentItemType() === 'message';
  const activePath = isMessage ? settings.emailFolderPath : settings.folderPath;
  const ok = !!(settings.vaultName && activePath);
  document.getElementById('configWarning').style.display = ok ? 'none' : 'block';
  document.getElementById('vaultDisplay').textContent    = settings.vaultName  || 'Not set';
  document.getElementById('pathDisplay').textContent     = activePath || 'Not set';
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

  const isMessage = getCurrentItemType() === 'message';
  showStatus(isMessage ? 'Extracting email…' : 'Extracting meeting details…', 'info');

  try {
    extractedData = isMessage
      ? await extractEmailDetails()
      : await extractMeetingDetails();

    // Apply [EXTERNAL] prefix when configured
    if (settings.includeExternal && extractedData.hasExternalAttendees) {
      extractedData.title = '[EXTERNAL] ' + extractedData.title;
      extractedData.note  = extractedData.note.replace(/^(# ).+$/m, `$1${extractedData.title}`);
    }

    renderDetails(extractedData);
    showStatus(isMessage ? 'Email extracted successfully.' : 'Meeting details extracted successfully.', 'success');

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
  const isEmail = data.kind === 'email';

  const rows = [];

  rows.push(`
    <div class="detail-row">
      <span class="detail-label">${isEmail ? 'Subject' : 'Title'}</span>
      <span class="detail-value">${esc(truncate(data.title, 60))}</span>
    </div>`);

  if (isEmail) {
    rows.push(`
    <div class="detail-row">
      <span class="detail-label">From</span>
      <span class="detail-value">${esc(truncate(data.fromName || '', 50))}</span>
    </div>`);
  }

  rows.push(`
    <div class="detail-row">
      <span class="detail-label">${isEmail ? 'Sent' : 'Date/Time'}</span>
      <span class="detail-value">${esc(data.time || 'Not specified')}</span>
    </div>`);

  if (!isEmail && data.location) {
    rows.push(`
    <div class="detail-row">
      <span class="detail-label">Location</span>
      <span class="detail-value">${esc(truncate(data.location, 50))}</span>
    </div>`);
  }

  rows.push(`
    <div class="detail-row">
      <span class="detail-label">${isEmail ? 'Recipients' : 'Attendees'}</span>
      <span class="detail-value">${data.totalAttendees}</span>
    </div>`);

  if (data.hasExternalAttendees) {
    rows.push(`
    <div class="detail-row">
      <span class="detail-label">External</span>
      <span class="detail-value external">Yes</span>
    </div>`);
  }

  document.getElementById('details').innerHTML = rows.join('');
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
  const basePath  = data.kind === 'email' ? settings.emailFolderPath : settings.folderPath;
  const folder    = `${basePath}/${data.meetingDate.year}/${data.meetingDate.month}`;
  const safeTitle = data.title.replace(/[<>:"/\\|?*]/g, '-');
  const file      = `${folder}/${data.meetingDate.full} ${safeTitle}`;
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

  /* ── Body / Description (HTML → Markdown) ─────────────────────────────── */
  const body = await getBodyMarkdown(item);

  /* ── Attendees with RSVP ──────────────────────────────────────────────── */
  let attendeesByStatus = {
    'Accepted':    new Map(),
    'Tentative':   new Map(),
    'Declined':    new Map(),
    'No Response': new Map(),
  };

  let ewsErrorMessage = '';
  if (isReadMode && item.itemId) {
    try {
      attendeesByStatus = await getAttendeesViaEws(item.itemId);
    } catch (ewsErr) {
      console.warn('EWS attendee fetch failed, using Office.js fallback:', ewsErr.message);
      ewsErrorMessage = ewsErr.message || 'unknown';
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

  return { kind: 'meeting', title, time, location, totalAttendees, hasExternalAttendees, meetingDate, safeTitle, note, ewsErrorMessage };
}

/* ═══════════════════════════════════════════════════════════════════════════
   CORE EXTRACTION — Email (Message read mode)
   ═══════════════════════════════════════════════════════════════════════════ */

async function extractEmailDetails() {
  const item = Office.context.mailbox.item;

  /* ── Subject ──────────────────────────────────────────────────────────── */
  const title = (await getItemProperty(item.subject)) || 'No Subject';

  /* ── Sent date — prefer dateTimeCreated, fall back to dateTimeModified ─── */
  const sentDate = (item.dateTimeCreated instanceof Date && !isNaN(item.dateTimeCreated))
    ? item.dateTimeCreated
    : (item.dateTimeModified instanceof Date ? item.dateTimeModified : new Date());

  const y  = sentDate.getFullYear();
  const mo = String(sentDate.getMonth() + 1).padStart(2, '0');
  const d  = String(sentDate.getDate()).padStart(2, '0');
  const meetingDate = { year: String(y), month: mo, day: d, full: `${y}-${mo}-${d}` };

  const time = `${sentDate.toLocaleDateString()} ${sentDate.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' })}`;

  /* ── From / Sender ────────────────────────────────────────────────────── */
  const fromObj    = item.from || item.sender || {};
  const fromName   = fromObj.displayName || fromObj.emailAddress || '';
  const fromEmail  = (fromObj.emailAddress || '').toLowerCase();

  /* ── To / Cc ──────────────────────────────────────────────────────────── */
  const toList = Array.isArray(item.to) ? item.to : [];
  const ccList = Array.isArray(item.cc) ? item.cc : [];

  function formatPerson(p) {
    const name  = p.displayName || p.emailAddress || '';
    const email = p.emailAddress || '';
    return email && name !== email ? `${name} <${email}>` : name;
  }

  const toFormatted = toList.map(formatPerson).filter(Boolean);
  const ccFormatted = ccList.map(formatPerson).filter(Boolean);

  /* ── Body (HTML → Markdown) ───────────────────────────────────────────── */
  const body = await getBodyMarkdown(item);

  /* ── External detection ───────────────────────────────────────────────── */
  const internalSuffix = '@' + (settings.internalDomain || 'syntax.com')
    .toLowerCase()
    .replace(/^@/, '');

  let hasExternalAttendees = !!(fromEmail && !fromEmail.endsWith(internalSuffix));
  if (!hasExternalAttendees) {
    for (const p of [...toList, ...ccList]) {
      const e = (p.emailAddress || '').toLowerCase();
      if (e && !e.endsWith(internalSuffix)) { hasExternalAttendees = true; break; }
    }
  }

  const totalAttendees = toList.length + ccList.length;

  /* ── Format note ──────────────────────────────────────────────────────── */
  const fromLine = fromName
    ? (fromEmail && fromName !== fromEmail ? `${fromName} <${fromEmail}>` : fromName)
    : 'Unknown sender';

  const noteLines = [
    '---',
    'tags: email',
    `date: ${meetingDate.full}`,
    'type: outlook-email',
    `from: "${fromLine.replace(/"/g, '\\"')}"`,
    `external: ${hasExternalAttendees}`,
    'summary: ',
    '---',
    '',
    `# ${title}`,
    '',
    '## Email Details',
    `**From:** ${fromLine}`,
    `**Sent:** ${time}`,
    toFormatted.length ? `**To (${toFormatted.length}):** ${toFormatted.join(', ')}` : null,
    ccFormatted.length ? `**Cc (${ccFormatted.length}):** ${ccFormatted.join(', ')}` : null,
    hasExternalAttendees ? '**External Email:** Yes' : null,
    '',
    '### Body',
    body.trim() || 'No body content',
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

  return {
    kind: 'email',
    title,
    time,
    fromName: fromLine,
    location: '',
    totalAttendees,
    hasExternalAttendees,
    meetingDate,
    safeTitle,
    note,
  };
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
 * Get the body as Markdown. Pulls HTML (line breaks survive HTML coercion in
 * New Outlook where plain-text coercion collapses them), strips Outlook-
 * specific cruft, then runs Turndown. Falls back to plain text if Turndown
 * is unavailable or HTML retrieval fails.
 */
function getBodyMarkdown(item) {
  return new Promise(resolve => {
    if (!item.body) { resolve(''); return; }

    item.body.getAsync(Office.CoercionType.Html, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !result.value) {
        getBodyText(item).then(resolve);
        return;
      }

      const html = flattenTableCellBlocks(
        promoteOutlookTableHeaders(cleanOutlookHtml(result.value))
      );

      if (typeof TurndownService === 'undefined') {
        // CDN blocked or still loading — fall back to a sanitized text dump.
        const tmp = document.createElement('div');
        tmp.innerHTML = html;
        resolve((tmp.textContent || '').trim());
        return;
      }

      try {
        const td = new TurndownService({
          headingStyle:     'atx',
          bulletListMarker: '-',
          codeBlockStyle:   'fenced',
          emDelimiter:      '*',
        });

        // GFM plugin: tables, strikethrough, task lists.
        if (typeof turndownPluginGfm !== 'undefined') {
          td.use(turndownPluginGfm.gfm);
        }

        // Drop embedded inline images (cid: refs aren't useful in the note).
        td.addRule('strip-cid-images', {
          filter: node => node.nodeName === 'IMG' &&
            (node.getAttribute('src') || '').toLowerCase().startsWith('cid:'),
          replacement: (_c, node) => {
            const alt = node.getAttribute('alt');
            return alt ? `*[${alt}]*` : '';
          },
        });

        const md = td.turndown(html)
          .replace(/ /g, ' ')      // non-breaking spaces → regular space
          .replace(/\n{3,}/g, '\n\n')    // collapse runs of blank lines
          .trim();

        resolve(md);
      } catch (e) {
        console.warn('Turndown conversion failed, using text fallback:', e.message);
        getBodyText(item).then(resolve);
      }
    });
  });
}

/**
 * Outlook tables don't use <th>/<thead> — the header row is just bolded
 * <td>s. Turndown's GFM table rule only fires when the first row is a
 * proper heading row, so we promote the first row's <td>s to <th>s
 * whenever the table has no existing header cells.
 */
function promoteOutlookTableHeaders(html) {
  const wrapper = document.createElement('div');
  wrapper.innerHTML = html;

  for (const table of wrapper.querySelectorAll('table')) {
    if (table.querySelector('th') || table.querySelector('thead')) continue;

    const firstRow = table.querySelector('tr');
    if (!firstRow) continue;

    const tdCells = Array.from(firstRow.children).filter(c => c.nodeName === 'TD');
    if (tdCells.length === 0) continue;

    for (const td of tdCells) {
      const th = document.createElement('th');
      for (const attr of Array.from(td.attributes)) {
        th.setAttribute(attr.name, attr.value);
      }
      while (td.firstChild) th.appendChild(td.firstChild);
      td.parentNode.replaceChild(th, td);
    }
  }

  return wrapper.innerHTML;
}

/**
 * Markdown tables require each cell on a single line. Outlook wraps cell
 * content in <p>…</p> (and sometimes <div>), which Turndown converts to
 * \n\n-padded paragraphs — breaking the row syntax. Unwrap any <p>/<div>
 * inside table cells so their inline content survives intact.
 */
function flattenTableCellBlocks(html) {
  const wrapper = document.createElement('div');
  wrapper.innerHTML = html;

  for (const cell of wrapper.querySelectorAll('td, th')) {
    let block;
    while ((block = cell.querySelector('p, div'))) {
      const parent = block.parentNode;
      while (block.firstChild) parent.insertBefore(block.firstChild, block);
      parent.removeChild(block);
    }
  }

  return wrapper.innerHTML;
}

/** Strip Office/Outlook-specific HTML cruft that confuses converters. */
function cleanOutlookHtml(html) {
  return html
    .replace(/<!--\[if[^>]*?\]>[\s\S]*?<!\[endif\]-->/gi, '')   // MSO conditional comments
    .replace(/<o:p[^>]*>[\s\S]*?<\/o:p>/gi, '')                 // <o:p>...</o:p>
    .replace(/<o:p[^>]*\/>/gi, '')                              // <o:p/>
    .replace(/<\/?(?:meta|link|style)\b[^>]*>/gi, '')           // strip head leftovers
    .replace(/<style[\s\S]*?<\/style>/gi, '');                  // and any inline <style>
}

/**
 * Attendee extraction using the Office.js object model.
 *
 * For meetings the current user organizes (read mode), each
 * EmailAddressDetails entry exposes appointmentResponse — the attendee's
 * RSVP. When that's populated we bucket attendees accordingly. When it
 * isn't (e.g. attendee-side view of someone else's meeting, or compose
 * mode), everyone falls into "No Response".
 */
function getAttendeesViaOfficeJs(item) {
  return new Promise(async resolve => {
    const result = {
      'Accepted':    new Map(),
      'Tentative':   new Map(),
      'Declined':    new Map(),
      'No Response': new Map(),
    };

    // Office.MailboxEnums.ResponseType → bucket
    const RESPONSE_MAP = {
      'accepted':   'Accepted',
      'tentative':  'Tentative',
      'declined':   'Declined',
      'none':       'No Response',
      'organizer':  null,        // skip — captured via item.organizer
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
      if (!name) continue;

      const rt = (a.appointmentResponse || 'none').toString().toLowerCase();
      const bucket = RESPONSE_MAP.hasOwnProperty(rt) ? RESPONSE_MAP[rt] : 'No Response';
      if (bucket === null) continue;
      result[bucket].set(name, email);
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

    // In New Outlook on Windows (and some Outlook on the web configurations)
    // item.itemId is REST-formatted. EWS requires the EWS-formatted ID.
    // convertToEwsId is a no-op when the ID is already EWS-formatted in
    // recent Outlook builds, but we guard with try/catch defensively.
    let ewsId = itemId;
    try {
      const converted = Office.context.mailbox.convertToEwsId(
        itemId,
        Office.MailboxEnums.RestVersion.v2_0
      );
      if (converted) ewsId = converted;
    } catch (_) {
      // Already EWS format, or conversion unavailable — use as-is.
    }

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
        <t:ItemId Id="${ewsId}"/>
      </m:ItemIds>
    </m:GetItem>
  </soap:Body>
</soap:Envelope>`;

    Office.context.mailbox.makeEwsRequestAsync(soap, result => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        const err  = result.error || {};
        const code = err.code != null ? ` code=${err.code}` : '';
        const name = err.name ? ` name=${err.name}` : '';
        const body = typeof result.value === 'string' && result.value.length > 0
          ? ' body=' + result.value.replace(/\s+/g, ' ').slice(0, 400)
          : '';
        reject(new Error((err.message || 'EWS request failed') + code + name + body));
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
