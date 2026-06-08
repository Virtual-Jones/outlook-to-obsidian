// Respond to ping messages to check if content script is loaded
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'ping') {
    sendResponse({status: 'ready'});
    return true;
  }

  if (request.action === 'extractMeeting') {
    extractMeetingDetails().then(sendResponse);
    return true; // Keep the message channel open for async response
  }
});

async function extractMeetingDetails() {
  try {

    /* ── TITLE / SUBJECT ────────────────────────────────────────────────
     * Priority: ARIA-labelled inputs (edit view) → heading in event panel
     * (read view) → Fluent UI text in panel → legacy hashed class names
     * ─────────────────────────────────────────────────────────────────── */
    let title = '';

    // Edit / compose view: the subject field is usually a labelled input
    const titleInputSelectors = [
      'input[aria-label="Subject"]',
      'input[aria-label*="subject" i]',
      'input[aria-label*="title" i]',
      'div[role="textbox"][aria-label*="subject" i]',
      'textarea[aria-label*="subject" i]',
      'input[placeholder*="subject" i]',
      'input[placeholder*="title" i]',
    ];
    for (const sel of titleInputSelectors) {
      try {
        const el = document.querySelector(sel);
        if (el) {
          const text = (el.value || el.textContent || '').trim();
          if (text) { title = text; break; }
        }
      } catch (_) {}
    }

    // Read view: prominent heading inside the event panel / dialog
    if (!title) {
      const panelSel = [
        '[role="dialog"]',
        '[role="complementary"]',
        '[data-app-section="eventPanel"]',
        '[class*="eventPanel" i]',
        '[class*="EventPanel"]',
        '[class*="event-details" i]',
      ].join(',');
      const panel = document.querySelector(panelSel);
      if (panel) {
        for (const lvl of [
          'h1',
          'h2',
          '[role="heading"][aria-level="1"]',
          '[role="heading"][aria-level="2"]',
          '[role="heading"]',
        ]) {
          const h = panel.querySelector(lvl);
          if (h) {
            const text = (h.textContent || '').trim();
            if (text && text.length < 300) { title = text; break; }
          }
        }
      }
    }

    // Legacy hashed class names — kept as last resort
    if (!title) {
      for (const sel of [
        'span.CWGkB',
        'div.FZzLA span',
        'span[class*="CWGkB"]',
        'div[class*="FZzLA"] span',
      ]) {
        try {
          const el = document.querySelector(sel);
          if (el) {
            const text = (el.value || el.textContent || '').trim();
            if (text) { title = text; break; }
          }
        } catch (_) {}
      }
    }

    if (!title) title = 'Untitled Meeting';

    /* ── DATE / TIME ────────────────────────────────────────────────────
     * Priority: ARIA-labelled date/time inputs or buttons → <time> element
     * → legacy hashed class names
     * ─────────────────────────────────────────────────────────────────── */
    let time = '';
    let meetingDate = null;

    const timeSelectors = [
      'input[aria-label="Start date"]',
      'input[aria-label*="start date" i]',
      'button[aria-label*="start date" i]',
      'input[aria-label*="start time" i]',
      'button[aria-label*="start time" i]',
      'input[aria-label*="Start" i]',
      'button[aria-label*="Start" i]',
      // Legacy
      'div.y79CD',
      'div[class*="y79CD"]',
    ];

    for (const sel of timeSelectors) {
      const el = document.querySelector(sel);
      if (el) {
        const text = (el.value || el.getAttribute('aria-label') || el.textContent || '').trim();
        if (text && text.length < 200 && !text.match(/SMTWTFS/)) {
          time = text;
          break;
        }
      }
    }

    // Fallback: <time datetime="…"> element
    if (!time) {
      const timeEl = document.querySelector('time[datetime]');
      if (timeEl) {
        time = (timeEl.textContent || timeEl.getAttribute('datetime') || '').trim();
      }
    }

    // Parse date from whatever time string we captured
    if (time) {
      // MM/DD/YYYY
      let m = time.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (m) {
        meetingDate = {
          year: m[3],
          month: m[1].padStart(2, '0'),
          day: m[2].padStart(2, '0'),
          full: `${m[3]}-${m[1].padStart(2, '0')}-${m[2].padStart(2, '0')}`,
        };
      }
      // YYYY-MM-DD (ISO / datetime attribute)
      if (!meetingDate) {
        m = time.match(/(\d{4})-(\d{2})-(\d{2})/);
        if (m) {
          meetingDate = { year: m[1], month: m[2], day: m[3], full: `${m[1]}-${m[2]}-${m[3]}` };
        }
      }
      // D Month YYYY  (e.g. "15 April 2026")
      if (!meetingDate) {
        m = time.match(/(\d{1,2})\s+(\w+)\s+(\d{4})/);
        if (m) {
          const months = { january:'01',february:'02',march:'03',april:'04',may:'05',june:'06',
                           july:'07',august:'08',september:'09',october:'10',november:'11',december:'12' };
          const mo = months[m[2].toLowerCase()];
          if (mo) {
            meetingDate = {
              year: m[3],
              month: mo,
              day: m[1].padStart(2, '0'),
              full: `${m[3]}-${mo}-${m[1].padStart(2, '0')}`,
            };
          }
        }
      }
    }

    if (!meetingDate) {
      const today = new Date();
      meetingDate = {
        year: today.getFullYear().toString(),
        month: (today.getMonth() + 1).toString().padStart(2, '0'),
        day: today.getDate().toString().padStart(2, '0'),
        full: today.toISOString().split('T')[0],
      };
    }

    /* ── LOCATION ────────────────────────────────────────────────────── */
    let location = '';
    const locationSelectors = [
      'input[aria-label="Location"]',
      'input[aria-label*="location" i]',
      'button[aria-label*="location" i]',
      'div[role="textbox"][aria-label*="location" i]',
      'div[aria-label*="location" i]:not([class*="calendar"])',
      'span[aria-label*="location" i]:not([class*="calendar"])',
      'input[placeholder*="location" i]',
    ];

    for (const sel of locationSelectors) {
      const el = document.querySelector(sel);
      if (el) {
        const text = (el.value || el.textContent || '').replace(/^Location\s*/i, '').trim();
        if (text) { location = text; break; }
      }
    }

    /* ── ORGANIZER ───────────────────────────────────────────────────── */
    let organizer = '';
    let organizerEmail = '';

    // Helper: extract first email address found in an element's ARIA labels or title attrs
    function extractEmail(el) {
      let anc = el;
      while (anc && anc !== document.body) {
        const label = anc.getAttribute('aria-label') || '';
        const m = label.match(/[\w.\-+]+@[\w.-]+\.\w+/);
        if (m) return m[0].toLowerCase();
        anc = anc.parentElement;
      }
      for (const t of el.querySelectorAll('[title]')) {
        const tv = t.getAttribute('title') || '';
        if (tv.includes('@')) {
          const m = tv.match(/[\w.\-+]+@[\w.-]+\.\w+/);
          if (m) return m[0].toLowerCase();
        }
      }
      const dataEmail = el.getAttribute('data-email') || '';
      return dataEmail ? dataEmail.toLowerCase() : '';
    }

    // Try to find organizer: look for a persona inside an "organizer" labelled region
    const organizerEl =
      document.querySelector('[aria-label*="organizer" i] [class*="Persona__primaryText"]') ||
      document.querySelector('[aria-label*="organizer" i] [class*="fui-Persona"]') ||
      // Legacy hashed classes
      document.querySelector('div.lpxUr.EKtKK span.fui-Persona__primaryText') ||
      document.querySelector('div[class*="EKtKK"] span[class*="Persona__primaryText"]');

    if (organizerEl) {
      organizer = organizerEl.textContent.trim();
      const emailHost = organizerEl.closest('[aria-label]') || organizerEl.closest('[title]');
      if (emailHost) {
        const label = emailHost.getAttribute('aria-label') || emailHost.getAttribute('title') || '';
        const m = label.match(/[\w.\-+]+@[\w.-]+\.\w+/);
        if (m) organizerEmail = m[0].toLowerCase();
      }
      if (!organizerEmail) organizerEmail = extractEmail(organizerEl.closest('div') || organizerEl);
    }

    /* ── MEETING BODY / DESCRIPTION ─────────────────────────────────── */
    let body = '';
    const bodySelectors = [
      'div[aria-label="Message body"][role="textbox"]',
      'div[role="textbox"][aria-label*="message body" i]',
      'div[aria-label*="description" i][role="textbox"]',
      'div[role="textbox"][aria-label*="description" i]',
      '[aria-label*="Message body"]',
      '[aria-label*="description" i]',
      'div[role="textbox"][contenteditable="true"]',
      'div[class*="body"][class*="content"]',
    ];

    for (const sel of bodySelectors) {
      const el = document.querySelector(sel);
      if (el) {
        const text = (el.innerText || el.textContent || '').trim();
        if (text) { body = text; break; }
      }
    }

    /* ── ATTENDEES WITH RSVP GROUPING ────────────────────────────────── */
    const attendeesByStatus = {
      'Accepted': new Map(),
      'Tentative': new Map(),
      'Declined': new Map(),
      'No Response': new Map(),
    };
    let hasExternalAttendees = false;
    const processedNames = new Set();

    // Map text → RSVP status bucket
    function getRSVPStatus(text) {
      const t = (text || '').toLowerCase();
      if (t.includes('accepted')) return 'Accepted';
      if (t.includes('tentative')) return 'Tentative';
      if (t.includes('declined')) return 'Declined';
      if (t.includes('no response') || t.includes('no-response') || t.includes('awaiting')) return 'No Response';
      return null;
    }

    // Add a persona element to the right bucket
    function processPersona(el, status) {
      // The name lives inside a Persona__primaryText span (Fluent UI, stable prefix)
      const nameEl =
        el.querySelector('[class*="Persona__primaryText"]') ||
        el.querySelector('[class*="persona-primary" i]');
      if (!nameEl) return;

      const name = nameEl.textContent.trim();
      if (!name || processedNames.has(name)) return;

      const email = extractEmail(el);
      if (email && !email.endsWith('@syntax.com')) hasExternalAttendees = true;

      attendeesByStatus[status].set(name, email);
      processedNames.add(name);
    }

    /* Find the scrollable attendee container.
     * Strategy (in order):
     *  1. role="tree"  — Outlook renders the RSVP list as an ARIA tree
     *  2. role="list" / aria-label*="attendee"
     *  3. Walk up from the first Persona__primaryText element until we find
     *     an overflow:auto/scroll ancestor
     *  4. Legacy: any overflow div that contains div.lpxUr
     */
    let scrollContainer = null;

    const containerCandidates = [
      document.querySelector('[role="tree"]'),
      document.querySelector('[aria-label*="attendee" i][role="list"]'),
      document.querySelector('[role="list"][aria-label*="attendee" i]'),
      document.querySelector('[aria-label*="attendee" i]'),
    ].filter(Boolean);

    if (containerCandidates.length === 0) {
      // Walk up from first visible persona element
      const firstPersona = document.querySelector('[class*="Persona__primaryText"]');
      if (firstPersona) {
        let anc = firstPersona.parentElement;
        while (anc && anc !== document.body) {
          const s = window.getComputedStyle(anc);
          if ((s.overflowY === 'auto' || s.overflowY === 'scroll') && anc.scrollHeight > anc.clientHeight) {
            containerCandidates.push(anc);
            break;
          }
          anc = anc.parentElement;
        }
      }
    }

    if (containerCandidates.length === 0) {
      // Legacy fallback: any scrollable div with div.lpxUr inside
      for (const div of document.querySelectorAll('div')) {
        const s = window.getComputedStyle(div);
        if ((s.overflowY === 'auto' || s.overflowY === 'scroll') && div.querySelector('div.lpxUr')) {
          containerCandidates.push(div);
          break;
        }
      }
    }

    for (const c of containerCandidates) {
      if (c && c.scrollHeight >= c.clientHeight) {
        scrollContainer = c;
        break;
      }
    }
    if (!scrollContainer && containerCandidates.length > 0) {
      scrollContainer = containerCandidates[0];
    }

    /* Process attendees from the currently-visible portion of the list.
     *
     * Element classification (in DOM order):
     *  • RSVP header  → element whose trimmed text is a short RSVP label
     *                    AND that does NOT itself contain a Persona span
     *  • Attendee row → element that DOES contain a Persona__primaryText span
     *
     * We also accept ARIA tree / list roles as structural signals.
     */
    function processAttendeeRoot(root) {
      let currentStatus = 'No Response';

      // Build a flat ordered list of relevant nodes
      const candidates = root.querySelectorAll([
        // ARIA structural roles
        '[role="treeitem"]',
        '[role="listitem"]',
        '[role="group"]',
        '[role="heading"]',
        // Fluent UI Persona rows (new Outlook)
        '[class*="Persona__primaryText"]',
        // Legacy hashed rows
        'div.ELl7R.TqkD6',
        'div.lpxUr.VCV1f.sKKBX',
        'div.lpxUr',
      ].join(','));

      for (const el of candidates) {
        // Is this a group / heading acting as an RSVP section header?
        const isGroupOrHeading =
          el.getAttribute('role') === 'group' ||
          el.getAttribute('role') === 'heading' ||
          el.classList.contains('ELl7R');

        if (isGroupOrHeading) {
          const status = getRSVPStatus(el.textContent);
          if (status) { currentStatus = status; }
          continue;
        }

        // Does it contain a persona name but is NOT itself just a name span?
        const hasPersona = !!el.querySelector('[class*="Persona__primaryText"]');
        const isPersonaSpan = el.matches('[class*="Persona__primaryText"]');

        if (hasPersona && !isPersonaSpan) {
          processPersona(el, currentStatus);
          continue;
        }

        // Short element whose text is purely an RSVP label (section header without role)
        const trimmed = el.textContent.trim();
        if (!hasPersona && !isPersonaSpan && trimmed.length < 60) {
          const status = getRSVPStatus(trimmed);
          if (status) { currentStatus = status; }
        }
      }
    }

    if (scrollContainer) {
      scrollContainer.scrollTop = 0;
      await new Promise(r => setTimeout(r, 200));

      let lastCount = 0;
      let noChangeCount = 0;
      let scrollAttempts = 0;
      const maxScrollAttempts = 50;

      while (scrollAttempts < maxScrollAttempts) {
        processAttendeeRoot(scrollContainer);

        const totalCount = processedNames.size;
        if (totalCount === lastCount) {
          noChangeCount++;
          if (noChangeCount > 3) break;
        } else {
          noChangeCount = 0;
          lastCount = totalCount;
        }

        const prevScrollTop = scrollContainer.scrollTop;
        scrollContainer.scrollTop += scrollContainer.clientHeight * 0.8;
        await new Promise(r => setTimeout(r, 300));

        if (scrollContainer.scrollTop === prevScrollTop) break;
        scrollAttempts++;
      }

      scrollContainer.scrollTop = 0;
    } else {
      // No dedicated container — scan the whole document
      processAttendeeRoot(document.body);
    }

    if (organizerEmail && !organizerEmail.endsWith('@syntax.com')) {
      hasExternalAttendees = true;
    }

    /* ── FORMAT NOTE ─────────────────────────────────────────────────── */
    const acceptedAttendeesList = Array.from(attendeesByStatus['Accepted'].keys()).sort();
    if (organizer) acceptedAttendeesList.unshift(organizer);

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
      `attendees: [${acceptedAttendeesList.map(n => `"${n}"`).join(', ')}]`,
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

    const note = noteLines.filter(line => line !== null).join('\r\n');
    const safeTitle = title.replace(/[<>:"/\\|?*]/g, '-');

    return {
      success: true,
      data: { title, time, location, totalAttendees, hasExternalAttendees, meetingDate, safeTitle, note },
    };

  } catch (error) {
    return { success: false, error: error.message };
  }
}
