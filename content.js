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
    /* TITLE/SUBJECT */
    let title = '';
    let titleSelectors = [
      'span.CWGkB',
      'div.FZzLA span',
      'span[class*="CWGkB"]',
      'div[class*="FZzLA"] span',
      'input[aria-label*="subject" i]',
      'input[placeholder*="subject" i]'
    ];
    
    for (let selector of titleSelectors) {
      try {
        let element = document.querySelector(selector);
        if (element) {
          let text = element.value || element.textContent;
          if (text && text.trim()) {
            title = text.trim();
            break;
          }
        }
      } catch(e) {}
    }
    
    if (!title) title = 'Untitled Meeting';
    
    /* DATE/TIME */
    let time = '';
    let meetingDate = null;
    let timeSelectors = [
      'div.y79CD',
      'div[class*="y79CD"]',
      'button[aria-label*="Start time" i]',
      'input[aria-label*="start" i][aria-label*="time" i]'
    ];
    
    for (let selector of timeSelectors) {
      let element = document.querySelector(selector);
      if (element) {
        let text = element.value || element.getAttribute('aria-label') || element.textContent;
        if (text && text.trim()) {
          text = text.trim();
          if (text.length < 200 && !text.match(/SMTWTFS/)) {
            time = text;
            
            let dateMatch = text.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
            if (dateMatch) {
              let month = dateMatch[1].padStart(2, '0');
              let day = dateMatch[2].padStart(2, '0');
              let year = dateMatch[3];
              meetingDate = {
                year: year,
                month: month,
                day: day,
                full: `${year}-${month}-${day}`
              };
            }
            break;
          }
        }
      }
    }
    
    if (!meetingDate) {
      let today = new Date();
      meetingDate = {
        year: today.getFullYear().toString(),
        month: (today.getMonth() + 1).toString().padStart(2, '0'),
        day: today.getDate().toString().padStart(2, '0'),
        full: today.toISOString().split('T')[0]
      };
    }
    
    /* LOCATION */
    let location = '';
    let locationSelectors = [
      'input[aria-label*="location" i]',
      'button[aria-label*="location" i]',
      'div[aria-label*="location" i]:not([class*="calendar"])',
      'span[class*="location" i]:not([class*="calendar"])',
      'input[placeholder*="location" i]'
    ];
    
    for (let selector of locationSelectors) {
      let element = document.querySelector(selector);
      if (element) {
        let text = element.value || element.textContent;
        if (text) {
          location = text.replace(/^Location\s*/i, '').trim();
          break;
        }
      }
    }
    
    /* ORGANIZER */
    let organizer = '';
    let organizerEmail = '';
    let organizerElement = document.querySelector('div.lpxUr.EKtKK span.fui-Persona__primaryText');
    if (!organizerElement) {
      organizerElement = document.querySelector('div[class*="EKtKK"] span[class*="Persona__primaryText"]');
    }
    if (organizerElement) {
      organizer = organizerElement.textContent.trim();
      let parentElement = organizerElement.closest('div.lpxUr.EKtKK') || organizerElement.closest('span[aria-label]');
      if (parentElement) {
        let ariaLabel = parentElement.getAttribute('aria-label') || '';
        let emailMatch = ariaLabel.match(/[\w\.-]+@[\w\.-]+\.\w+/);
        if (emailMatch) {
          organizerEmail = emailMatch[0].toLowerCase();
        }
      }
    }
    
    /* MEETING BODY */
    let body = '';
    let bodySelectors = [
      'div[role="textbox"][aria-label*="message" i]',
      'div[role="textbox"][contenteditable="true"]',
      'div[class*="body"][class*="content"]',
      'div[aria-label*="description" i]',
      'div[aria-label*="Message body"]'
    ];
    
    for (let selector of bodySelectors) {
      let element = document.querySelector(selector);
      if (element) {
        body = element.innerText || element.textContent;
        if (body) break;
      }
    }
    
    /* ATTENDEES WITH SCROLLING */
    let attendeesByStatus = {
      'Accepted': new Map(),
      'Tentative': new Map(),
      'Declined': new Map(),
      'No Response': new Map()
    };
    
    let hasExternalAttendees = false;
    
    /* Find scrollable container */
    let scrollContainer = null;
    let possibleContainers = [
      document.querySelector('div[role="tree"]'),
      document.querySelector('div[class*="scrollable"]'),
      ...Array.from(document.querySelectorAll('div')).filter(el => {
        let style = window.getComputedStyle(el);
        return (style.overflowY === 'auto' || style.overflowY === 'scroll') && 
               el.querySelector('div.lpxUr');
      })
    ].filter(Boolean);
    
    for (let container of possibleContainers) {
      if (container && container.scrollHeight > container.clientHeight) {
        scrollContainer = container;
        break;
      }
    }
    
    let currentRSVPStatus = 'No Response';
    let processedNames = new Set();
    
    if (scrollContainer) {
      scrollContainer.scrollTop = 0;
      await new Promise(resolve => setTimeout(resolve, 200));
      
      let lastCount = 0;
      let totalCount = 0;
      let noChangeCount = 0;
      let scrollAttempts = 0;
      const maxScrollAttempts = 50;
      
      while (scrollAttempts < maxScrollAttempts) {
        let allVisibleElements = scrollContainer.querySelectorAll('div.ELl7R.TqkD6, div.lpxUr.VCV1f.sKKBX');
        
        allVisibleElements.forEach(element => {
          if (element.classList.contains('ELl7R')) {
            let headerText = element.textContent || element.getAttribute('aria-label') || '';
            
            if (headerText.toLowerCase().includes('accepted')) {
              currentRSVPStatus = 'Accepted';
            } else if (headerText.toLowerCase().includes('tentative')) {
              currentRSVPStatus = 'Tentative';
            } else if (headerText.toLowerCase().includes('declined')) {
              currentRSVPStatus = 'Declined';
            } else if (headerText.toLowerCase().includes('no response') || headerText.toLowerCase().includes('no-response')) {
              currentRSVPStatus = 'No Response';
            }
          } else if (element.classList.contains('lpxUr')) {
            let nameElement = element.querySelector('span.fui-Persona__primaryText, span[class*="Persona__primaryText"]');
            
            if (nameElement) {
              let name = nameElement.textContent.trim();
              
              if (name && name.length > 0 && !processedNames.has(name)) {
                let email = '';
                
                let parentSpan = element.closest('span[aria-label]');
                if (parentSpan) {
                  let ariaLabel = parentSpan.getAttribute('aria-label') || '';
                  let emailMatch = ariaLabel.match(/[\w\.-]+@[\w\.-]+\.\w+/);
                  if (emailMatch) {
                    email = emailMatch[0].toLowerCase();
                  }
                }
                
                if (!email) {
                  let elementsWithTitle = element.querySelectorAll('[title]');
                  for (let titleEl of elementsWithTitle) {
                    let titleText = titleEl.getAttribute('title');
                    if (titleText && titleText.includes('@')) {
                      let emailMatch = titleText.match(/[\w\.-]+@[\w\.-]+\.\w+/);
                      if (emailMatch) {
                        email = emailMatch[0].toLowerCase();
                        break;
                      }
                    }
                  }
                }
                
                if (email && !email.endsWith('@syntax.com')) {
                  hasExternalAttendees = true;
                }
                
                attendeesByStatus[currentRSVPStatus].set(name, email);
                processedNames.add(name);
              }
            }
          }
        });
        
        totalCount = processedNames.size;
        
        if (totalCount === lastCount) {
          noChangeCount++;
          if (noChangeCount > 3) break;
        } else {
          noChangeCount = 0;
          lastCount = totalCount;
        }
        
        let previousScrollTop = scrollContainer.scrollTop;
        scrollContainer.scrollTop += scrollContainer.clientHeight * 0.8;
        
        await new Promise(resolve => setTimeout(resolve, 300));
        
        if (scrollContainer.scrollTop === previousScrollTop) break;
        
        scrollAttempts++;
      }
      
      scrollContainer.scrollTop = 0;
    } else {
      // Process visible elements without scrolling
      let allElements = document.querySelectorAll('div.ELl7R.TqkD6, div.lpxUr.VCV1f.sKKBX');
      
      allElements.forEach(element => {
        if (element.classList.contains('ELl7R')) {
          let headerText = element.textContent || element.getAttribute('aria-label') || '';
          
          if (headerText.toLowerCase().includes('accepted')) {
            currentRSVPStatus = 'Accepted';
          } else if (headerText.toLowerCase().includes('tentative')) {
            currentRSVPStatus = 'Tentative';
          } else if (headerText.toLowerCase().includes('declined')) {
            currentRSVPStatus = 'Declined';
          } else if (headerText.toLowerCase().includes('no response')) {
            currentRSVPStatus = 'No Response';
          }
        } else if (element.classList.contains('lpxUr')) {
          let nameElement = element.querySelector('span.fui-Persona__primaryText, span[class*="Persona__primaryText"]');
          
          if (nameElement) {
            let name = nameElement.textContent.trim();
            if (name && name.length > 0) {
              let email = '';
              let parentSpan = element.closest('span[aria-label]');
              if (parentSpan) {
                let ariaLabel = parentSpan.getAttribute('aria-label') || '';
                let emailMatch = ariaLabel.match(/[\w\.-]+@[\w\.-]+\.\w+/);
                if (emailMatch) {
                  email = emailMatch[0].toLowerCase();
                  if (!email.endsWith('@syntax.com')) {
                    hasExternalAttendees = true;
                  }
                }
              }
              attendeesByStatus[currentRSVPStatus].set(name, email);
            }
          }
        }
      });
    }
    
    if (organizerEmail && !organizerEmail.endsWith('@syntax.com')) {
      hasExternalAttendees = true;
    }
    
    /* Format the note */
    let acceptedAttendeesList = Array.from(attendeesByStatus['Accepted'].keys()).sort();
    if (organizer) {
      acceptedAttendeesList.unshift(organizer);
    }
    
    let attendeesFormatted = [];
    let totalAttendees = (organizer ? 1 : 0);
    
    if (organizer) {
      attendeesFormatted.push(`**Organizer:** ${organizer}`);
    }
    
    let statusOrder = ['Accepted', 'Tentative', 'No Response', 'Declined'];
    
    for (let status of statusOrder) {
      let attendeeList = Array.from(attendeesByStatus[status].keys()).sort();
      if (attendeeList.length > 0) {
        totalAttendees += attendeeList.length;
        attendeesFormatted.push(`**${status} (${attendeeList.length}):** ${attendeeList.join(', ')}`);
      }
    }
    
    let attendeesStr = attendeesFormatted.join('\n\n');
    if (!attendeesStr) {
      attendeesStr = 'No attendees found';
    }
    
    let noteLines = [
      '---',
      'tags: meeting',
      `date: ${meetingDate.full}`,
      'type: outlook-meeting',
      `external: ${hasExternalAttendees}`,
      `attendees: [${acceptedAttendeesList.map(name => `"${name}"`).join(', ')}]`,
	  'summary: ',
      '---',
      '',
      `# ${title}`,
      '',
      '## Meeting Details',
      `**Date/Time:** ${time || 'Not specified'}`,
      location ? `**Location:** ${location}` : '',
      `**Total Attendees:** ${totalAttendees}`,
      hasExternalAttendees ? `**External Meeting:** Yes` : '',
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
      ''
    ];
    
    let note = noteLines.filter(line => line !== null).join('\r\n');
    
    let safeTitle = title.replace(/[<>:"/\\|?*]/g, '-');
    
    return {
      success: true,
      data: {
        title: title,
        time: time,
        location: location,
        totalAttendees: totalAttendees,
        hasExternalAttendees: hasExternalAttendees,
        meetingDate: meetingDate,
        safeTitle: safeTitle,
        note: note
      }
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}