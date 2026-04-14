let extractedData = null;
let currentSettings = {};

document.addEventListener('DOMContentLoaded', function() {
  loadAndDisplaySettings();
  
  // Settings button
  document.getElementById('settingsBtn').addEventListener('click', function() {
    chrome.runtime.openOptionsPage();
  });
});

function loadAndDisplaySettings() {
  chrome.storage.sync.get({
    vaultName: '',
    folderPath: '',
    autoClose: true,
    copyToClipboard: true,
    includeExternal: false
  }, function(items) {
    currentSettings = items;
    
    if (!items.vaultName || !items.folderPath) {
      document.getElementById('configWarning').style.display = 'block';
      document.getElementById('vaultDisplay').textContent = 'Not configured';
      document.getElementById('pathDisplay').textContent = 'Not configured';
      document.getElementById('extractBtn').disabled = true;
    } else {
      document.getElementById('configWarning').style.display = 'none';
      document.getElementById('vaultDisplay').textContent = items.vaultName;
      document.getElementById('pathDisplay').textContent = items.folderPath;
      document.getElementById('extractBtn').disabled = false;
    }
  });
}

// Function to inject content script if needed
async function ensureContentScript(tabId) {
  try {
    // Try to send a ping message first
    return new Promise((resolve) => {
      chrome.tabs.sendMessage(tabId, {action: 'ping'}, function(response) {
        if (chrome.runtime.lastError) {
          // Content script not loaded, inject it
          chrome.scripting.executeScript({
            target: {tabId: tabId},
            files: ['content.js']
          }, () => {
            if (chrome.runtime.lastError) {
              resolve(false);
            } else {
              // Give it a moment to initialize
              setTimeout(() => resolve(true), 100);
            }
          });
        } else {
          // Content script already loaded
          resolve(true);
        }
      });
    });
  } catch (error) {
    console.error('Error ensuring content script:', error);
    return false;
  }
}

document.getElementById('extractBtn').addEventListener('click', async function() {
  const btn = this;
  const statusDiv = document.getElementById('status');
  const detailsDiv = document.getElementById('details');
  const createBtn = document.getElementById('createObsidianBtn');
  
  btn.disabled = true;
  statusDiv.className = 'status info';
  statusDiv.textContent = 'Checking page...';
  detailsDiv.style.display = 'none';
  createBtn.style.display = 'none';
  
  try {
    const [tab] = await chrome.tabs.query({active: true, currentWindow: true});
    
    // Check if we're on an Outlook page
    if (!tab.url || !tab.url.match(/https:\/\/(outlook\.office\.com|outlook\.office365\.com|outlook\.live\.com)/)) {
      statusDiv.className = 'status error';
      statusDiv.textContent = 'Please navigate to an Outlook Web meeting invitation first';
      btn.disabled = false;
      return;
    }
    
    // Ensure content script is injected
    statusDiv.textContent = 'Preparing extraction...';
    const scriptReady = await ensureContentScript(tab.id);
    
    if (!scriptReady) {
      statusDiv.className = 'status error';
      statusDiv.textContent = 'Could not access the page. Please refresh and try again.';
      btn.disabled = false;
      return;
    }
    
    statusDiv.textContent = 'Extracting meeting details...';
    
    // Send extraction message
    chrome.tabs.sendMessage(tab.id, {action: 'extractMeeting'}, function(response) {
      if (chrome.runtime.lastError) {
        console.error('Runtime error:', chrome.runtime.lastError);
        statusDiv.className = 'status error';
        statusDiv.textContent = 'Failed to extract. Please refresh the page and try again.';
        btn.disabled = false;
        return;
      }
      
      if (response && response.success) {
        extractedData = response.data;
        
        // Add external marker to title if configured
        if (currentSettings.includeExternal && response.data.hasExternalAttendees) {
          extractedData.title = '[EXTERNAL] ' + extractedData.title;
          // Update the note content with the modified title
          let noteLines = extractedData.note.split('\r\n');
          for (let i = 0; i < noteLines.length; i++) {
            if (noteLines[i].startsWith('# ')) {
              noteLines[i] = `# ${extractedData.title}`;
              break;
            }
          }
          extractedData.note = noteLines.join('\r\n');
        }
        
        statusDiv.className = 'status success';
        statusDiv.textContent = '✓ Meeting details extracted!';
        
        // Show details
        detailsDiv.style.display = 'block';
        detailsDiv.innerHTML = `
          <div class="detail-row">
            <span class="detail-label">Title:</span>
            <span class="detail-value">${response.data.title.substring(0, 50)}${response.data.title.length > 50 ? '...' : ''}</span>
          </div>
          <div class="detail-row">
            <span class="detail-label">Date:</span>
            <span class="detail-value">${response.data.time || 'Not specified'}</span>
          </div>
          <div class="detail-row">
            <span class="detail-label">Attendees:</span>
            <span class="detail-value">${response.data.totalAttendees}</span>
          </div>
          <div class="detail-row">
            <span class="detail-label">External:</span>
            <span class="detail-value">${response.data.hasExternalAttendees ? 'Yes' : 'No'}</span>
          </div>
        `;
        
        // Copy to clipboard if configured
        if (currentSettings.copyToClipboard) {
          copyToClipboard(extractedData.note);
        }
        
        // Show create button
        createBtn.style.display = 'block';
      } else {
        statusDiv.className = 'status error';
        statusDiv.textContent = response ? (response.error || 'Failed to extract meeting details') : 'No response from page';
      }
      
      btn.disabled = false;
    });
  } catch (error) {
    console.error('Extraction error:', error);
    statusDiv.className = 'status error';
    statusDiv.textContent = 'Error: ' + error.message;
    btn.disabled = false;
  }
});

document.getElementById('createObsidianBtn').addEventListener('click', function() {
  if (!extractedData) return;
  
  // Create folder path with year/month
  const folderPath = `${currentSettings.folderPath}/${extractedData.meetingDate.year}/${extractedData.meetingDate.month}`;
  
  // Create filename
  const safeTitle = extractedData.title.replace(/[<>:"/\\|?*]/g, '-');
  const filename = `${extractedData.meetingDate.full} ${safeTitle}`;
  
  // Create Obsidian URI
  const obsidianUri = `obsidian://new?vault=${encodeURIComponent(currentSettings.vaultName)}&file=${encodeURIComponent(folderPath + '/' + filename)}&content=${encodeURIComponent(extractedData.note)}`;
  
  // Open Obsidian
  chrome.tabs.create({url: obsidianUri});
  
  // Close popup if configured
  if (currentSettings.autoClose) {
    window.close();
  }
});

function copyToClipboard(text) {
  const textarea = document.createElement('textarea');
  textarea.value = text;
  document.body.appendChild(textarea);
  textarea.select();
  document.execCommand('copy');
  document.body.removeChild(textarea);
}