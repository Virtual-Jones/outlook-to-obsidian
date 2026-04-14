// Default settings
const DEFAULT_SETTINGS = {
  vaultName: 'Syntax',
  folderPath: '02. Meeting Notes',
  autoClose: true,
  copyToClipboard: true,
  includeExternal: false
};

// Load settings on page load
document.addEventListener('DOMContentLoaded', loadSettings);

// Update preview when folder path changes
document.getElementById('folderPath').addEventListener('input', updatePreview);
document.getElementById('vaultName').addEventListener('input', updatePreview);

// Handle preset buttons
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

// Save button
document.getElementById('saveBtn').addEventListener('click', saveSettings);

// Reset button
document.getElementById('resetBtn').addEventListener('click', function() {
  if (confirm('Reset all settings to defaults?')) {
    chrome.storage.sync.set(DEFAULT_SETTINGS, function() {
      loadSettings();
      showStatus('Settings reset to defaults', 'success');
    });
  }
});

function loadSettings() {
  chrome.storage.sync.get(DEFAULT_SETTINGS, function(items) {
    document.getElementById('vaultName').value = items.vaultName;
    document.getElementById('folderPath').value = items.folderPath;
    document.getElementById('autoClose').checked = items.autoClose;
    document.getElementById('copyToClipboard').checked = items.copyToClipboard;
    document.getElementById('includeExternal').checked = items.includeExternal;
    updatePreview();
  });
}

function saveSettings() {
  const settings = {
    vaultName: document.getElementById('vaultName').value.trim(),
    folderPath: document.getElementById('folderPath').value.trim(),
    autoClose: document.getElementById('autoClose').checked,
    copyToClipboard: document.getElementById('copyToClipboard').checked,
    includeExternal: document.getElementById('includeExternal').checked
  };
  
  // Validate
  if (!settings.vaultName) {
    showStatus('Please enter a vault name', 'error');
    return;
  }
  
  if (!settings.folderPath) {
    showStatus('Please enter a folder path', 'error');
    return;
  }
  
  // Save
  chrome.storage.sync.set(settings, function() {
    showStatus('Settings saved successfully!', 'success');
  });
}

function updatePreview() {
  const vaultName = document.getElementById('vaultName').value || 'YourVault';
  const folderPath = document.getElementById('folderPath').value || 'Meeting Notes';
  const today = new Date();
  const year = today.getFullYear();
  const month = String(today.getMonth() + 1).padStart(2, '0');
  
  const preview = `
📁 ${vaultName}/
  📁 ${folderPath}/
    📁 ${year}/
      📁 ${month}/
        📄 ${year}-${month}-15 Example Meeting.md
        📄 ${year}-${month}-20 Another Meeting.md
  `;
  
  document.getElementById('folderPreview').textContent = preview;
}

function showStatus(message, type) {
  const statusDiv = document.getElementById('status');
  statusDiv.textContent = message;
  statusDiv.className = `status ${type}`;
  
  setTimeout(() => {
    statusDiv.className = 'status';
  }, 3000);
}