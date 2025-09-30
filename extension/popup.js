const SERVER_URL = 'http://localhost:5000';

async function checkServerStatus() {
  try {
    const response = await fetch(`${SERVER_URL}/ping`, { method: 'GET' });
    if (response.ok) {
      document.getElementById('status').className = 'status connected';
      document.getElementById('status').textContent = 'Server: Connected ✓';
      document.getElementById('scrapeBtn').disabled = false;
      return true;
    }
  } catch (error) {
    document.getElementById('status').className = 'status disconnected';
    document.getElementById('status').textContent = 'Server: Not Running ✗';
    document.getElementById('scrapeBtn').disabled = true;
    return false;
  }
}

document.getElementById('scrapeBtn').addEventListener('click', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  
  const supportedDomains = ['grainger.com', 'mcmaster.com', 'festo.com', 'zoro.com'];
  const isSupported = supportedDomains.some(domain => tab.url.includes(domain));
  
  if (isSupported) {
    chrome.tabs.sendMessage(tab.id, { action: 'scrapeNow' }, (response) => {
      if (response && response.success) {
        document.getElementById('status').textContent = 'Scraped successfully!';
        setTimeout(checkServerStatus, 2000);
      }
    });
  } else {
    alert('Navigate to a supported vendor page first');
  }
});

document.getElementById('testBtn').addEventListener('click', checkServerStatus);

// Check on load
checkServerStatus();