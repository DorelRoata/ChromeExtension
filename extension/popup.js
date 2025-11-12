const SERVER_URL = 'http://127.0.0.1:5000';
let serverConnected = false;

async function checkServerStatus() {
  try {
    const response = await fetch(`${SERVER_URL}/ping`, { method: 'GET' });
    if (response.ok) {
      serverConnected = true;
      document.getElementById('status').className = 'status connected';
      document.getElementById('status').textContent = 'Server: Connected ✓ (Optional)';
      return true;
    }
  } catch (error) {
    serverConnected = false;
    document.getElementById('status').className = 'status disconnected';
    document.getElementById('status').textContent = 'Server: Not Connected (Standalone Mode)';
    return false;
  }
}

function displayScrapedData(data) {
  const resultsDiv = document.getElementById('results');
  const resultsContent = document.getElementById('resultsContent');

  resultsContent.innerHTML = `
    <strong>Vendor:</strong> ${data.vendor || 'N/A'}<br>
    <strong>Part Number:</strong> ${data.partNumber || 'N/A'}<br>
    <strong>Description:</strong> ${data.description || 'N/A'}<br>
    <strong>Price:</strong> ${data.price || 'N/A'}<br>
    <strong>Unit:</strong> ${data.unit || 'N/A'}<br>
    <strong>MFR Number:</strong> ${data.mfrNumber || 'N/A'}<br>
    <strong>Brand:</strong> ${data.brand || 'N/A'}<br>
    <strong>URL:</strong> <a href="${data.url}" target="_blank">View Page</a>
  `;

  resultsDiv.style.display = 'block';
}

document.getElementById('scrapeBtn').addEventListener('click', async () => {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });

  const supportedDomains = ['grainger.com', 'mcmaster.com', 'festo.com', 'zoro.com'];
  const isSupported = supportedDomains.some(domain => tab.url.includes(domain));

  if (isSupported) {
    // Hide previous results
    document.getElementById('results').style.display = 'none';

    chrome.tabs.sendMessage(tab.id, { action: 'scrapeNow' }, async (response) => {
      if (response && response.success) {
        if (serverConnected) {
          document.getElementById('status').textContent = 'Scraped & sent to server ✓';
        } else {
          document.getElementById('status').textContent = 'Data scraped successfully ✓';
        }

        // Get the scraped data from content script
        chrome.tabs.sendMessage(tab.id, { action: 'getData' }, (dataResponse) => {
          if (dataResponse && dataResponse.data) {
            displayScrapedData(dataResponse.data);
          }
        });

        setTimeout(checkServerStatus, 2000);
      } else {
        document.getElementById('status').textContent = 'Error: Could not scrape page';
      }
    });
  } else {
    alert('Navigate to a supported vendor page first (Grainger, McMaster-Carr, Festo, or Zoro)');
  }
});

document.getElementById('testBtn').addEventListener('click', checkServerStatus);

// Check on load
checkServerStatus();
