// Local server config
const SERVER_URL = 'http://127.0.0.1:5000';

// Keep service worker alive
chrome.runtime.onInstalled.addListener(() => {
  console.log('Advantage Multi-Vendor Price Scraper installed');
});

// Badge management
chrome.tabs.onUpdated.addListener((tabId, changeInfo, tab) => {
  if (changeInfo.status === 'complete') {
    const supportedDomains = ['grainger.com', 'mcmaster.com', 'festo.com', 'zoro.com'];
    const isSupported = supportedDomains.some(domain => tab.url?.includes(domain));

    if (isSupported) {
      chrome.action.setBadgeText({ text: '●', tabId: tabId });
      chrome.action.setBadgeBackgroundColor({ color: '#4CAF50', tabId: tabId });
    }
  }
});

// Message handlers for tab management
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'getTabId') {
    // Return the tab ID to the content script
    sendResponse(sender.tab?.id);
    return true;
  }

  if (request.action === 'closeSelf' && sender.tab) {
    // Close the tab that sent this message
    chrome.tabs.remove(sender.tab.id);
    sendResponse({ success: true });
    return true;
  }

  return false;
});

// Polling management for close checks
const closePollers = new Map(); // tabId -> intervalId

async function registerTabWithServer(tabId, url) {
  try {
    const resp = await fetch(`${SERVER_URL}/register-tab`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ tabId, url })
    });
    if (!resp.ok) {
      console.warn('register-tab failed:', resp.status, resp.statusText);
    }
    return resp.ok;
  } catch (e) {
    console.warn('register-tab error:', e);
    return false;
  }
}

async function sendScrapeToServer(data) {
  try {
    const resp = await fetch(`${SERVER_URL}/scrape`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!resp.ok) {
      console.warn('scrape failed:', resp.status, resp.statusText);
    }
    return resp.ok;
  } catch (e) {
    console.warn('scrape error:', e);
    return false;
  }
}

function startClosePolling(tabId) {
  if (closePollers.has(tabId)) return; // already polling
  const intervalId = setInterval(async () => {
    try {
      const resp = await fetch(`${SERVER_URL}/should-close/${tabId}`);
      if (!resp.ok) return;
      const data = await resp.json();
      if (data && data.shouldClose) {
        // Close the tab
        chrome.tabs.remove(tabId);
        stopClosePolling(tabId);
      }
    } catch (_) {
      // ignore polling errors
    }
  }, 1500);
  closePollers.set(tabId, intervalId);
}

function stopClosePolling(tabId) {
  const id = closePollers.get(tabId);
  if (id) {
    clearInterval(id);
    closePollers.delete(tabId);
  }
}

// Cleanup when tab is removed
chrome.tabs.onRemoved.addListener((tabId) => {
  stopClosePolling(tabId);
  // Tell server tab closed (best-effort)
  fetch(`${SERVER_URL}/tab-closed`, { method: 'POST' }).catch(() => {});
});

// Message handling for server interactions
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'registerTab') {
    const { tabId, url } = request;
    registerTabWithServer(tabId, url).then((ok) => {
      if (ok) startClosePolling(tabId);
      sendResponse({ success: ok });
    });
    return true;
  }
  if (request.action === 'sendScrape') {
    const payload = Object.assign({}, request.data || {});
    if (!payload.tabId && sender && sender.tab && typeof sender.tab.id === 'number') {
      payload.tabId = sender.tab.id;
    }
    sendScrapeToServer(payload).then((ok) => {
      sendResponse({ success: ok });
    });
    return true;
  }
  if (request.action === 'startClosePolling') {
    if (typeof request.tabId === 'number') startClosePolling(request.tabId);
    sendResponse({ success: true });
    return true;
  }
  if (request.action === 'stopClosePolling') {
    if (typeof request.tabId === 'number') stopClosePolling(request.tabId);
    sendResponse({ success: true });
    return true;
  }
});
