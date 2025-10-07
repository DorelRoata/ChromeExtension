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