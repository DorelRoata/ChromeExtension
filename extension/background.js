// Keep service worker alive
chrome.runtime.onInstalled.addListener(() => {
  console.log('Multi-Vendor Price Scraper installed');
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