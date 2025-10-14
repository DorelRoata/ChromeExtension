(function() {
  'use strict';

  const SERVER_URL = 'http://localhost:5000';
  let extractionAttempted = false;

  // Detect vendor from URL
  function detectVendor() {
    const hostname = window.location.hostname;
    if (hostname.includes('grainger.com')) return 'grainger';
    if (hostname.includes('mcmaster.com')) return 'mcmaster';
    if (hostname.includes('festo.com')) return 'festo';
    if (hostname.includes('zoro.com')) return 'zoro';
    return null;
  }

  // Wait for page to fully load
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

  function init() {
    const vendor = detectVendor();
    if (!vendor) return;

    // Register this tab immediately (before scraping)
    registerTab();

    // Wait a bit for dynamic content
    setTimeout(() => {
      if (!extractionAttempted) {
        extractionAttempted = true;
        extractAndSend(vendor);
      }
    }, 2000);
  }

  async function registerTab() {
    try {
      // Get our tab ID
      const tabId = await new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'getTabId' }, resolve);
      });

      if (tabId) {
        // Send a registration message to server
        await fetch(`${SERVER_URL}/register-tab`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ tabId: tabId, url: window.location.href })
        });

        console.log('Tab registered with server');

        // Start polling for close signal immediately
        startPollingForClose(tabId);
      }
    } catch (error) {
      // Server may not be running, fail silently
      console.log('Could not register tab with server');
    }
  }

  function extractAndSend(vendor) {
    try {
      let data;
      
      switch(vendor) {
        case 'grainger':
          data = extractGraingerData();
          break;
        case 'mcmaster':
          data = extractMcMasterData();
          break;
        case 'festo':
          data = extractFestoData();
          break;
        case 'zoro':
          data = extractZoroData();
          break;
        default:
          return;
      }

      if (data) {
        data.vendor = vendor;
        console.log('Scraped data:', data);
        sendToServer(data);
      }
    } catch (error) {
      console.error('Extraction failed:', error);
    }
  }

  // Helper function to try multiple selectors
  function findText(selectors) {
    for (const selector of selectors) {
      const elem = document.querySelector(selector);
      if (elem && elem.textContent.trim()) {
        return elem.textContent.trim();
      }
    }
    return "Not Found";
  }

  function extractPartNumber() {
    const path = window.location.pathname;
    const match = path.match(/\/(?:product|i|a)\/([^\/\?]+)/);
    return match ? match[1] : "";
  }

  // GRAINGER SCRAPER
  function extractGraingerData() {
    // Find price with dollar sign
    let price = "Not Found";
    const priceElements = document.querySelectorAll('.HANkB');
    for (const elem of priceElements) {
      if (elem.textContent.includes('$')) {
        price = elem.textContent.trim();
        break;
      }
    }

    // Find MFR number - look for item number without "Item" prefix
    let mfrNumber = "Not Found";
    const itemLabel = Array.from(document.querySelectorAll('dt')).find(
      dt => dt.textContent.trim().toLowerCase().includes('item')
    );
    if (itemLabel && itemLabel.nextElementSibling) {
      mfrNumber = itemLabel.nextElementSibling.textContent.trim();
    }

    return {
      partNumber: extractPartNumber(),
      description: findText([
        '.ue2KE .tliLs',
        'h1[class*="product"]',
        'h1'
      ]),
      price: price,
      unit: findText([
        '.G32gdF',
        '.FuAZI6 .G32gdF',
        '[class*="unit"]'
      ]),
      mfrNumber: mfrNumber,
      brand: (() => {
        const brandLabel = Array.from(document.querySelectorAll('dt')).find(
          dt => dt.textContent.trim() === 'Brand'
        );
        return brandLabel && brandLabel.nextElementSibling
          ? brandLabel.nextElementSibling.textContent.trim()
          : "Not Found";
      })(),
      url: window.location.href,
      timestamp: new Date().toISOString()
    };
  }

  // MCMASTER SCRAPER
  function extractMcMasterData() {
    const primaryHeader = document.querySelector('[class*="_productDetailHeaderPrimary_"]');
    const secondaryHeader = document.querySelector('[class*="_productDetailHeaderSecondary_"]');
    
    let description = "Not Found";
    if (primaryHeader) {
      description = primaryHeader.textContent.trim();
      if (secondaryHeader) {
        description += '\n' + secondaryHeader.textContent.trim();
      }
    }

    // Try multiple price selectors
    let price = "Not Found";
    const priceSelectors = ['PrceTxt', 'Price_price', 'PrceTierPrceCol', 'Prce', '_price_'];
    for (const sel of priceSelectors) {
      const elem = document.querySelector(`[class*="${sel}"]`);
      if (elem) {
        price = elem.textContent.trim();
        break;
      }
    }

    return {
      partNumber: extractPartNumber(),
      description: description,
      price: price,
      unit: findText([
        '[class*="wrapper--qty-input"]',
        '[class*="UnitOfMeasure"]'
      ]),
      mfrNumber: findText(['[class*="MfrNumber"]']),
      brand: findText(['[class*="Brand"]']),
      url: window.location.href,
      timestamp: new Date().toISOString()
    };
  }

  // FESTO SCRAPER
  function extractFestoData() {
    let description = "Not Found";
    
    // Try different page layouts
    const headlineSelectors = [
      "[class*='product-page-headline']",
      ".main-headline",
      "[class*='article-detail-page-header-headline']"
    ];
    
    const orderCodeSelectors = [
      "[class*='order-code']",
      ".product-summary-article__order-code"
    ];

    for (const sel of headlineSelectors) {
      const headline = document.querySelector(sel);
      if (headline) {
        description = headline.textContent.trim();
        
        for (const codeSel of orderCodeSelectors) {
          const code = document.querySelector(codeSel);
          if (code) {
            description += ' ' + code.textContent.trim();
            break;
          }
        }
        break;
      }
    }

    return {
      partNumber: extractPartNumber(),
      description: description,
      price: findText(["[class*='price-display-text']"]),
      unit: findText(["[class*='quantity-selector-label']"]),
      mfrNumber: findText(["[class*='code']"]),
      brand: "Festo",
      qty: (() => {
        const qtyInput = document.querySelector("[class*='quantity-selector-input']");
        return qtyInput ? qtyInput.value : "Not Found";
      })(),
      url: window.location.href,
      timestamp: new Date().toISOString()
    };
  }

  // ZORO SCRAPER
  function extractZoroData() {
    return {
      partNumber: extractPartNumber(),
      description: findText([
        '[data-za*="product-name"]',
        'h1[class*="product"]',
        'h1'
      ]),
      price: findText([
        '[data-za*="product-price"]',
        '[class*="price"]'
      ]),
      unit: findText([
        '.FuAZI6 .G32gdF',
        '[class*="unit"]'
      ]),
      mfrNumber: findText([
        '[data-za="PDPMfrNo"]',
        '[class*="mfr"]'
      ]),
      brand: findText([
        '[data-za*="product-brand-name"]',
        '[class*="brand"]'
      ]),
      url: window.location.href,
      timestamp: new Date().toISOString()
    };
  }

  async function sendToServer(data) {
    try {
      // Get our tab ID from background script
      const tabId = await new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'getTabId' }, resolve);
      });

      // Include tab ID in data
      data.tabId = tabId;

      const response = await fetch(`${SERVER_URL}/scrape`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      if (response.ok) {
        console.log('Data sent successfully');
        showNotification('✓ Price data captured');
        // Note: Polling already started in registerTab(), no need to start again
      }
    } catch (error) {
      console.error('Server error:', error);
      // Fail silently if server not running
    }
  }

  function startPollingForClose(tabId) {
    let pollCount = 0;
    const maxPolls = 600; // Stop after 10 minutes (600 * 1 second)
    let pollInterval = null;

    const cleanup = () => {
      if (pollInterval) {
        clearInterval(pollInterval);
        pollInterval = null;
      }
      window.removeEventListener('beforeunload', handleBeforeUnload);
    };

    const handleBeforeUnload = () => {
      cleanup();
      notifyServerTabClosed(tabId);
    };

    pollInterval = setInterval(async () => {
      pollCount++;

      // Stop polling after max attempts
      if (pollCount >= maxPolls) {
        cleanup();
        console.log('Stopped polling for tab close');
        notifyServerTabClosed(tabId);
        return;
      }

      try {
        const response = await fetch(`${SERVER_URL}/should-close/${tabId}`);

        if (response.ok) {
          const data = await response.json();

          if (data.shouldClose) {
            cleanup();
            console.log('Closing tab as requested by server');

            // Close this tab
            chrome.runtime.sendMessage({ action: 'closeSelf' });
          }
        }
      } catch (error) {
        // Server may be down, continue polling
      }
    }, 1000); // Poll every 1 second

    // Cleanup when page unloads (user manually closes tab or navigates away)
    window.addEventListener('beforeunload', handleBeforeUnload);
  }

  function notifyServerTabClosed(tabId) {
    // Use sendBeacon for reliable delivery even during page unload
    const data = JSON.stringify({ tabId: tabId });
    navigator.sendBeacon(`${SERVER_URL}/tab-closed`, data);
  }

  function showNotification(message) {
    const notification = document.createElement('div');
    notification.textContent = message;
    notification.style.cssText = `
      position: fixed;
      top: 20px;
      right: 20px;
      background: #4CAF50;
      color: white;
      padding: 12px 20px;
      border-radius: 5px;
      z-index: 999999;
      font-family: Arial, sans-serif;
      font-size: 14px;
      box-shadow: 0 2px 5px rgba(0,0,0,0.2);
      animation: slideIn 0.3s ease-out;
    `;

    // Only add style once to prevent accumulation
    if (!document.getElementById('vendor-scraper-notification-style')) {
      const style = document.createElement('style');
      style.id = 'vendor-scraper-notification-style';
      style.textContent = `
        @keyframes slideIn {
          from { transform: translateX(400px); opacity: 0; }
          to { transform: translateX(0); opacity: 1; }
        }
      `;
      document.head.appendChild(style);
    }

    document.body.appendChild(notification);
    setTimeout(() => notification.remove(), 3000);
  }

  // Store last scraped data
  let lastScrapedData = null;

  // Listen for manual trigger from popup
  const messageListener = (request, sender, sendResponse) => {
    if (request.action === 'scrapeNow') {
      const vendor = detectVendor();
      if (vendor) {
        try {
          let data;
          switch(vendor) {
            case 'grainger':
              data = extractGraingerData();
              break;
            case 'mcmaster':
              data = extractMcMasterData();
              break;
            case 'festo':
              data = extractFestoData();
              break;
            case 'zoro':
              data = extractZoroData();
              break;
          }

          if (data) {
            data.vendor = vendor;
            lastScrapedData = data;
            sendToServer(data);
            sendResponse({ success: true });
          } else {
            sendResponse({ success: false });
          }
        } catch (error) {
          console.error('Extraction failed:', error);
          sendResponse({ success: false });
        }
      } else {
        sendResponse({ success: false });
      }
      return true;
    }

    if (request.action === 'getData') {
      sendResponse({ data: lastScrapedData });
      return true;
    }
  };
  chrome.runtime.onMessage.addListener(messageListener);

  // Listen for Ctrl+Shift+X to trigger manual timeout
  const keydownHandler = async (e) => {
    if (e.ctrlKey && e.shiftKey && (e.key === 'X' || e.key === 'x')) {
      e.preventDefault();
      console.log('Manual timeout triggered by user');

      const vendor = detectVendor();
      if (!vendor) return;

      // Send a timeout notification to server with whatever data we have
      const tabId = await new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'getTabId' }, resolve);
      });

      if (tabId) {
        try {
          await fetch(`${SERVER_URL}/scrape`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
              tabId: tabId,
              vendor: vendor,
              partNumber: extractPartNumber(),
              description: 'Not Found',
              price: 'Not Found',
              unit: 'Not Found',
              mfrNumber: 'Not Found',
              brand: 'Not Found',
              url: window.location.href,
              timestamp: new Date().toISOString(),
              manualTimeout: true
            })
          });
          showNotification('⏭ Manual timeout - proceeding with form');
        } catch (error) {
          console.error('Failed to send manual timeout:', error);
        }
      }
    }
  };
  document.addEventListener('keydown', keydownHandler);

  // Cleanup listeners on page unload
  window.addEventListener('beforeunload', () => {
    chrome.runtime.onMessage.removeListener(messageListener);
    document.removeEventListener('keydown', keydownHandler);
  });
})();