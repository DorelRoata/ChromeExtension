(function() {
  'use strict';

  // All network calls are proxied via the background service worker
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
    }, 3000);
  }

  async function registerTab() {
    try {
      // Get our tab ID
      const tabId = await new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'getTabId' }, resolve);
      });

      if (tabId) {
        // Ask background to register and start close polling
        chrome.runtime.sendMessage({ action: 'registerTab', tabId, url: window.location.href }, () => {});
        console.log('Requested background to register tab');
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
    const itemNumberNode = document.querySelector('dd[data-testid^="product-item-number"] span');
    const mainItemId = itemNumberNode ? itemNumberNode.textContent.trim() : "";
    let price = "Not Found";
    let priceSource = "";

    const setPriceFromElement = (elem, source) => {
      if (elem && elem.textContent && elem.textContent.trim()) {
        price = elem.textContent.trim();
        priceSource = source;
        return true;
      }
      return false;
    };

    if (mainItemId) {
      const directCandidate = document.querySelector(`span[data-testid="pricing-component-${mainItemId}"]`);
      setPriceFromElement(directCandidate, "direct");
    }

    if (price === "Not Found") {
      const purchaseAside = document.querySelector('aside[aria-label*="Purchase"]');
      if (purchaseAside) {
        const asideCandidate = purchaseAside.querySelector('span[data-testid^="pricing-component"]');
        setPriceFromElement(asideCandidate, "purchase-aside");
      }
    }

    if (price === "Not Found") {
      const dataTestIdPrices = document.querySelectorAll('span[data-testid^="pricing-component"]');
      for (const candidate of dataTestIdPrices) {
        const dataTestId = candidate.getAttribute('data-testid') || '';
        if (!mainItemId || dataTestId === `pricing-component-${mainItemId}`) {
          if (setPriceFromElement(candidate, "data-testid-scan")) {
            break;
          }
        }
      }
    }

    if (price === "Not Found") {
      const priceElements = document.querySelectorAll('.HANkB');
      for (const elem of priceElements) {
        const dataTestId = elem.getAttribute('data-testid') || '';
        if (dataTestId && mainItemId && dataTestId !== `pricing-component-${mainItemId}`) {
          continue;
        }
        if (elem.textContent && elem.textContent.includes('$') && setPriceFromElement(elem, "class-fallback")) {
          break;
        }
      }
    }

    if (priceSource && priceSource !== "direct") {
      console.debug('Grainger price fallback used:', priceSource);
    }

    let mfrNumber = "Not Found";
    const headerSection = document.querySelector('.ue2KE');
    if (headerSection) {
      const mfrLabel = Array.from(headerSection.querySelectorAll('dt')).find(dt =>
        dt.textContent.trim().toLowerCase().includes('mfr')
      );
      if (mfrLabel && mfrLabel.nextElementSibling) {
        mfrNumber = mfrLabel.nextElementSibling.textContent.trim();
      }
    }
    if (mfrNumber === "Not Found" && itemNumberNode) {
      const fallbackLabel = Array.from(document.querySelectorAll('dt')).find(dt =>
        dt.textContent.trim().toLowerCase().includes('item')
      );
      if (fallbackLabel && fallbackLabel.nextElementSibling) {
        mfrNumber = fallbackLabel.nextElementSibling.textContent.trim();
      }
    }

    let brand = "Not Found";
    const techSection = document.querySelector('[data-testid="product-techs"]');
    if (techSection) {
      const brandLabel = Array.from(techSection.querySelectorAll('dt')).find(dt =>
        dt.textContent.trim() === 'Brand'
      );
      if (brandLabel && brandLabel.nextElementSibling) {
        brand = brandLabel.nextElementSibling.textContent.trim();
      }
    }

    return {
      partNumber: mainItemId || extractPartNumber(),
      description: findText([
        'h1[data-testid="product-title"]',
        '.ue2KE .tliLs',
        'h1[class*="product"]',
        'h1'
      ]),
      price: price,
      unit: findText([
        'aside[aria-label*="Purchase"] .G32gdF',
        '.FuAZI6 .G32gdF',
        '.G32gdF',
        '[class*="unit"]'
      ]),
      mfrNumber: mfrNumber,
      brand: brand,
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

  function sendToServer(data, callback) {
    // callback(boolean) -> indicates whether server accepted the payload
    (async () => {
      try {
        const tabId = await new Promise((resolve) => {
          chrome.runtime.sendMessage({ action: 'getTabId' }, resolve);
        });
        data.tabId = tabId;
        chrome.runtime.sendMessage({ action: 'sendScrape', data }, (resp) => {
          const ok = !!(resp && resp.success);
          if (ok) {
            showNotification('✓ Price data captured');
          } else {
            console.warn('Background failed to send scrape payload');
          }
          if (typeof callback === 'function') callback(ok);
        });
      } catch (error) {
        console.warn('Failed to send scrape to background:', error);
        if (typeof callback === 'function') callback(false);
      }
    })();
  }

  function startPollingForClose(tabId) {
    // Delegate polling to background service worker
    chrome.runtime.sendMessage({ action: 'startClosePolling', tabId }, () => {});
  }

  // Polling and tab closed notifications are handled in the background script

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
            // send to server via background, respond to popup only when server accepts
            sendToServer(data, (ok) => {
              if (ok) lastScrapedData = data;
              sendResponse({ success: ok });
            });
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

      const tabId = await new Promise((resolve) => {
        chrome.runtime.sendMessage({ action: 'getTabId' }, resolve);
      });

      if (tabId) {
        const payload = {
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
        };
        chrome.runtime.sendMessage({ action: 'sendScrape', data: payload }, () => {});
        showNotification('⏭ Manual timeout - proceeding with form');
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
