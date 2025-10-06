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

    // Wait a bit for dynamic content
    setTimeout(() => {
      if (!extractionAttempted) {
        extractionAttempted = true;
        extractAndSend(vendor);
      }
    }, 2000);
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

      if (data && data.price !== "Not Found") {
        data.vendor = vendor;
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
    return match ? match[1] : "Unknown";
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
    const priceSelectors = ['PrceTxt', 'Price_price', 'PrceTierPrceCol', 'Prce'];
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
      const response = await fetch(`${SERVER_URL}/scrape`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(data)
      });

      if (response.ok) {
        console.log('Data sent successfully');
        showNotification('✓ Price data captured');
      }
    } catch (error) {
      console.error('Server error:', error);
      // Fail silently if server not running
    }
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
    
    const style = document.createElement('style');
    style.textContent = `
      @keyframes slideIn {
        from { transform: translateX(400px); opacity: 0; }
        to { transform: translateX(0); opacity: 1; }
      }
    `;
    document.head.appendChild(style);
    
    document.body.appendChild(notification);
    setTimeout(() => notification.remove(), 3000);
  }

  // Listen for manual trigger from popup
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.action === 'scrapeNow') {
      const vendor = detectVendor();
      extractAndSend(vendor);
      sendResponse({ success: true });
    }
  });
})();