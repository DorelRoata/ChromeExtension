# Advantage Multi-Vendor Scraper Extension

Chrome extension for scraping pricing data from multiple vendor websites.

## What's New

- Background script proxies all network calls to the local app at `http://127.0.0.1:5000` for reliability.
- Popup indicates server status clearly: Connected (Optional) vs. Standalone Mode.
- Success only reported when the background receives a 2xx response from the app.
- Host permissions include both `127.0.0.1` and `localhost`.

## Supported Vendors
- Grainger
- McMaster-Carr
- Festo
- Zoro

## Privacy Policy
[View Privacy Policy](https://github.com/DorelRoata/AdvantagePriceScraperPrivatePolicy/blob/main/privacy_policy.markdown)

## Features
- Automatic price scraping from vendor pages
- Tab auto-close after data confirmation
- Manual timeout with Ctrl+Shift+X
- Integration with Python desktop application

## Installation
1. Open Chrome
2. Go to `chrome://extensions`
3. Enable "Developer mode"
4. Click "Load unpacked"
5. Select the extension folder

## Version
2.1.0
