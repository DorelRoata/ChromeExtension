# Multi-Vendor Price Scraper

## Installation

### 1. Install Chrome Extension
1. Open Chrome
2. Go to `chrome://extensions/`
3. Enable "Developer mode" (top right)
4. Click "Load unpacked"
5. Select the `extension` folder
6. Pin the extension to your toolbar

### 2. Run the Application
- Double-click `PriceScraper.exe`
- The Flask server will start automatically
- Extension popup should show "Server: Connected ✓"

## Usage

1. **Enter ACI Number**: Type the part number when prompted
2. **Browser Opens**: The vendor page opens automatically
3. **Extension Scrapes**: Data is captured and sent to the app
4. **Review Data**: Confirm or modify the data in the GUI
5. **Save**: Click Submit to update the Excel file

## Supported Vendors
- Grainger
- McMaster-Carr
- Festo
- Zoro

## Troubleshooting

**Extension shows "Server: Not Running"**
- Make sure PriceScraper.exe is running
- Check Windows Firewall isn't blocking port 5000

**No data received after 30 seconds**
- Check if you're logged into the vendor website
- Refresh the page and try the extension popup "Scrape This Page"
- Check browser console (F12) for errors

**Excel file locked**
- Close Excel before running the scraper
- Check if file is open on network drive