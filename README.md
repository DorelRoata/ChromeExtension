# Multi-Vendor Price Scraper

A powerful tool that combines a Python Flask backend with a Chrome extension to automatically scrape and update pricing data from multiple industrial suppliers.

## What's New

- Robust Chrome detection across Windows/macOS/Linux; falls back gracefully if Chrome isn’t found.
- Excel typing fixes: ACI saved as numeric when digits-only, dates stored as real dates, prices parsed to numbers, and consistent number formats applied.
- Vendor parsing cleanup for Grainger, McMaster‑Carr, Festo, and Zoro to normalize unit/qty/price.
- Batch update refinements: vendor-aware part matching (skip strict match for McMaster), tolerant unit checks, and hyphen rule limited to McMaster only.
- Extension reliability: background proxies all requests to `127.0.0.1`; “success” only when the app accepts the payload; popup clarifies server is optional.
- Stability: explicit bind to `127.0.0.1`, private‑network CORS header, and minor cleanup.
- Executable naming unified: app binary is `AdvantageScraper.exe` (formerly `PriceScraper.exe`).

## New Features

### Add New ACI Numbers
When searching for a non-existent ACI number, you can now add it to the database:
- Enter vendor and vendor part number
- Automatic data scraping for supported vendors
- Manual entry for other vendors

### Batch Update
Process multiple ACI numbers at once:
- **GUI Mode**: Click "Batch Update" button and paste ACI list
- **CLI Mode**: Run from command line or Excel macros
- Smart validation: only updates if description, part number, and unit match
- Safety limit: only updates prices within ±15% change
- Detailed summary report

## Installation

### 1. Install Chrome Extension
1. Open Chrome
2. Go to `chrome://extensions/`
3. Enable "Developer mode" (top right)
4. Click "Load unpacked"
5. Select the `extension` folder
6. Pin the extension to your toolbar

### 2. Run the Application
- Double-click `AdvantageScraper.exe`
- The Flask server will start automatically
- Extension popup should show "Server: Connected ✓"

## Usage

### Single Update Mode
1. **Enter ACI Number**: Type the part number when prompted
2. **Browser Opens**: The vendor page opens automatically
3. **Extension Scrapes**: Data is captured and sent to the app
4. **Review Data**: Confirm or modify the data in the GUI
5. **Save**: Click Submit to update the Excel file

### Batch Update Mode (GUI)
1. Click **"Batch Update"** button on the main dialog
2. Paste or type ACI numbers (one per line or comma-separated)
3. Click **"Start Batch"**
4. Review the summary report when complete

### Command-Line Batch Update
For automation via Excel macros or scripts:

```bash
# Comma-separated list
AdvantageScraper.exe --batch "ACI001,ACI002,ACI003"

# From file (one ACI per line)
AdvantageScraper.exe --batch-file "aci_list.txt"

# Show version / help
AdvantageScraper.exe --version
AdvantageScraper.exe --help
```

**Excel VBA Macro Example:**
```vba
Sub BatchUpdatePrices()
    Dim aciList As String
    Dim exePath As String

    ' Build list from selected cells
    aciList = Join(Application.Transpose(Selection), ",")

    ' Run batch update
    exePath = "C:\Path\To\AdvantageScraper.exe"
    Shell exePath & " --batch """ & aciList & """"
End Sub
```

## Supported Vendors
- Grainger
- McMaster-Carr
- Festo
- Zoro

## Troubleshooting

**Extension shows "Server: Not Running"**
- Make sure AdvantageScraper.exe is running
- Check Windows Firewall isn't blocking port 5000

**No data received after 30 seconds**
- Check if you're logged into the vendor website
- Refresh the page and try the extension popup "Scrape This Page"
- Check browser console (F12) for errors

**Excel file locked**
- Close Excel before running the scraper
- Check if file is open on network drive
