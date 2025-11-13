# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a **Multi-Vendor Price Scraper** system that combines a Python Flask backend with a Chrome extension frontend to automatically scrape and update pricing data from multiple industrial suppliers (Grainger, McMaster-Carr, Festo, Zoro). The system maintains pricing history in an Excel workbook with VBA macros.

**New Features (2024):**
- **Add New ACI Number**: Ability to create new ACI entries when searching for non-existent items
- **Batch Update**: Process multiple ACI numbers automatically with validation and ±15% price change limits
- **CLI Batch Mode**: Support for Excel macro integration via command-line arguments

## Architecture

### Two-Part System

**1. Python Application (`main.py`)**
- Flask server running on `localhost:5000` that acts as a data broker
- Tkinter GUI for data entry and confirmation
- Excel integration using `openpyxl` (VBA-preserving mode)
- Browser automation for opening vendor pages

**2. Chrome Extension (`extension/`)**
- Content scripts that scrape vendor pages automatically
- Background service worker for tab management
- Popup UI for manual scraping and server status

### Communication Flow

```
User Input → Python GUI → Browser Opens → Extension Scrapes → Flask Receives → GUI Displays → User Confirms → Excel Updates
```

The extension automatically extracts data and POSTs it to the Flask server's `/scrape` endpoint. The Python app queues this data, displays it in a Tkinter form for confirmation, then saves to Excel.

### Tab Lifecycle Management

The system tracks browser tabs to automatically close them after data processing:
- Tabs register themselves via `/register-tab` when opened
- Extension polls `/should-close/<tab_id>` every 1 second
- Python app adds tab IDs to `TABS_TO_CLOSE` set after form submission
- Extension closes tabs when signaled or after 10-minute timeout

## Key Components

### Flask Server (`main.py:58-129`)

**Endpoints:**
- `GET /ping` - Health check for extension
- `POST /scrape` - Receives scraped data from extension
- `POST /register-tab` - Registers a new browser tab
- `GET /should-close/<tab_id>` - Polling endpoint for tab closure
- `POST /tab-closed` - Tab cleanup notification

**Important:** Server runs in a daemon thread and checks for port conflicts on startup (main.py:131-165).

### Content Scripts (`extension/content.js`)

**Vendor-Specific Scrapers:**
- `extractGraingerData()` (line 114)
- `extractMcMasterData()` (line 162)
- `extractFestoData()` (line 201)
- `extractZoroData()` (line 249)

Each scraper uses CSS selectors to extract: description, price, unit, MFR number, brand. Selectors are fragile and may break with vendor site updates.

**Manual Trigger:** Extension popup can manually trigger scraping via `chrome.runtime.sendMessage({action: 'scrapeNow'})`.

**Timeout Override:** Press `Ctrl+Shift+X` on vendor pages to manually timeout and proceed with "Not Found" data (content.js:449).

### Data Processing (`main.py:273-419`)

**Vendor-Specific Cleanup Functions:**
- `cleanup_grainger_data()` - Removes "$", parses "X of Y" units
- `cleanup_mcmaster_data()` - Handles "per pack" pricing
- `cleanup_festo_data()` - Extracts quantity from input field
- `cleanup_zoro_data()` - Parses CSV-like price format

**Price Change Detection:**
- Calculates percentage change (main.py:421)
- Alerts on changes ≥20% or decreases ≥10%
- Updates price history in CSV format when change ≥1%

### Excel Integration (`main.py:464-513`)

**File Path Priority:**
1. Server paths: `Z:\ACOD\MMLV2.xlsm` (or `.xlsx`)
2. Local paths: `MML.xlsm` in program directory

**Critical:** Excel operations use `keep_vba=True` to preserve macros. Always close workbooks in `finally` blocks to prevent file locks.

**Sheet:** All operations target the "Purchase Parts" sheet, searching column A (ACI #) for matches.

### GUI Form (`main.py:649-888`)

Displays 15 fields side-by-side:
- Left: Scraped data (editable)
- Right: Current Excel data (read-only)
- Checkboxes: Toggle to use current data instead of scraped

**Highlighting:**
- Yellow: Values differ between scraped and current
- Red: "Not Found" values from scraper
- White: Unchanged or using current data

**Keyboard Shortcuts:**
- `Ctrl+S`: Submit
- `Esc`: Cancel (entry form)
- `Enter` (search dialog): Submit
- `Esc` (search dialog): Cancel

## New Features

### Feature 1: Add New ACI Number (main.py:546-650, 1477-1519)

When a user searches for an ACI number that doesn't exist in the Excel database:

**Flow:**
1. `process_excel()` returns `"NOT_FOUND"` instead of showing error
2. `prompt_add_new_aci()` asks user if they want to add the new ACI
3. If yes, `get_new_aci_details()` dialog collects:
   - Vendor (dropdown: Grainger, McMaster-Carr, McMaster, Festo, Zoro, Other)
   - Vendor Part Number (text entry)
4. `add_new_row_to_excel()` creates new row with ACI#, Vendor, Vendor Part#, and Date
5. If vendor is auto-supported (Grainger, McMaster, Festo, Zoro):
   - Automatically opens browser to scrape data
   - Pre-fills form with scraped data
6. If manual vendor:
   - Shows empty form for manual data entry
7. User confirms and saves via standard form

**Key Functions:**
- `add_new_row_to_excel(file_path, aci_number, vendor, vendor_part_number)` - Creates new Excel row
- `prompt_add_new_aci(aci_number)` - Yes/No dialog
- `get_new_aci_details(aci_number)` - Vendor and part number collection
- `is_vendor_auto(vendor)` - Checks if vendor supports auto-scraping

### Feature 2: Batch Update (main.py:1085-1433, 1447-1464)

Allows automatic updating of multiple ACI numbers with strict validation to prevent errors.

**GUI Mode:**
1. Click "Batch Update" button on main search dialog
2. `batch_update_dialog()` shows text area for ACI list
   - Supports newline-separated or comma-separated ACIs
   - Can paste directly from Excel
3. `batch_update_worker()` processes each ACI:
   - Looks up in Excel
   - Checks if vendor is auto-supported
   - Opens browser and scrapes data
   - Validates with `validate_batch_match()`
   - Checks price change is within ±15%
   - Updates Excel if all validations pass
   - Closes browser tab automatically
4. `show_batch_summary()` displays categorized results:
   - ✓ Updated (with old → new price)
   - ⊘ Skipped (with reason)
   - ✗ Errors (with error message)
   - ? Not Found

**CLI Mode (for Excel Macros):**

Command-line syntax:
```bash
# Comma-separated list
python main.py --batch "ACI001,ACI002,ACI003"

# From file (one ACI per line, # for comments)
python main.py --batch-file "aci_list.txt"

# From Excel VBA macro
Shell "AdvantageScraper.exe --batch ""ACI001,ACI002,ACI003"""
```

**Validation Rules (validate_batch_match):**
1. **Description Match**: Scraped description must contain or be contained in current description (fuzzy match)
2. **Part Number Match**: MFR part numbers must match exactly (case-insensitive), if available
3. **Unit Match**: Units must match exactly (case-insensitive), if available
4. **Price Change Limit**: New price must be within ±15% of old price

**Skipping Criteria:**
- Not found in Excel
- Manual vendor (not Grainger/McMaster/Festo/Zoro)
- Description mismatch
- Part number mismatch
- Unit mismatch
- Price change >±15%
- Price not found on vendor site
- Scraping timeout

**Key Functions:**
- `batch_update_dialog()` - GUI for ACI list entry
- `batch_update_worker(file_path, aci_list)` - Main batch processor
- `validate_batch_match(current_data, scraped_data)` - Returns (is_match, reason)
- `show_batch_summary(results)` - Display results GUI
- Command-line argument parsing in `main()` (lines 1556-1699)

## Common Development Tasks

### Building Executable

```bash
python build.py
```

This creates `dist/AdvantageScraper.exe` using PyInstaller with:
- Icon bundling (`icon.ico`, `icon.png`)
- Windowed mode (no console)
- Hidden imports for Flask, openpyxl, tkinter

**Spec file:** `AdvantageScraper.spec` contains the PyInstaller configuration. Alternatively, use `build.py` to generate `dist/AdvantageScraper.exe`.

### Installing Chrome Extension

1. Navigate to `chrome://extensions/`
2. Enable "Developer mode"
3. Click "Load unpacked"
4. Select the `extension/` folder

### Running Development Server

```bash
pip install -r requirements.txt
python main.py
```

Ensure Excel file exists at expected path or in program directory.

**Command-line Options:**
```bash
# Normal GUI mode (default)
python main.py

# Batch update from command line (comma-separated)
python main.py --batch "ACI001,ACI002,ACI003"

# Batch update from file
python main.py --batch-file path/to/aci_list.txt

# Show help
python main.py --help
```

**Excel VBA Macro Integration:**
```vba
' Call batch update from Excel VBA
Sub BatchUpdatePrices()
    Dim aciList As String
    Dim exePath As String

    ' Build comma-separated list from selected cells
    aciList = Join(Application.Transpose(Selection), ",")

    ' Path to compiled executable
    exePath = "C:\Path\To\AdvantageScraper.exe"

    ' Run batch update
    Shell exePath & " --batch """ & aciList & """"
End Sub
```

### Adding New Vendor Support

1. **Add URL pattern** to `VENDOR_URLS` dict (main.py:37)
2. **Update manifest.json** with host permissions and content script matches
3. **Create scraper function** in `content.js` following existing patterns
4. **Add vendor case** to `detectVendor()` and extraction switch statements
5. **Create cleanup function** in `main.py` (e.g., `cleanup_newvendor_data()`)
6. **Update `parse_vendor_data()`** to call cleanup function

## Important Patterns

### Queue Management
- `DATA_QUEUE` has maxsize=50 to prevent memory issues
- Oldest items discarded when full (main.py:74-79)

### Tab Tracking
- `TABS_TO_CLOSE` is a set for O(1) lookups
- `REGISTERED_TABS` dict stores tab metadata with timestamps
- Stale tabs (>30 minutes) cleaned up automatically (main.py:48)

### Error Handling
- Extension failures are silent (don't break user workflow)
- Server connection failures in extension logged but not alerted
- Excel errors show messageboxes to user

### Icon Management
- `_APP_ICON_PHOTO` global prevents Tkinter memory leaks
- ICO format preferred on Windows, fallback to PNG (main.py:531-546)

## Data Schema (15 Fields)

```
0:  ACI #
1:  MFR Part #
2:  MFR (Brand)
3:  Description
4:  QTY
5:  Per (Unit)
6:  Vendor
7:  Vendor Part #
8:  Legacy
9:  Unit Price
10: Change %
11: Date (current date on update)
12: Last Updated Price
13: Last Updated Date
14: Price History (CSV format: "Date: X Price: Y, Date: X2 Price: Y2...")
```

## Chrome Extension Manifest V3

Uses service worker (`background.js`) instead of persistent background page. Content scripts run at `document_idle` to ensure page is fully loaded before scraping.

## Testing Scraper Selectors

To test if selectors still work after vendor site changes:
1. Open browser DevTools (F12)
2. Check console for "Extraction failed" errors
3. Use `document.querySelector('[selector]')` in console to verify elements exist
4. Update selectors in `content.js` if needed

## Dependencies

- **Flask 3.0.0** - Web server framework
- **flask-cors 4.0.0** - Cross-origin support for extension
- **openpyxl 3.1.2** - Excel file manipulation (VBA-preserving)
- **pyinstaller 6.3.0** - EXE creation

## Known Limitations

- Extension selectors may break when vendors update their HTML
- Only supports Chrome/Chromium browsers
- Requires local Flask server to be running
- Excel file must not be open in Excel during updates
- No authentication/security on Flask endpoints (localhost only)
