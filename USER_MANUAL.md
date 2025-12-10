# Advantage Multi-Vendor Price Scraper
## User Manual

**Version 2.1.1**

---

## Table of Contents

1. [Introduction](#introduction)
2. [System Requirements](#system-requirements)
3. [Installation](#installation)
   - [Installing the Chrome Extension](#installing-the-chrome-extension)
   - [Setting Up the Application](#setting-up-the-application)
4. [Getting Started](#getting-started)
5. [Using the Application](#using-the-application)
   - [Single ACI Update](#single-aci-update)
   - [Batch Update](#batch-update)
   - [Adding New ACI Numbers](#adding-new-aci-numbers)
6. [Keyboard Shortcuts](#keyboard-shortcuts)
7. [Supported Vendors](#supported-vendors)
8. [Understanding the Data Form](#understanding-the-data-form)
9. [Excel Integration](#excel-integration)
10. [Troubleshooting](#troubleshooting)
11. [Frequently Asked Questions](#frequently-asked-questions)

---

## Introduction

The **Advantage Multi-Vendor Price Scraper** is a productivity tool designed to streamline the process of updating pricing information from industrial suppliers. Instead of manually visiting vendor websites, copying prices, and pasting them into Excel, this tool automates the entire workflow.

### What This Tool Does

- Automatically opens vendor websites based on part numbers in your Excel database
- Extracts pricing data, descriptions, and part information from vendor pages
- Displays the scraped data alongside your current Excel data for comparison
- Updates your Excel workbook with new pricing information
- Tracks price history and alerts you to significant price changes

### How It Works

The system has two parts that work together:

1. **Desktop Application (AdvantageScraper.exe)** - The main program you interact with. It manages the Excel file and displays the data entry forms.

2. **Chrome Extension** - Runs in the background and automatically extracts data from vendor websites when you visit them.

---

## System Requirements

### Hardware
- Windows 10 or later
- 4 GB RAM minimum
- 100 MB free disk space

### Software
- **Google Chrome** browser (required for the extension)
- **Microsoft Excel** (for viewing/editing the MML workbook)
- Network access to vendor websites

### Network
- Access to the following vendor websites:
  - www.grainger.com
  - www.mcmaster.com
  - www.festo.com
  - www.zoro.com
- Access to `Z:\ACOD\` network drive (for the shared Excel file)

---

## Installation

### Installing the Chrome Extension

The Chrome extension is what allows the tool to automatically read data from vendor websites.

#### Step 1: Open Chrome Extensions Page

1. Open **Google Chrome**
2. Type `chrome://extensions/` in the address bar
3. Press **Enter**

#### Step 2: Enable Developer Mode

1. Look for the **"Developer mode"** toggle in the top-right corner
2. Click to turn it **ON** (the toggle should turn blue)

#### Step 3: Load the Extension

1. Click the **"Load unpacked"** button (appears after enabling Developer mode)
2. A folder selection dialog will open
3. Navigate to the folder containing the extension files (the `extension` folder)
4. Select the folder and click **"Select Folder"**

#### Step 4: Pin the Extension (Recommended)

1. Click the **puzzle piece icon** in Chrome's toolbar (Extensions menu)
2. Find **"Advantage Multi-Vendor Scraper"** in the list
3. Click the **pin icon** next to it
4. The extension icon will now appear in your toolbar

#### Verifying Installation

After installation, you should see:
- The extension listed on the `chrome://extensions/` page
- The extension icon in your Chrome toolbar (if pinned)

---

### Setting Up the Application

#### Step 1: Locate the Application

The application file is named `AdvantageScraper.exe`. It should be located in a shared folder or provided to you by your administrator.

#### Step 2: Create a Shortcut (Optional)

1. Right-click on `AdvantageScraper.exe`
2. Select **"Create shortcut"**
3. Move the shortcut to your Desktop for easy access

#### Step 3: First Launch

1. Double-click `AdvantageScraper.exe` to start the application
2. The application will start a local server on your computer
3. A search dialog will appear asking for an ACI number

#### Verifying the Server Connection

1. Click the extension icon in Chrome
2. You should see **"Server: Connected"** with a green checkmark
3. If it shows "Not Running", make sure the application is open

---

## Getting Started

### Your First Price Update

Let's walk through updating the price for a single part:

1. **Launch the Application**
   - Double-click `AdvantageScraper.exe`
   - Wait for the search dialog to appear

2. **Enter an ACI Number**
   - Type the ACI number you want to update (e.g., `12345`)
   - Press **Enter** or click **"Search"**

3. **Wait for the Browser**
   - Chrome will automatically open to the vendor's website
   - The extension will extract the pricing data
   - This usually takes 5-15 seconds

4. **Review the Data**
   - A form will appear showing two columns:
     - **Left column**: New data from the website
     - **Right column**: Current data from Excel
   - Different values are highlighted in yellow

5. **Make Your Decision**
   - Review the differences
   - Use checkboxes to keep current values if needed
   - Click **"Submit"** to save, or **"Cancel"** to discard

6. **Done!**
   - The Excel file is updated
   - The browser tab closes automatically
   - You can search for another ACI number

---

## Using the Application

### Single ACI Update

This is the standard mode for updating one part at a time.

#### Starting an Update

1. Launch `AdvantageScraper.exe`
2. In the search dialog, enter the ACI number
3. Press **Enter** or click **"Search"**

#### What Happens Next

The application will:
1. Look up the ACI number in the Excel database
2. Find the vendor associated with that part
3. Open the vendor's website in Chrome
4. Wait for the extension to scrape the data
5. Display a comparison form

#### The Data Entry Form

The form shows 15 fields organized in two columns:

| Field | Description |
|-------|-------------|
| ACI # | Your internal part number |
| MFR Part # | Manufacturer's part number |
| MFR | Manufacturer/Brand name |
| Description | Product description |
| QTY | Quantity |
| Per | Unit of measure (ea, pk, etc.) |
| Vendor | Supplier name |
| Vendor Part # | Vendor's part number |
| Legacy | Legacy reference |
| Unit Price | Current unit price |
| Date | Date of update |
| Change % | Price change percentage |
| Last Updated Price | Previous price |
| Last Updated Date | Date of previous update |
| Price History | Historical price data |

#### Using the Checkboxes

Each field has a checkbox labeled **"KEEP"**:
- **Unchecked**: Use the new scraped value (left column)
- **Checked**: Keep the current Excel value (right column)

There are also convenience buttons:
- **"Check All"**: Keep all current values
- **"Uncheck All"**: Use all new values

#### Color Coding

- **Yellow background**: Value differs between scraped and current
- **Red background**: "Not Found" - the scraper couldn't find this data
- **White background**: Values match or using current data

#### Submitting the Form

- Click **"Submit"** or press **Ctrl+S** to save changes
- Click **"Cancel"** or press **Esc** to discard changes

---

### Batch Update

Batch update allows you to process multiple ACI numbers automatically.

#### When to Use Batch Update

- Updating prices for multiple parts at once
- Regular price maintenance
- End-of-month pricing updates

#### Starting a Batch Update (GUI Method)

1. Launch `AdvantageScraper.exe`
2. In the search dialog, click **"Batch Update"**
3. A text area will appear
4. Enter your ACI numbers:
   - One per line, OR
   - Comma-separated (e.g., `ACI001, ACI002, ACI003`)
5. Click **"Start Batch"**

#### What Happens During Batch Processing

For each ACI number, the system will:
1. Look up the part in Excel
2. Check if it's an auto-supported vendor
3. Open the vendor website
4. Scrape the data
5. Validate the data matches (description, part number, unit)
6. Check the price change is within ±15%
7. Update Excel if all validations pass
8. Close the browser tab
9. Move to the next ACI number

#### Batch Validation Rules

The batch update is stricter than single updates to prevent errors:

| Validation | Rule |
|------------|------|
| Description | Must match (fuzzy matching) |
| Part Number | Must match exactly |
| Unit | Must match exactly |
| Price Change | Must be within ±15% |

If any validation fails, that ACI is **skipped** (not updated).

#### Batch Summary Report

After processing, you'll see a summary with four categories:

- **Updated**: Successfully updated with old → new price shown
- **Skipped**: Validation failed (reason shown)
- **Errors**: Something went wrong (error shown)
- **Not Found**: ACI number not in Excel database

#### Command-Line Batch Update

For automation or Excel macro integration:

```
AdvantageScraper.exe --batch "ACI001,ACI002,ACI003"
```

Or from a file:
```
AdvantageScraper.exe --batch-file "C:\path\to\aci_list.txt"
```

The file should have one ACI number per line. Lines starting with `#` are treated as comments.

#### Excel VBA Macro Example

You can call the batch update directly from Excel:

```vba
Sub BatchUpdatePrices()
    Dim aciList As String
    Dim exePath As String

    ' Build list from selected cells
    aciList = Join(Application.Transpose(Selection), ",")

    ' Path to the executable
    exePath = "C:\Path\To\AdvantageScraper.exe"

    ' Run batch update
    Shell exePath & " --batch """ & aciList & """"
End Sub
```

---

### Adding New ACI Numbers

When you search for an ACI number that doesn't exist in the database, you can add it.

#### Step-by-Step Process

1. **Search for the ACI Number**
   - Enter the new ACI number in the search dialog
   - Click Search or press Enter

2. **Confirm You Want to Add It**
   - A dialog will ask: "ACI number not found. Would you like to add it?"
   - Click **"Yes"** to proceed

3. **Enter Vendor Information**
   - Select the **Vendor** from the dropdown:
     - Grainger
     - McMaster-Carr
     - Festo
     - Zoro
     - Other (for manual entry)
   - Enter the **Vendor Part Number**
   - Click **"OK"**

4. **Data Entry**
   - **For supported vendors**: The browser opens and scrapes data automatically
   - **For "Other" vendors**: A blank form appears for manual entry

5. **Complete the Form**
   - Fill in or verify all fields
   - Click **"Submit"** to save

#### What Gets Created

A new row is added to the Excel file with:
- The ACI number you entered
- The vendor you selected
- The vendor part number you provided
- Today's date
- All other data from the form

---

## Keyboard Shortcuts

### In the Data Entry Form

| Shortcut | Action |
|----------|--------|
| **Ctrl+S** | Submit/Save the form |
| **Esc** | Cancel and close the form |

### In the Search Dialog

| Shortcut | Action |
|----------|--------|
| **Enter** | Search for the entered ACI number |
| **Esc** | Close the dialog |

### On Vendor Pages (Chrome)

| Shortcut | Action |
|----------|--------|
| **Ctrl+Shift+X** | Force timeout - proceed with "Not Found" |

Use Ctrl+Shift+X if the page is stuck or won't load properly.

---

## Supported Vendors

The extension can automatically scrape data from these vendors:

### Grainger (www.grainger.com)
- Full support for product pages
- Extracts: Description, Price, Unit, MFR Part #, Brand

### McMaster-Carr (www.mcmaster.com)
- Full support for product pages
- Extracts: Description, Price, Unit, MFR Part #
- Note: Handles "per pack" pricing automatically

### Festo (www.festo.com)
- Full support for US product pages
- Extracts: Description, Price, Unit, Part Number

### Zoro (www.zoro.com)
- Full support for product pages
- Extracts: Description, Price, Unit, MFR Part #, Brand

### Other Vendors
- Not auto-supported
- Must enter data manually
- Can still track in the system

---

## Understanding the Data Form

### Field-by-Field Explanation

#### ACI #
Your internal Advantage Conveyor part number. This is the primary identifier.

#### MFR Part #
The manufacturer's original part number. Used to verify you're looking at the correct product.

#### MFR (Manufacturer)
The brand or manufacturer name (e.g., "3M", "Brady", "Dorner").

#### Description
The product description from the vendor. Used in batch validation to ensure correct product.

#### QTY
The quantity value. Often "1" for single items.

#### Per (Unit)
The unit of measure:
- **ea** = each (single item)
- **pk** = pack
- **bx** = box
- **cs** = case
- **pr** = pair

#### Vendor
The supplier you purchase from (Grainger, McMaster-Carr, etc.).

#### Vendor Part #
The part number used by the vendor (may differ from MFR Part #).

#### Legacy
Legacy reference field for backward compatibility.

#### Unit Price
The current price per unit from the vendor.

#### Date
The date of this update. Defaults to today's date.

#### Change %
Calculated percentage change from the last price:
- Positive = price increased
- Negative = price decreased

#### Last Updated Price
The previous unit price (before this update).

#### Last Updated Date
The date of the previous update.

#### Price History
A record of historical prices in CSV format:
```
Date: 01-15-24 Price: 12.50, Date: 06-20-24 Price: 13.25
```

---

## Excel Integration

### File Location

The application looks for the Excel file in this order:
1. `Z:\ACOD\MMLV2.xlsm` (network drive, with macros)
2. `Z:\ACOD\MMLV2.xlsx` (network drive, without macros)
3. `MML.xlsm` in the same folder as the application
4. `MML.xlsx` in the same folder as the application

### Required Sheet

The application works with the **"Purchase Parts"** sheet.

### Important Notes

- **Close Excel First**: The Excel file cannot be open in Excel while updating. Close it before running updates.
- **VBA Macros Preserved**: If your file has macros (.xlsm), they will be preserved during updates.
- **Network Access**: Ensure you have write access to the network drive.

### Column Layout

The application expects these columns in order:
1. ACI #
2. MFR Part #
3. MFR
4. Description
5. QTY
6. Per
7. Vendor
8. Vendor Part #
9. Legacy
10. Unit Price
11. Date
12. Change %
13. Last Updated Price
14. Last Updated Date
15. Price History

---

## Troubleshooting

### Extension Shows "Server: Not Running"

**Cause**: The desktop application isn't running or can't start the server.

**Solutions**:
1. Make sure `AdvantageScraper.exe` is running
2. Check if another program is using port 5000
3. Restart the application
4. Check Windows Firewall isn't blocking the application

### No Data Received After 30 Seconds

**Cause**: The extension couldn't scrape the page.

**Solutions**:
1. Check if you're logged into the vendor website
2. Refresh the page and wait for it to fully load
3. Click the extension icon and try "Scrape This Page"
4. Press **Ctrl+Shift+X** to force timeout and proceed with manual entry
5. Check browser console (F12) for errors

### Excel File Locked Error

**Cause**: The Excel file is open elsewhere.

**Solutions**:
1. Close Excel completely (check Task Manager)
2. Check if another user has the file open on the network
3. Wait a moment and try again

### Wrong Data Scraped

**Cause**: Vendor website layout may have changed.

**Solutions**:
1. Verify you're on the correct product page
2. Manually edit the data in the form before submitting
3. Report the issue so selectors can be updated

### Browser Doesn't Open

**Cause**: Chrome isn't installed or can't be found.

**Solutions**:
1. Ensure Google Chrome is installed
2. Try opening Chrome manually first
3. The application will fall back to your default browser

### Price Change Alert

**Cause**: Price changed more than expected (≥20% increase or ≥10% decrease).

**This is informational**: The system alerts you to significant changes so you can verify they're correct before saving.

### Batch Update Skipping Items

**Cause**: Validation rules are failing.

**Check the summary for specific reasons**:
- "Description mismatch" - Product may have changed
- "Part number mismatch" - Wrong product page loaded
- "Price change >15%" - Unusual price change, review manually
- "Manual vendor" - Vendor doesn't support auto-scraping

---

## Frequently Asked Questions

### Q: Do I need to keep the application running?

**A**: Yes, `AdvantageScraper.exe` must be running for the extension to send data. You can minimize it, but don't close it.

### Q: Can I update the Excel file while using this tool?

**A**: No, close Excel before running updates. The tool needs exclusive access to write changes.

### Q: What if the price change is more than 15% in batch mode?

**A**: The item will be skipped. Use single-update mode to manually review and approve large price changes.

### Q: Can I add vendors that aren't in the list?

**A**: Yes, select "Other" when adding a new ACI. You'll need to enter data manually, but it will still be tracked in the system.

### Q: Does this work with Microsoft Edge?

**A**: No, the extension only works with Google Chrome. The data scraping features require the Chrome extension.

### Q: What happens if the vendor website is down?

**A**: The scraping will timeout after about 30 seconds. You can press Ctrl+Shift+X to skip waiting and enter data manually.

### Q: Can multiple people use this at the same time?

**A**: Not recommended. Only one person should update the Excel file at a time to avoid conflicts.

### Q: How do I update the extension?

**A**:
1. Go to `chrome://extensions/`
2. Find the extension and click "Remove"
3. Load the new version using "Load unpacked"

### Q: Where is price history stored?

**A**: Price history is stored in column 15 (Price History) of the Excel file as a comma-separated list.

### Q: Can I undo an update?

**A**: There's no automatic undo. You can manually edit the Excel file to restore previous values, or use the Last Updated Price/Date fields as reference.

---

## Getting Help

If you encounter issues not covered in this manual:

1. Check the [Troubleshooting](#troubleshooting) section
2. Verify your installation following the [Installation](#installation) steps
3. Contact your system administrator

---

*Document Version: 2.1.1*
*Last Updated: December 2024*
