# Privacy Policy for Advantage Multi-Vendor Price Scraper

**Effective Date**: October 7, 2024

## Introduction
Advantage Multi-Vendor Price Scraper ("the Extension") is a Chrome extension designed to scrape pricing data from supported vendor websites (Grainger, McMaster-Carr, Festo, and Zoro) and send it to a user-controlled local server for processing. This Privacy Policy outlines how the Extension handles data, including collection, use, and sharing practices.

## Data Collection
The Extension collects the following data when you use it on supported vendor websites:
- **Product Information**: Part number, description, price, unit, manufacturer number, brand, URL of the product page, and timestamp.
- **Tab Information**: Tab ID and URL of the webpage being scraped to manage scraping sessions.
- **No Personal Data**: The Extension does not collect personally identifiable information (e.g., name, email, or payment details) unless explicitly provided by you in the scraped data (which is not expected or required).

## Data Usage
- **Purpose**: Collected data is used solely to extract pricing and product details from supported vendor websites and send them to a local server running on your device (`http://localhost:5000`) for further processing.
- **Local Processing**: All data is sent to your local server, which you control. The Extension does not transmit data to external servers or third parties unless explicitly configured by you.
- **Temporary Storage**: Data is temporarily held in memory during scraping and transmission to the local server. No persistent storage occurs within the Extension itself.

## Data Sharing
- **No Third-Party Sharing**: The Extension does not share data with third parties, including the developer or external services, unless you configure the local server to do so.
- **Local Server**: Data is sent to your local server (`http://localhost:5000`). You are responsible for securing this server and any subsequent data handling.
- **Browser APIs**: The Extension uses Chrome APIs (e.g., `chrome.runtime`, `chrome.tabs`) to manage tabs and communicate with the local server, but no data is sent to Google or other entities via these APIs.

## Data Security
- **Transmission**: Data is sent to your local server using HTTP POST requests. As the server runs locally (`http://localhost`), it is not exposed to external networks unless you configure otherwise.
- **No Cloud Storage**: The Extension does not store data in the cloud or on external servers.
- **Your Responsibility**: You are responsible for securing your local server and ensuring it uses appropriate security measures (e.g., HTTPS if extended beyond localhost).

## User Control
- **Access**: You can review the scraped data by inspecting the requests sent to your local server.
- **Deletion**: Since data is not stored by the Extension, deletion is managed by your local server. You can stop the server or clear its data to remove collected information.
- **Opt-Out**: To stop data collection, uninstall the Extension or refrain from using it on supported vendor websites.

## Cookies and Tracking
- The Extension does not use cookies, tracking technologies, or analytics services.
- It only accesses the content of supported vendor webpages to extract product details.

## Compliance with Chrome Web Store Policies
This Extension complies with Chrome Web Store policies by:
- Clearly disclosing data collection and usage practices.
- Limiting data collection to what is necessary for functionality.
- Not sharing data with third parties without user consent.

## Changes to This Policy
We may update this Privacy Policy to reflect changes in the Extension’s functionality or legal requirements. Updates will be reflected in the hosted policy document, and you are encouraged to review it periodically.

## Contact
For questions about this Privacy Policy, contact the developer via GitHub Issues at https://github.com/DorelRoata/ChromeExtension/issues