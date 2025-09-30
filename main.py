import os
import sys

import time
import queue
import logging
import threading
import webbrowser
import subprocess
import tkinter as tk
from tkinter import simpledialog, messagebox
import tkinter.font as tkFont
from datetime import datetime
from openpyxl import load_workbook
from flask import Flask, request, jsonify
from flask_cors import CORS

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

# Global configuration
SERVER_PORT = 5000
DATA_QUEUE = queue.Queue()
FLASK_APP = None

# Schema fields
FIELDS = [
    'ACI #', 'MFR Part #', 'MFR', 'Description', 'QTY', 'Per', 
    'Vendor', 'Vendor Part #', 'Legacy', 'Unit Price', 'Change %', 
    'Date', 'Last Updated Price', 'Last Updated Date', 'Price History'
    
]

VENDOR_URLS = {
    'grainger': 'https://www.grainger.com/product/{}/',
    'mcmaster-carr': 'https://www.mcmaster.com/{}/',
    'mcmaster': 'https://www.mcmaster.com/{}/',
    'festo': 'https://www.festo.com/us/en/a/{}',
    'zoro': 'https://www.zoro.com/i/{}/'
}

#
# FLASK SERVER
#
def create_flask_app():
    app = Flask(__name__)
    CORS(app)
    
    @app.route('/ping', methods=['GET'])
    def ping():
        return jsonify({"status": "ok", "timestamp": time.time()})
    
    @app.route('/scrape', methods=['POST'])
    def receive_scrape():
        try:
            data = request.json
            logger.info(f"Received data from extension: {data.get('vendor')} - {data.get('partNumber')}")
            DATA_QUEUE.put(data)
            return jsonify({"status": "success"}), 200
        except Exception as e:
            logger.error(f"Error receiving data: {e}")
            return jsonify({"status": "error", "message": str(e)}), 500
    
    return app

def start_flask_server():
    """Start Flask server in background thread"""
    global FLASK_APP

    # Check if port is already in use
    import socket
    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    result = sock.connect_ex(('localhost', SERVER_PORT))
    sock.close()

    if result == 0:
        logger.warning(f"Port {SERVER_PORT} is already in use. Another instance may be running.")
        from tkinter import messagebox
        response = messagebox.askyesno(
            "Port Already in Use",
            f"Port {SERVER_PORT} is already in use.\n\n"
            "This usually means another instance is running.\n\n"
            "Continue anyway? (May cause errors)"
        )
        if not response:
            sys.exit(0)

    FLASK_APP = create_flask_app()

    def run_server():
        logger.info(f"Starting Flask server on port {SERVER_PORT}...")
        try:
            FLASK_APP.run(host='localhost', port=SERVER_PORT, debug=False, use_reloader=False, threaded=True)
        except OSError as e:
            logger.error(f"Failed to start server: {e}")

    server_thread = threading.Thread(target=run_server, daemon=True)
    server_thread.start()
    time.sleep(1)  # Wait for server to start
    logger.info("Flask server started")

#
# BROWSER CONTROLLER
#
class BrowserController:
    @staticmethod
    def open_vendor_page(vendor_name, part_number):
        """Open vendor page in Chrome"""
        vendor_key = vendor_name.lower().strip()

        if vendor_key not in VENDOR_URLS:
            logger.error(f"Unknown vendor: {vendor_name}")
            return False

        url = VENDOR_URLS[vendor_key].format(part_number)
        logger.info(f"Opening {url} in Chrome...")

        try:
            # Force Chrome browser
            chrome_path = None
            if sys.platform == 'win32':
                chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
                if not os.path.exists(chrome_path):
                    chrome_path = r'C:\Program Files (x86)\Google\Chrome\Application\chrome.exe'
            elif sys.platform == 'darwin':
                chrome_path = 'open -a /Applications/Google\ Chrome.app'
            else:  # Linux
                chrome_path = '/usr/bin/google-chrome'

            if chrome_path and os.path.exists(chrome_path):
                webbrowser.register('chrome', None, webbrowser.BackgroundBrowser(chrome_path))
                webbrowser.get('chrome').open(url)
            else:
                logger.warning("Chrome not found at default location, using default browser")
                webbrowser.open(url)
            return True
        except Exception as e:
            logger.error(f"Failed to open browser: {e}")
            return False
    
    @staticmethod
    def wait_for_scraped_data(timeout=30):
        """Wait for extension to send scraped data"""
        logger.info(f"Waiting up to {timeout}s for scraped data...")
        start_time = time.time()
        
        while time.time() - start_time < timeout:
            try:
                data = DATA_QUEUE.get(timeout=0.5)
                logger.info("Data received from extension")
                return data
            except queue.Empty:
                continue
        
        logger.warning("Timeout waiting for scraped data")
        return None

#
# DATA PROCESSING
#
def sanitize_string(value):
    """Clean string data"""
    if isinstance(value, str):
        return value.replace('\xa0', ' ').replace('\u200e', '').strip()
    return value

def parse_vendor_data(raw_data, vendor_name):
    """Parse raw scraped data into standardized format"""
    if not raw_data:
        return None
    
    vendor_key = vendor_name.lower().strip()
    
    # Extract common fields
    description = raw_data.get('description', 'Not Found')
    price_raw = raw_data.get('price', 'Not Found')
    unit = raw_data.get('unit', 'Not Found')
    mfr_number = raw_data.get('mfrNumber', 'Not Found')
    brand = raw_data.get('brand', 'Not Found')
    
    # Vendor-specific cleanup
    if vendor_key == 'grainger':
        price, unit, qty = cleanup_grainger_data(price_raw, unit, mfr_number)
    elif vendor_key in ['mcmaster', 'mcmaster-carr']:
        price, unit, qty = cleanup_mcmaster_data(price_raw, unit)
        # For McMaster, use part number as MFR number
        mfr_number = raw_data.get('partNumber', mfr_number)
    elif vendor_key == 'festo':
        price, unit, qty = cleanup_festo_data(price_raw, unit, raw_data.get('qty'))
    elif vendor_key == 'zoro':
        price, unit, qty = cleanup_zoro_data(price_raw, unit)
    else:
        price = price_raw
        qty = 1
    
    return {
        'description': description,
        'price': price,
        'unit': unit,
        'qty': qty,
        'mfr_number': mfr_number,
        'brand': brand
    }

def cleanup_grainger_data(price, unit, mfr):
    """Clean Grainger-specific data"""
    # Remove $ and clean price
    price = price.replace('$', '').strip() if price != "Not Found" else "Not Found"
    
    # Parse unit and quantity
    unit = unit.replace('/', '').strip()
    qty = 1
    
    if " of " in unit:
        parts = unit.split(" of ")
        unit = parts[0].strip()
        try:
            qty = int(parts[1].strip())
        except ValueError:
            qty = 1
    
    # Clean MFR number
    if ' ' in mfr:
        mfr = mfr.split()[-1]
    
    return price, unit, qty

def cleanup_mcmaster_data(price, unit):
    """Clean McMaster-specific data"""
    # Parse price
    if "Not Found" not in price:
        if " per " in price.lower():
            parts = price.lower().split("per", 1)
            price = parts[0].replace('$', '').strip()
        elif "each" in price.lower():
            parts = price.lower().split("each", 1)
            price = parts[0].replace('$', '').strip()
        else:
            price = price.replace('$', '').strip()
    
    # Parse unit and quantity
    qty = 1
    if "each" in unit.lower():
        qty = 1
    elif "pack" in unit.lower() and "of" in unit.lower():
        parts = unit.split(" of ", 1)
        unit = parts[0].strip()
        try:
            qty = int(parts[1].strip())
        except ValueError:
            qty = 1
    
    return price, unit, qty

def cleanup_festo_data(price, unit, qty_raw):
    """Clean Festo-specific data"""
    price = price.replace('$', '').strip() if price != "Not Found" else "Not Found"
    
    qty = 1
    if qty_raw and qty_raw != "Not Found":
        try:
            qty = int(qty_raw)
        except ValueError:
            qty = 1
    
    return price, unit, qty

def cleanup_zoro_data(price, unit):
    """Clean Zoro-specific data"""
    qty = 1
    
    if "Not Found" not in price:
        # Remove newlines and extra text
        price = price.replace('\r\n', ' ').replace('\n', ' ')
        price = price.replace('product price:', '').strip()
        
        # Parse format: "$123.45 / pk 10" or "$45.67 / ea"
        parts = price.split(',')
        
        for part in parts:
            part = part.strip().replace('$', '')
            detail_parts = part.split('/')
            
            if len(detail_parts) >= 2:
                try:
                    price = float(detail_parts[0].strip())
                except ValueError:
                    continue
                
                unit_part = detail_parts[1].strip()
                
                if 'pk' in unit_part:
                    unit = 'pk'
                    try:
                        qty = int(unit_part.split()[1])
                    except:
                        qty = 1
                elif 'ea' in unit_part:
                    unit = 'each'
                    qty = 1
                elif 'pr' in unit_part:
                    unit = 'pair'
                    qty = 2
                
                break
    
    return str(price), unit, qty

def calculate_percentage_change(old_price, new_price):
    """Calculate percentage change between prices"""
    def clean_price(price):
        if price is None or price == '' or price == 'Not Found':
            return None
        if isinstance(price, (int, float)):
            return float(price)
        cleaned = ''.join(c for c in str(price) if c.isdigit() or c == '.')
        return float(cleaned) if cleaned else None
    
    try:
        old = clean_price(old_price)
        new = clean_price(new_price)
        
        if old is None or new is None:
            return None
        
        if old == 0:
            return float('inf') if new != 0 else 0.0
        
        return round(((new - old) / old) * 100, 2)
    except Exception as e:
        logger.error(f"Error calculating percentage change: {e}")
        return None

def update_price_history(entry_data, current_data):
    """Update price history in CSV format"""
    if current_data[14] is None:
        current_data[14] = ""
    
    if current_data[14] != "":
        entry_data[14] = current_data[14] + f", Date: {current_data[11]} Price: {current_data[9]}"
    else:
        entry_data[14] = f"Date: {current_data[11]} Price: {current_data[9]}"
    
    entry_data[12] = current_data[9]  # Last updated price
    entry_data[13] = current_data[11]  # Last updated date
    
    return entry_data

#
# EXCEL OPERATIONS
#
def process_excel(file_path, search_string):
    """Search for ACI number and return row data"""
    workbook = None
    try:
        logger.info(f"Searching Excel for ACI#: {search_string}")
        workbook = load_workbook(file_path, read_only=False, keep_vba=True)
        sheet = workbook["Purchase Parts"]
        
        for row in sheet.iter_rows(min_row=1, max_col=15, max_row=sheet.max_row):
            cell_value = row[0].value
            # Cast both to string and strip whitespace for comparison
            if cell_value is not None:
                # Handle both numeric and text values
                cell_str = str(sanitize_string(cell_value)).strip()
                search_str = str(search_string).strip()
                if cell_str == search_str:
                    row_index = row[0].row
                    current_data = [sanitize_string(cell.value) for cell in row[:15]]
                    logger.info(f"Found match at row {row_index}")
                    return current_data, row_index
        
        logger.info("No match found")
        messagebox.showinfo("Info", "No match found for that ACI number")
        return None, None
    except Exception as e:
        logger.error(f"Excel error: {e}")
        return None, None
    finally:
        if workbook:
            workbook.close()

def save_to_excel(file_path, row_index, data):
    """Save updated data to Excel"""
    workbook = None
    try:
        workbook = load_workbook(file_path, read_only=False, keep_vba=True)
        sheet = workbook["Purchase Parts"]
        
        for idx, value in enumerate(data):
            sheet.cell(row=row_index, column=idx + 1, value=value)
        
        workbook.save(filename=file_path)
        logger.info("Excel file saved successfully")
        return True
    except Exception as e:
        logger.error(f"Error saving Excel: {e}")
        return False
    finally:
        if workbook:
            workbook.close()

#
# GUI COMPONENTS
#
def get_search_string():
    """Prompt user for ACI number"""
    root = tk.Tk()
    root.title("Advantage Conveyor")

    # Set window size
    window_width = 320
    window_height = 160

    # Get screen dimensions
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Calculate center position
    center_x = int((screen_width - window_width) / 2)
    center_y = int((screen_height - window_height) / 2)

    # Set geometry with centered position
    root.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

    # Status label
    status_label = tk.Label(root, text="Server Status: Running ✓", font=("Arial", 9), fg="green")
    status_label.pack(pady=(12, 8))

    # Instruction label
    instruction_label = tk.Label(root, text="Enter ACI Number:", font=("Arial", 10))
    instruction_label.pack(pady=(8, 5))

    # Entry field
    entry = tk.Entry(root, font=("Arial", 10), width=25)
    entry.pack(pady=8)
    entry.focus_set()

    result = {'value': None}

    def on_submit():
        result['value'] = entry.get()
        root.destroy()

    def on_cancel():
        result['value'] = None
        root.destroy()

    # Bind Enter key to submit
    entry.bind('<Return>', lambda e: on_submit())

    # Buttons
    button_frame = tk.Frame(root)
    button_frame.pack(pady=12)

    cancel_btn = tk.Button(button_frame, text="Cancel", command=on_cancel, font=("Arial", 10), bg="#f44336", fg="white", width=10)
    cancel_btn.pack(side=tk.LEFT, padx=5)

    submit_btn = tk.Button(button_frame, text="Submit", command=on_submit, font=("Arial", 10), bg="#4CAF50", fg="white", width=10)
    submit_btn.pack(side=tk.LEFT, padx=5)

    root.protocol("WM_DELETE_WINDOW", on_cancel)
    root.mainloop()

    return result['value']

def compare_and_highlight(widget, current_value, new_value):
    """Highlight widget if values differ"""
    if str(new_value).strip() != str(current_value).strip():
        widget.config(bg="yellow")
    else:
        widget.config(bg="white")

def switch_checkbox_state(index, checkboxes, text_boxes, current_text_boxes, fields, entry_data, current_data):
    """Toggle checkbox and update display"""
    checkbox = checkboxes[index]
    # Get the state AFTER the checkbox has been toggled (it's already changed by the time this is called)
    state = checkbox.var.get()

    if state:  # Checked - use current data
        if fields[index] == "Description":
            text_boxes[index].delete("1.0", tk.END)
            text_boxes[index].insert(tk.END, current_data[index] if current_data[index] is not None else "")
            text_boxes[index].config(bg="white")
        else:
            text_boxes[index].delete(0, tk.END)
            text_boxes[index].insert(0, str(current_data[index]) if current_data[index] is not None else "")
            text_boxes[index].config(bg="white")
    else:  # Unchecked - use entry data
        if fields[index] == "Description":
            text_boxes[index].delete("1.0", tk.END)
            text_boxes[index].insert(tk.END, entry_data[index] if entry_data[index] is not None else "")
            compare_and_highlight(text_boxes[index], current_data[index], entry_data[index])
        else:
            text_boxes[index].delete(0, tk.END)
            text_boxes[index].insert(0, str(entry_data[index]) if entry_data[index] is not None else "")
            compare_and_highlight(text_boxes[index], current_data[index], entry_data[index])

def user_form(current_data, entry_data, fields, file_path, row_index):
    """Display GUI form for user confirmation"""
    root = tk.Tk()
    root.title("Update Data - ACI# " + str(current_data[0]))

    def on_closing():
        root.quit()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    large_font = tkFont.Font(family="Arial", size=10)

    # Headers
    tk.Label(root, text="Entry Data", font=large_font).grid(row=0, column=1, padx=5, pady=5)
    tk.Label(root, text="Current Data", font=large_font).grid(row=0, column=2, padx=2, pady=5)
    tk.Label(root, text="KEEP", font=large_font).grid(row=0, column=3, padx=(2, 10), pady=2, sticky="w")

    text_boxes = []
    current_text_boxes = []
    checkboxes = []

    for i, field in enumerate(fields):
        tk.Label(root, text=field, font=large_font).grid(row=i + 1, column=0, padx=5, pady=5, sticky="e")

        is_not_found = (entry_data[i] == "Not Found")

        if field == "Description":
            # Multi-line text for description
            text_box = tk.Text(root, font=large_font, height=7, width=40, wrap="word")
            text_box.grid(row=i + 1, column=1, padx=5, pady=5)
            entry_text = "" if is_not_found or entry_data[i] is None else str(entry_data[i])
            text_box.insert(tk.END, entry_text)

            current_text_box = tk.Text(root, font=large_font, height=7, width=30, wrap="word", state="normal", takefocus=0)
            current_text_box.grid(row=i + 1, column=2, padx=50, pady=5)
            current_text = "" if current_data[i] is None else str(current_data[i])
            current_text_box.insert(tk.END, current_text)

            if entry_text.strip() != current_text.strip():
                text_box.config(bg="yellow")
        else:
            # Single-line entry
            text_box = tk.Entry(root, font=large_font, width=40)
            text_box.grid(row=i + 1, column=1, padx=5, pady=5)
            entry_text = "" if is_not_found or entry_data[i] is None else str(entry_data[i])
            text_box.insert(0, entry_text)

            current_text_box = tk.Entry(root, font=large_font, state='normal', width=30, takefocus=0)
            current_text_box.grid(row=i + 1, column=2, padx=50, pady=5)
            current_text = "" if current_data[i] is None else str(current_data[i])
            current_text_box.insert(0, current_text)
            current_text_box.config(state="readonly")

        if is_not_found:
            text_box.config(bg="red")

        compare_and_highlight(text_box, entry_data[i], current_data[i])

        text_boxes.append(text_box)
        current_text_boxes.append(current_text_box)

        # Checkbox
        var = tk.BooleanVar()
        checkbox = tk.Checkbutton(root, text="", variable=var, onvalue=True, offvalue=False, takefocus=0)
        checkbox.var = var
        checkbox.config(command=lambda idx=i: switch_checkbox_state(
            idx, checkboxes, text_boxes, current_text_boxes, fields, entry_data, current_data
        ))
        checkbox.grid(row=i + 1, column=3, padx=(2, 10), pady=2, sticky="w")
        checkboxes.append(checkbox)

    def submit():
        try:
            # Update entry_data based on checkboxes
            for i, field in enumerate(fields):
                checkbox_state = checkboxes[i].var.get()
                if checkbox_state:
                    # Use current data
                    if field == "Description":
                        entry_data[i] = current_text_boxes[i].get("1.0", tk.END).strip()
                    else:
                        entry_data[i] = current_text_boxes[i].get()
                else:
                    # Use entry data
                    if field == "Description":
                        entry_data[i] = text_boxes[i].get("1.0", tk.END).strip()
                    else:
                        entry_data[i] = text_boxes[i].get()
            
            # Calculate price change
            percent_change = 0
            if current_data[9] not in ['Legacy', 'None', None] and entry_data[9] not in ['Legacy', 'None', None]:
                try:
                    percent_change = calculate_percentage_change(current_data[9], entry_data[9])
                    
                    if percent_change and abs(percent_change) >= 1:
                        update_price_history(entry_data, current_data)
                except ValueError as e:
                    logger.error(f"Price conversion error: {e}")
            
            # Alert on significant price changes
            if percent_change and abs(percent_change) >= 20:
                approve = messagebox.askyesno(
                    "Price Change Alert", 
                    f"Price change is significant: {percent_change}%. Proceed?"
                )
                if not approve:
                    return
            elif percent_change and percent_change <= -10:
                approve = messagebox.askyesno(
                    "Price Change Alert", 
                    f"Price decrease: {percent_change}%. Proceed?"
                )
                if not approve:
                    return
            
            # Save to Excel
            if save_to_excel(file_path, row_index, entry_data):
                root.destroy()
            else:
                messagebox.showerror("Error", "Failed to save data")
        
        except Exception as e:
            logger.error(f"Error in submit: {e}")
            messagebox.showerror("Error", str(e))
    
    def cancel():
        root.quit()
        root.destroy()

    # Check All button and action buttons on same row
    check_all_state = {'checked': False}

    def toggle_all_checkboxes():
        check_all_state['checked'] = not check_all_state['checked']
        for checkbox in checkboxes:
            checkbox.var.set(check_all_state['checked'])
            # Trigger the visual update by updating the display
            idx = checkboxes.index(checkbox)
            if check_all_state['checked']:
                # Use current data
                if fields[idx] == "Description":
                    text_boxes[idx].delete("1.0", tk.END)
                    text_boxes[idx].insert(tk.END, current_data[idx] if current_data[idx] is not None else "")
                    text_boxes[idx].config(bg="white")
                else:
                    text_boxes[idx].delete(0, tk.END)
                    text_boxes[idx].insert(0, str(current_data[idx]) if current_data[idx] is not None else "")
                    text_boxes[idx].config(bg="white")
            else:
                # Use entry data
                if fields[idx] == "Description":
                    text_boxes[idx].delete("1.0", tk.END)
                    text_boxes[idx].insert(tk.END, entry_data[idx] if entry_data[idx] is not None else "")
                    compare_and_highlight(text_boxes[idx], current_data[idx], entry_data[idx])
                else:
                    text_boxes[idx].delete(0, tk.END)
                    text_boxes[idx].insert(0, str(entry_data[idx]) if entry_data[idx] is not None else "")
                    compare_and_highlight(text_boxes[idx], current_data[idx], entry_data[idx])

        # Update button text
        if check_all_state['checked']:
            check_all_button.config(text="Uncheck All")
        else:
            check_all_button.config(text="Check All")

    check_all_button = tk.Button(
        root,
        text="Check All",
        command=toggle_all_checkboxes,
        font=large_font,
        takefocus=0,
        width=10
    )
    check_all_button.grid(row=len(fields) + 1, column=2, padx=10, pady=10, sticky="e")

    cancel_button = tk.Button(root, text="Cancel", command=cancel, font=large_font, bg="#f44336", fg="white", takefocus=0)
    cancel_button.grid(row=len(fields) + 1, column=1, padx=20, pady=10)

    submit_button = tk.Button(root, text="Submit", command=submit, font=large_font, bg="#4CAF50", fg="white", takefocus=0)
    submit_button.grid(row=len(fields) + 1, column=2, padx=10, pady=10)
    
    root.protocol("WM_DELETE_WINDOW", cancel)
    root.transient()
    root.grab_set()
    root.focus_set()

    # Center the window after all widgets are added
    root.update_idletasks()
    window_width = root.winfo_width()
    window_height = root.winfo_height()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    center_x = int((screen_width - window_width) / 2)
    center_y = int((screen_height - window_height) / 2)
    root.geometry(f"+{center_x}+{center_y}")

    root.mainloop()

#
# MAIN WORKFLOW
#
def process_item(vendor_name, part_number, current_data, entry_data):
    """Orchestrate the scraping workflow"""
    logger.info(f"Processing {vendor_name} part {part_number}")
    
    # Open browser
    if not BrowserController.open_vendor_page(vendor_name, part_number):
        logger.error("Failed to open browser")
        return False
    
    # Wait for scraped data
    raw_data = BrowserController.wait_for_scraped_data(timeout=30)
    
    if not raw_data:
        logger.warning("No data received from extension")
        messagebox.showwarning("Timeout", "No data received. Extension may not be installed or page took too long.")
        return False
    
    # Parse data
    parsed_data = parse_vendor_data(raw_data, vendor_name)
    
    if not parsed_data:
        logger.error("Failed to parse vendor data")
        return False
    
    # Update entry_data
    entry_data[1] = parsed_data['mfr_number']
    entry_data[2] = parsed_data['brand']
    entry_data[3] = parsed_data['description']
    entry_data[4] = parsed_data['qty']
    entry_data[5] = parsed_data['unit']
    entry_data[9] = parsed_data['price']
    entry_data[10] = calculate_percentage_change(current_data[9], parsed_data['price'])
    entry_data[11] = datetime.now().strftime("%m/%d/%Y")
    
    logger.info(f"Data parsed successfully: Price=${parsed_data['price']}, Qty={parsed_data['qty']}")
    return True

def is_vendor_auto(vendor):
    """Check if vendor supports automatic scraping"""
    return vendor.lower() in ['grainger', 'mcmaster-carr', 'mcmaster', 'zoro', 'festo']

def main_loop(file_path):
    """Main application loop"""
    logger.info("Starting main application loop")
    
    while True:
        search_string = get_search_string()
        
        if search_string is None:
            logger.info("User cancelled, exiting")
            break
        
        search_string = sanitize_string(search_string).upper()
        logger.info(f"Searching for: {search_string}")
        
        try:
            # Search Excel
            result = process_excel(file_path, search_string)
            if result is None or result[0] is None:
                continue
            
            current_data, row_index = result
            entry_data = current_data.copy()
            
            # Check if vendor supports auto-scraping
            vendor_name = current_data[6]
            if is_vendor_auto(vendor_name):
                part_number = current_data[7]
                logger.info(f"Auto-scraping enabled for {vendor_name}")
                process_item(vendor_name, part_number, current_data, entry_data)
            else:
                logger.info(f"Manual vendor: {vendor_name}")
            
            # Show user form
            user_form(current_data, entry_data, FIELDS, file_path, row_index)
        
        except Exception as e:
            logger.error(f"Error in main loop: {e}", exc_info=True)
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            continue

#
# APPLICATION ENTRY POINT
#
def main():
    """Application entry point"""
    try:
        # Determine file path
        server_file_path = r'Z:\ACOD\MMLV2.xlsm'
        local_file_path = 'MML.xlsm'

        if os.path.exists(server_file_path):
            file_path = server_file_path
            logger.info(f"Using server file: {file_path}")
        else:
            file_path = local_file_path
            logger.info(f"Using local file: {file_path}")
            messagebox.showwarning("Warning", "Server file not found. Using local file.")

        # Start Flask server
        start_flask_server()

        # Start main loop
        main_loop(file_path)

    except KeyboardInterrupt:
        logger.info("Program terminated by user")
        sys.exit(0)
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        messagebox.showerror("Fatal Error", str(e))
        sys.exit(1)
    finally:
        logger.info("Cleaning up and exiting...")
        # Daemon threads will automatically terminate when main thread exits

if __name__ == "__main__":
    main()