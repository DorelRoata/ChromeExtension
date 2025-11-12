# Code Improvements Summary

## Implemented Changes

### 1. Memory Icon Handling (Issue #8)
**Problem**: Repeated creation of `PhotoImage` objects can leak memory if not referenced.

**Current State**:
- Added helper `set_window_icon_safe()` to attach icon refs to window objects and avoid globals.
- The code currently still uses a cached global `_APP_ICON_PHOTO` to prevent GC, which mitigates leaks by reusing a single image.

**Next Step**:
- Adopt `set_window_icon_safe()` in all GUI windows to remove the remaining global usage.

**Files**: `main.py`
- Helper present near top of file.

### 2. Error Handling Improvement (Issue #12)
**Problem**: Inconsistent error messages and handling patterns throughout codebase.

**Solution**:
- Created `handle_error()` centralized error handling function
- Provides user-friendly messages based on exception type
- Consistent logging with proper context
- Updated 15+ error handling blocks throughout codebase

**Features**:
- Context-aware error messages (FileNotFound, PermissionError, Excel errors, etc.)
- Consistent logging with exc_info for debugging
- Optional user notification with parent window support
- Different log levels (error, warning, info)

### 3. Window Position Tracking (Issue #13)
**Problem**: Windows don't remember their positions between sessions.

**Solution**:
- Added JSON-based configuration system
- `load_window_positions()` and `save_window_positions()` functions
- `setup_window_position()` function for automatic position management
- Removed manual window centering code from all GUI functions

**Features**:
- Remembers position and size for each window type
- Falls back to center screen if no saved position
- Saves position on window close
- Non-critical feature (fails gracefully)

## New Functions Added

```python
# Memory Management
def set_window_icon_safe(window, icon_ico_path=None, icon_png_path=None)

# Error Handling  
def handle_error(exception, context="", show_user=True, parent_window=None, log_level="error")

# Window Position Management
def load_window_positions()
def save_window_positions(positions)
def setup_window_position(window, window_name, default_size=(400, 200))
```

## Configuration File

Creates `window_config.json` in program directory:
```json
{
  "main": {"x": 100, "y": 100, "width": 400, "height": 200},
  "batch": {"x": 150, "y": 150, "width": 500, "height": 400},
  "form": {"x": 200, "y": 200, "width": 900, "height": 700},
  "summary": {"x": 250, "y": 250, "width": 700, "height": 500},
  "new_aci": {"x": 300, "y": 300, "width": 400, "height": 300}
}
```

## Testing

Created `test_improvements.py` to verify all improvements:
- Memory management test (multiple window creation/destruction)
- Error handling test (various exception types)
- Window position tracking test
- Configuration file test

All tests pass successfully.

## Benefits

1. **Memory**: No more PhotoImage memory leaks
2. **User Experience**: Better error messages with actionable guidance
3. **Convenience**: Windows remember positions between sessions
4. **Maintainability**: Centralized error handling and icon management
5. **Zero Dependencies**: No new packages required
6. **Performance**: Minimal impact on application speed

## Backward Compatibility

- All existing functionality preserved
- No breaking changes to user workflow
- Configuration file is optional (created automatically)
- Error messages are improved but still functional

## Files Modified

- `main.py`: Core improvements
- `test_improvements.py`: New test suite (optional)
- `IMPROVEMENTS_SUMMARY.md`: This documentation

## Backup

Original file backed up as `main.py.backup` before changes.

---

## Recent Changes (2025-11-06)

### 4. Batch Update: Vendor-specific Hyphen Rule
**Problem**: Batch update skipped any ACI containing a hyphen, even for vendors that don’t require it.

**Solution**:
- Only skip hyphenated ACI numbers when the vendor is McMaster or McMaster‑Carr.
- All other vendors process hyphenated ACIs normally.

**Files Modified**: `main.py`
- Adjusted logic in `batch_update_worker()` to check `vendor_name` before skipping.

### 5. Excel: Preserve Numeric Type for ACI (VLOOKUP)
**Problem**: ACI numbers were sometimes written as text, breaking Excel VLOOKUP against numeric keys.

**Solution**:
- Added `prepare_aci_for_excel()` to coerce digit‑only ACI values to integers while preserving mixed/alpha IDs.
- Ensured column 1 (ACI #) is written as numeric when possible in both updates and inserts.

**Files Modified**: `main.py`
- `prepare_aci_for_excel()` added.
- `save_to_excel()` writes ACI as numeric when digit‑only.
- `add_new_row_to_excel()` writes new ACI as numeric when digit‑only.

### 6. Add Vendors to New ACI Dropdown
**Change**: Added non‑auto vendors to the vendor selection list when adding a new ACI:
- ABB Baldor, Allen Bradley, Habasit, Etcetera

**Files Modified**: `main.py`
- Updated `get_new_aci_details()` vendor options list.

### 7. Allow Back‑Dated Entry Dates
**Problem**: The Date field was always overwritten with today’s date on submit.

**Solution**:
- Preserve the user‑entered date if valid (parsed via `prepare_date_for_excel()`); otherwise, default to today.

**Files Modified**: `main.py`
- Updated submit logic in `user_form()` to keep user date.

### 8. Build Script Improvements
**Enhancements**:
- OS‑aware `--add-data` separator handling.
- Added `--clean` and argument logging.
- Added defensive hidden imports (`et_xmlfile`, `jdcal`) and `--collect-all=openpyxl` for reliability.

**Files Modified**: `build.py`

### 9. Extension: Move Localhost Calls to Background
**Problem**: Content script network calls to localhost triggered Chrome local network prompts and were unreliable.

**Solution**:
- Proxy all server calls via the background service worker.
- Use 127.0.0.1 consistently; add PNA header server-side.
- Tightened success reporting: popup only shows success when the server accepts payload.

**Files Modified**: `extension/background.js`, `extension/content.js`, `extension/popup.js`, `extension/manifest.json`, `main.py`

### 10. Extension Reliability Follow‑up (Timeout Unchanged)
**Problem**: Popup showed success even when server didn’t receive data; server timed out waiting for data.

**Solution**:
- Gate “success” on actual server 2xx response from background fetch.
- Bind Flask to `127.0.0.1` and align all extension URLs to `127.0.0.1` (manifest allows both 127.0.0.1 and localhost).
- Keep wait timeout at 15s; do not increase.
- Add lightweight background warnings on failed network calls.

**Key References**:
- `main.py`:148 (PNA header), `main.py`:229 (port check to 127.0.0.1), `main.py`:249 (Flask bind to 127.0.0.1), `main.py`:345/1355/1566 (15s waits)
- `extension/background.js`:2 (SERVER_URL), 41–115 (register/scrape/poll with result checks)
- `extension/content.js`: send via background and only report success on 2xx
- `extension/popup.js`:1 (SERVER_URL)
- `extension/manifest.json`: host_permissions include 127.0.0.1 and localhost

---

## Recent Changes (2025-11-12)

### 11. Browser: Robust Chrome Detection (Windows/macOS/Linux)
**Problem**: Opening vendor pages failed on some systems when Chrome wasn’t on PATH.

**Solution**:
- Added `BrowserController.find_chrome_path()` with Windows Registry lookup and common fallbacks; supports macOS/Linux paths.
- `open_vendor_page()` uses detected Chrome when available; falls back to default browser.

**Files**: `main.py:262` (find_chrome_path), `main.py:317` (open_vendor_page)

### 12. Excel Typing + Formatting Consistency
**Problem**: Some fields were saved as strings, causing VLOOKUP/format issues (dates/prices/ACI).

**Solution**:
- `prepare_aci_for_excel()` stores digit-only ACIs as integers.
- `prepare_date_for_excel()` writes true date objects (no time); display formatting controlled by column formats.
- `format_price_value()` ensures numeric prices; `clean_value_for_excel()` converts placeholders like "Not Found" to blanks.
- `save_to_excel()` applies type-aware conversions and column number formats.
- `add_new_row_to_excel()` writes Date as `datetime.now()` (typed), preserving formats by copying the last row.

**Files**: `main.py:422` (prepare_date_for_excel), `main.py:461` (prepare_aci_for_excel), `main.py:734` (save_to_excel), `main.py:777` (add_new_row_to_excel)

### 13. Vendor Data Normalization
**Problem**: Units/quantities/prices differed per site and required cleanup.

**Solution**:
- Centralized parsing via `parse_vendor_data()` plus per-vendor cleaners:
  - Grainger: split unit "each/pack of N", clean MFR.
  - McMaster: parse "$X per pack of N" and "each" variants.
  - Festo: type price; infer qty.
  - Zoro: parse "$X / pk N", "$X / ea", "$X / pr"; set unit and qty.

**Files**: `main.py:522` (cleanup_grainger_data), `main.py:545` (cleanup_mcmaster_data), `main.py:594` (cleanup_festo_data), `main.py:607` (cleanup_zoro_data)

### 14. Batch Matching Refinements
**Change**:
- Part-number match is skipped for McMaster/McMaster‑Carr (their product/MFR patterns differ).
- Unit match check is tolerant of unavailable values ("Not Found").
- Hyphenated ACI skip remains vendor-specific to McMaster only (see 2025‑11‑06 #4).

**Files**: `main.py:1482` (validate_batch_match), `main.py:1524` (batch rules)

### 15. Extension UX: Optional Server / Standalone Messaging
**Change**:
- Popup shows "Server: Connected ✓ (Optional)" and "Not Connected (Standalone Mode)" to reflect background-proxy behavior.
- Success now reflects background 2xx acceptance; content script shows a toast only when payload accepted.

**Files**: `extension/popup.js`, `extension/content.js`, `extension/background.js`

### 16. Misc
- Flask binds to `127.0.0.1` explicitly; PNA header added in `after_request` hook.
- Removed stale `main.py.backup` file to avoid drift.
- Minor CRLF normalization across several files.

**Files**: `main.py:145` (after_request header), `main.py:249` (Flask bind), repo cleanup
