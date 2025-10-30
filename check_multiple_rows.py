"""
Check date formats across multiple rows
"""
import openpyxl
from datetime import datetime

# Open the Excel file
file_path = 'MML.xlsm'
workbook = openpyxl.load_workbook(file_path, read_only=False, keep_vba=True)
sheet = workbook["Purchase Parts"]

print("=" * 80)
print("CHECKING DATE COLUMNS (11 and 14) ACROSS MULTIPLE ROWS")
print("=" * 80)

# Check first 20 rows
for row_num in range(2, min(22, sheet.max_row + 1)):
    row = sheet[row_num]

    # Column 11 (index 10) - Date field
    date_cell = row[10]
    date_value = date_cell.value
    date_type = type(date_value).__name__
    date_format = date_cell.number_format

    # Column 14 (index 13) - Last Updated Date field
    last_date_cell = row[13]
    last_date_value = last_date_cell.value
    last_date_type = type(last_date_value).__name__
    last_date_format = last_date_cell.number_format

    print(f"\nRow {row_num}:")
    print(f"  Col 11 (Date): value={repr(date_value)[:40]}, type={date_type}, format={date_format}")
    print(f"  Col 14 (Last Date): value={repr(last_date_value)[:40]}, type={last_date_type}, format={last_date_format}")

workbook.close()

print("\n" + "=" * 80)
print("SUMMARY")
print("=" * 80)
print("\nKEY FINDINGS:")
print("1. Check if existing dates are stored as datetime objects or strings")
print("2. Check the number_format applied to date columns")
print("3. The code currently writes strings like '10/14/2025'")
print("4. It SHOULD write datetime objects like: datetime(2025, 10, 14)")
