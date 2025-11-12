"""
Diagnostic script to check date formats in Excel file
"""
import openpyxl
from datetime import datetime

# Open the Excel file
file_path = 'MML.xlsm'
workbook = openpyxl.load_workbook(file_path, read_only=False, keep_vba=True)
sheet = workbook["Purchase Parts"]

# Check row 2 (first data row after header)
print("=" * 80)
print("ROW 2 INSPECTION (First Data Row)")
print("=" * 80)

row = sheet[2]
for idx, cell in enumerate(row[:15], start=1):
    value = cell.value
    value_type = type(value).__name__
    number_format = cell.number_format

    # Special attention to date columns (11 and 14 in 1-indexed)
    if idx in [11, 14]:  # Date columns
        print(f"\n*** Column {idx} (Date Field) ***")
        print(f"  Raw Value: {repr(value)}")
        print(f"  Python Type: {value_type}")
        print(f"  Excel Number Format: {number_format}")
        print(f"  Is datetime object: {isinstance(value, datetime)}")
    else:
        print(f"\nColumn {idx}:")
        print(f"  Value: {value}")
        print(f"  Type: {value_type}")
        print(f"  Number Format: {number_format}")

print("\n" + "=" * 80)
print("WHAT THE CODE IS CURRENTLY DOING")
print("=" * 80)

# Show what the code does
test_date_string = datetime.now().strftime("%m/%d/%Y")
print(f"\nCurrent code writes: {repr(test_date_string)}")
print(f"Type: {type(test_date_string).__name__}")
print(f"This is a STRING, not a datetime object!")

print("\n" + "=" * 80)
print("WHAT IT SHOULD BE")
print("=" * 80)
print(f"\nShould write: datetime.now()")
print(f"Type: {type(datetime.now()).__name__}")
print(f"Excel will automatically format this as a date")

workbook.close()
