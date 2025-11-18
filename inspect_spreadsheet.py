import xlwings as xw
import pandas as pd

# Open the workbook
wb = xw.Book("microsoft_Sales forecast tracker small business.xlsx")

print("=" * 80)
print("DETAILED INSPECTION OF FORECAST INPUT TAB")
print("=" * 80)

input_sheet = wb.sheets['Forecast input']

# Read a larger range to see the structure
print("\nFirst 15 rows of data (columns A-K):")
data = input_sheet.range('A1:K15').value
for i, row in enumerate(data, 1):
    print(f"Row {i:2d}: {row}")

print("\n" + "=" * 80)
print("DETAILED INSPECTION OF SALES FORECAST TAB")
print("=" * 80)

forecast_sheet = wb.sheets['Sales forecast']

# Read a larger range to see the structure
print("\nFirst 20 rows of data (columns A-R):")
data = forecast_sheet.range('A1:R20').value
for i, row in enumerate(data, 1):
    print(f"Row {i:2d}: {row}")

wb.close()
