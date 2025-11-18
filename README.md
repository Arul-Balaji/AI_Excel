# Sales Forecast Tracker - xlwings Integration

This project provides Python code to programmatically interact with the Microsoft Sales Forecast Tracker Excel spreadsheet using the `xlwings` module.

## Files

- **`sales_forecast_model.py`** - Main module with the `SalesForecastModel` class
- **`demo_add_row.py`** - Demonstration script showing how to add a new forecast input row
- **`inspect_spreadsheet.py`** - Utility script for detailed spreadsheet inspection
- **`microsoft_Sales forecast tracker small business.xlsx`** - The Excel workbook

## Installation

Install required dependencies:

```bash
pip install xlwings pandas
```

## Spreadsheet Structure

### Forecast Input Tab
- **Headers**: Row 6, Columns B-J
- **Data starts**: Row 7
- **Columns**:
  - Opportunity name
  - Sales agent
  - Sales region
  - Sales category
  - Forecast amount
  - Sales phase
  - Probability of sale
  - Forecast close (date)
  - Weighted forecast

### Sales Forecast Tab
- **Headers**: Row 6, Columns P-R
- **Data starts**: Row 7
- **Columns**:
  - Month (date)
  - Monthly Forecast (amount)
  - Cumulative (running total)

## Usage

### 1. Inspect the Spreadsheet

```python
from sales_forecast_model import SalesForecastModel

model = SalesForecastModel("microsoft_Sales forecast tracker small business.xlsx")
model.inspect_spreadsheet()
```

### 2. Add a New Forecast Input Row

```python
from sales_forecast_model import SalesForecastModel
from datetime import datetime

model = SalesForecastModel("microsoft_Sales forecast tracker small business.xlsx")

new_data = {
    'Opportunity name': 'Tech Innovations Inc',
    'Sales \nagent': 'Jane Smith',
    'Sales \nregion': 'US - West',
    'Sales \ncategory': 'Products',
    'Forecast amount': 250000,
    'Sales \nphase': 'Proposal submitted',
    'Probability of sale': 0.6,
    'Forecast \nclose': datetime(2026, 11, 15),
    'Weighted forecast': 150000
}

row_number = model.add_forecast_input_row(new_data)
print(f"Added row at position {row_number}")
```

**Note**: Some column headers contain `\n` (newline characters) in the Excel file, so you must use the exact header names including `\n`.

### 3. Read Sales Forecast Data

```python
from sales_forecast_model import SalesForecastModel

model = SalesForecastModel("microsoft_Sales forecast tracker small business.xlsx")

# Read as pandas DataFrame
forecast_df = model.read_sales_forecast()
print(forecast_df)

# Access specific data
total_forecast = forecast_df['Monthly Forecast'].sum()
final_cumulative = forecast_df['Cumulative'].iloc[-1]

print(f"Total Monthly Forecast: ${total_forecast:,.2f}")
print(f"Final Cumulative: ${final_cumulative:,.2f}")
```

### 4. Read Specific Range from Sales Forecast

```python
from sales_forecast_model import SalesForecastModel

model = SalesForecastModel("microsoft_Sales forecast tracker small business.xlsx")

# Read a specific range
data = model.read_sales_forecast_range('P7:R12')
print(data)
```

## Running the Examples

### Inspect the spreadsheet:
```bash
python sales_forecast_model.py
```

### Add a new row (demonstration):
```bash
python demo_add_row.py
```

### Detailed inspection:
```bash
python inspect_spreadsheet.py
```

## Class Reference

### `SalesForecastModel`

#### Methods

- **`__init__(file_path)`** - Initialize with path to Excel file
- **`inspect_spreadsheet()`** - Print detailed information about both tabs
- **`add_forecast_input_row(data_dict)`** - Add a new row to the Forecast input tab
  - Returns: Row number where data was inserted
- **`read_sales_forecast(as_dataframe=True)`** - Read forecast outputs
  - Returns: pandas DataFrame or list of lists
- **`read_sales_forecast_range(range_address)`** - Read specific range from Sales forecast tab
  - Returns: List of data from the specified range

## Important Notes

1. **Excel must be installed** on Windows for xlwings to work
2. The workbook will be **automatically saved** when adding new rows
3. Column headers in the Excel file contain newline characters (`\n`) - use exact header names
4. The code handles the non-standard structure where headers are in row 6, not row 1
5. Data is written starting from column B (column A is empty in the template)

## Example Output

When reading sales forecast data:

```
        Month  Monthly Forecast  Cumulative
0  2026-01-01            151600           0
1  2026-02-01            160320      151600
2  2026-03-01            313500      311920
...

Total Monthly Forecast: $2,018,845.00
Final Cumulative: $1,775,445.00
```
