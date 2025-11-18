import xlwings as xw
import pandas as pd
from datetime import datetime
import os


class SalesForecastModel:
    """
    A class to interact with the Microsoft Sales Forecast Tracker spreadsheet.
    Provides methods to add forecast inputs and read forecast outputs.
    """
    
    def __init__(self, file_path="microsoft_Sales forecast tracker small business.xlsx"):
        """
        Initialize the SalesForecastModel with the path to the Excel file.
        
        Args:
            file_path (str): Path to the Excel file
        """
        self.file_path = file_path
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")
    
    def inspect_spreadsheet(self):
        """
        Inspect the spreadsheet structure and print information about both tabs.
        """
        print(f"Inspecting spreadsheet: {self.file_path}\n")
        
        # Open the workbook (visible=False means it opens in background)
        wb = xw.Book(self.file_path)
        
        try:
            # List all sheets
            print("Available sheets:")
            for sheet in wb.sheets:
                print(f"  - {sheet.name}")
            print()
            
            # Inspect 'Forecast input' tab
            if 'Forecast input' in [s.name for s in wb.sheets]:
                input_sheet = wb.sheets['Forecast input']
                print("=== FORECAST INPUT TAB ===")
                
                # Headers are in row 6, starting from column B
                headers = input_sheet.range('B6:J6').value
                print(f"Headers (Row 6, Columns B-J): {headers}")
                
                # Find the last row with data in column B
                last_row = input_sheet.range('B7').end('down').row
                print(f"Last row with data: {last_row}")
                print(f"Number of data rows: {last_row - 6}")
                
                # Show first few rows of data
                data_range = input_sheet.range(f'B6:J{min(last_row, 10)}').value
                df = pd.DataFrame(data_range[1:], columns=data_range[0])
                print(f"\nFirst few rows:\n{df}\n")
            
            # Inspect 'Sales forecast' tab
            if 'Sales forecast' in [s.name for s in wb.sheets]:
                forecast_sheet = wb.sheets['Sales forecast']
                print("=== SALES FORECAST TAB ===")
                
                # Headers are in row 6, columns B-D contain the main forecast data
                headers = forecast_sheet.range('B6:D6').value
                print(f"Headers (Row 6, Columns B-D): {headers}")
                
                # Find the last row with data in column B
                last_row = forecast_sheet.range('B7').end('down').row
                print(f"Last row with data: {last_row}")
                print(f"Number of forecast months: {last_row - 6}")
                
                # Show forecast data
                data_range = forecast_sheet.range(f'B6:D{last_row}').value
                df = pd.DataFrame(data_range[1:], columns=data_range[0])
                print(f"\nForecast data:\n{df}\n")
        
        finally:
            # Close the workbook without saving
            wb.close()
    
    def add_forecast_input_row(self, data_dict):
        """
        Add a new row to the 'Forecast input' tab.
        
        Args:
            data_dict (dict): Dictionary with column names as keys and values to insert
                             Example: {
                                 'Opportunity name': 'New Client Corp',
                                 'Sales agent': 'John Doe',
                                 'Sales region': 'US - Northeast',
                                 'Sales category': 'Consulting',
                                 'Forecast amount': 200000,
                                 'Sales phase': 'Needs analysis',
                                 'Probability of sale': 0.5,
                                 'Forecast close': datetime(2026, 12, 1),
                                 'Weighted forecast': 100000
                             }
        
        Returns:
            int: Row number where data was inserted
        """
        wb = xw.Book(self.file_path)
        
        try:
            input_sheet = wb.sheets['Forecast input']
            
            # Get headers from row 6, columns B-J
            headers = input_sheet.range('B6:J6').value
            
            # Find the last row with data in column B
            last_row = input_sheet.range('B7').end('down').row
            new_row = last_row + 1
            
            # Prepare data in the correct column order
            row_data = []
            for header in headers:
                row_data.append(data_dict.get(header, ''))
            
            # Write the new row starting from column B
            input_sheet.range(f'B{new_row}').value = row_data
            
            # Save the workbook
            wb.save()
            
            print(f"Successfully added new row at row {new_row}")
            print(f"Data: {dict(zip(headers, row_data))}")
            
            return new_row
        
        finally:
            wb.close()
    
    def read_sales_forecast(self, as_dataframe=True):
        """
        Read forecast outputs from the 'Sales forecast' tab.
        
        Args:
            as_dataframe (bool): If True, return as pandas DataFrame; if False, return as list of lists
        
        Returns:
            pd.DataFrame or list: Forecast data with columns: Month, Monthly Forecast, Cumulative
        """
        wb = xw.Book(self.file_path)
        
        try:
            forecast_sheet = wb.sheets['Sales forecast']
            
            # Get forecast data from columns B-D, starting from row 6
            # Find the last row with data in column B
            last_row = forecast_sheet.range('B7').end('down').row
            
            # Read the data including headers
            data = forecast_sheet.range(f'B6:D{last_row}').value
            
            if as_dataframe:
                # Convert to DataFrame (first row as headers)
                if data and len(data) > 1:
                    df = pd.DataFrame(data[1:], columns=data[0])
                    print(f"Read {len(df)} forecast rows from 'Sales forecast' tab")
                    return df
                else:
                    print("No data found in 'Sales forecast' tab")
                    return pd.DataFrame()
            else:
                print(f"Read {len(data)-1} forecast rows from 'Sales forecast' tab")
                return data
        
        finally:
            wb.close()
    
    def read_sales_forecast_range(self, range_address):
        """
        Read a specific range from the 'sales forecast' tab.
        
        Args:
            range_address (str): Excel range address (e.g., 'A1:D10')
        
        Returns:
            list: Data from the specified range
        """
        wb = xw.Book(self.file_path)
        
        try:
            forecast_sheet = wb.sheets['Sales forecast']
            data = forecast_sheet.range(range_address).value
            
            print(f"Read range {range_address} from 'sales forecast' tab")
            return data
        
        finally:
            wb.close()


# Example usage
if __name__ == "__main__":
    # Initialize the model
    model = SalesForecastModel("microsoft_Sales forecast tracker small business.xlsx")
    
    # Inspect the spreadsheet structure
    print("=" * 80)
    print("INSPECTING SPREADSHEET")
    print("=" * 80)
    model.inspect_spreadsheet()
    
    # Example: Add a new row to forecast input
    # Uncomment and modify the data_dict based on your actual column structure
    print("\n" + "=" * 80)
    print("ADDING NEW FORECAST INPUT ROW")
    print("=" * 80)
    new_data = {
        'Opportunity name': 'New Client Corp',
        'Sales \nagent': 'John Doe',
        'Sales \nregion': 'US - Northeast',
        'Sales \ncategory': 'Consulting',
        'Forecast amount': 200000,
        'Sales \nphase': 'Needs analysis',
        'Probability of sale': 0.5,
        'Forecast \nclose': datetime(2026, 12, 1),
        'Weighted forecast': 100000
    }
    model.add_forecast_input_row(new_data)
    
    # Example: Read sales forecast data
    print("\n" + "=" * 80)
    print("READING SALES FORECAST DATA")
    print("=" * 80)
    forecast_df = model.read_sales_forecast()
    print(forecast_df)
    print(f"\nTotal Monthly Forecast: ${forecast_df['Monthly \nforecast'].sum():,.2f}")
    print(f"Final Cumulative: ${forecast_df['Cumulative'].iloc[-1]:,.2f}")
