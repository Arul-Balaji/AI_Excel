import xlwings as xw
import pandas as pd
from datetime import datetime
import os


class ShipManagementModel:
    """
    A class to interact with the Ship Management Financial Model spreadsheet.
    Provides methods to add ship types and read revenue calculations.
    """
    
    def __init__(self, file_path="Ship Mgt Financial Model v1 - Populated Example.xlsx"):
        """
        Initialize the ShipManagementModel with the path to the Excel file.
        
        Args:
            file_path (str): Path to the Excel file
        """
        self.file_path = file_path
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"Excel file not found: {file_path}")
    
    def inspect_spreadsheet(self):
        """
        Inspect the spreadsheet structure and print information about all tabs.
        """
        print(f"Inspecting spreadsheet: {self.file_path}\n")
        
        # Open the workbook
        wb = xw.Book(self.file_path)
        
        try:
            # List all sheets
            print("Available sheets:")
            for sheet in wb.sheets:
                print(f"  - {sheet.name}")
            print()
            
            # Inspect 'i_Setup' tab
            if 'i_Setup' in [s.name for s in wb.sheets]:
                setup_sheet = wb.sheets['i_Setup']
                print("=" * 80)
                print("=== i_Setup TAB ===")
                print("=" * 80)
                
                # Read a large range to understand the structure
                print("\nFirst 30 rows and columns A-M:")
                data = setup_sheet.range('A1:M30').value
                for i, row in enumerate(data, 1):
                    print(f"Row {i:2d}: {row}")
                
                print("\n" + "-" * 80)
                print("Columns N-Z (rows 1-30):")
                data = setup_sheet.range('N1:Z30').value
                for i, row in enumerate(data, 1):
                    print(f"Row {i:2d}: {row}")
            
            # Inspect 'c_Calculations' tab
            if 'c_Calculations' in [s.name for s in wb.sheets]:
                calc_sheet = wb.sheets['c_Calculations']
                print("\n" + "=" * 80)
                print("=== c_Calculations TAB ===")
                print("=" * 80)
                
                # Read a large range to understand the structure
                print("\nFirst 30 rows and columns A-M:")
                data = calc_sheet.range('A1:M30').value
                for i, row in enumerate(data, 1):
                    print(f"Row {i:2d}: {row}")
                
                print("\n" + "-" * 80)
                print("Columns N-Z (rows 1-30):")
                data = calc_sheet.range('N1:Z30').value
                for i, row in enumerate(data, 1):
                    print(f"Row {i:2d}: {row}")
        
        finally:
            # Close the workbook without saving
            wb.close()
    
    def add_ship_type(self, ship_type_name, ship_type_slot=None):
        """
        Add a new ship type to the 'i_Setup' tab.
        
        The ship types are defined in the i_Setup tab starting at row 32.
        Ship Type 1 (ST1) is at row 32, column M
        Ship Type 2 (ST2) is at row 33, column M
        ... and so on up to Ship Type 10 (ST10) at row 41, column M
        
        Args:
            ship_type_name (str): Name of the ship type (e.g., 'VLCC')
            ship_type_slot (int): Optional slot number (1-10). If None, finds first empty slot.
        
        Returns:
            dict: Status information with slot number and message
        """
        wb = xw.Book(self.file_path)
        
        try:
            setup_sheet = wb.sheets['i_Setup']
            
            # Ship types are in rows 32-41, column M (13th column)
            # Row 32 = ST1, Row 33 = ST2, ..., Row 41 = ST10
            ship_type_start_row = 32
            ship_type_column = 'M'
            
            if ship_type_slot is not None:
                # Use specified slot
                if ship_type_slot < 1 or ship_type_slot > 10:
                    raise ValueError("Ship type slot must be between 1 and 10")
                target_row = ship_type_start_row + ship_type_slot - 1
                slot_num = ship_type_slot
            else:
                # Find first empty slot
                slot_num = None
                for i in range(10):
                    row = ship_type_start_row + i
                    current_value = setup_sheet.range(f'{ship_type_column}{row}').value
                    if current_value is None or current_value == '':
                        target_row = row
                        slot_num = i + 1
                        break
                
                if slot_num is None:
                    return {
                        'success': False,
                        'message': 'All ship type slots (ST1-ST10) are already filled'
                    }
            
            # Write the ship type name
            setup_sheet.range(f'{ship_type_column}{target_row}').value = ship_type_name
            # Set 'Yes' in the cell three columns to the right
            yes_column = chr(ord(ship_type_column) + 3)  # M + 2 = P
            setup_sheet.range(f'{yes_column}{target_row}').value = 'Yes'
            
            # Save the workbook
            wb.save()
            
            print(f"Successfully added ship type '{ship_type_name}' to slot ST{slot_num} (row {target_row})")
            
            return {
                'success': True,
                'slot': slot_num,
                'row': target_row,
                'message': f"Successfully added ship type '{ship_type_name}' to slot ST{slot_num}"
            }
        
        finally:
            wb.close()

    def add_data_in_i_assumptions_tab(self, ship_type_name):
        wb = xw.Book(self.file_path)
        
        try:
            assumptions_sheet = wb.sheets['i_Assumptions']
            
            # Populate number of ships: cells M20 to Q20 with values 1,2,3,4,5
            values = [1, 2, 3, 4, 5]
            assumptions_sheet.range('M20:Q20').value = values

            # Populate service rate per month per ship: cell M34 with value 120000
            assumptions_sheet.range('M34').value = 120000
            # Populate cells to the right with formula (previous cell * 1.05)
            for col_offset in range(1, 5):  # N, O, P, Q columns
                col_letter = chr(ord('M') + col_offset)
                prev_col_letter = chr(ord('M') + col_offset - 1)
                assumptions_sheet.range(f'{col_letter}34').formula = f'={prev_col_letter}34*1.05'
            
            
            # Populate direct capex cost for the ship: Y266:Y268 with values 200000, 100000, 50000
            cost_values = [200000, 100000, 50000]
            assumptions_sheet.range('Y266').value = cost_values[0]
            assumptions_sheet.range('Y267').value = cost_values[1]
            assumptions_sheet.range('Y268').value = cost_values[2]
            # Populate direct capex cost values for AF266:AF268 
            assumptions_sheet.range('AF266').value = cost_values[0]
            assumptions_sheet.range('AF267').value = cost_values[1]
            assumptions_sheet.range('AF268').value = cost_values[2]
            # Populate direct capex cost values for AR266:AR268 
            assumptions_sheet.range('AR266').value = cost_values[0]
            assumptions_sheet.range('AR267').value = cost_values[1]
            assumptions_sheet.range('AR268').value = cost_values[2]
            
            # Populate opex per month per ship: M275:M279 with values 1500, 2000, 500, 1000, 4000
            opex_values = [-1500, -2000, -500, -1000, -4000]
            assumptions_sheet.range('M275').value = opex_values[0]
            # Populate cells to the right with formula (previous cell * 1.03)
            for col_offset in range(1, 5):  # N, O, P, Q columns
                col_letter = chr(ord('M') + col_offset)
                prev_col_letter = chr(ord('M') + col_offset - 1)
                assumptions_sheet.range(f'{col_letter}{275 + 0}').formula = f'={prev_col_letter}{275 + 0}*1.03'
            
            assumptions_sheet.range('M276').value = opex_values[1]
            # Populate cells to the right with formula (previous cell * 1.03)
            for col_offset in range(1, 5):  # N, O, P, Q columns
                col_letter = chr(ord('M') + col_offset)
                prev_col_letter = chr(ord('M') + col_offset - 1)
                assumptions_sheet.range(f'{col_letter}{275 + 1}').formula = f'={prev_col_letter}{275 + 1}*1.03'
            
            assumptions_sheet.range('M277').value = opex_values[2]
            # Populate cells to the right with formula (previous cell * 1.03)
            for col_offset in range(1, 5):  # N, O, P, Q columns
                col_letter = chr(ord('M') + col_offset)
                prev_col_letter = chr(ord('M') + col_offset - 1)
                assumptions_sheet.range(f'{col_letter}{275 + 2}').formula = f'={prev_col_letter}{275 + 2}*1.03'
            
            assumptions_sheet.range('M278').value = opex_values[3]
            # Populate cells to the right with formula (previous cell * 1.03)
            for col_offset in range(1, 5):  # N, O, P, Q columns
                col_letter = chr(ord('M') + col_offset)
                prev_col_letter = chr(ord('M') + col_offset - 1)
                assumptions_sheet.range(f'{col_letter}{275 + 3}').formula = f'={prev_col_letter}{275 + 3}*1.03'
            
            assumptions_sheet.range('M279').value = opex_values[4]
            # Populate cells to the right with formula (previous cell * 1.03)
            for col_offset in range(1, 5):  # N, O, P, Q columns
                col_letter = chr(ord('M') + col_offset)
                prev_col_letter = chr(ord('M') + col_offset - 1)
                assumptions_sheet.range(f'{col_letter}{275 + 4}').formula = f'={prev_col_letter}{275 + 4}*1.03'
            
            # Populate direct staff numbers per ship: M287:M292 with values 1.0, 2.0, 2.0, 5.0, 3.0, 2.0
            staff_values = [1.0, 2.0, 2.0, 5.0, 3.0, 2.0]
            # Populate direct staff numbers per ship: M287:M292 with values distributed across columns M-Q
            for i, value in enumerate(staff_values):
                row = 287 + i
                assumptions_sheet.range(f'M{row}:Q{row}').value = [value] * 5
            
            # Save the workbook
            wb.save()
    
            return {
                'success': True,
                'message': "Successfully populated number of ships, service rate per month per ship, direct cost, opex, and direct staff numbers in i_Assumptions tab"
            }

        finally:
            wb.close()
        
    def read_total_revenue(self, ship_type_name, include_tax=False):
        """
        Read the total revenue for a specific ship type from the 'c_Calculations' tab.
        
        The revenue data is located in the c_Calculations tab:
        - Total Revenue (Excl Tax) for each ship type: rows 44-53, column M
          ST1 at row 44, ST2 at row 45, ..., ST10 at row 53
        - Total Revenue (Incl Tax) for each ship type: rows 61-70, column M
          ST1 at row 61, ST2 at row 62, ..., ST10 at row 70
        
        Args:
            ship_type_name (str): Name of the ship type (e.g., 'VLCC', 'Container Ships')
            include_tax (bool): If True, return revenue including sales tax; if False, exclude tax
        
        Returns:
            dict: Revenue information including amount, ship type slot, and whether tax is included
        """
        wb = xw.Book(self.file_path)
        
        try:
            setup_sheet = wb.sheets['i_Setup']
            calc_sheet = wb.sheets['c_Calculations']
            
            # First, find which slot (ST1-ST10) this ship type is in
            ship_type_start_row = 32
            ship_type_column = 'M'
            slot_num = None
            
            for i in range(10):
                row = ship_type_start_row + i
                current_value = setup_sheet.range(f'{ship_type_column}{row}').value
                if current_value and current_value.strip().upper() == ship_type_name.strip().upper():
                    slot_num = i + 1
                    break
            
            if slot_num is None:
                return {
                    'success': False,
                    'message': f"Ship type '{ship_type_name}' not found in i_Setup tab"
                }
            
            # Now read the revenue from c_Calculations
            if include_tax:
                # Revenue including tax: rows 61-70
                revenue_row = 61 + (slot_num - 1)
            else:
                # Revenue excluding tax: rows 44-53
                revenue_row = 44 + (slot_num - 1)
            
            revenue_column = 'M'
            revenue_value = calc_sheet.range(f'{revenue_column}{revenue_row}').value
            
            # Also get the label to confirm
            label_column = 'H'
            label = calc_sheet.range(f'{label_column}{revenue_row}').value
            
            print(f"Found ship type '{ship_type_name}' in slot ST{slot_num}")
            print(f"  Revenue ({'incl. tax' if include_tax else 'excl. tax'}): ${revenue_value:,.2f}")
            
            return {
                'success': True,
                'ship_type': ship_type_name,
                'slot': slot_num,
                'revenue': revenue_value,
                'include_tax': include_tax,
                'label': label,
                'message': f"Total revenue for '{ship_type_name}': ${revenue_value:,.2f} ({'incl. tax' if include_tax else 'excl. tax'})"
            }
        
        finally:
            wb.close()
    
    def get_all_ship_types(self):
        """
        Get all defined ship types from the i_Setup tab.
        
        Returns:
            list: List of dictionaries with ship type information
        """
        wb = xw.Book(self.file_path)
        
        try:
            setup_sheet = wb.sheets['i_Setup']
            
            ship_types = []
            ship_type_start_row = 32
            ship_type_column = 'M'
            
            for i in range(10):
                row = ship_type_start_row + i
                ship_name = setup_sheet.range(f'{ship_type_column}{row}').value
                
                if ship_name and ship_name.strip():
                    ship_types.append({
                        'slot': i + 1,
                        'name': ship_name,
                        'row': row
                    })
            
            return ship_types
        
        finally:
            wb.close()


# Main script
if __name__ == "__main__":
    # Initialize the model
    model = ShipManagementModel("Ship Mgt Financial Model v1 - Populated Example.xlsx")
    
    print("=" * 80)
    print("SHIP MANAGEMENT FINANCIAL MODEL - DEMO")
    print("=" * 80)
    
    # Show current ship types
    print("\n1. Current Ship Types:")
    print("-" * 80)
    current_ships = model.get_all_ship_types()
    for ship in current_ships:
        print(f"  ST{ship['slot']}: {ship['name']}")
    
    # Add VLCC ship type
    print("\n2. Adding VLCC Ship Type:")
    print("-" * 80)
    result = model.add_ship_type("VLCC")
    if result['success']:
        print(f"  {result['message']}")
    else:
        print(f"  Error: {result['message']}")

    # Add VLCC ship type
    print("\n3. Adding revenue and costs for VLCC Ship Type:")
    print("-" * 80)
    result = model.add_data_in_i_assumptions_tab("VLCC")
    if result['success']:
        print(f"  {result['message']}")
    else:
        print(f"  Error: {result['message']}")
    
    # Read total revenue for VLCC
    print("\n4. Reading Total Revenue for VLCC:")
    print("-" * 80)
    
    # Try to read revenue (excluding tax)
    revenue_result = model.read_total_revenue("VLCC", include_tax=False)
    if revenue_result['success']:
        print(f"  {revenue_result['message']}")
    else:
        print(f"  {revenue_result['message']}")
        print("  Note: Revenue will be calculated once ship parameters are configured in the model")
    
    # Also try with tax included
    revenue_result_tax = model.read_total_revenue("VLCC", include_tax=True)
    if revenue_result_tax['success']:
        print(f"  Revenue (incl. tax): ${revenue_result_tax['revenue']:,.2f}")
    
    print("\n" + "=" * 80)
    print("DEMO COMPLETE")
    print("=" * 80)
