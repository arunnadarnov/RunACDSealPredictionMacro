import pandas as pd
import xlwings as xw
import re
import sys

class ExcelColumnFinder:
    """Finds a column in an Excel file and writes the cell details to another sheet."""

    def __init__(self, workbook, columns_to_find):
        """
        Initializes the ExcelColumnFinder class with a workbook and a list of columns to find.
        """
        self.workbook = workbook
        self.columns_to_find = columns_to_find

    def find_columns(self):
        """
        Finds the columns in the workbook and writes the cell details to another sheet.
        """
        # Load a sheet into a DataFrame by name
        sheet1 = self.workbook.sheets['Data']
        df1 = sheet1.range('A1').options(pd.DataFrame, expand='table').value

        # Write the cell details to another sheet
        if 'vars' in [sheet.name.lower() for sheet in self.workbook.sheets]:
            sheet2 = self.workbook.sheets['Vars']  # Use the correct case here
        else:
            raise ValueError("Sheet named 'Vars' not found in workbook")

        # List to store columns without unit of measurement
        columns_without_unit = []

        # Search for each column name
        for column_info in self.columns_to_find:
            column_name = column_info['column_name']
            #print (f"Column Name - {column_name}")
            
            # Get full column name from Excel sheet (case insensitive)
            full_column_name = next((col for col in df1.columns if col.lower().startswith(column_name.lower())), None)
            
            if full_column_name:
                # Get the cell details (row number and column number)
                row_num = df1.columns.get_loc(full_column_name) + 2  # +2 because pandas' index is 0-based and Excel's index is 1-based, and the header is in the second row
                sheet2.range(column_info['cell_to_update']).value = row_num
                
                # Extract unit of measurement from column name
                match = re.search(r'\((.*?)\)', full_column_name)
                if match:
                    unit_of_measurement = match.group(1)
                    print(f"Unit of measurement for '{full_column_name}' is '{unit_of_measurement}'")
                    sheet2.range(column_info['unit_cell_to_update']).value = unit_of_measurement
                else:
                    columns_without_unit.append(full_column_name)
            else:
                print(f"Column '{column_name}' not found.")
        
        # Check if there are any columns without unit of measurement
        if columns_without_unit:
            print(f"Unit of measurement not found for the following columns: {', '.join(columns_without_unit)}. Exiting the program.")
            sys.exit()
        
        # Save the workbook after writing the cell details
        self.workbook.save()