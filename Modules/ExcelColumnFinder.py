import pandas as pd
import xlwings as xw
import re
import sys

class ExcelColumnFinder:
    """Finds a column in an Excel file and writes the cell details to another sheet."""

    def __init__(self, workbook, columns_to_find, dynamodb_client, well_name):
        """
        Initializes the ExcelColumnFinder class with a workbook, a list of columns to find, a DynamoDB client, and a well name.
        """
        self.workbook = workbook
        self.columns_to_find = columns_to_find
        self.dynamodb_client = dynamodb_client
        self.well_name = well_name

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

        # Fetch the item from DynamoDB using the well_name
        item_key = {'well_name': self.well_name}
        item = self.dynamodb_client.get_single_item(item_key)

        #print(item)
        # Search for each column name
        for column_info in self.columns_to_find:
            column_name = None
            unit_of_measurement = None
            if 'attribute_name' in column_info:
                print(f"Looking for attribute: {column_info['attribute_name']}")
                #print(f"Available attributes: {item['Item'].keys()}")
            if column_info['column_name_key'] == 'placeholder':
                # This is a special case where we have a value_to_update key
                # We just write this value to the specified cell
                sheet2.range(column_info['cell_to_update']).value = column_info['value_to_update']
            else:
                attribute = item['Item'].get(column_info['attribute_name'], {}).get('M')
                if attribute is not None:
                    column_name = attribute.get(column_info['column_name_key'].split('.')[0], {}).get(column_info['column_name_key'].split('.')[1])
                    unit_of_measurement = attribute.get(column_info['unit_of_measurement_key'].split('.')[0], {}).get(column_info['unit_of_measurement_key'].split('.')[1])

            # Get full column name from Excel sheet (case insensitive)
            full_column_name = next((col for col in df1.columns if col.lower().startswith(column_name.lower())), None) if column_name else None

            if full_column_name:
                # Get the cell details (row number and column number)
                row_num = df1.columns.get_loc(full_column_name) + 2  # +2 because pandas' index is 0-based and Excel's index is 1-based, and the header is in the second row
                sheet2.range(column_info['cell_to_update']).value = row_num
                sheet2.range(column_info['unit_cell_to_update']).value = unit_of_measurement
            else:
                print(f"Column '{column_name}' not found.")
        
        # Save the workbook after writing the cell details
        self.workbook.save()
