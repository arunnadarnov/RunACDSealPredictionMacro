import pandas as pd
import xlwings as xw

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

        # Search for each column name
        for column_info in self.columns_to_find:
            column_name = column_info['column_name']
            if column_name in df1.columns:
                # Get the cell details (row number and column number)
                row_num = df1.columns.get_loc(column_name) + 2  # +2 because pandas' index is 0-based and Excel's index is 1-based, and the header is in the second row
                sheet2.range(column_info['cell_to_update']).value = row_num
            else:
                print(f"Column '{column_name}' not found.")