import os
import pandas as pd
import json
import xlwings as xw

class CsvDataTransfer:
    """Handles data transfer from CSV files."""

    def transfer_data(self, config):
        """
        Transfers data from a source CSV file to a destination file.
        """

        # Source File
        source_file = config['source_file']
        print (f"Source File - {source_file}")

        # Load CSV file into a DataFrame
        df1 = pd.read_csv(config['source_file'], index_col='Timestamp')

        # Convert 'Timestamp' column to string format
        #df1['Timestamp'] = df1['Timestamp'].astype(str)

        # Check if there is any data in the source CSV file
        if df1.empty:
            print(f"The file '{config['source_file']}' has no data.")
            return

        # Check if the destination file exists
        if os.path.exists(config['destination_file']):
            destination_wb = xw.Book(config['destination_file'])
        else:
            print(f"Error: The destination file '{config['destination_file']}' does not exist.")
            return

        # Check if the sheet exists in the destination file
        if config['destination_sheet'] in [sheet.name for sheet in destination_wb.sheets]:
            destination_sheet = destination_wb.sheets[config['destination_sheet']]
        else:
            destination_sheet = destination_wb.sheets.add(config['destination_sheet'])

        # Write the dataframe object into excel file
        destination_sheet.range('A1').value = df1

        print(f"Data has been written to '{config['destination_file']}'")

        # Save and close the workbook
        destination_wb.save()
        destination_wb.close()