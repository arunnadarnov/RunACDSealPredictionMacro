import boto3
import json
import os
import glob
import pywintypes
import xlwings as xw
from Modules.CsvDataTransfer import CsvDataTransfer
from Modules.DynamodbClient import DynamoDBClient
from Modules.ExcelDataTransfer import ExcelDataTransfer
from Modules.ExcelColumnFinder import ExcelColumnFinder
from Modules.ExcelCellUpdater import ExcelCellUpdater
from Modules.JsonDataToDeltaLake import TransferJsonDataToDatabricks
from Modules.ConfigValidator import ValidateJsonConfigFile
from Modules.S3Downloader import S3Downloader

class DataTransfer:
    """Controls data transfer based on file type."""

    def __init__(self, config_file):
        """
        Initializes the DataTransfer class with a config file.
        """
        with open(config_file, 'r') as f:
            self.configs = json.load(f)

        # Validate the config file
        validator = ValidateJsonConfigFile(self.configs)
        validator.validate()

        self.excel_transfer = ExcelDataTransfer()
        self.csv_transfer = CsvDataTransfer()

    def transfer_data(self):
        """
        Transfers data based on the file type specified in the config file.
        """
        first_run = True
        wb = None
        for config in self.configs:
            # Get all the files in the input_files_folder
            files = glob.glob(os.path.join(config['input_files_folder'], '*'))

            # Check if there are any files in the folder
            if not files:
                print(f"Error: No files found in folder '{config['input_files_folder']}'.")
                return

            # Check the number of CSV and Excel files
            csv_files = [file for file in files if file.lower().endswith('.csv')]
            excel_files = [file for file in files if file.lower().endswith('.xlsx')]
            if len(csv_files) > 1:
                print(f"Error: More than one CSV file found in folder '{config['input_files_folder']}'.")
                return
            if len(excel_files) > 1:
                print(f"Error: More than one Excel file found in folder '{config['input_files_folder']}'.")
                return
            if len(csv_files) == 1 and len(excel_files) == 1:
                print(f"Error: The folder '{config['input_files_folder']}' contains both a CSV file and an Excel file. It should contain either a single CSV file or a single Excel file.")
                return

            for file in files:
                # Get the file extension
                _, file_extension = os.path.splitext(file)
                file_extension = file_extension.lower()

                # Check the file type and transfer data
                if file_extension == '.xlsx':
                    config['source_file'] = file
                    if first_run:
                        wb = xw.Book(config['destination_file'])
                        if config['destination_sheet'] in [sheet.name for sheet in wb.sheets]:
                            wb.sheets[config['destination_sheet']].clear_contents()
                        first_run = False
                    self.excel_transfer.transfer_data(config)
                elif file_extension == '.csv':
                    config['source_file'] = file
                    if first_run:
                        wb = xw.Book(config['destination_file'])
                        if config['destination_sheet'] in [sheet.name for sheet in wb.sheets]:
                            wb.sheets[config['destination_sheet']].clear_contents()
                        first_run = False
                    self.csv_transfer.transfer_data(config)
                    
                else:
                    print(f"Unsupported file type '{file_extension}' in folder '{config['input_files_folder']}'.")

            # Get the data from DynamoDB
            table_name = "app04d2-dev-acd-predict-well-configs"  # replace with your actual table name
            well_name = "test_well"  # replace with your actual well name

            # Create the session and dynamodb_resource
            session = boto3.Session(profile_name='226268956475_ApplicationAdministratorAccess')
            dynamodb_resource = session.resource('dynamodb', region_name='us-east-1')
            dynamodb_client = DynamoDBClient(table_name, dynamodb_resource)
            item_key = {"well_name": well_name}
            response = dynamodb_client.get_single_item(item_key=item_key)
            item = response["Item"]

            finder = ExcelColumnFinder(wb, config['columns_to_find'], dynamodb_client, well_name)
            finder.find_columns()

            # Pass the data from DynamoDB to the ExcelCellUpdater
            updater = ExcelCellUpdater(wb, config['cells_to_update'], item)
            updater.update_cells()

            try:
                # Run the specified macro
                if 'macro_to_run' in config:
                    macro_name = config['macro_to_run']
                    macro = wb.app.macro(macro_name)
                    print (f"Executing Macro - {macro_name}")
                    macro()
            except pywintypes.com_error as e:
                print(f"An error occurred: {e}")

            # Save the workbook
            wb.save()

        # value extractor config file
        value_extractor_config_file = r"C:\Arun\scripts\python\RunACDSealPredictionMacro\Application\ConfigFiles\field_to_cell_mapping.json"

        databricks_instance = "dbc-eebd97f8-a4cd.cloud.databricks.com"
        access_token = "dapiaf2c4555394b2427f3095d6a92513a8a"
        sql_warehouse_id = "87aca86eba59ce73"
        table_name = "acd_seal_prediction_results"

        # Create an instance of class TransferJsonDataToDatabricks and call function run
        transfer = TransferJsonDataToDatabricks(wb, value_extractor_config_file, databricks_instance, access_token, sql_warehouse_id, table_name)
        transfer.run()

        # Close the destination workbook after all files have been processed
        wb.close()

# Config file
config_file = r"C:\Arun\scripts\python\RunACDSealPredictionMacro\Application\ConfigFiles\global_configuration.json"

# Create an instance of class DataTransfer and call function transfer_data
data_transfer = DataTransfer(config_file)
data_transfer.transfer_data()
