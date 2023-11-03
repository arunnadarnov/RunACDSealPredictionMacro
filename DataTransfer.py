import json
import os
import xlwings as xw
from Modules.CsvDataTransfer import CsvDataTransfer
from Modules.ExcelDataTransfer import ExcelDataTransfer
from Modules.ExcelColumnFinder import ExcelColumnFinder
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
        for config in self.configs:
            # Download the source files from S3
            #downloader = S3Downloader(config['s3_bucket'], config['input_files_folder'])
            #downloader.download_files()

            # Get the file extension
            _, file_extension = os.path.splitext(config['source_file'])
            file_extension = file_extension.lower()

            # Check the file type and transfer data
            if file_extension == '.xlsx':
                self.excel_transfer.transfer_data(config)
                wb = xw.Book(config['destination_file'])
                finder = ExcelColumnFinder(wb, config['columns_to_find'])
                finder.find_columns()
            elif file_extension == '.csv':
                self.csv_transfer.transfer_data(config)
                wb = xw.Book(config['destination_file'])
                finder = ExcelColumnFinder(wb, config['columns_to_find'])
                finder.find_columns()
            else:
                print(f"Unsupported file type '{file_extension}' in the config file.")

# Config file
config_file = r"C:\Arun\scripts\python\RunACDSealPredictionMacro\Application\ConfigFiles\global_configuration.json"

# Create an instance of class DataTransfer and call function transfer_data
data_transfer = DataTransfer(config_file)
data_transfer.transfer_data()
