# Data Transfer and Macro Execution Application

This application transfers data based on the file type specified in a configuration file. It supports both Excel (.xlsx) and CSV (.csv) files. An upcoming feature will allow the application to execute macros in Excel files.

## Folder Structure

The application has the following folder structure:

```Application
ConfigFiles\global_configuration.json
Modules\
    ConfigValidator.py
    CsvDataTransfer.py
    ExcelColumnFinder.py
    ExcelDataTransfer.py
    S3Downloader.py
DataTransfer.py
```
## Modules
The application consists of several modules:

`DataTransfer`: This is the main module that controls data transfer based on file type. It uses the other modules to perform specific tasks based on the configuration file.

`ValidateJsonConfigFile`: This module validates a JSON configuration file for the data transfer script. It checks for required keys and validates that specified files exist.

`CsvDataTransfer`: This module handles data transfer from CSV files. It reads data from a source CSV file and writes it to a specified sheet in a destination Excel file.

`ExcelColumnFinder`: This module finds a specified column in an Excel file and writes the cell details to another sheet. It’s used to locate and record where specific data is located in your spreadsheets.

`ExcelDataTransfer`: This module handles data transfer from Excel files. It reads data from a source Excel file and writes it to a specified sheet in a destination Excel file.

`S3Downloader`: This module downloads files from an S3 bucket. It’s used to retrieve your source files before data transfer.

Each of these modules plays a crucial role in the application, and they work together to transfer data according to your specifications.

## Configuration File

The application uses a JSON configuration file to specify details about the data transfer. Here's an example of what the configuration file might look like:

```json
[
    {
        "s3_bucket": "<S3 bucket name>",
        "input_files_folder": "<Path to input files folder>",
        "source_file": "<Path to source file>",
        "destination_file": "<Path to destination file>",
        "destination_sheet": "<Name of destination sheet in destination file>",
        "columns_to_find": [
            {
                "column_name": "<Name of column to find>",
                "cell_to_update": "<Cell in which to write column details>"
            }
        ]
    }
]
```
Each object in the array represents a data transfer task. The columns_to_find array specifies which columns to find in the source file and where to write their details in the destination file.

## How to Use

Specify the source and destination files, and other configuration details, in a JSON config file.
Create an instance of the DataTransfer class with the config file as an argument.
Call the transfer_data method on the DataTransfer instance.
Example:
```python
# Config file
config_file = "<Path to your config file>"

# Create an instance of class DataTransfer and call function transfer_data
data_transfer = DataTransfer(config_file)
data_transfer.transfer_data()
```

## Upcoming Feature
The application will soon support executing macros in Excel files.

## Requirements
This application requires the following Python libraries:

json
os
pandas
xlwings
boto3
You can install these libraries using pip:
```
pip install pandas xlwings boto3
```
