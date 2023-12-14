import os
import logging
import requests
import json
import xlwings as xw
from .ExcelValueExtractor import ExcelValueExtractor

# Get the directory of the current script
script_dir = os.path.dirname(os.path.abspath(__file__))

# Create a logs directory inside the script directory
logs_dir = os.path.join(script_dir, 'logs')
os.makedirs(logs_dir, exist_ok=True)

# Create a log file inside the logs directory
log_file = os.path.join(logs_dir, 'databricks_output.log')

# Configure the logger
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s %(message)s')

class DatabricksAPI:
    def __init__(self, instance, token, warehouse_id):
        self.url = f"https://{instance}/api/2.0/sql/statements"
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json"
        }
        self.warehouse_id = warehouse_id

    def insert_data(self, statement):
        data = {
            "statement": statement,
            "warehouse_id": self.warehouse_id
        }
        response = requests.post(self.url, headers=self.headers, data=json.dumps(data), verify=False)
        logging.info(f"Inserted data into table with statement: {statement}")
        return response.json()

    def insert_json_to_table(self, json_object, table_name):
        for section, records in json_object.items():
            for record in records:
                # Convert Python None to SQL NULL
                window_start = 'NULL' if record['window_start'] is None else f"'{record['window_start']}'"
                window_end = 'NULL' if record['window_end'] is None else f"'{record['window_end']}'"
                field_name = 'NULL' if record['field_name'] is None else f"'{record['field_name']}'"
                sub_field = 'NULL' if record['sub_field'] is None else f"'{record['sub_field']}'"
                value = 'NULL' if record['value'] is None else f"'{record['value']}'"
                unit_of_measurement = 'NULL' if record['unit_of_measurement'] is None else f"'{record['unit_of_measurement']}'"
                section = 'NULL' if record['section'] is None else f"'{record['section']}'"

                statement = f"INSERT INTO {table_name} VALUES ({window_start}, {window_end}, {field_name}, {sub_field}, {value}, {unit_of_measurement}, {section})"
                self.insert_data(statement)


class TransferJsonDataToDatabricks:
    def __init__(self, workbook, config_file_path, databricks_instance, access_token, sql_warehouse_id, table_name):
        self.wb = workbook
        self.config_file_path = config_file_path
        self.databricks_instance = databricks_instance
        self.access_token = access_token
        self.sql_warehouse_id = sql_warehouse_id
        self.table_name = table_name

    def run(self):
        
        # Workbook
        wb = self.wb

        # Create an instance of class ExcelValueExtractor
        value_extractor = ExcelValueExtractor(wb, self.config_file_path)

        # Call function _extract_and_create_json
        json_object = value_extractor.extract_and_create_json()

        # Create an instance of class DatabricksAPI
        db = DatabricksAPI(self.databricks_instance, self.access_token, self.sql_warehouse_id)

        # Call function insert_json_to_table
        db.insert_json_to_table(json_object, self.table_name)

# Close the logger
logging.shutdown()

excel_file_path = r"C:\Arun\scripts\python\TransferJsonDataToDatabricks\InputFiles\DAT SET4 2023A.xlsm"
config_file_path = r"C:\Arun\scripts\python\TransferJsonDataToDatabricks\Application\ConfigFiles\field_to_cell_mapping.json"
databricks_instance = "dbc-eebd97f8-a4cd.cloud.databricks.com"
access_token = "dapiaf2c4555394b2427f3095d6a92513a8a"
sql_warehouse_id = "87aca86eba59ce73"
table_name = "acd_seal_prediction_results"

# Usage:
#transfer = TransferJsonDataToDatabricks(excel_file_path, config_file_path, databricks_instance, access_token, sql_warehouse_id, table_name)
#json_object = transfer.run()
