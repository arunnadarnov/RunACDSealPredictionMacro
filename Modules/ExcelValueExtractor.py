import json
import xlwings as xw
from datetime import datetime

class ExcelValueExtractor:
    """Extracts specified values from an Excel file."""

    def __init__(self, workbook, config_file):
        """
        Initializes the ExcelValueExtractor class with a workbook and a config file.
        """
        self.wb = workbook
        self.sheet = self.wb.sheets['Plotting']  # Access the 'Plotting' sheet
        with open(config_file, 'r') as f:
            self.configs = json.load(f)

    def extract_and_create_json(self):
        """
        Extracts and creates a JSON object for all fields.
        """
        json_object = {}

        for section in self.configs:
            json_object[section] = []
            for field in self.configs[section]:
                field_name = field['field_name']
                data_cell = field.get('data_cell')  # Use the get method to avoid KeyError
                uom_cell = field.get('uom_cell')  # Use the get method to avoid KeyError

                # Get the value and unit of measure
                value = self.sheet.range(data_cell).value if data_cell else None
                if isinstance(value, datetime):
                    value = value.strftime('%Y-%m-%d %H:%M:%S')  # Convert datetime to string
                uom = self.sheet.range(uom_cell).value if uom_cell and self.sheet.range(uom_cell).value != '' else None  # Assign None instead of ''

                # Check if the field is 'Window Start' or 'Window End'
                if field_name == 'Window Start':
                    window_start = value
                elif field_name == 'Window End':
                    window_end = value
                else:
                    if 'pivot_table' in field:
                        for key, cell in field['pivot_table'].items():
                            sub_field_value = self.sheet.range(cell).value
                            field_object = {
                                "window_start": window_start,
                                "window_end": window_end,
                                "field_name": field_name,
                                "sub_field": key,
                                "value": sub_field_value,
                                "unit_of_measurement": uom,
                                "section": section
                            }
                            json_object[section].append(field_object)
                    else:
                        field_object = {
                            "window_start": window_start,
                            "window_end": window_end,
                            "field_name": field_name,
                            "sub_field": None,
                            "value": value,
                            "unit_of_measurement": uom,
                            "section": section
                        }
                        json_object[section].append(field_object)

        # Print the JSON object in a pretty format
        #print(json.dumps(json_object, indent=4))

        return json_object