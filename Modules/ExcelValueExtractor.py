import xlwings as xw
import json

class ExcelValueExtractor:
    """Extracts specified values from an Excel file."""

    def __init__(self, workbook, config_file):
        """
        Initializes the ExcelValueExtractor class with a workbook and a config file.
        """
        self.wb = workbook
        with open(config_file, 'r') as f:
            self.configs = json.load(f)

    def extract_acd_values(self):
        """
        Extracts and prints the ACD usage statistics specified in the config file.
        """
        print("ACD Usage Statistics")
        for field in self.configs['acd']:
            self._extract_and_print_field(field)

    def extract_dsit_values(self):
        """
        Extracts and prints the DSIT usage statistics specified in the config file.
        """
        print("\nDSIT Usage Statistics")
        for field in self.configs['dsit']:
            self._extract_and_print_field(field)

    def _extract_and_print_field(self, field):
        """
        Extracts and prints a single field.
        """
        field_name = field['field_name']
        
        if 'pivot_table' in field:
            print(f"\n{field_name}: ", end="")
            for key, cell in field['pivot_table'].items():
                value = self.wb.sheets[0].range(cell).value
                print(f"{key}: {value}\t", end="")
            print()
        else:
            data_cell = field['data_cell']
            uom_cell = field['uom_cell']

            # Get the value and unit of measure
            value = self.wb.sheets[0].range(data_cell).value
            uom = self.wb.sheets[0].range(uom_cell).value if uom_cell else ''

            print(f"{field_name}: {value} {uom}")