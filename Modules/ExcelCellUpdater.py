class ExcelCellUpdater:
    """Updates cells in an Excel file with specified values."""

    def __init__(self, workbook, cells_to_update, item):
        """
        Initializes the ExcelCellUpdater class with a workbook, a list of cells to update, and the data from DynamoDB.
        """
        self.workbook = workbook
        self.cells_to_update = cells_to_update
        self.item = item  # data from DynamoDB
        
    def update_cells(self):
        """
        Updates the cells in the workbook with the specified values.
        """
        # Write the values to the specified cells
        if 'Vars' in [sheet.name for sheet in self.workbook.sheets]:
            sheet = self.workbook.sheets['Vars']  # Use the correct case here
        else:
            raise ValueError("Sheet named 'Vars' not found in workbook")

        # Update each cell with the specified value
        for cell_info in self.cells_to_update:
            cell_to_update = cell_info['cell_to_update']
            dynamodb_key = cell_info['dynamodb_key']
            value_to_update = self.item.get(dynamodb_key)  # get the value from the DynamoDB data
            sheet.range(cell_to_update).value = value_to_update

        # Save the workbook after updating the cells
        self.workbook.save()
