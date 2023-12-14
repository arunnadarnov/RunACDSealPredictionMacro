class DynamoDBClient:
    def __init__(self, table_name, dynamodb_resource):
        self.dynamodb_resource = dynamodb_resource
        self.table = self.dynamodb_resource.Table(table_name)

    def get_single_item(self, item_key):
        response = self.table.get_item(Key=item_key)
        return response

