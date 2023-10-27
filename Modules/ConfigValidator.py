import os
import pandas as pd

class ValidateJsonConfigFile:
    """Validates a JSON config file for a data transfer script."""

    def __init__(self, json_config):
        """Initializes the validator with a JSON config file."""
        self.json_config = json_config
        self.error_messages = []

    def validate_keys(self, config):
        """Validates that all required keys are present in the config."""
        required_keys = ["source_file", "destination_file", "destination_sheet"]
        for key in required_keys:
            if key not in config:
                self.error_messages.append(f"Error: Missing key: {key}")

    def validate_file_exists(self, file_path):
        """Validates that a file exists at the given path."""
        if not os.path.isfile(file_path):
            self.error_messages.append(f"Error: {file_path} is not a valid file.")

    def validate(self):
        """Validates the JSON config file."""
        for config in self.json_config:
            self.validate_keys(config)
            self.validate_file_exists(config["source_file"])
            self.validate_file_exists(config["destination_file"])

        # If there are any error messages, print them and exit
        if self.error_messages:
            print("\n".join(self.error_messages))
            exit(1)