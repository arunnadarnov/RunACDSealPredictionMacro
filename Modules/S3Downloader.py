import boto3
import os

class S3Downloader:
    """Downloads files from an S3 bucket."""

    def __init__(self, bucket_name, local_folder):
        """
        Initializes the S3Downloader class with a bucket name and a local folder.
        """
        self.bucket_name = bucket_name
        self.local_folder = local_folder

    def download_files(self):
        """
        Downloads all files from the S3 bucket to the local folder.
        """
        s3 = boto3.client('s3')

        # List all files in the S3 bucket
        files = s3.list_objects(Bucket=self.bucket_name)['Contents']

        # Download each file
        for file in files:
            s3_file = file['Key']
            local_file = os.path.join(self.local_folder, os.path.basename(s3_file))
            s3.download_file(self.bucket_name, s3_file, local_file)
