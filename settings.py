from pathlib import Path
from dotenv import load_dotenv
import os
import boto3

class Settings:
    def __init__(self):
        load_dotenv()
        self.aws_access_key_id = os.getenv('AWS_ACCESS_KEY_ID')
        self.aws_secret_access_key = os.getenv('AWS_SECRET_ACCESS_KEY')
        self.aws_session_token = os.getenv('AWS_SESSION_TOKEN')
        self.aws_default_region = os.getenv('AWS_DEFAULT_REGION', 'us-east-1')
        self.bedrock_model_id = os.getenv('BEDROCK_MODEL_ID', 'anthropic.claude-3-sonnet-20240229-v1:0')
        self.bedrock_region = os.getenv('BEDROCK_REGION', 'us-east-1')
        self.upload_folder = Path(os.getenv('UPLOAD_FOLDER', 'uploads'))
        self.output_folder = Path(os.getenv('OUTPUT_FOLDER', 'output'))
        self.max_file_size = int(os.getenv('MAX_FILE_SIZE', 10485760))  # 10MB
        self.allowed_extensions = os.getenv('ALLOWED_EXTENSIONS', 'pdf').split(',')
        self.default_source_language = os.getenv('DEFAULT_SOURCE_LANGUAGE', 'auto')
        self.default_target_language = os.getenv('DEFAULT_TARGET_LANGUAGE', 'en')
        self.supported_languages = os.getenv('SUPPORTED_LANGUAGES', 'en,es,fr,de,it,pt,nl,ru,zh,ja,ko,ar,hi').split(',')
        self.template_path = Path(os.getenv('TEMPLATE_PATH', 'templates/monthly_template.pptx'))
        self.domain_focus = ['Logistics Vehicles', 'Cargo Management']

        # Create directories
        self.upload_folder.mkdir(exist_ok=True)
        self.output_folder.mkdir(exist_ok=True)

        # Set up AWS session
        try:
            boto3.setup_default_session(
                aws_access_key_id=self.aws_access_key_id,
                aws_secret_access_key=self.aws_secret_access_key,
                aws_session_token=self.aws_session_token,
                region_name=self.aws_default_region
            )
        except Exception as e:
            print(f"AWS setup error: {str(e)}")