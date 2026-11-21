"""
Configuration for Voice-Controlled Excel Assistant
"""
import os
from dotenv import load_dotenv

load_dotenv()

class Config:
    # OpenAI API Key (optional - can work without it using rule-based NLP)
    OPENAI_API_KEY = os.getenv('OPENAI_API_KEY', '')
    
    # Application Settings
    SECRET_KEY = os.getenv('SECRET_KEY')
    DEBUG = os.getenv('DEBUG', 'True').lower() == 'true'
    
    # Upload Settings
    UPLOAD_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'uploads')
    OUTPUT_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024  # 16MB max file size
    
    # Allowed Extensions
    ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
    
    @classmethod
    def init_folders(cls):
        """Create necessary folders if they don't exist"""
        os.makedirs(cls.UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(cls.OUTPUT_FOLDER, exist_ok=True)

