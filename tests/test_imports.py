"""
Test script to verify all imports work correctly
"""

import sys
import os
from pathlib import Path
from dotenv import load_dotenv

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

dotenv_path = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path)

from src.colours import GREEN, RED, RESET

print("Testing imports...")

try:
    # Built-in modules
    import imaplib
    import email
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import smtplib
    import re
    from datetime import datetime, timedelta
    import os
    import json
    import sys
    print(f"{GREEN}Built-in modules imported successfully{RESET}")
    
    # External modules
    from googleapiclient.discovery import build
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    import requests
    print(f"{GREEN}Google API modules imported successfully{RESET}")
    print(f"{GREEN}Requests module imported successfully{RESET}")
    
    # Environment variable handling
    print(f"{GREEN}Python-dotenv imported successfully{RESET}")
    
    # Test environment variable loading
    if os.getenv('GMAIL_USER'):
        print(f"{GREEN}Environment variables loaded successfully{RESET}")
    else:
        print(f"{RED}No .env file found or GMAIL_USER not set (this is okay for testing){RESET}")
    
    print(f"\nPython version: {sys.version}")
    print(f"Current directory: {os.getcwd()}")
    print(f"\n{GREEN}All imports successful! Your environment is ready.{RESET}")
    
except ImportError as e:
    print(f"{RED}Import error: {e}{RESET}")
    print("Run: pip install -r requirements.txt")
except Exception as e:
    print(f"{RED}Unexpected error: {e}{RESET}")
