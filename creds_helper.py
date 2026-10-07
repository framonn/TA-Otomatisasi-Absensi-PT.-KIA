import os
import json
import base64
import tempfile
import gspread

def get_gspread_client():
    """Load gspread client from GOOGLE_CREDENTIALS_B64 env var or credentials.json file"""
    # Try environment variable first
    if os.getenv("GOOGLE_CREDENTIALS_B64"):
        try:
            creds_json = base64.b64decode(os.getenv("GOOGLE_CREDENTIALS_B64")).decode()
            creds = json.loads(creds_json)
            # Use temporary file for gspread
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                json.dump(creds, f)
                return gspread.service_account(filename=f.name)
        except Exception as e:
            print(f"Failed to load from env var: {e}")
    
    # Fallback to credentials.json
    if os.path.exists("credentials.json"):
        return gspread.service_account(filename="credentials.json")
    
    raise FileNotFoundError("No credentials found: set GOOGLE_CREDENTIALS_B64 env var or place credentials.json in /app")
