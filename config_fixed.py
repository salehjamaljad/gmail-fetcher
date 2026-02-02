import os
import json
import re
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import requests

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

def safe_json_load(file_path):
    """Load JSON with proper error handling and cleanup"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Remove BOM if exists
        content = content.lstrip('\ufeff')
        
        # Remove any control characters
        content = re.sub(r'[\x00-\x1f\x7f-\x9f]', '', content)
        
        # Strip whitespace
        content = content.strip()
        
        # Try to parse
        return json.loads(content)
        
    except json.JSONDecodeError as e:
        print(f"\n❌ JSON Error in {file_path}:")
        print(f"   Position: {e.pos}, Line: {e.lineno}, Column: {e.colno}")
        print(f"   Message: {e.msg}")
        
        # Show content for debugging
        lines = content.split('\n')
        if e.lineno <= len(lines):
            print(f"   Problem line: {lines[e.lineno-1][:100]}")
        
        raise Exception(f"Invalid JSON in {file_path}: {e.msg}")
    
    except Exception as e:
        print(f"\n❌ Error reading {file_path}: {e}")
        raise

def authenticate_gmail():
    """Authenticate with Gmail API using token.json and credentials.json"""
    creds = None
    
    # Check if token.json exists
    if os.path.exists('token.json'):
        print("Loading token.json...")
        try:
            token_data = safe_json_load('token.json')
            creds = Credentials.from_authorized_user_info(token_data, SCOPES)
        except Exception as e:
            print(f"Failed to load token.json: {e}")
            creds = None
    
    # If credentials are invalid or don't exist, refresh or create new ones
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("Refreshing expired credentials...")
            creds.refresh(Request())
        else:
            print("Creating new credentials from credentials.json...")
            creds_data = safe_json_load('credentials.json')
            flow = InstalledAppFlow.from_client_config(creds_data, SCOPES)
            creds = flow.run_local_server(port=0)
        
        # Save the credentials for the next run
        print("Saving refreshed token...")
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    print("✅ Gmail authentication successful")
    return build('gmail', 'v1', credentials=creds)

def upload_order_and_metadata(file_bytes, filename, client, order_type, 
                              order_date, delivery_date, status, city=None, po_number=None):
    """Upload file and metadata to Supabase"""
    
    SUPABASE_URL = os.environ.get('SUPABASE_URL')
    SUPABASE_KEY = os.environ.get('SUPABASE_KEY')
    
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise Exception("Missing SUPABASE_URL or SUPABASE_KEY environment variables")
    
    # Upload file to storage
    storage_url = f"{SUPABASE_URL}/storage/v1/object/orders/{filename}"
    headers = {
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "Content-Type": "application/octet-stream"
    }
    
    response = requests.post(storage_url, headers=headers, data=file_bytes)
    
    if response.status_code not in [200, 201]:
        raise Exception(f"Storage upload failed: {response.text}")
    
    # Insert metadata into database
    db_url = f"{SUPABASE_URL}/rest/v1/purchase_orders"
    headers = {
        "Authorization": f"Bearer {SUPABASE_KEY}",
        "apikey": SUPABASE_KEY,
        "Content-Type": "application/json",
        "Prefer": "return=representation"
    }
    
    metadata = {
        "file_path": filename,
        "client": client,
        "order_type": order_type,
        "order_date": order_date,
        "delivery_date": delivery_date,
        "status": status,
    }
    
    if city:
        metadata["city"] = city
    if po_number:
        metadata["po_number"] = po_number
    
    db_response = requests.post(db_url, headers=headers, json=metadata)
    
    if db_response.status_code not in [200, 201]:
        raise Exception(f"Database insert failed: {db_response.text}")
    
    return db_response.json()
