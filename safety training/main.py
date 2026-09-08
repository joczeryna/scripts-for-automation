from datetime import datetime
from pathlib import Path
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google_auth_oauthlib.flow import Flow
import gspread
import os
import time
import pandas as pd
import openpyxl
from safetytraining import SafetyReport
from dotenv import load_dotenv

class Sheets:
    def __init__(self, oauth_credentials_file, token_file='token.json'):
        """
        Initialize the Google Sheets appender with OAuth
        
        Args:
            oauth_credentials_file (str): Path to your OAuth client credentials JSON file
            spreadsheet_id (str): The ID of your Google Spreadsheet
            token_file (str): Path to store/retrieve the access token
        """
        self.oauth_credentials_file = oauth_credentials_file
        self.token_file = token_file
        self.client = None
        self.sheet = None

    def authenticate(self):
        """Authenticate with Google Sheets API using OAuth"""
        try:
            # Define the scope for Google Sheets API
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            creds = None
            
            # Check if we have saved credentials
            if os.path.exists(self.token_file):
                try:
                    creds = Credentials.from_authorized_user_file(self.token_file, scopes)
                    print("✓ Found existing credentials")
                except Exception as e:
                    print(f"⚠ Could not load existing credentials: {e}")
            
            # If credentials are not valid, get new ones
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    try:
                        creds.refresh(Request())
                        print("✓ Refreshed expired credentials")
                    except Exception as e:
                        print(f"⚠ Could not refresh credentials: {e}")
                        creds = None
                
                if not creds:
                    # Run OAuth flow
                    flow = Flow.from_client_secrets_file(
                        self.oauth_credentials_file, scopes)
                    flow.redirect_uri = 'http://localhost:8080/callback'
                    
                    # Get authorization URL
                    auth_url, _ = flow.authorization_url(prompt='consent')
                    
                    print("\n" + "="*60)
                    print("🔐 OAUTH AUTHENTICATION REQUIRED")
                    print("="*60)
                    print(f"Please visit this URL to authorize the application:")
                    print(f"\n{auth_url}\n")
                    print("After authorization, you'll be redirected to a URL.")
                    print("Copy the ENTIRE redirect URL and paste it below.")
                    print("="*60)
                    
                    # Get the authorization response
                    authorization_response = input("Enter the full redirect URL: ").strip()
                    
                    # Exchange the authorization code for credentials
                    flow.fetch_token(authorization_response=authorization_response)
                    creds = flow.credentials
                    
                    # Save credentials for future use
                    with open(self.token_file, 'w') as token:
                        token.write(creds.to_json())
                    
                    print("✓ Successfully obtained new credentials")
            
            # Initialize the client
            self.client = gspread.authorize(creds)
            print("✓ Successfully authenticated with Google Sheets API")
            return True
            
        except Exception as e:
            print(f"✗ Authentication failed: {e}")
            return False
    
    def authenticate_with_local_server(self):
        """Alternative authentication method using local server (more user-friendly)"""
        try:
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            creds = None
            
            # Check if we have saved credentials
            if os.path.exists(self.token_file):
                try:
                    creds = Credentials.from_authorized_user_file(self.token_file, scopes)
                except Exception as e:
                    print(f"⚠ Could not load existing credentials: {e}")
            
            # If credentials are not valid, get new ones
            if not creds or not creds.valid:
                if creds and creds.expired and creds.refresh_token:
                    try:
                        creds.refresh(Request())
                        print("✓ Refreshed expired credentials")
                    except Exception as e:
                        print(f"⚠ Could not refresh credentials: {e}")
                        creds = None
                
                if not creds:
                    # Run OAuth flow with local server
                    flow = InstalledAppFlow.from_client_secrets_file(
                        self.oauth_credentials_file, scopes)
                    
                    print("\n🔐 Opening browser for OAuth authentication...")
                    creds = flow.run_local_server(port=8080)
                    
                    # Save credentials for future use
                    with open(self.token_file, 'w') as token:
                        token.write(creds.to_json())
                    
                    print("✓ Successfully obtained new credentials")
            
            # Initialize the client
            self.client = gspread.authorize(creds)
            print("✓ Successfully authenticated with Google Sheets API")
            return True
            
        except ImportError:
            print("⚠ google-auth-oauthlib not available for local server method")
            print("  Falling back to manual method...")
            return self.authenticate()
        except Exception as e:
            print(f"✗ Authentication failed: {e}")
            return False

def get_excel_column_count(file_path):
    """
    Get the number of columns in the first sheet of an Excel file.
    
    Args:
        file_path (Path): Path to the Excel file
    
    Returns:
        tuple: (number_of_columns, sheet_name, error_message)
    """
    try:
        # Try using openpyxl first (works with .xlsx files)
        if file_path.suffix.lower() in ['.xlsx', '.xlsm']:
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            first_sheet = workbook.active
            sheet_name = first_sheet.title
            
            # Find the maximum column with data
            max_col = first_sheet.max_column
            
            # If max_column is 1 but the cell is empty, there might be no data
            if max_col == 1 and first_sheet.cell(1, 1).value is None:
                max_col = 0
                
            workbook.close()
            return max_col, sheet_name, None
            
        # For older Excel formats (.xls, .xlsb), use pandas
        else:
            # Read just the first few rows to determine column count
            df = pd.read_excel(file_path, sheet_name=0, engine="xlrd")
            
            # Get sheet names to find the first sheet name
            xl_file = pd.ExcelFile(file_path)
            first_sheet_name = xl_file.sheet_names[0] if xl_file.sheet_names else "Unknown"
            xl_file.close()
            
            return len(df.columns), first_sheet_name, None
            
    except FileNotFoundError:
        return 0, "", "File not found"
    except PermissionError:
        return 0, "", "Permission denied"
    except Exception as e:
        return 0, "", f"Error reading file: {str(e)}"

def is_excel_file(file_path):
    """
    Check if a file is an Excel file based on its extension.

    Args:
        file_path (str): Path to the file to check

    Returns: 
        True for .xlsx, .xls, .xlsm, .xlsb files
    """
    excel_extensions = {'.xlsx', '.xls', '.xlsm', '.xlsb', '.csv'}
    return file_path.suffix.lower() in excel_extensions

def was_modified_today(file_path):
    """
    Check if a file was modified today.

    Args:
        file_path (str): Path to the file to check

    Returns: 
        datetime: When the file was created.
        bool: True if the file's modification time is today.
    """
    try:
        # Get file modification time
        mod_time = datetime.fromtimestamp(file_path.stat().st_mtime)
        today = datetime.today().strftime("%Y-%m-%d")
        
        # Check if modification date is today
        return datetime.fromtimestamp(file_path.stat().st_ctime), mod_time.strftime("%Y-%m-%d") == today
    except (OSError, ValueError):
        return False
    
def get_new_excel_files(directory_path):
    """
    Check all files in a directory for Excel files modified today.
    
    Args:
        directory_path (str): Path to the directory to check
    
    Returns:
        list: List of Excel files modified today
    """
    directory = Path(directory_path)
    
    if not directory.exists():
        print(f"Error: Directory '{directory_path}' does not exist.")
        return []
    
    if not directory.is_dir():
        print(f"Error: '{directory_path}' is not a directory.")
        return []
    
    excel_files_today = []
    total_files = 0
    excel_files_total = 0
    
    print(f"Checking files in: {directory.absolute()}")
    print(f"Today's date: {datetime.today().strftime('%Y-%m-%d')}")
    print("-" * 50)
    
    try:
        # Iterate through all files in the directory
        for file_path in directory.iterdir():
            if file_path.is_file():
                total_files += 1
                
                # Check if it's an Excel file
                if is_excel_file(file_path):
                    excel_files_total += 1
                    
                    # Check if it was modified today
                    ctime, modified_today = was_modified_today(file_path)
                    if modified_today:
                        
                      # Get sheet name for files modified today
                        error = None

                        file_info = {
                            'path': file_path,
                            'name': file_path.name,
                            'ctime': time.ctime,
                            'error': error
                        }
                                                
                        excel_files_today.append(file_info)
                        
                        if error:
                            print(f"✓ Excel file found (today): {file_path.name} - Error: {error}")
                        else:
                         print(f"  Excel file found (today): {file_path.name}")
    
    except PermissionError:
        print(f"Error: Permission denied accessing directory '{directory_path}'")
        return []
    
    print("-" * 50)
    print(f"Summary:")
    print(f"  Total files: {total_files}")
    print(f"  Total Excel files: {excel_files_total}")
    print(f"  Excel files modified today: {len(excel_files_today)}")
    
    return excel_files_today

def main():
    load_dotenv()

    start_time = time.perf_counter()

    directory = os.getenv("PROD_DIR")
    excel_files = get_new_excel_files(directory)

    print("FILES FOUND:", [f['name'] for f in excel_files])

    if len(excel_files) == 0:
        return

    credentials_file = os.getenv("CREDS_FILE")
    srv = Sheets(credentials_file)

    print("Attempting Sheets authentication with local server...")
    if not srv.authenticate_with_local_server():
        print("Local server method failed, trying manual method...")
        if not srv.authenticate():
            print("All authentication methods failed!")
            return

    prod_report = SafetyReport(srv)

    for f in excel_files:
        print(f"File Name: {repr(f['name'])}")
        prod_report.process_file(f['path'], f['name'])

    prod_report.update_report_date()

    end_time = time.perf_counter()
    duration = end_time - start_time
    print("-" * 50)
    print(f"Script completed in {duration:.2f} seconds")

if __name__ == "__main__":
    main()