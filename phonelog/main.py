import os
from dotenv import load_dotenv
from smartsheets import Smartsheets

# Load our environment variables
load_dotenv()

# Retrieve the Smartsheet Client ID
client_id = os.getenv("SMARTSHEET_CLIENT_ID")
sheet_id = os.getenv("SMARTSHEET_SHEET_ID")

sheet = Smartsheets(client_id, sheet_id)
sheet.get_smartsheet_data('./internal/data.json')
