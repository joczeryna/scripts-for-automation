from pathlib import Path
import pandas as pd
import os


class ManHourReport:
    def __init__(self, srv):
        if os.getenv("PRODUCTION") == "true":
            self.ssid = os.getenv("PROD_SSID")
        else:
            self.ssid = os.getenv("DEV_SSID")

        self.srv = srv

    def process_file(self, file_path, file_name): 
        suffix = Path(file_name).suffix.lower() # Get the file extension in lowercase

        if suffix in [".xlsx", ".xls", ".xlsm", ".xlsb"]: # Check if the file is an Excel file
            return self.mh_report(file_path, suffix) # Process the Excel file and update the Google Sheet
        else:
            print(f"Unsupported file: {file_name}") 

    def df_to_values(self, df: pd.DataFrame): # Convert a DataFrame to a list of lists, handling NaN and datetime values
        values = [] # Initialize an empty list to store the cleaned rows

        for row in df.itertuples(index=False, name=None): # Iterate over each row in the DataFrame as a tuple, without including the index and without naming the tuple
            cleaned_row = [] # Initialize an empty list to store the cleaned values for the current row

            for v in row:
                if pd.isna(v): # Check if the value is NaN (Not a Number) using pandas' isna function
                    cleaned_row.append("") # If the value is NaN, append an empty string to the cleaned row
                elif isinstance(v, pd.Timestamp): # Check if the value is a pandas Timestamp (datetime) object
                    cleaned_row.append(v.strftime("%m/%d/%Y")) # If the value is a Timestamp, format it as a string in the format "MM/DD/YYYY" and append it to the cleaned row
                elif hasattr(v, "item"): # Check if the value has an "item" method (this is often the case for numpy scalar types)
                    cleaned_row.append(v.item()) # If the value has an "item" method, call it to get the underlying Python scalar value and append it to the cleaned row
                else:
                    cleaned_row.append(v) # If the value is not NaN, not a Timestamp, and does not have an "item" method, append it as is to the cleaned row

            values.append(cleaned_row) 

        return values

    def clear_and_write_excel(self, worksheet_name: str, clear_range: str, file_path: str, file_type: str): # Clear a specified range in a Google Sheet worksheet and write data from an Excel file to that range
        spreadsheet = self.srv.client.open_by_key(self.ssid) # Open the Google Sheet using the service's client and the spreadsheet ID (ssid)
        worksheet = spreadsheet.worksheet(worksheet_name) # Get the specified worksheet by name from the opened spreadsheet

        if file_type == ".xls": # If the file type is ".xls", use the "xlrd" engine to read the Excel file into a DataFrame (this is necessary because newer versions of pandas do not support reading .xls files with the default engine)
            df = pd.read_excel(file_path, engine="xlrd")
        else:
            df = pd.read_excel(file_path)

        df_values = self.df_to_values(df) # Convert the DataFrame to a list of lists, handling NaN and datetime values

        worksheet.batch_clear([clear_range]) 

        worksheet.update(
            clear_range,
            df_values,
            value_input_option="RAW"
        )

        print(f"Wrote {len(df_values)} rows to {worksheet_name}")

    def mh_report(self, file_path, file_type): # Process an Excel file and update the "Data" worksheet in the Google Sheet with its contents
        return self.clear_and_write_excel( 
            worksheet_name="Data",
            clear_range="A2:AB",
            file_path=file_path,
            file_type=file_type
        )