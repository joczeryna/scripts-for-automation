from datetime import datetime
from googleapiclient.errors import HttpError
import pandas as pd
import os

class ProductionReport:
    def __init__(self, srv):
        """
        Initializes the Production Report object

        Args:
            spreadsheet_id (str): The ID of the Google Spreadsheet
            srv (Sheets): The sheets client
        """
        if os.getenv("PRODUCTION") == "true":
            self.ssid = os.getenv("PROD_SSID")
        else:
            self.ssid = os.getenv("DEV_SSID")
        self.srv = srv
        self.last_approved = None
        self.last_daily = None
        self.last_daily_approval = None

    def set_last_approved(self):
        """
        Retreives the last approval date found on the Google Sheet
        """
        try:
            # Open the spreadsheet by ID
            spreadsheet = self.srv.client.open_by_key(self.ssid)
            sheet = spreadsheet.worksheet("Info")
            vals = sheet.get("B2:B4")
            last_approved_string = vals[0][0]
            last_daily_string = vals[1][0]
            self.last_daily = vals[2][0]
            
            if last_approved_string == None:
                return
            
            format_string = "%m/%d/%Y %H:%M:%S"

            self.last_approved = datetime.strptime(last_approved_string, format_string)
            self.last_daily_approval = datetime.strptime(last_daily_string, format_string)

            print(f"Last approval date: {self.last_approved} | Last daily approved: {self.last_daily_approval}")
        except HttpError as err:
            print(f"Error getting values from sheet: {str(err)}")

    def update_report_date(self, ctime):
        """
        Updates the date the data was last pulled based off of the created time of the file
        """
        try:
            # Open the spreadsheet by ID
            spreadsheet = self.srv.client.open_by_key(self.ssid)
            sheet = spreadsheet.worksheet("Info")
            sheet.update_acell("B1", ctime.strftime('%Y-%m-%d %H:%M:%S'))
            
            print(f"Successfully updated the last updated time to {ctime}")
        except HttpError as err:
            print(f"Error updating report date on sheet: {str(err)}")
    
    def read_employee_production_report(self, file_path, sheet_name):
        """
        Reads the Employee Production Report data
        
        Args:
            file_path (Path): Path to the Excel file
            sheet_name (str): The name of the sheet the data is in 
        
        Returns:
            string: Any errors received
        """
        try:
            filter_and_sort_excel(file_path, sheet_name, self.last_approved, file_path, 13)

            # Open the spreadsheet by ID
            spreadsheet = self.srv.client.open_by_key(self.ssid)
            worksheet = spreadsheet.worksheet("Data")

            # Read the data in the file into a data frame
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Convert DataFrame to list of lists
            values_to_append = []

            # Add data rows
            for _, row in df.iterrows():
                # Convert each value to string and handle NaN/None values
                row_values = []
                for value in row:
                    if pd.isna(value):
                        row_values.append("")
                    elif isinstance(value, datetime):
                        row_values.append(value.strftime('%Y-%m-%d %H:%M:%S'))
                    else:
                        row_values.append(str(value))
                values_to_append.append(row_values)
            
            # Append all rows at once (more efficient)
            if values_to_append:
                worksheet.append_rows(values_to_append, "USER_ENTERED")
                print(f"Successfully appended {len(values_to_append)} rows to Google Sheet")
        except FileNotFoundError:
            return "File not found"
        except PermissionError:
            return "Permission denied"
        except HttpError as err:
            print(f"Error updating data on sheet: {str(err)}")
        except Exception as e:
            return f"Error reading file: {str(e)}"
    
    def read_production_approval(self, file_path, sheet_name):
        """
        Reads the Production Approval Analysis data
        
        Args:
            file_path (Path): Path to the Excel file
            sheet_name (str): The name of the sheet the data is in 
        
        Returns:
            string: Any errors received
        """
        try:
            filter_and_sort_excel(file_path, sheet_name, self.last_daily_approval, file_path, 8)

            # Open the spreadsheet by ID
            spreadsheet = self.srv.client.open_by_key(self.ssid)
            worksheet = spreadsheet.worksheet("ProductionApproval")

            # Read the data in the file into a data frame
            df = pd.read_excel(file_path, sheet_name=sheet_name)

            # Convert DataFrame to list of lists
            values_to_append = []

            # Add data rows
            for _, row in df.iterrows():
                # Convert each value to string and handle NaN/None values
                row_values = []
                for value in row:
                    if pd.isna(value):
                        row_values.append("")
                    elif isinstance(value, datetime):
                        row_values.append(value.strftime('%Y-%m-%d %H:%M:%S'))
                    else:
                        row_values.append(str(value))
                values_to_append.append(row_values)
            
            # Append all rows at once (more efficient)
            if values_to_append:
                worksheet.append_rows(values_to_append, "USER_ENTERED")
                print(f"Successfully appended {len(values_to_append)} rows to Google Sheet")
        except FileNotFoundError:
            return "File not found"
        except PermissionError:
            return "Permission denied"
        except HttpError as err:
            print(f"Error updating data on sheet: {str(err)}")
        except Exception as e:
            return f"Error reading file: {str(e)}"
        
def filter_and_sort_excel(file_path, sheet_name=None, filter_date=None, output_path=None, column_index=None):
    """
    Opens an Excel file, filters by column H (date > filter_date), and sorts by the specified column
    
    Parameters:
    file_path (str): Path to the input Excel file
    sheet_name (str): Name of the sheet to process (None for first sheet)
    filter_date (str or datetime): Date to filter by (format: 'YYYY-MM-DD')
    output_path (str): Path for output file (None to overwrite original)
    column_index (int): The index of the column to filter by
    """
    
    # Set default filter date if none provided
    if filter_date is None:
        filter_date = '2024-01-01'
    
    # Convert string date to datetime if necessary
    if isinstance(filter_date, str):
        filter_date = datetime.strptime(filter_date, '%Y-%m-%d')
    
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Ensure column H exists
        if df.shape[1] < column_index+1:
            raise ValueError("File doesn't have enough columns")
        
        # Get column
        col_name = df.columns[column_index]
        
        # Convert column H to datetime if it isn't already
        df[col_name] = pd.to_datetime(df[col_name], errors='coerce')
        
        # Filter rows where column H date > filter_date
        filtered_df = df[df[col_name] > filter_date]
        
        # Sort by column H (ascending order)
        filtered_df = filtered_df.sort_values(by=col_name, ascending=True)
        
        # Reset index
        filtered_df = filtered_df.reset_index(drop=True)
        
        # Determine output path
        if output_path is None:
            output_path = file_path
        
        # Save the filtered and sorted data
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, sheet_name=sheet_name or 'Sheet1', index=False)
        
        print(f"Successfully filtered {len(filtered_df)} rows")
        print(f"Filtered data saved to: {output_path}")
        
        return filtered_df
        
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        return None