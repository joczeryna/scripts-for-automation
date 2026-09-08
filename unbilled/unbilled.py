from datetime import datetime
from os import devnull
import xlrd
import pandas as pd
import os


class UnbilledReport:
    def __init__(self, srv):
        if os.getenv("PRODUCTION") == "true":
            self.ssid = os.getenv("PROD_SSID")
        else:
            self.ssid = os.getenv("DEV_SSID")

        self.srv = srv

        self.division_map = {
            "20": "20 Unbilled Sales WC",
            "28": "28 Unbilled Sales WC",
            "46": "46 Unbilled Sales WC",
            "47": "47 Unbilled Sales WC",
            "55": "55 Unbilled Sales WC",
            "67": "67 Unbilled Sales WC",
            "70": "70 Unbilled Sales WC",
            "78": "78 Unbilled Sales WC",
        }

    def _df_to_values(self, df: pd.DataFrame):
        values = []
        data = df.values.tolist()

        for row in data:
            row_values = []

            for i in range(len(row)):
                # O and R -> force as text and KEEP "NA"
                if i in [14, 17]:  # O = 14, R = 17
                    v = row[i]
                    if pd.isna(v):
                        row_values.append("NA")
                    else:
                        row_values.append(str(v))

                else:
                    v = row[i]

                    if pd.isna(v):
                        row_values.append("")
                    elif isinstance(v, datetime):
                        row_values.append(v.strftime("%m/%d/%Y"))
                    elif isinstance(v, float):
                        row_values.append(f"{v:.2f}")
                    else:
                        row_values.append(str(v))

            values.append(row_values)

        return values

    def clear_and_write(self, worksheet_name: str, clear_range: str, file_path): #
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet(worksheet_name)

        ws.batch_clear([clear_range])

        # Always read the first sheet in the Excel file
        book = xlrd.open_workbook(file_path, logfile=open(devnull, "w"))
        df = pd.read_excel(book, sheet_name=0, engine="xlrd", keep_default_na=False)
        values = self._df_to_values(df)

        if not values:
            print(f"No rows to write for '{worksheet_name}'")
            return

        end_row = 1 + len(values)
        write_range = f"{clear_range}{end_row}"

        ws.update(
            values=values,
            range_name=write_range,
            value_input_option="USER_ENTERED"
        )
        print(f"Wrote {len(values)} rows to '{worksheet_name}'")

    def process_file(self, file_path, file_name: str): 
        file_name = file_name.strip().lower() # Normalize the file name to lowercase and remove leading/trailing whitespace

        division = None
        if "unbilledsaleswc-" in file_name:  # Extract the division number from the file name
            division = (
                file_name
                .replace("unbilledsaleswc-", "") # Remove the prefix to isolate the division number
                .replace(".xls", "") # Remove the .xls extension if present
                .replace(".xlsx", "") # Remove the .xlsx extension if present
            )

        worksheet_name = self.division_map.get(division)

        if not worksheet_name:
            print(f"No worksheet mapping found for: {file_name}")
            return

        self.clear_and_write(
            worksheet_name=worksheet_name,
            clear_range="A2:BN",
            file_path=file_path
        )

    def update_report_date(self):
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet("Instructions")
        ws.update_cell(1, 5, datetime.now().strftime("%m/%d/%Y %H:%M:%S"))