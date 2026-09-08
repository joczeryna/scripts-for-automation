import os
from pathlib import Path
from datetime import datetime
import pandas as pd


class SafetyReport:

    def __init__(self, srv):
        if os.getenv("PRODUCTION") == "true":
            self.ssid = os.getenv("PROD_SSID")
        else:
            self.ssid = os.getenv("DEV_SSID")

        self.srv = srv
        
    def process_file(self, file_path, file_name):
        suffix = Path(file_path).suffix.lower()

        if suffix == ".csv":
            return self.safety_worksheet(file_path, file_type="csv")

        elif suffix in [".xlsx", ".xls", ".xlsm", ".xlsb"]:
            return self.safety_worksheet(file_path, file_type="excel")

        else:
            print(f"Unsupported file: {file_name}")

    def _df_to_values(self, df: pd.DataFrame):
        values = []
        data = df.values.tolist()

        for row in data:
            row_values = []

            for v in row:
                if pd.isna(v):
                    row_values.append("")
                elif isinstance(v, datetime):
                    row_values.append(v.strftime("%Y-%m-%d %H:%M:%S"))
                elif isinstance(v, float):
                    if v.is_integer():
                        row_values.append(str(int(v)))
                    else:
                        row_values.append(str(v))
                else:
                    row_values.append(str(v))

            values.append(row_values)

        return values

    def clear_and_write(self, worksheet_name: str, clear_range: str, file_path, file_type):
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet(worksheet_name)

        ws.batch_clear([clear_range])

        if file_type == "csv":
            df = pd.read_csv(file_path, encoding="cp1252", sep=None, engine="python")
        else:
            df = pd.read_excel(file_path)

        values = self._df_to_values(df)

        if not values:
            print(f"No rows to write for '{worksheet_name}'")
            return

        end_row = 2 + len(values)
        write_range = f"A3:AF{end_row}"

        ws.update(
            values=values,
            range_name=write_range,
            value_input_option="USER_ENTERED"
        )

        print(f"Wrote {len(values)} rows to '{worksheet_name}'")

    def safety_worksheet(self, file_path, file_type):
        return self.clear_and_write(
            worksheet_name="Import",
            clear_range="A3:AF",
            file_path=file_path,
            file_type=file_type
        )

    def update_report_date(self):
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet("Import")
        ws.update_cell(1, 4, datetime.now().strftime("%m/%d/%Y"))