import os
from pathlib import Path
from datetime import datetime
import pandas as pd


class LytxReport: 
    def __init__(self, srv):
        if os.getenv("PRODUCTION") == "true":
            self.ssid = os.getenv("PROD_SSID")
        else:
            self.ssid = os.getenv("DEV_SSID")

        self.srv = srv
        
    def process_file(self, file_path, file_name):
        suffix = Path(file_path).suffix.lower()

        if suffix == ".csv":
            return self.lytx_report(file_path, file_type="csv")

        elif suffix in [".xlsx", ".xls", ".xlsm", ".xlsb"]:
            return self.lytx_report(file_path, file_type="excel")

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

        suffix = Path(file_path).suffix.lower()

        if suffix == ".csv":
            df = pd.read_csv(file_path, encoding="cp1252", sep=None, engine="python") #
        elif suffix == ".xls":
            df = pd.read_excel(file_path, engine="xlrd")
        else:
            df = pd.read_excel(file_path) 

        mapped_df = pd.DataFrame({
            0: df.iloc[:, 3],  # Sheets A = Excel D
            1: "",             # Sheets B = blank
            2: df.iloc[:, 0],  # Sheets C = Excel A
            3: df.iloc[:, 1],  # Sheets D = Excel B
            4: df.iloc[:, 2],  # Sheets E = Excel C
            5: df.iloc[:, 6],  # Sheets F = Excel G
            6: df.iloc[:, 5],  # Sheets G = Excel F
        })

        df_values = self._df_to_values(mapped_df)

        df_values.append(["", "", "", "", "", "", ""])
        
        df_values.append(["KEEP", "", "2304", "VCI - VC067 - San Marcos, CA", "", "", ""])
        df_values.append(["KEEP", "", "3315C", "VCI - VC095 - Upland - Corporate", "", "", ""])
        df_values.append(["KEEP", "", "2128", "VCI - VC095 - Upland - Corporate", "", "", ""])
        df_values.append(["KEEP", "", "3391", "VCI - VC073 - Henderson, NV", "", "", ""])

        ws.update(
            range_name=clear_range,
            values=df_values,
            value_input_option="USER_ENTERED"
        )

        print(f"Wrote {len(df_values)} rows to '{worksheet_name}'")


    def lytx_report(self, file_path, file_type):
        return self.clear_and_write(
            worksheet_name="Trucks",
            clear_range="A3:G",
            file_path=file_path,
            file_type=file_type
        )
    
    def update_report_date(self):
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet("Trucks") 
        ws.update_cell(1,2, datetime.now().strftime("%m/%d/%Y"))