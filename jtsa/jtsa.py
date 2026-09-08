from datetime import datetime
import os
import pandas as pd


class JtsaReport:
    def __init__(self, srv):
        if os.getenv("PRODUCTION") == "true":
            self.ssid = os.getenv("PROD_SSID")
        else:
            self.ssid = os.getenv("DEV_SSID")

        self.srv = srv

    def _read_excel_file(self, file_path, sheet_name: str) -> pd.DataFrame:
        """
        Read the Excel file as closely as possible to source data.
        Keeps blanks as blanks and avoids over-converting values.
        """
        ext = os.path.splitext(str(file_path))[1].lower()

        if ext == ".xls":
            df = pd.read_excel(file_path, sheet_name=sheet_name, engine="xlrd")
        else:
            df = pd.read_excel(file_path, sheet_name=sheet_name)

        return df

    def _normalize_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """
        Normalize only the columns that need special handling.
        Leave everything else as untouched as possible.
        """
        df = df.copy()

        # Keep Worker Number as 6-digit text
        if "Worker Number" in df.columns:
            def fmt_worker(x):
                if pd.isna(x) or x == "":
                    return ""
                try:
                    return str(int(float(x))).zfill(6)
                except (ValueError, TypeError):
                    return str(x).strip()

            df["Worker Number"] = df["Worker Number"].apply(fmt_worker)

        return df

    def _convert_cell_value(self, value, col_name=None):
        """
        Convert a single cell value for Google Sheets upload.
        Try to preserve native types where possible.
        """
        if pd.isna(value):
            return ""

        # Preserve Worker Number exactly as text
        if col_name == "Worker Number":
            return str(value)

        # Pandas Timestamp / Python datetime
        if isinstance(value, pd.Timestamp):
            if pd.isna(value):
                return ""
            return value.strftime("%m/%d/%Y %H:%M:%S")

        if isinstance(value, datetime):
            return value.strftime("%m/%d/%Y %H:%M:%S")

        # Convert floats that are whole numbers to int
        if isinstance(value, float):
            if value.is_integer():
                return int(value)
            return value

        return value

    def _df_to_values(self, df: pd.DataFrame):
        """
        Convert DataFrame to list-of-lists for gspread.
        Does NOT skip rows.
        """
        columns = list(df.columns)
        values = []

        for _, row in df.iterrows():
            row_values = []
            for col_name, v in zip(columns, row):
                row_values.append(self._convert_cell_value(v, col_name))
            values.append(row_values)

        return values

    def clear_and_write(self, worksheet_name: str, clear_range: str, file_path, sheet_name: str):
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet(worksheet_name)

        # Read file
        df = self._read_excel_file(file_path, sheet_name)
        df = self._normalize_dataframe(df)
        values = self._df_to_values(df)

        # Clear old data only
        ws.batch_clear([clear_range])

        if not values:
            print(f"No rows to write for '{worksheet_name}'")
            return

        # Write starting at A2
        ws.update(
            range_name="A2",
            values=values,
            value_input_option="RAW"
        )

        print(f"Wrote {len(values)} rows to '{worksheet_name}'")

    def sha_worksheet(self, file_path, sheet_name):
        return self.clear_and_write(
            worksheet_name="Shift Allocation",
            clear_range="A2:AB",
            file_path=file_path,
            sheet_name=sheet_name
        )

    def shapw_worksheet(self, file_path, sheet_name):
        return self.clear_and_write(
            worksheet_name="Shift Allocation PW",
            clear_range="A2:AH",
            file_path=file_path,
            sheet_name=sheet_name
        )

    def jtsa_worksheet(self, file_path, sheet_name):
        return self.clear_and_write(
            worksheet_name="JTSA Details Auditor",
            clear_range="A2:G",
            file_path=file_path,
            sheet_name=sheet_name
        )

    def update_report_date(self):
        spreadsheet = self.srv.client.open_by_key(self.ssid)
        ws = spreadsheet.worksheet("JTSA Summary")
        ws.update_cell(1, 6, datetime.now().strftime("%m/%d/%Y"))