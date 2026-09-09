from __future__ import annotations

from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def _as_date(value: Any) -> date | None:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if value is None:
        return None
    text = str(value).strip()
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d", "%m-%d-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    return None


def validate_workbook(path: Path) -> None:
    if not path.exists() or path.stat().st_size == 0:
        raise RuntimeError(f"Downloaded workbook is missing or empty: {path}")
    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        if not workbook.sheetnames:
            raise RuntimeError(f"Workbook has no worksheets: {path.name}")
        sheet = workbook[workbook.sheetnames[0]]
        max_row = sheet.max_row
        if max_row is None:
            # Some Power BI paginated exports omit worksheet dimension metadata.
            # Force OpenPyXL to discover the used range from the actual cells.
            sheet.reset_dimensions()
            sheet.calculate_dimension(force=True)
            max_row = sheet.max_row

        populated_cells = 0
        inspected_rows = 0
        for row in sheet.iter_rows():
            inspected_rows += 1
            populated_cells += sum(cell.value is not None for cell in row)
            if inspected_rows >= 100 or populated_cells >= 2:
                break

        if populated_cells < 2:
            raise RuntimeError(f"Workbook contains no usable report content: {path.name}")
    finally:
        workbook.close()


def clean_dvir_date_window(
    path: Path,
    date_header: str,
    start_date: date,
    run_date: date,
) -> int:
    """Keep dated records from the last prior export date through the run date."""
    if start_date > run_date:
        raise ValueError("DVIR cleanup start date cannot be after the run date")

    workbook = load_workbook(path)
    sheet = workbook[workbook.sheetnames[0]]

    header_row = None
    date_column = None
    for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 25)):
        for cell in row:
            if str(cell.value).strip().casefold() == date_header.strip().casefold():
                header_row = cell.row
                date_column = cell.column
                break
        if date_column:
            break

    if not header_row or not date_column:
        workbook.close()
        raise RuntimeError(f"Could not find '{date_header}' in {path.name}")

    rows_to_delete: list[int] = []
    for row_number in range(header_row + 1, sheet.max_row + 1):
        row_date = _as_date(sheet.cell(row=row_number, column=date_column).value)
        if row_date and not (start_date <= row_date <= run_date):
            rows_to_delete.append(row_number)

    for row_number in reversed(rows_to_delete):
        sheet.delete_rows(row_number)

    workbook.save(path)
    workbook.close()
    return len(rows_to_delete)
