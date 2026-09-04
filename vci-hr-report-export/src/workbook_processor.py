from datetime import datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


def normalize_header(value: Any) -> str:
    """Normalize a header for reliable comparison."""
    if value is None:
        return ""

    return "".join(
        character
        for character in str(value).strip().upper()
        if character.isalnum()
    )


def find_header_row(
    worksheet,
    required_headers: list[str],
    rows_to_scan: int = 20,
) -> int | None:
    """Find the row containing all required headers."""
    normalized_required = {
        normalize_header(header)
        for header in required_headers
    }

    maximum_row = min(
        worksheet.max_row,
        rows_to_scan,
    )

    for row_number in range(1, maximum_row + 1):
        normalized_values = {
            normalize_header(cell.value)
            for cell in worksheet[row_number]
            if cell.value is not None
        }

        if normalized_required.issubset(normalized_values):
            return row_number

    return None


def parse_worksheet_date(
    worksheet_name: str,
    date_format: str,
):
    """Return the date represented by an MMDDYYYY worksheet name."""
    try:
        return datetime.strptime(
            worksheet_name.strip(),
            date_format,
        ).date()
    except ValueError:
        return None


def clean_workbook(
    source_path: Path,
    output_path: Path,
    date_format: str,
    required_headers: list[str],
) -> dict:
    """
    Keep only the most recent dated worksheet, validate it,
    and save the cleaned workbook.
    """
    workbook = load_workbook(
        source_path,
        read_only=False,
        data_only=False,
    )

    dated_worksheets = []

    for worksheet in workbook.worksheets:
        worksheet_date = parse_worksheet_date(
            worksheet.title,
            date_format,
        )

        if worksheet_date is not None:
            dated_worksheets.append(
                (
                    worksheet_date,
                    worksheet.title,
                )
            )

    if not dated_worksheets:
        workbook.close()
        raise ValueError(
            f"No worksheet in {source_path.name} has a valid "
            f"MMDDYYYY name."
        )

    latest_date = max(
        worksheet_date
        for worksheet_date, _ in dated_worksheets
    )

    latest_matches = [
        worksheet_name
        for worksheet_date, worksheet_name in dated_worksheets
        if worksheet_date == latest_date
    ]

    if len(latest_matches) != 1:
        workbook.close()
        raise ValueError(
            f"More than one worksheet represents the latest date "
            f"{latest_date:%m%d%Y} in {source_path.name}."
        )

    latest_worksheet_name = latest_matches[0]
    latest_worksheet = workbook[latest_worksheet_name]

    header_row = find_header_row(
        latest_worksheet,
        required_headers,
    )

    if header_row is None:
        workbook.close()
        raise ValueError(
            f"The latest worksheet {latest_worksheet_name} in "
            f"{source_path.name} does not contain the required headers: "
            f"{', '.join(required_headers)}."
        )

    for worksheet in list(workbook.worksheets):
        if worksheet.title != latest_worksheet_name:
            workbook.remove(worksheet)

    latest_worksheet.sheet_state = "visible"

    output_path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    workbook.save(
        output_path
    )

    workbook.close()

    return {
        "latest_worksheet": latest_worksheet_name,
        "latest_date": latest_date.isoformat(),
        "header_row": header_row,
        "output_path": str(output_path),
    }