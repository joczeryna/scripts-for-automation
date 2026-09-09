from __future__ import annotations

import os
import csv
from pathlib import Path

from openpyxl import Workbook, load_workbook


def prepare_current_tab(source: Path, prepared: Path, source_tab: str) -> tuple[int, int]:
    with source.open("r", encoding="utf-8-sig", newline="") as handle:
        values = list(csv.reader(handle))

    populated_rows = sum(1 for row in values if any(value != "" for value in row))
    if populated_rows < 2:
        raise ValueError("CURRENT does not contain a header and at least one data row.")

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = source_tab
    for row in values:
        worksheet.append(row)

    prepared.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(prepared)

    verification = load_workbook(prepared, read_only=True, data_only=False)
    try:
        if verification.sheetnames != [source_tab]:
            raise ValueError("Prepared workbook contains an unexpected worksheet layout.")
        verified_sheet = verification[source_tab]
        rows, columns = verified_sheet.max_row, verified_sheet.max_column
        if rows < 2 or columns < 1:
            raise ValueError("Prepared CURRENT worksheet failed row/column validation.")
        if any(
            cell.data_type == "f"
            for row in verified_sheet.iter_rows()
            for cell in row
        ):
            raise ValueError("Prepared CURRENT contains formulas; static values were required.")
    finally:
        verification.close()

    return rows, columns


def replace_destination(prepared: Path, destination: Path) -> None:
    if not destination.parent.exists():
        raise FileNotFoundError(
            f"Destination folder is unavailable: {destination.parent}. Confirm W: is mapped."
        )

    staged = destination.with_name(f".{destination.stem}.pending{destination.suffix}")
    try:
        staged.write_bytes(prepared.read_bytes())
        os.replace(staged, destination)
    finally:
        if staged.exists():
            staged.unlink()
