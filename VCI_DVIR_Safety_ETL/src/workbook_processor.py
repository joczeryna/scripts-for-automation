from __future__ import annotations

import re
import zipfile
from pathlib import Path

from openpyxl import load_workbook


def _normalize(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").casefold())


def validate_workbook(
    path: Path,
    expected_sheet: str,
    required_headers: list[str | list[str]],
    minimum_columns: int | None = None,
    data_start_row: int = 2,
    first_column_value: str | None = None,
    required_data_columns: list[int] | None = None,
) -> tuple[int, int]:
    if not path.exists() or path.stat().st_size == 0:
        raise RuntimeError(f"Downloaded workbook is missing or empty: {path}")
    if not zipfile.is_zipfile(path):
        raise RuntimeError(f"Downloaded file is not a valid Excel workbook: {path.name}")

    workbook = load_workbook(path, read_only=True, data_only=False)
    try:
        matching_sheet = next(
            (name for name in workbook.sheetnames if name.casefold() == expected_sheet.casefold()),
            None,
        )
        if matching_sheet is None:
            raise RuntimeError(
                f"Expected worksheet '{expected_sheet}' was not found in {path.name}. "
                f"Available worksheets: {workbook.sheetnames}"
            )

        sheet = workbook[matching_sheet]
        row_count = sheet.max_row or 0
        column_count = sheet.max_column or 0
        cached_rows: list[tuple[object, ...]] | None = None
        # Power BI paginated-report exports can contain valid worksheet XML while
        # declaring an empty or undersized worksheet dimension. In read-only mode,
        # openpyxl trusts that metadata unless we explicitly reset it.
        expected_minimum = minimum_columns or 0
        if row_count < data_start_row or column_count < expected_minimum:
            sheet.reset_dimensions()
            cached_rows = list(sheet.iter_rows(values_only=True))
            row_count = len(cached_rows)
            column_count = max((len(row) for row in cached_rows), default=0)
        required_groups = [
            {_normalize(value) for value in (header if isinstance(header, list) else [header])}
            for header in required_headers
        ]
        if required_groups:
            rows = (
                cached_rows[:25]
                if cached_rows is not None
                else list(
                    sheet.iter_rows(
                        min_row=1,
                        max_row=min(row_count or 1, 25),
                        values_only=True,
                    )
                )
            )
            if not any(
                all(group & {_normalize(value) for value in row if value is not None}
                    for group in required_groups)
                for row in rows
            ):
                raise RuntimeError(
                    f"Required headers {required_headers} were not found in {path.name}."
                )

        required_column_count = minimum_columns or len(required_groups)
        if row_count < data_start_row or column_count < required_column_count:
            raise RuntimeError(
                f"Workbook contains no usable report rows: {path.name} "
                f"(worksheet rows={row_count}, columns={column_count}; "
                f"required start row={data_start_row}, minimum columns={required_column_count})"
            )

        data_rows = (
            iter(cached_rows[data_start_row - 1 :])
            if cached_rows is not None
            else sheet.iter_rows(
                min_row=data_start_row,
                max_row=row_count,
                values_only=True,
            )
        )
        expected = _normalize(first_column_value) if first_column_value is not None else None
        required_indexes = required_data_columns or []
        first_data_row = next(
            (
                row
                for row in data_rows
                if any(value not in (None, "") for value in row)
                and (expected is None or _normalize(row[0] if row else "") == expected)
                and all(index < len(row) and row[index] not in (None, "") for index in required_indexes)
            ),
            None,
        )
        if first_data_row is None:
            signature = []
            if first_column_value is not None:
                signature.append(f"column A = '{first_column_value}'")
            if required_indexes:
                signature.append(
                    "populated column(s) " + ", ".join(str(index + 1) for index in required_indexes)
                )
            detail = f" matching {', '.join(signature)}" if signature else ""
            raise RuntimeError(f"Workbook contains no populated report rows{detail}: {path.name}")
        if first_column_value is not None:
            actual = _normalize(first_data_row[0] if first_data_row else "")
            if actual != expected:
                raise RuntimeError(
                    f"Expected '{first_column_value}' in the first column of the first data row "
                    f"in {path.name}; found '{first_data_row[0]}'."
                )
        return row_count, column_count
    finally:
        workbook.close()
