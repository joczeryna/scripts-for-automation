import csv
import re
from pathlib import Path


def normalize_header(
    value: object,
) -> str:
    """
    Normalize a CSV header for reliable matching, including removal
    of the UTF-8 BOM sometimes attached to the Vehicle header.
    """
    if value is None:
        return ""

    return re.sub(
        r"[^a-z0-9]+",
        "",
        str(value)
        .replace("\ufeff", "")
        .strip()
        .lower(),
    )


def clean_value(
    value: object,
) -> str:
    """
    Convert a CSV value into clean text.
    """
    if value is None:
        return ""

    return str(value).strip()


def process_vehicle_csv(
    csv_path: Path,
    column_mapping: dict[str, str],
) -> dict:
    """
    Read the downloaded Lytx CSV and map:

    Vehicle -> Truck #
    Group   -> Division - Data
    Status  -> Status - Data
    """
    csv_path = Path(
        csv_path
    )

    if not csv_path.exists():
        raise FileNotFoundError(
            f"The downloaded CSV was not found: {csv_path}"
        )

    if csv_path.stat().st_size == 0:
        raise ValueError(
            "The downloaded Lytx CSV is empty."
        )

    with csv_path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as csv_file:
        reader = csv.DictReader(
            csv_file
        )

        if not reader.fieldnames:
            raise ValueError(
                "The downloaded CSV does not contain a header row."
            )

        normalized_source_columns = {}

        for source_column in reader.fieldnames:
            normalized_name = normalize_header(
                source_column
            )

            if not normalized_name:
                continue

            if normalized_name in normalized_source_columns:
                raise ValueError(
                    "The downloaded CSV contains duplicate or "
                    f"ambiguous headers: {source_column}"
                )

            normalized_source_columns[
                normalized_name
            ] = source_column

        resolved_mapping = {}
        missing_source_columns = []

        for (
            configured_source,
            destination_header,
        ) in column_mapping.items():
            normalized_source = normalize_header(
                configured_source
            )

            actual_source_column = (
                normalized_source_columns.get(
                    normalized_source
                )
            )

            if actual_source_column is None:
                missing_source_columns.append(
                    configured_source
                )
            else:
                resolved_mapping[
                    actual_source_column
                ] = destination_header

        if missing_source_columns:
            available_columns = ", ".join(
                str(column)
                for column in reader.fieldnames
            )

            raise ValueError(
                "The following required Lytx CSV columns were not "
                f"found: {', '.join(missing_source_columns)}. "
                f"Available columns: {available_columns}"
            )

        destination_rows = []
        skipped_blank_rows = 0

        vehicle_source_column = (
            normalized_source_columns[
                normalize_header(
                    "Vehicle"
                )
            ]
        )

        for source_row_number, source_row in enumerate(
            reader,
            start=2,
        ):
            if not any(
                clean_value(value)
                for value in source_row.values()
            ):
                skipped_blank_rows += 1
                continue

            vehicle_value = clean_value(
                source_row.get(
                    vehicle_source_column
                )
            )

            if not vehicle_value:
                raise ValueError(
                    "A populated Lytx CSV row does not contain a "
                    f"Vehicle value. Source row: {source_row_number}"
                )

            destination_row = {}

            for (
                actual_source_column,
                destination_header,
            ) in resolved_mapping.items():
                destination_row[
                    destination_header
                ] = clean_value(
                    source_row.get(
                        actual_source_column
                    )
                )

            destination_rows.append(
                destination_row
            )

    if not destination_rows:
        raise ValueError(
            "The Lytx CSV does not contain any populated vehicle "
            "records."
        )

    return {
        "target_headers": list(
            column_mapping.values()
        ),
        "rows": destination_rows,
        "row_count": len(
            destination_rows
        ),
        "skipped_blank_rows": skipped_blank_rows,
        "source_columns": list(
            reader.fieldnames
        ),
    }