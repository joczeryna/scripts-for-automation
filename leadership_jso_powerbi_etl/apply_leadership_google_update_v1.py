from __future__ import annotations

import csv
import json
import logging
import sys
import time
from datetime import date
from pathlib import Path

from google_sheets_jso_load_preview_v2 import (
    comparable,
    final_export,
    google_client,
    google_retry,
    latest_export_directory,
    load_settings,
    read_csv,
    worksheet_by_gid,
)


LOG = logging.getLogger("leadership_google_controlled_write")


def write_csv(path: Path, rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle).writerows(rows)


def padded(row: list[str], width: int) -> list[str]:
    return [str(value or "").strip() for value in (row + [""] * width)[:width]]


def rows_equivalent(
    source: list[list[str]], destination: list[list[str]], headers: list[str]
) -> bool:
    if len(source) != len(destination):
        return False
    for source_row, destination_row in zip(source, destination):
        left = padded(source_row, len(headers))
        right = padded(destination_row, len(headers))
        for index, header in enumerate(headers):
            if comparable(left[index], header) != comparable(right[index], header):
                return False
    return True


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    settings_path = Path(
        sys.argv[1] if len(sys.argv) > 1 else "google_sheets_preview_settings.json"
    ).resolve()
    try:
        settings = load_settings(settings_path)

        base = settings_path.parent
        export_dir = latest_export_directory((base / settings.get("export_root", "exports")).resolve())
        leadership_path = final_export(export_dir, "Leadership_JSOs")
        headers, rows = read_csv(leadership_path)
        rows = [
            row
            for row in rows
            if not any(str(value).strip().casefold() == "grand total" for value in row)
        ]
        if len(headers) != 7 or any(len(row) != 7 for row in rows):
            raise RuntimeError("Leadership export must contain exactly seven columns")

        client = google_client(settings, base)
        book = google_retry(
            "Open Leadership workbook",
            lambda: client.open_by_key(settings["leadership_spreadsheet_id"]),
        )
        tabs = settings["tabs"]
        gids = settings["tab_gids"]
        import_ws = worksheet_by_gid(
            book, gids["leadership_import"], tabs["leadership_import"]
        )
        before = google_retry(
            "Read Leadership pre-write backup",
            lambda: book.values_batch_get(
                [
                    f"'{import_ws.title}'!A1:J",
                ],
                params={"valueRenderOption": "FORMULA"},
            ),
        ).get("valueRanges", [])
        if len(before) != 1:
            raise RuntimeError("Could not capture the Leadership pre-write backup range")
        before_import = before[0].get("values", [])
        stamp = time.strftime("%Y-%m-%d_%H%M%S")
        backup_dir = base / "backups" / stamp
        write_csv(backup_dir / "Leadership_Data_Import_before.csv", before_import)

        report_date = f"{date.today().month}/{date.today().day}/{date.today().year}"
        output = [headers] + rows
        end_row = len(output) + 1
        preservation_end_row = max(len(before_import), end_row)
        preserved_hj_before = [
            padded(before_import[index][7:10] if index < len(before_import) else [], 3)
            for index in range(preservation_end_row)
        ]

        # Values-only operations preserve all existing formatting. Columns H:J
        # are never cleared or written.
        google_retry(
            "Clear old Leadership import values",
            lambda: import_ws.batch_clear(["A2:G"]),
        )
        google_retry(
            "Write Leadership import values",
            lambda: import_ws.update(
                values=output,
                range_name=f"A2:G{end_row}",
                value_input_option="USER_ENTERED",
            ),
        )
        google_retry(
            "Write Leadership report date",
            lambda: import_ws.update(
                values=[[report_date]],
                range_name="B1",
                value_input_option="USER_ENTERED",
            ),
        )

        verification = google_retry(
            "Verify Leadership update",
            lambda: book.values_batch_get(
                [
                    f"'{import_ws.title}'!A2:G{end_row}",
                    f"'{import_ws.title}'!B1",
                ]
            ),
        ).get("valueRanges", [])
        if len(verification) != 2:
            raise RuntimeError("Leadership post-write verification ranges were incomplete")
        written_rows = verification[0].get("values", [])
        written_date = verification[1].get("values", [[""]])[0][0]
        hj_formula_values = google_retry(
            "Verify Leadership H:J formulas and values",
            lambda: book.values_batch_get(
                [f"'{import_ws.title}'!H1:J{preservation_end_row}"],
                params={"valueRenderOption": "FORMULA"},
            ),
        ).get("valueRanges", [])
        if len(hj_formula_values) != 1:
            raise RuntimeError("Leadership H:J formula verification range was incomplete")
        hj_after_rows = hj_formula_values[0].get("values", [])
        preserved_hj_after = [
            padded(hj_after_rows[index] if index < len(hj_after_rows) else [], 3)
            for index in range(preservation_end_row)
        ]
        result = {
            "mode": "CONTROLLED LEADERSHIP DATA IMPORT WRITE",
            "backupDirectory": str(backup_dir),
            "sourceFile": str(leadership_path),
            "dataImport": {
                "gid": import_ws.id,
                "headersWritten": 7,
                "dataRowsWritten": len(rows),
                "rangeWritten": f"A2:G{end_row}",
                "reportDateCell": "B1",
                "reportDateWritten": written_date,
                "sourceValuesVerified": rows_equivalent(output, written_rows, headers),
                "columnsHJPreserved": preserved_hj_before == preserved_hj_after,
                "formattingChanged": False,
            },
            "allowedWriteRanges": ["Data Import!A2:G", "Data Import!B1"],
            "otherRangesWritten": 0,
            "summaryModified": False,
            "outOfScope": ["Summary", "All JSOs", "Deviation Statuses"],
        }
        if not result["dataImport"]["sourceValuesVerified"]:
            raise RuntimeError("Leadership post-write source-value verification failed")
        if not result["dataImport"]["columnsHJPreserved"]:
            raise RuntimeError("Leadership columns H:J changed unexpectedly")
        LOG.info("LEADERSHIP_GOOGLE_WRITE_RESULT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Controlled Leadership Google Sheet write failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
