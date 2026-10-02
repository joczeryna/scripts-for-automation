from __future__ import annotations

import csv
import json
import logging
import sys
import time
from pathlib import Path

from google_sheets_jso_load_preview_v2 import (
    comparable,
    final_export,
    google_client,
    google_retry,
    latest_export_directory,
    load_settings,
    padded,
    preview,
    read_csv,
    worksheet_by_gid,
)


LOG = logging.getLogger("deviation_statuses_google_controlled_write")
CONTROLLED_COLUMNS = 14


def write_csv(path: Path, rows: list[list[str]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8-sig", newline="") as handle:
        csv.writer(handle).writerows(rows)


def rows_equivalent(
    source: list[list[str]], destination: list[list[str]], headers: list[str]
) -> bool:
    if len(source) != len(destination):
        return False
    for source_row, destination_row in zip(source, destination):
        left = padded(source_row, CONTROLLED_COLUMNS)
        right = padded(destination_row, CONTROLLED_COLUMNS)
        for index in range(CONTROLLED_COLUMNS):
            header = headers[index] if index < len(headers) else ""
            if comparable(left[index], header) != comparable(right[index], header):
                return False
    return True


def copy_row_format_to_new_rows(book, worksheet, template_row: int, first_new_row: int, last_new_row: int) -> None:
    if first_new_row > last_new_row:
        return
    google_retry(
        "Copy Deviation Statuses row formatting",
        lambda: book.batch_update(
            {
                "requests": [
                    {
                        "copyPaste": {
                            "source": {
                                "sheetId": worksheet.id,
                                "startRowIndex": template_row - 1,
                                "endRowIndex": template_row,
                                "startColumnIndex": 0,
                                "endColumnIndex": CONTROLLED_COLUMNS,
                            },
                            "destination": {
                                "sheetId": worksheet.id,
                                "startRowIndex": first_new_row - 1,
                                "endRowIndex": last_new_row,
                                "startColumnIndex": 0,
                                "endColumnIndex": CONTROLLED_COLUMNS,
                            },
                            "pasteType": "PASTE_FORMAT",
                            "pasteOrientation": "NORMAL",
                        }
                    }
                ]
            }
        ),
    )


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    settings_path = Path(
        sys.argv[1] if len(sys.argv) > 1 else "google_sheets_preview_settings.json"
    ).resolve()
    gate_path = None
    if "--gate-file" in sys.argv:
        gate_index = sys.argv.index("--gate-file")
        if gate_index + 1 >= len(sys.argv):
            raise SystemExit("--gate-file requires a JSON path")
        gate_path = Path(sys.argv[gate_index + 1]).resolve()

    try:
        # Mandatory read-only gate. This revalidates workbook identity, headers,
        # Audit Id overlap, and the existing overlap row before any write.
        gate = (
            json.loads(gate_path.read_text(encoding="utf-8"))
            if gate_path
            else preview(settings_path)
        )
        deviation_gate = gate["deviationStatuses"]
        if not deviation_gate["overlapVerified"]:
            raise RuntimeError("Deviation Statuses overlap validation did not pass")

        base = settings_path.parent
        settings = load_settings(settings_path)
        export_dir = latest_export_directory(
            (base / settings.get("export_root", "exports")).resolve()
        )
        source_path = final_export(export_dir, "JSO_Deviation_Statuses")
        headers, source_rows = read_csv(source_path)
        if len(headers) != CONTROLLED_COLUMNS:
            raise RuntimeError(
                f"Deviation Statuses export must contain exactly {CONTROLLED_COLUMNS} columns"
            )
        if any(len(row) != CONTROLLED_COLUMNS for row in source_rows):
            raise RuntimeError("Deviation Statuses export contains a malformed data row")

        source_overlap_index = int(deviation_gate["sourceOverlapRow"]) - 2
        destination_overlap_row = int(deviation_gate["destinationOverlapRow"])
        write_rows = source_rows[source_overlap_index:]
        if not write_rows:
            raise RuntimeError("No verified Deviation Statuses overlap row was available")
        if len(write_rows) != int(deviation_gate["proposedRowsFromOverlap"]):
            raise RuntimeError("Deviation Statuses source changed after the validation gate")

        client = google_client(settings, base)
        book = google_retry(
            "Open JSO workbook for controlled write",
            lambda: client.open_by_key(settings["jso_spreadsheet_id"]),
        )
        tabs = settings["tabs"]
        gids = settings["tab_gids"]
        worksheet = worksheet_by_gid(
            book,
            gids["deviation_statuses"],
            tabs["deviation_statuses"],
        )

        last_existing_row = int(deviation_gate["lastDestinationRow"])
        existing_rows_from_overlap = last_existing_row - destination_overlap_row + 1
        live_overlap = google_retry(
            "Recheck complete Deviation Statuses overlap group immediately before write",
            lambda: worksheet.get(
                f"A{destination_overlap_row}:N{last_existing_row}"
            ),
        )
        if not rows_equivalent(
            write_rows[:existing_rows_from_overlap],
            live_overlap,
            headers,
        ):
            raise RuntimeError(
                "Deviation Statuses destination changed after validation; no writes were made"
            )

        # Capture the complete controlled destination range before writing.
        backup_values = google_retry(
            "Read Deviation Statuses pre-write backup",
            lambda: worksheet.get(f"A1:N{last_existing_row}"),
        )
        stamp = time.strftime("%Y-%m-%d_%H%M%S")
        backup_dir = base / "backups" / stamp
        write_csv(
            backup_dir / "Deviation_Statuses_A_to_N_before.csv",
            backup_values,
        )

        target_end_row = destination_overlap_row + len(write_rows) - 1
        if target_end_row > worksheet.row_count:
            google_retry(
                "Expand Deviation Statuses row capacity",
                lambda: worksheet.add_rows(target_end_row - worksheet.row_count),
            )

        first_new_row = last_existing_row + 1
        copy_row_format_to_new_rows(
            book,
            worksheet,
            destination_overlap_row,
            first_new_row,
            target_end_row,
        )

        # This is a values-only write. It does not write formulas or replace
        # formatting. Only A:N from the verified overlap row onward is touched.
        google_retry(
            "Write Deviation Statuses verified overlap and new rows",
            lambda: worksheet.update(
                values=write_rows,
                range_name=(
                    f"A{destination_overlap_row}:N{target_end_row}"
                ),
                value_input_option="USER_ENTERED",
            ),
        )

        verification = google_retry(
            "Verify Deviation Statuses controlled write",
            lambda: worksheet.get(
                f"A{destination_overlap_row}:N{target_end_row}"
            ),
        )
        verified = rows_equivalent(write_rows, verification, headers)
        if not verified:
            raise RuntimeError("Deviation Statuses post-write verification failed")

        overwritten_rows = min(len(write_rows), existing_rows_from_overlap)
        new_rows = max(0, len(write_rows) - existing_rows_from_overlap)
        result = {
            "mode": "CONTROLLED DEVIATION STATUSES GOOGLE SHEET WRITE",
            "sourceFile": str(source_path),
            "backupDirectory": str(backup_dir),
            "destination": {
                "spreadsheetId": settings["jso_spreadsheet_id"],
                "tab": worksheet.title,
                "gid": worksheet.id,
                "controlledColumns": "A:N",
                "firstOverlapRow": destination_overlap_row,
                "existingRowsOverwritten": overwritten_rows,
                "newRowsAppended": new_rows,
                "rangeWritten": (
                    f"A{destination_overlap_row}:N{target_end_row}"
                ),
                "valuesOnly": True,
                "newRowFormattingCopied": new_rows > 0,
                "postWriteVerified": verified,
            },
            "otherTabsModified": False,
            "leadershipModified": False,
            "allJsosModified": False,
        }
        LOG.info("DEVIATION_STATUSES_GOOGLE_WRITE_RESULT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Controlled Deviation Statuses Google Sheet write failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
