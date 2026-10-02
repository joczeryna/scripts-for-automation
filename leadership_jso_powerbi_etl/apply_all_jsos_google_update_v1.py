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


LOG = logging.getLogger("all_jsos_google_controlled_write")
CONTROLLED_COLUMNS = 15


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
            if comparable(left[index], headers[index]) != comparable(
                right[index], headers[index]
            ):
                return False
    return True


def copy_format(book, worksheet, template_row: int, first_row: int, last_row: int) -> None:
    if first_row > last_row:
        return
    google_retry(
        "Copy All JSOs formatting to appended rows",
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
                                "startRowIndex": first_row - 1,
                                "endRowIndex": last_row,
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
        gate = (
            json.loads(gate_path.read_text(encoding="utf-8"))
            if gate_path
            else preview(settings_path)
        )
        all_gate = gate["allJsos"]
        if not all_gate.get("fullRowOverlapVerified"):
            raise RuntimeError("All JSOs complete-row overlap validation did not pass")

        base = settings_path.parent
        settings = load_settings(settings_path)
        export_dir = latest_export_directory(
            (base / settings.get("export_root", "exports")).resolve()
        )
        source_path = final_export(export_dir, "JSO_All_Audits")
        headers, source_rows = read_csv(source_path)
        if len(headers) != CONTROLLED_COLUMNS:
            raise RuntimeError(
                f"All JSOs export must contain exactly {CONTROLLED_COLUMNS} columns"
            )
        if any(len(row) != CONTROLLED_COLUMNS for row in source_rows):
            raise RuntimeError("All JSOs export contains a malformed data row")

        source_overlap_index = int(all_gate["sourceOverlapRow"]) - 2
        destination_overlap_row = int(all_gate["lastDestinationRow"])
        source_overlap_row = source_rows[source_overlap_index]
        append_rows = source_rows[source_overlap_index + 1 :]
        if len(append_rows) != int(all_gate["proposedAppendRows"]):
            raise RuntimeError("All JSOs source changed after the validation gate")

        if not append_rows:
            result = {
                "mode": "CONTROLLED ALL JSOS GOOGLE SHEET APPEND",
                "sourceFile": str(source_path),
                "lastDestinationRow": destination_overlap_row,
                "lastAuditId": all_gate["lastAuditId"],
                "fullRowOverlapVerified": True,
                "newRowsAppended": 0,
                "writesPerformed": False,
                "reason": "No source rows exist after the verified overlap row.",
            }
            LOG.info("ALL_JSOS_GOOGLE_WRITE_RESULT=%s", json.dumps(result))
            return 0

        client = google_client(settings, base)
        book = google_retry(
            "Open JSO workbook for All JSOs append",
            lambda: client.open_by_key(settings["jso_spreadsheet_id"]),
        )
        tabs = settings["tabs"]
        gids = settings["tab_gids"]
        worksheet = worksheet_by_gid(
            book,
            gids["all_jsos"],
            tabs["all_jsos"],
        )

        live_overlap = google_retry(
            "Recheck complete All JSOs overlap row immediately before append",
            lambda: worksheet.get(
                f"A{destination_overlap_row}:O{destination_overlap_row}"
            ),
        )
        if not rows_equivalent([source_overlap_row], live_overlap, headers):
            raise RuntimeError(
                "All JSOs destination changed after validation; no writes were made"
            )

        backup_start_row = max(1, destination_overlap_row - 99)
        backup_values = google_retry(
            "Read All JSOs pre-append backup tail",
            lambda: worksheet.get(
                f"A{backup_start_row}:O{destination_overlap_row}"
            ),
        )
        stamp = time.strftime("%Y-%m-%d_%H%M%S")
        backup_dir = base / "backups" / stamp
        write_csv(backup_dir / "All_JSOs_tail_before.csv", backup_values)

        first_new_row = destination_overlap_row + 1
        last_new_row = destination_overlap_row + len(append_rows)
        if last_new_row > worksheet.row_count:
            google_retry(
                "Expand All JSOs row capacity",
                lambda: worksheet.add_rows(last_new_row - worksheet.row_count),
            )
        copy_format(
            book,
            worksheet,
            destination_overlap_row,
            first_new_row,
            last_new_row,
        )

        google_retry(
            "Append new All JSOs values",
            lambda: worksheet.update(
                values=append_rows,
                range_name=f"A{first_new_row}:O{last_new_row}",
                value_input_option="USER_ENTERED",
            ),
        )

        verification = google_retry(
            "Verify All JSOs appended rows",
            lambda: worksheet.get(f"A{first_new_row}:O{last_new_row}"),
        )
        verified = rows_equivalent(append_rows, verification, headers)
        if not verified:
            raise RuntimeError("All JSOs post-append verification failed")

        result = {
            "mode": "CONTROLLED ALL JSOS GOOGLE SHEET APPEND",
            "sourceFile": str(source_path),
            "backupDirectory": str(backup_dir),
            "destination": {
                "spreadsheetId": settings["jso_spreadsheet_id"],
                "tab": worksheet.title,
                "gid": worksheet.id,
                "controlledColumns": "A:O",
                "verifiedOverlapRow": destination_overlap_row,
                "newRowsAppended": len(append_rows),
                "rangeWritten": f"A{first_new_row}:O{last_new_row}",
                "valuesOnly": True,
                "newRowFormattingCopied": True,
                "postWriteVerified": verified,
            },
            "otherTabsModified": False,
            "leadershipModified": False,
            "deviationStatusesModified": False,
        }
        LOG.info("ALL_JSOS_GOOGLE_WRITE_RESULT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Controlled All JSOs Google Sheet append failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
