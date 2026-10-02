from __future__ import annotations

import json
import logging
import sys
from pathlib import Path

from google_sheets_jso_load_preview_v2 import (
    google_client,
    google_retry,
    load_settings,
    worksheet_by_gid,
)


LOG = logging.getLogger("leadership_google_layout")


def compact_rows(values: list[list[str]]) -> list[dict]:
    return [
        {"row": index + 1, "values": row}
        for index, row in enumerate(values)
        if any(str(value).strip() for value in row)
    ]


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    settings_path = Path(
        sys.argv[1] if len(sys.argv) > 1 else "google_sheets_preview_settings.json"
    ).resolve()
    try:
        settings = load_settings(settings_path)
        client = google_client(settings, settings_path.parent)
        book = google_retry(
            "Open Leadership workbook",
            lambda: client.open_by_key(settings["leadership_spreadsheet_id"]),
        )
        tabs = settings["tabs"]
        gids = settings["tab_gids"]
        import_ws = worksheet_by_gid(
            book, gids["leadership_import"], tabs["leadership_import"]
        )
        summary_ws = google_retry(
            "Open Leadership Summary tab",
            lambda: book.worksheet(tabs["leadership_summary"]),
        )
        ranges = google_retry(
            "Read Leadership layout",
            lambda: book.values_batch_get(
                [
                    f"'{import_ws.title}'!A1:J20",
                    f"'{summary_ws.title}'!A1:J20",
                ]
            ),
        ).get("valueRanges", [])
        if len(ranges) != 2:
            raise RuntimeError(f"Expected two layout ranges; received {len(ranges)}")
        result = {
            "mode": "READ ONLY LEADERSHIP GOOGLE SHEET LAYOUT INSPECTION",
            "dataImport": {
                "gid": import_ws.id,
                "nonblankRows1To20": compact_rows(ranges[0].get("values", [])),
                "requiredReportDateCell": "B2",
            },
            "summary": {
                "gid": summary_ws.id,
                "nonblankRows1To20": compact_rows(ranges[1].get("values", [])),
                "requiredSort": "Column E ascending",
            },
            "writesPerformed": False,
        }
        LOG.info("LEADERSHIP_GOOGLE_LAYOUT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Leadership Google Sheet layout inspection failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
