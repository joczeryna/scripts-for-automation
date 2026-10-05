from __future__ import annotations

import csv
import json
import logging
import re
import sys
import time
from datetime import date, datetime
from pathlib import Path
from typing import Any

import gspread
from gspread.exceptions import APIError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow


SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
LOG = logging.getLogger("leadership_jso_google_preview")


class PreviewError(RuntimeError):
    pass


def google_retry(label: str, action, attempts: int = 2):
    retryable = {429, 500, 502, 503, 504}
    for attempt in range(1, attempts + 1):
        try:
            return action()
        except APIError as error:
            status = getattr(getattr(error, "response", None), "status_code", None)
            if status not in retryable or attempt == attempts:
                raise
            delay = 5 * attempt
            LOG.warning(
                "%s returned Google API %s; retrying in %s seconds (%s/%s)",
                label,
                status,
                delay,
                attempt,
                attempts,
            )
            time.sleep(delay)


def load_settings(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def google_client(settings: dict[str, Any], base: Path) -> gspread.Client:
    oauth = settings["google_oauth"]
    credentials_path = (base / oauth["credentials_file"]).resolve()
    token_path = (base / oauth["authorized_user_file"]).resolve()
    credentials = None
    if token_path.exists():
        credentials = Credentials.from_authorized_user_file(str(token_path), SCOPES)
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(Request())
    if not credentials or not credentials.valid:
        if not credentials_path.exists():
            raise PreviewError(f"Google OAuth credentials not found: {credentials_path}")
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
        credentials = flow.run_local_server(port=0)
        token_path.write_text(credentials.to_json(), encoding="utf-8")
    client = gspread.authorize(credentials)
    timeout_seconds = int(settings.get("google_api_timeout_seconds", 60))
    if hasattr(client, "set_timeout"):
        client.set_timeout(timeout_seconds)
    return client


def latest_export_directory(root: Path) -> Path:
    candidates = sorted(
        (item for item in root.iterdir() if item.is_dir()),
        key=lambda item: item.name,
        reverse=True,
    )
    if not candidates:
        raise PreviewError(f"No timestamped export directories found in {root}")
    return candidates[0]


def final_export(folder: Path, prefix: str) -> Path:
    pattern = re.compile(rf"^{re.escape(prefix)}_\d{{4}}-\d{{2}}-\d{{2}}_\d{{6}}\.csv$")
    matches = [item for item in folder.glob(f"{prefix}_*.csv") if pattern.match(item.name)]
    if len(matches) != 1:
        raise PreviewError(
            f"Expected one final {prefix} CSV in {folder}; found {len(matches)}"
        )
    return matches[0]


def read_csv(path: Path) -> tuple[list[str], list[list[str]]]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.reader(handle)
        rows = list(reader)
    if len(rows) < 2:
        raise PreviewError(f"CSV contains no data rows: {path}")
    return [value.strip() for value in rows[0]], rows[1:]


def normalized(value: Any) -> str:
    return str(value or "").strip()


def find_header_row(values: list[list[str]], required: str, label: str) -> tuple[int, list[str]]:
    key = required.casefold()
    for index, row in enumerate(values[:20]):
        if any(normalized(cell).casefold() == key for cell in row):
            return index, [normalized(cell) for cell in row]
    raise PreviewError(f"{label} does not contain the required header {required!r}")


def header_index(headers: list[str], required: str, label: str) -> int:
    wanted = required.casefold()
    for index, value in enumerate(headers):
        if value.casefold() == wanted:
            return index
    raise PreviewError(f"{label} is missing required header {required!r}")


def last_data_row(values: list[list[str]], header_row: int, column: int, label: str) -> tuple[int, str]:
    for index in range(len(values) - 1, header_row, -1):
        value = normalized(values[index][column] if column < len(values[index]) else "")
        if value:
            return index, value
    raise PreviewError(f"No Audit Id was found in {label}")


def last_nonblank_sheet_cell(
    worksheet,
    column_letter: str,
    header_sheet_row: int,
    label: str,
    chunk_size: int = 5000,
) -> tuple[int, str]:
    """Return the last populated sheet row without requesting an entire column."""
    end_row = max(int(worksheet.row_count), header_sheet_row + 1)
    while end_row > header_sheet_row:
        start_row = max(header_sheet_row + 1, end_row - chunk_size + 1)
        values = google_retry(
            f"Read {label} Audit Id tail {start_row}:{end_row}",
            lambda start=start_row, end=end_row: worksheet.get(
                f"{column_letter}{start}:{column_letter}{end}"
            ),
        )
        for offset in range(len(values) - 1, -1, -1):
            value = normalized(values[offset][0] if values[offset] else "")
            if value:
                return start_row + offset, value
        end_row = start_row - 1
    raise PreviewError(f"No Audit Id was found in {label}")


def first_row_of_trailing_value_group(
    worksheet,
    column_letter: str,
    last_sheet_row: int,
    value: str,
    header_sheet_row: int,
    label: str,
    chunk_size: int = 1000,
) -> int:
    """Find the first row of the final contiguous Audit Id group."""
    group_start = last_sheet_row
    end_row = last_sheet_row
    while end_row > header_sheet_row:
        start_row = max(header_sheet_row + 1, end_row - chunk_size + 1)
        values = google_retry(
            f"Read {label} trailing Audit Id group {start_row}:{end_row}",
            lambda start=start_row, end=end_row: worksheet.get(
                f"{column_letter}{start}:{column_letter}{end}"
            ),
        )
        for offset in range(len(values) - 1, -1, -1):
            cell = normalized(values[offset][0] if values[offset] else "")
            sheet_row = start_row + offset
            if cell == value:
                group_start = sheet_row
                continue
            if sheet_row < group_start:
                return group_start
        if group_start > start_row:
            return group_start
        end_row = start_row - 1
    return group_start


def matching_positions(rows: list[list[str]], column: int, audit_id: str) -> list[int]:
    return [
        index
        for index, row in enumerate(rows)
        if column < len(row) and normalized(row[column]) == audit_id
    ]


def padded(row: list[str], width: int) -> list[str]:
    return [normalized(value) for value in (row[:width] + [""] * width)[:width]]


def rows_equivalent(
    source: list[list[str]],
    destination: list[list[str]],
    headers: list[str],
    width: int,
) -> bool:
    if len(source) != len(destination):
        return False
    for source_row, destination_row in zip(source, destination):
        left = padded(source_row, width)
        right = padded(destination_row, width)
        for index in range(width):
            header = headers[index] if index < len(headers) else ""
            if comparable(left[index], header) != comparable(right[index], header):
                return False
    return True


def comparable(value: Any, header: str) -> str:
    text = normalized(value)
    key = normalized(header).casefold()
    if not text:
        return ""
    if key == "date":
        date_text = text.split(" ", 1)[0]
        for pattern in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                return datetime.strptime(date_text, pattern).date().isoformat()
            except ValueError:
                pass
    if key in {"audit id", "auditor number", "worker number"} and text.isdigit():
        return str(int(text))
    if key in {"deviation flag", "jso flag"}:
        folded = text.casefold()
        if folded in {"true", "1", "yes"}:
            return "1"
        if folded in {"false", "0", "no"}:
            return "0"
        return folded
    return text


def one_column(values: list[list[str]]) -> list[str]:
    return [normalized(row[0]) if row else "" for row in values]


def tail_value_info(
    values: list[list[str]],
    start_sheet_row: int,
    header_sheet_row: int,
    label: str,
) -> tuple[int, str, int]:
    """Return last row/value and first row of its trailing contiguous group."""
    last_offset = None
    last_value = ""
    for offset in range(len(values) - 1, -1, -1):
        value = normalized(values[offset][0] if values[offset] else "")
        if value:
            last_offset = offset
            last_value = value
            break
    if last_offset is None:
        raise PreviewError(f"No Audit Id was found in {label}")

    group_offset = last_offset
    while group_offset > 0:
        previous = normalized(
            values[group_offset - 1][0] if values[group_offset - 1] else ""
        )
        if previous != last_value:
            break
        group_offset -= 1

    last_sheet_row = start_sheet_row + last_offset
    group_start_row = start_sheet_row + group_offset
    if group_start_row <= header_sheet_row:
        raise PreviewError(f"Invalid trailing Audit Id group in {label}")
    return last_sheet_row, last_value, group_start_row


def worksheet_by_gid(book, gid: int, expected_title: str):
    worksheets = google_retry(
        f"List worksheets for {book.title}",
        lambda: book.worksheets(),
    )
    for worksheet in worksheets:
        if int(worksheet.id) == int(gid):
            if worksheet.title != expected_title:
                raise PreviewError(
                    f"Sheet ID {gid} is named {worksheet.title!r}; "
                    f"expected {expected_title!r}"
                )
            return worksheet
    raise PreviewError(f"Worksheet ID {gid} was not found in {book.title}")


def worksheets_by_gid(book, expected: dict[int, str]) -> dict[int, Any]:
    worksheets = google_retry(
        f"List worksheets for {book.title}",
        lambda: book.worksheets(),
    )
    by_gid = {int(worksheet.id): worksheet for worksheet in worksheets}
    result = {}
    for gid, title in expected.items():
        worksheet = by_gid.get(int(gid))
        if worksheet is None:
            raise PreviewError(f"Worksheet ID {gid} was not found in {book.title}")
        if worksheet.title != title:
            raise PreviewError(
                f"Sheet ID {gid} is named {worksheet.title!r}; expected {title!r}"
            )
        result[int(gid)] = worksheet
    return result


def preview(settings_path: Path) -> dict[str, Any]:
    settings = load_settings(settings_path)
    base = settings_path.parent
    export_root = (base / settings.get("export_root", "exports")).resolve()
    export_dir = latest_export_directory(export_root)

    leadership_path = final_export(export_dir, "Leadership_JSOs")
    deviations_path = final_export(export_dir, "JSO_Deviation_Statuses")
    audits_path = final_export(export_dir, "JSO_All_Audits")

    leadership_headers, leadership_rows = read_csv(leadership_path)
    deviation_headers, deviation_rows = read_csv(deviations_path)
    audit_headers, audit_rows = read_csv(audits_path)
    audit_header = settings.get("audit_id_header", "Audit Id")
    source_all_id_col = header_index(audit_headers, audit_header, "All Audits export")
    source_dev_id_col = header_index(deviation_headers, audit_header, "Deviation export")

    client = google_client(settings, base)
    jso_book = google_retry(
        "Open JSO workbook",
        lambda: client.open_by_key(settings["jso_spreadsheet_id"]),
    )
    leadership_book = google_retry(
        "Open Leadership workbook",
        lambda: client.open_by_key(settings["leadership_spreadsheet_id"]),
    )
    tabs = settings["tabs"]
    gids = settings["tab_gids"]
    jso_worksheets = worksheets_by_gid(
        jso_book,
        {
            int(gids["all_jsos"]): tabs["all_jsos"],
            int(gids["deviation_statuses"]): tabs["deviation_statuses"],
        },
    )
    leadership_worksheets = worksheets_by_gid(
        leadership_book,
        {int(gids["leadership_import"]): tabs["leadership_import"]},
    )
    all_ws = jso_worksheets[int(gids["all_jsos"])]
    dev_ws = jso_worksheets[int(gids["deviation_statuses"])]
    import_ws = leadership_worksheets[int(gids["leadership_import"])]

    # One bounded batch obtains both headers and both Audit Id tails. This
    # replaces repeated large-sheet reads that previously caused multi-minute
    # 503 retry cycles.
    tail_size = int(settings.get("audit_tail_rows", 300))
    all_tail_start = max(2, int(all_ws.row_count) - tail_size + 1)
    dev_tail_start = max(2, int(dev_ws.row_count) - tail_size + 1)
    batch = google_retry(
        "Read JSO validation ranges",
        lambda: jso_book.values_batch_get(
            [
                f"'{tabs['all_jsos']}'!A1:Z20",
                f"'{tabs['deviation_statuses']}'!A1:N20",
                f"'{tabs['all_jsos']}'!D{all_tail_start}:D{all_ws.row_count}",
                f"'{tabs['deviation_statuses']}'!D{dev_tail_start}:D{dev_ws.row_count}",
            ]
        ),
    )
    value_ranges = batch.get("valueRanges", [])
    if len(value_ranges) != 4:
        raise PreviewError(
            f"Expected four Google Sheets validation ranges; received {len(value_ranges)}"
        )
    all_header_values = value_ranges[0].get("values", [])
    all_header_row, all_headers = find_header_row(
        all_header_values, audit_header, "All JSOs"
    )
    all_id_col = header_index(all_headers, audit_header, "All JSOs")
    if all_id_col != 3:
        raise PreviewError(
            f"All JSOs Audit Id moved from column D to column {all_id_col + 1}"
        )
    last_all_sheet_row, last_audit_id, _ = tail_value_info(
        value_ranges[2].get("values", []),
        all_tail_start,
        all_header_row + 1,
        "All JSOs",
    )
    last_all_index = last_all_sheet_row - 1

    source_all_matches = matching_positions(audit_rows, source_all_id_col, last_audit_id)
    if not source_all_matches:
        raise PreviewError(
            f"Last All JSOs Audit Id {last_audit_id} was not found in All Audits export"
        )
    destination_all_overlap_result = google_retry(
        "Read All JSOs complete overlap row",
        lambda: all_ws.get(f"A{last_all_index + 1}:O{last_all_index + 1}"),
    )
    destination_all_overlap_row = padded(
        destination_all_overlap_result[0] if destination_all_overlap_result else [],
        15,
    )
    comparable_destination_all = [
        comparable(value, audit_headers[index])
        for index, value in enumerate(destination_all_overlap_row)
    ]
    equivalent_source_positions = []
    for position in source_all_matches:
        source_candidate = padded(audit_rows[position], 15)
        comparable_candidate = [
            comparable(value, audit_headers[index])
            for index, value in enumerate(source_candidate)
        ]
        if comparable_candidate == comparable_destination_all:
            equivalent_source_positions.append(position)
    if not equivalent_source_positions:
        candidate_diagnostics = []
        for position in source_all_matches[:10]:
            source_candidate = padded(audit_rows[position], 15)
            differences = []
            for index in range(15):
                source_comparable = comparable(
                    source_candidate[index], audit_headers[index]
                )
                destination_comparable = comparable(
                    destination_all_overlap_row[index], audit_headers[index]
                )
                if source_comparable != destination_comparable:
                    differences.append(
                        {
                            "column": index + 1,
                            "header": audit_headers[index],
                            "sourceValue": source_candidate[index],
                            "destinationValue": destination_all_overlap_row[index],
                        }
                    )
            candidate_diagnostics.append(
                {
                    "sourceRow": position + 2,
                    "differences": differences,
                }
            )
        diagnostic = {
            "mode": "READ ONLY ALL JSOS OVERLAP DIAGNOSTIC",
            "auditId": last_audit_id,
            "destinationRow": last_all_index + 1,
            "candidateCount": len(source_all_matches),
            "candidates": candidate_diagnostics,
            "writesPerformed": False,
        }
        LOG.error(
            "LEADERSHIP_JSO_ALL_AUDITS_OVERLAP_DIAGNOSTIC=%s",
            json.dumps(diagnostic),
        )
        raise PreviewError(
            f"Last All JSOs row for Audit Id {last_audit_id} did not match any complete source row"
        )
    if len(equivalent_source_positions) != 1:
        raise PreviewError(
            f"Last All JSOs row for Audit Id {last_audit_id} matched "
            f"{len(equivalent_source_positions)} source rows; append position is ambiguous"
        )
    source_all_overlap = equivalent_source_positions[0]
    all_new_rows = audit_rows[source_all_overlap + 1 :]

    dev_header_values = value_ranges[1].get("values", [])
    dev_header_row, dev_destination_headers = find_header_row(
        dev_header_values, audit_header, "Deviation Statuses"
    )
    dev_destination_id_col = header_index(
        dev_destination_headers, audit_header, "Deviation Statuses"
    )
    if dev_destination_id_col != 3:
        raise PreviewError(
            "Deviation Statuses Audit Id is no longer in controlled column D"
        )
    last_dev_sheet_row, last_dev_audit_id, destination_dev_group_start = tail_value_info(
        value_ranges[3].get("values", []),
        dev_tail_start,
        dev_header_row + 1,
        "Deviation Statuses",
    )
    last_dev_index = last_dev_sheet_row - 1
    source_dev_matches = matching_positions(
        deviation_rows, source_dev_id_col, last_dev_audit_id
    )
    if not source_dev_matches:
        raise PreviewError(
            f"Last Deviation Statuses Audit Id {last_dev_audit_id} was not found in Deviation export"
        )
    # Start at the lowest/first row for the overlap Audit Id in both places.
    # An Audit Id can occupy multiple rows; selecting the final occurrence
    # would skip earlier records from the same Audit Id group.
    source_dev_overlap = source_dev_matches[0]
    destination_dev_overlap = destination_dev_group_start - 1
    destination_dev_group_rows = google_retry(
        "Read Deviation Statuses overlap group",
        lambda: dev_ws.get(
            f"A{destination_dev_group_start}:N{last_dev_sheet_row}"
        ),
    )
    destination_overlap_result = destination_dev_group_rows[:1]
    destination_overlap_row = (
        destination_overlap_result[0] if destination_overlap_result else []
    )
    source_overlap_row = padded(deviation_rows[source_dev_overlap], 14)
    destination_overlap_row = padded(destination_overlap_row, 14)
    comparable_source = [
        comparable(value, deviation_headers[index])
        for index, value in enumerate(source_overlap_row)
    ]
    comparable_destination = [
        comparable(value, deviation_headers[index])
        for index, value in enumerate(destination_overlap_row)
    ]
    formatting_differences = [
        {
            "column": index + 1,
            "header": deviation_headers[index],
            "sourceValue": source_overlap_row[index],
            "destinationValue": destination_overlap_row[index],
        }
        for index in range(14)
        if source_overlap_row[index] != destination_overlap_row[index]
        and comparable_source[index] == comparable_destination[index]
    ]
    dev_overlap_matches = comparable_source == comparable_destination
    if not dev_overlap_matches:
        differences = []
        for index, (source_value, destination_value) in enumerate(
            zip(source_overlap_row, destination_overlap_row)
        ):
            if comparable_source[index] != comparable_destination[index]:
                differences.append(
                    {
                        "column": index + 1,
                        "header": deviation_headers[index]
                        if index < len(deviation_headers)
                        else "",
                        "sourceValue": source_value,
                        "destinationValue": destination_value,
                    }
                )
        diagnostic = {
            "mode": "READ ONLY DEVIATION OVERLAP DIAGNOSTIC",
            "auditId": last_dev_audit_id,
            "sourceRow": source_dev_overlap + 2,
            "destinationRow": destination_dev_overlap + 1,
            "differences": differences,
            "writesPerformed": False,
        }
        LOG.error(
            "LEADERSHIP_JSO_DEVIATION_OVERLAP_DIAGNOSTIC=%s",
            json.dumps(diagnostic),
        )
        raise PreviewError(
            f"Deviation overlap row mismatch for Audit Id {last_dev_audit_id}; no writes are allowed"
        )
    deviation_write_rows = deviation_rows[source_dev_overlap:]
    existing_deviation_tail_rows = last_dev_index - destination_dev_overlap + 1
    proposed_deviation_new_rows = max(
        0,
        len(deviation_write_rows) - existing_deviation_tail_rows,
    )
    existing_deviation_rows_match = rows_equivalent(
        deviation_write_rows[:existing_deviation_tail_rows],
        destination_dev_group_rows,
        deviation_headers,
        14,
    )
    deviation_write_required = (
        proposed_deviation_new_rows > 0 or not existing_deviation_rows_match
    )

    grand_total_rows = sum(
        1
        for row in leadership_rows
        if any(normalized(value).casefold() == "grand total" for value in row)
    )
    leadership_clean = [
        row
        for row in leadership_rows
        if not any(normalized(value).casefold() == "grand total" for value in row)
    ]

    result = {
        "mode": "READ ONLY GOOGLE SHEETS LOAD PREVIEW",
        "exportDirectory": str(export_dir),
        "sourceFiles": {
            "leadership": str(leadership_path),
            "deviationStatuses": str(deviations_path),
            "allAudits": str(audits_path),
        },
        "leadership": {
            "sourceRows": len(leadership_rows),
            "grandTotalRowsExcluded": grand_total_rows,
            "proposedImportRows": len(leadership_clean),
            "sourceColumns": len(leadership_headers),
            "destinationTab": import_ws.title,
            "destinationGid": import_ws.id,
            "reportDateCell": "B1",
            "reportDateValue": date.today().isoformat(),
            "summaryTab": tabs["leadership_summary"],
            "summarySort": "Column E ascending",
        },
        "allJsos": {
            "lastDestinationRow": last_all_index + 1,
            "destinationGid": all_ws.id,
            "lastAuditId": last_audit_id,
            "sourceOverlapRow": source_all_overlap + 2,
            "fullRowOverlapVerified": True,
            "sourceRowsAfterOverlap": len(all_new_rows),
            "proposedAppendRows": len(all_new_rows),
            "sourceColumns": len(audit_headers),
        },
        "deviationStatuses": {
            "lastDestinationRow": last_dev_index + 1,
            "lastAuditId": last_dev_audit_id,
            "destinationGid": dev_ws.id,
            "sourceOverlapRow": source_dev_overlap + 2,
            "destinationOverlapRow": destination_dev_overlap + 1,
            "overlapVerified": True,
            "formattingDifferencesAccepted": formatting_differences,
            "proposedRowsFromOverlap": len(deviation_write_rows),
            "existingRowsFromOverlap": existing_deviation_tail_rows,
            "proposedOverwriteRows": min(
                len(deviation_write_rows),
                existing_deviation_tail_rows,
            ),
            "proposedNewRows": proposed_deviation_new_rows,
            "existingRowsMatchSource": existing_deviation_rows_match,
            "writeRequired": deviation_write_required,
            "columns": "A:N",
        },
        "writesPerformed": False,
    }
    LOG.info("LEADERSHIP_JSO_GOOGLE_LOAD_PREVIEW=%s", json.dumps(result))
    return result


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    settings_path = Path(sys.argv[1] if len(sys.argv) > 1 else "settings.json").resolve()
    try:
        preview(settings_path)
        return 0
    except Exception as error:
        LOG.exception("Google Sheets load preview failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
