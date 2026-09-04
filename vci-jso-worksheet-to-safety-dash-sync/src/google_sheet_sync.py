import csv
import io
import re
import time
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import BrowserContext, Page, sync_playwright


PASTE_BATCH_SIZE = 500
PASTE_WAIT_MS = 1_250
FINAL_SAVE_WAIT_MS = 12_000
VERIFICATION_ATTEMPTS = 12
VERIFICATION_WAIT_MS = 5_000


def column_letter_to_number(column_letter: str) -> int:
    result = 0
    for character in column_letter.upper():
        if not character.isalpha():
            raise ValueError(f"Invalid column letter: {column_letter}")
        result = result * 26 + ord(character) - ord("A") + 1
    return result


def normalize_value(value: object) -> str:
    if value is None:
        return ""

    # A tab or line break inside one source value can cause Google
    # Sheets to interpret an entire clipboard batch as a single cell.
    # Flatten those characters while preserving the displayed text.
    cleaned_value = str(value).replace("\r\n", "\n").replace(
        "\r",
        "\n",
    )
    cleaned_value = cleaned_value.replace("\t", " ")
    cleaned_value = re.sub(r"\s*\n\s*", " ", cleaned_value)
    return cleaned_value


def normalize_rows(
    rows: list[list[str]],
    expected_column_count: int,
) -> list[list[str]]:
    normalized_rows = []

    for row in rows:
        normalized_row = [
            normalize_value(value)
            for value in row[:expected_column_count]
        ]
        normalized_row.extend(
            [""] * (expected_column_count - len(normalized_row))
        )
        normalized_rows.append(normalized_row)

    while normalized_rows and not any(normalized_rows[-1]):
        normalized_rows.pop()

    return normalized_rows


def rows_to_tsv(rows: list[list[str]]) -> str:
    # Values have already been made clipboard-safe by normalize_value.
    # Avoid CSV quoting because Google Sheets can treat a quoted batch
    # containing multiline data as one cell.
    return "\n".join(
        "\t".join(normalize_value(value) for value in row)
        for row in rows
    )


def setup_google_login(
    browser_profile_path: Path,
    spreadsheet_id: str,
) -> None:
    browser_profile_path.mkdir(parents=True, exist_ok=True)
    spreadsheet_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/edit"
    )

    with sync_playwright() as playwright:
        browser_context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(browser_profile_path),
            channel="chrome",
            headless=False,
            accept_downloads=True,
        )
        page = (
            browser_context.pages[0]
            if browser_context.pages
            else browser_context.new_page()
        )
        page.goto(
            spreadsheet_url,
            wait_until="domcontentloaded",
            timeout=120_000,
        )

        print()
        print("Google Chrome is now open for Google Sheets.")
        print("Sign in with the authorized VCI Google account.")
        print("Confirm that the spreadsheet opens and is editable.")
        print(
            "Return to this terminal and press Enter only after "
            "the spreadsheet is visible."
        )
        input()
        browser_context.close()


def read_worksheet_values(
    browser_context: BrowserContext,
    spreadsheet_id: str,
    worksheet_name: str,
    start_column: str,
    end_column: str,
    data_start_row: int,
) -> list[list[str]]:
    source_range = f"{start_column}{data_start_row}:{end_column}"
    cache_buster = int(time.time() * 1000)
    export_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/gviz/tq"
        "?tqx=out:csv"
        f"&sheet={quote(worksheet_name)}"
        f"&range={quote(source_range)}"
        "&headers=0"
        f"&cache_bust={cache_buster}"
    )

    response = browser_context.request.get(export_url, timeout=180_000)

    if not response.ok:
        raise RuntimeError(
            f"Could not read worksheet '{worksheet_name}'. "
            f"Google returned HTTP {response.status}."
        )

    response_text = response.text()
    if (
        "<html" in response_text.lower()
        or "accounts.google.com" in response_text.lower()
    ):
        raise RuntimeError(
            "The Google session is not authorized to read "
            f"worksheet '{worksheet_name}'."
        )

    parsed_rows = list(csv.reader(io.StringIO(response_text)))
    expected_column_count = (
        column_letter_to_number(end_column)
        - column_letter_to_number(start_column)
        + 1
    )
    return normalize_rows(parsed_rows, expected_column_count)


def get_name_box(page: Page):
    selectors = [
        "#t-name-box",
        "input[aria-label='Name box']",
        "input[aria-label*='Name box']",
    ]

    for selector in selectors:
        locator = page.locator(selector).first
        try:
            if locator.is_visible(timeout=2_000):
                return locator
        except Exception:
            continue

    raise RuntimeError("Could not find the Google Sheets name box.")


def select_range(page: Page, cell_range: str) -> None:
    name_box = get_name_box(page)
    name_box.click()
    name_box.fill(cell_range)
    name_box.press("Enter")
    page.wait_for_timeout(400)


def open_destination_worksheet(
    page: Page,
    spreadsheet_id: str,
    worksheet_name: str,
) -> None:
    spreadsheet_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/edit"
    )
    page.goto(
        spreadsheet_url,
        wait_until="domcontentloaded",
        timeout=180_000,
    )
    page.wait_for_timeout(5_000)

    worksheet_tab = page.get_by_text(worksheet_name, exact=True).last
    try:
        worksheet_tab.wait_for(state="visible", timeout=60_000)
        worksheet_tab.click()
        page.wait_for_timeout(2_000)
    except Exception as error:
        raise RuntimeError(
            f"Could not select destination worksheet "
            f"'{worksheet_name}'. Confirm the tab exists."
        ) from error


def parse_google_sheet_date(
    value: str,
) -> date:
    """Parse common Google Sheets date display formats."""
    cleaned_value = value.strip()

    supported_formats = [
        "%m/%d/%Y",
        "%m/%d/%y",
        "%Y-%m-%d",
        "%m/%d/%Y %H:%M:%S",
    ]

    for date_format in supported_formats:
        try:
            return datetime.strptime(
                cleaned_value,
                date_format,
            ).date()
        except ValueError:
            continue

    raise ValueError(
        f"Google Sheets returned an unrecognized date: {value}"
    )


def update_and_verify_run_date(
    browser_context: BrowserContext,
    page: Page,
    spreadsheet_id: str,
    worksheet_name: str,
    cell_reference: str,
    run_date: date,
) -> None:
    """Write the current date as a value and verify the result."""
    cell_match = re.fullmatch(
        r"([A-Za-z]+)(\d+)",
        cell_reference.strip(),
    )

    if cell_match is None:
        raise ValueError(
            f"Invalid run-date cell reference: {cell_reference}"
        )

    column_letter = cell_match.group(1).upper()
    row_number = int(cell_match.group(2))
    date_text = run_date.strftime("%m/%d/%Y")

    open_destination_worksheet(
        page=page,
        spreadsheet_id=spreadsheet_id,
        worksheet_name=worksheet_name,
    )

    select_range(
        page=page,
        cell_range=f"{column_letter}{row_number}",
    )

    page.evaluate(
        """
        async (text) => {
            await navigator.clipboard.writeText(text);
        }
        """,
        date_text,
    )

    page.keyboard.press(
        "Control+Shift+V"
    )
    page.wait_for_timeout(2_000)

    last_value = ""

    for attempt in range(1, 7):
        returned_rows = read_worksheet_values(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name,
            start_column=column_letter,
            end_column=column_letter,
            data_start_row=row_number,
        )

        if returned_rows:
            last_value = returned_rows[0][0]

            try:
                if parse_google_sheet_date(last_value) == run_date:
                    print(
                        f"Updated and verified "
                        f"'{worksheet_name}'!{cell_reference} "
                        f"as {date_text}."
                    )
                    return
            except ValueError:
                pass

        if attempt < 6:
            time.sleep(2)

    raise RuntimeError(
        f"Run-date verification failed for "
        f"'{worksheet_name}'!{cell_reference}. "
        f"Expected {date_text}; found '{last_value}'."
    )


def apply_run_date_updates(
    browser_context: BrowserContext,
    page: Page,
    job_config: dict,
) -> None:
    """Apply optional configured run-date values."""
    updates = job_config.get(
        "run_date_updates",
        [],
    )

    if not updates:
        return

    run_date = datetime.now().date()
    spreadsheet_ids = {
        "source": job_config["source_spreadsheet_id"],
        "destination": job_config["destination_spreadsheet_id"],
    }

    print()
    print(
        f"Applying run date {run_date.strftime('%m/%d/%Y')}..."
    )

    for update in updates:
        spreadsheet_role = str(
            update["spreadsheet"]
        ).strip().lower()

        if spreadsheet_role not in spreadsheet_ids:
            raise ValueError(
                "Run-date spreadsheet must be 'source' or "
                f"'destination'. Received: {spreadsheet_role}"
            )

        update_and_verify_run_date(
            browser_context=browser_context,
            page=page,
            spreadsheet_id=spreadsheet_ids[spreadsheet_role],
            worksheet_name=update["worksheet"],
            cell_reference=update["cell"],
            run_date=run_date,
        )


def clear_destination_range(
    page: Page,
    start_column: str,
    end_column: str,
    data_start_row: int,
) -> None:
    # Open-ended range prevents the range-limit warning.
    destination_range = f"{start_column}{data_start_row}:{end_column}"
    select_range(page, destination_range)
    page.keyboard.press("Delete")
    page.wait_for_timeout(1_000)


def confirm_google_sheet_expansion_dialog(
    page: Page,
) -> bool:
    """
    Confirm the Google Sheets grid-expansion dialog that can appear
    when a values-only paste extends beyond the worksheet's current
    row count.
    """
    page.wait_for_timeout(500)

    modal_background = page.locator(
        ".modal-dialog-bg"
    ).first

    try:
        if not modal_background.is_visible(
            timeout=1_000
        ):
            return False
    except Exception:
        return False

    ok_candidates = [
        page.get_by_role(
            "button",
            name=re.compile(
                r"^OK$",
                re.IGNORECASE,
            ),
        ),
        page.locator(
            ".modal-dialog button"
        ).filter(
            has_text=re.compile(
                r"^\s*OK\s*$",
                re.IGNORECASE,
            )
        ),
        page.locator(
            ".modal-dialog"
        ).get_by_text(
            re.compile(
                r"^\s*OK\s*$",
                re.IGNORECASE,
            ),
            exact=True,
        ),
    ]

    for candidate in ok_candidates:
        try:
            if candidate.count() == 0:
                continue

            button = candidate.last

            if button.is_visible():
                button.click(
                    force=True
                )
                page.wait_for_timeout(1_500)
                print(
                    "  Confirmed Google Sheets row expansion."
                )
                return True
        except Exception:
            continue

    raise RuntimeError(
        "Google Sheets displayed a modal dialog during paste, "
        "but its OK confirmation button could not be found."
    )


def paste_values_only_in_batches(
    page: Page,
    start_column: str,
    data_start_row: int,
    rows: list[list[str]],
) -> None:
    if not rows:
        raise ValueError(
            "The source range contains no populated rows. "
            "The destination was not intentionally populated."
        )

    total_rows = len(rows)
    total_batches = (total_rows + PASTE_BATCH_SIZE - 1) // PASTE_BATCH_SIZE

    for batch_number, offset in enumerate(
        range(0, total_rows, PASTE_BATCH_SIZE),
        start=1,
    ):
        batch_rows = rows[offset : offset + PASTE_BATCH_SIZE]
        starting_row = data_start_row + offset
        starting_cell = f"{start_column}{starting_row}"

        # Exit any lingering cell-edit or paste-processing state before
        # selecting the next batch destination.
        page.keyboard.press("Escape")
        page.wait_for_timeout(250)

        select_range(page, starting_cell)
        clipboard_text = rows_to_tsv(batch_rows)
        page.evaluate(
            """
            async (text) => {
                await navigator.clipboard.writeText(text);
            }
            """,
            clipboard_text,
        )

        # Explicit Paste values only.
        page.keyboard.press("Control+Shift+V")
        page.wait_for_timeout(500)

        confirm_google_sheet_expansion_dialog(
            page
        )

        page.wait_for_timeout(PASTE_WAIT_MS)

        # Some large grid-expansion prompts render only after Google
        # Sheets has processed the initial clipboard payload.
        confirm_google_sheet_expansion_dialog(
            page
        )

        print(
            f"  Pasted batch {batch_number} of {total_batches} "
            f"({len(batch_rows)} rows)."
        )


def compare_rows(
    expected_rows: list[list[str]],
    actual_rows: list[list[str]],
    worksheet_name: str,
) -> None:
    if len(actual_rows) != len(expected_rows):
        raise RuntimeError(
            f"Verification failed for '{worksheet_name}'. "
            f"Expected {len(expected_rows)} rows but found "
            f"{len(actual_rows)} rows."
        )

    for row_index, (expected_row, actual_row) in enumerate(
        zip(expected_rows, actual_rows),
        start=2,
    ):
        if expected_row != actual_row:
            differing_columns = [
                column_index
                for column_index, (expected_value, actual_value)
                in enumerate(zip(expected_row, actual_row), start=1)
                if expected_value != actual_value
            ]
            raise RuntimeError(
                f"Verification failed for '{worksheet_name}' "
                f"at destination row {row_index}. Differing relative "
                "column positions: "
                + ", ".join(str(column) for column in differing_columns)
            )


def verify_with_retries(
    browser_context: BrowserContext,
    destination_spreadsheet_id: str,
    destination_worksheet: str,
    start_column: str,
    end_column: str,
    data_start_row: int,
    expected_rows: list[list[str]],
) -> None:
    last_error = None

    for attempt in range(1, VERIFICATION_ATTEMPTS + 1):
        actual_rows = read_worksheet_values(
            browser_context=browser_context,
            spreadsheet_id=destination_spreadsheet_id,
            worksheet_name=destination_worksheet,
            start_column=start_column,
            end_column=end_column,
            data_start_row=data_start_row,
        )

        try:
            compare_rows(expected_rows, actual_rows, destination_worksheet)
            return
        except RuntimeError as error:
            last_error = error

            if attempt < VERIFICATION_ATTEMPTS:
                print(
                    f"  Verification attempt {attempt} is not ready; "
                    "waiting for Google Sheets to finish saving..."
                )
                time.sleep(VERIFICATION_WAIT_MS / 1000)

    raise last_error


def sync_worksheet(
    browser_context: BrowserContext,
    destination_page: Page,
    source_spreadsheet_id: str,
    destination_spreadsheet_id: str,
    mapping: dict,
    data_start_row: int,
) -> dict:
    source_worksheet = mapping["source_worksheet"]
    destination_worksheet = mapping["destination_worksheet"]
    start_column = mapping["start_column"].upper()
    end_column = mapping["end_column"].upper()

    print()
    print(f"Reading source worksheet '{source_worksheet}'...")
    source_rows = read_worksheet_values(
        browser_context,
        source_spreadsheet_id,
        source_worksheet,
        start_column,
        end_column,
        data_start_row,
    )

    if not source_rows:
        raise RuntimeError(
            f"Source worksheet '{source_worksheet}' returned no "
            "populated rows. The destination was not changed."
        )

    print(
        f"Validated {len(source_rows)} source rows from "
        f"'{source_worksheet}'."
    )

    open_destination_worksheet(
        destination_page,
        destination_spreadsheet_id,
        destination_worksheet,
    )

    print(
        f"Clearing destination range {start_column}{data_start_row}:"
        f"{end_column} on '{destination_worksheet}'..."
    )
    clear_destination_range(
        destination_page,
        start_column,
        end_column,
        data_start_row,
    )

    print(f"Pasting values only into '{destination_worksheet}'...")
    paste_values_only_in_batches(
        destination_page,
        start_column,
        data_start_row,
        source_rows,
    )

    print("Waiting for Google Sheets to save the completed range...")
    destination_page.wait_for_timeout(FINAL_SAVE_WAIT_MS)

    print(
        f"Completed values-only paste to "
        f"'{destination_worksheet}': "
        f"{len(source_rows)} rows."
    )

    return {
        "source_worksheet": source_worksheet,
        "destination_worksheet": destination_worksheet,
        "rows_written": len(source_rows),
        "range": f"{start_column}{data_start_row}:{end_column}",
    }


def run_google_sheet_sync(
    job_config: dict,
    browser_profile_path: Path,
) -> list[dict]:
    browser_profile_path.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser_context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(browser_profile_path),
            channel="chrome",
            headless=False,
            permissions=["clipboard-read", "clipboard-write"],
        )

        try:
            existing_pages = list(browser_context.pages)

            destination_page = (
                existing_pages[0]
                if existing_pages
                else browser_context.new_page()
            )

            # Persistent Chrome profiles restore tabs from earlier runs.
            # Close every extra restored tab so an old or blank report
            # cannot be mistaken for the report selected by this job.
            for stale_page in existing_pages[1:]:
                try:
                    stale_page.close()
                except Exception:
                    pass

            source_spreadsheet_url = (
                "https://docs.google.com/spreadsheets/d/"
                f"{job_config['source_spreadsheet_id']}/edit"
            )

            print(
                "Opening the configured source spreadsheet..."
            )
            destination_page.goto(
                source_spreadsheet_url,
                wait_until="domcontentloaded",
                timeout=180_000,
            )
            destination_page.wait_for_timeout(5_000)

            if job_config["source_spreadsheet_id"] not in destination_page.url:
                raise RuntimeError(
                    "The configured source spreadsheet did not open. "
                    f"Current URL: {destination_page.url}"
                )

            results = []

            apply_run_date_updates(
                browser_context=browser_context,
                page=destination_page,
                job_config=job_config,
            )

            for mapping in job_config["mappings"]:
                results.append(
                    sync_worksheet(
                        browser_context=browser_context,
                        destination_page=destination_page,
                        source_spreadsheet_id=(
                            job_config["source_spreadsheet_id"]
                        ),
                        destination_spreadsheet_id=(
                            job_config["destination_spreadsheet_id"]
                        ),
                        mapping=mapping,
                        data_start_row=job_config.get("data_start_row", 2),
                    )
                )

            return results
        finally:
            browser_context.close()
