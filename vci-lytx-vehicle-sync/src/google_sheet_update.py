import csv
import io
import re
import time
from pathlib import Path
from urllib.parse import quote

from playwright.sync_api import (
    BrowserContext,
    Page,
    sync_playwright,
)


def column_number_to_letter(
    column_number: int,
) -> str:
    """
    Convert a one-based column number to an Excel/Google Sheets
    column letter.
    """
    if column_number < 1:
        raise ValueError(
            "Column numbers must be greater than zero."
        )

    letters = ""

    while column_number:
        column_number, remainder = divmod(
            column_number - 1,
            26,
        )
        letters = chr(
            65 + remainder
        ) + letters

    return letters


def normalize_header(
    value: object,
) -> str:
    """
    Normalize a header for reliable matching.
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


def setup_google_login(
    browser_profile_path: Path,
    spreadsheet_id: str,
    worksheet_name: str,
) -> None:
    """
    Open a dedicated Chrome profile so the user can sign in to
    the authorized Google account.
    """
    browser_profile_path.mkdir(
        parents=True,
        exist_ok=True,
    )

    spreadsheet_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/edit"
    )

    with sync_playwright() as playwright:
        browser_context = (
            playwright.chromium.launch_persistent_context(
                user_data_dir=str(browser_profile_path),
                channel="chrome",
                headless=False,
                accept_downloads=True,
            )
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
        print(
            "Sign in with the authorized VCI Google account if "
            "requested."
        )
        print(
            f"Open the '{worksheet_name}' worksheet and confirm "
            "that the sheet is editable."
        )
        print(
            "Return to this terminal and press Enter only after "
            "the worksheet is visible."
        )

        input()

        try:
            browser_context.close()
        except Exception:
            # Chrome may already have been closed manually after login.
            pass


def is_google_login_page(
    page: Page,
) -> bool:
    """Return True when Google redirected the browser to sign-in."""
    current_url = page.url.lower()

    return (
        "accounts.google.com" in current_url
        or "/signin/" in current_url
        or "servicelogin" in current_url
    )


def get_or_create_google_sheet_page(
    browser_context: BrowserContext,
    spreadsheet_id: str,
) -> Page:
    """Reuse a Google Sheets tab when possible instead of opening extras."""
    spreadsheet_marker = f"/spreadsheets/d/{spreadsheet_id}/"

    for page in browser_context.pages:
        if (
            not page.is_closed()
            and spreadsheet_marker in page.url
        ):
            return page

    # Use a fresh page instead of a restored about:blank tab. Chrome can
    # automatically discard or close restored tabs while they navigate.
    return browser_context.new_page()


def open_google_spreadsheet_page(
    browser_context: BrowserContext,
    spreadsheet_id: str,
) -> Page:
    """Open the spreadsheet, recreating its tab if Chrome closes it."""
    spreadsheet_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/edit"
    )
    last_error = None

    for _ in range(3):
        page = get_or_create_google_sheet_page(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
        )

        try:
            page.goto(
                spreadsheet_url,
                wait_until="domcontentloaded",
                timeout=180_000,
            )
            time.sleep(3)

            if page.is_closed():
                continue

            return page
        except Exception as error:
            last_error = error

            if "Target page, context or browser has been closed" not in str(
                error
            ):
                raise

    raise RuntimeError(
        "Google repeatedly closed the destination spreadsheet tab "
        "while it was opening."
    ) from last_error


def ensure_google_sheet_access(
    browser_context: BrowserContext,
    spreadsheet_id: str,
    worksheet_name: str,
    force_confirmation: bool = False,
) -> Page:
    """
    Open the destination spreadsheet and pause for interactive sign-in
    only when the saved Google session is no longer authorized.
    """
    page = get_or_create_google_sheet_page(
        browser_context=browser_context,
        spreadsheet_id=spreadsheet_id,
    )

    if (
        page.is_closed()
        or spreadsheet_id not in page.url
        or force_confirmation
    ):
        page = open_google_spreadsheet_page(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
        )
    else:
        time.sleep(3)

        if page.is_closed():
            page = open_google_spreadsheet_page(
                browser_context=browser_context,
                spreadsheet_id=spreadsheet_id,
            )

    if is_google_login_page(page) or force_confirmation:
        print()
        print("Google Sheets authentication is required.")
        print(
            "Complete Google sign-in in Chrome using the authorized "
            "VCI account."
        )
        print(
            f"Confirm that the '{worksheet_name}' worksheet is "
            "visible and editable."
        )
        print(
            "Return to this terminal and press Enter only after "
            "the worksheet is visible."
        )
        input()

        # The user may have completed authentication in a newly opened
        # tab, so find the spreadsheet tab again before continuing.
        page = get_or_create_google_sheet_page(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
        )

        if page.is_closed() or spreadsheet_id not in page.url:
            page = open_google_spreadsheet_page(
                browser_context=browser_context,
                spreadsheet_id=spreadsheet_id,
            )
        else:
            time.sleep(3)

        if is_google_login_page(page):
            raise RuntimeError(
                "Google sign-in was not completed. The destination "
                "spreadsheet is still on the login page."
            )

    return page


def get_worksheet_csv(
    browser_context: BrowserContext,
    spreadsheet_id: str,
    worksheet_name: str,
) -> list[list[str]]:
    """
    Export the destination worksheet as CSV using the authenticated
    Google browser session.
    """
    export_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/gviz/tq"
        "?tqx=out:csv"
        f"&sheet={quote(worksheet_name)}"
    )

    response = browser_context.request.get(
        export_url,
        timeout=120_000,
    )

    if not response.ok:
        raise RuntimeError(
            "Could not read the Google Sheet. "
            f"Google returned HTTP {response.status}."
        )

    response_text = response.text()

    if (
        "<html" in response_text.lower()
        or "accounts.google.com" in response_text.lower()
    ):
        raise RuntimeError(
            "The Google session is not authorized to read the "
            "destination spreadsheet."
        )

    return list(
        csv.reader(
            io.StringIO(response_text)
        )
    )


def get_worksheet_csv_with_login_retry(
    browser_context: BrowserContext,
    spreadsheet_id: str,
    worksheet_name: str,
) -> tuple[list[list[str]], Page]:
    """Read worksheet CSV and recover interactively from HTTP 401."""
    page = ensure_google_sheet_access(
        browser_context=browser_context,
        spreadsheet_id=spreadsheet_id,
        worksheet_name=worksheet_name,
    )

    try:
        worksheet_rows = get_worksheet_csv(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name,
        )
    except RuntimeError as error:
        error_text = str(error).lower()

        if (
            "http 401" not in error_text
            and "not authorized" not in error_text
        ):
            raise

        page = ensure_google_sheet_access(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name,
            force_confirmation=True,
        )

        worksheet_rows = get_worksheet_csv(
            browser_context=browser_context,
            spreadsheet_id=spreadsheet_id,
            worksheet_name=worksheet_name,
        )

    return worksheet_rows, page


def find_target_columns(
    worksheet_rows: list[list[str]],
    target_headers: list[str],
) -> dict[str, int]:
    """
    Locate each target header and return its one-based column number.
    """
    if not worksheet_rows:
        raise ValueError(
            "The destination Google worksheet is empty."
        )

    header_row = worksheet_rows[0]

    normalized_destination_headers = {
        normalize_header(header): index
        for index, header in enumerate(
            header_row,
            start=1,
        )
        if normalize_header(header)
    }

    target_columns = {}
    missing_headers = []

    for target_header in target_headers:
        normalized_target = normalize_header(
            target_header
        )

        column_number = (
            normalized_destination_headers.get(
                normalized_target
            )
        )

        if column_number is None:
            missing_headers.append(
                target_header
            )
        else:
            target_columns[target_header] = (
                column_number
            )

    if missing_headers:
        raise ValueError(
            "The following destination headers were not found on "
            f"the Google Sheet: {', '.join(missing_headers)}"
        )

    return target_columns


def open_target_worksheet(
    page: Page,
    spreadsheet_id: str,
    worksheet_name: str,
) -> None:
    """
    Open the spreadsheet and select the required worksheet tab.
    """
    spreadsheet_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{spreadsheet_id}/edit"
    )

    page.goto(
        spreadsheet_url,
        wait_until="domcontentloaded",
        timeout=180_000,
    )

    page.wait_for_timeout(
        5_000
    )

    worksheet_tab = page.get_by_text(
        worksheet_name,
        exact=True,
    ).last

    try:
        worksheet_tab.wait_for(
            state="visible",
            timeout=60_000,
        )
        worksheet_tab.click()
        page.wait_for_timeout(
            2_000
        )
    except Exception as error:
        raise RuntimeError(
            f"Could not select the '{worksheet_name}' worksheet "
            "tab. Confirm that the tab exists and is visible."
        ) from error


def get_name_box(
    page: Page,
):
    """
    Find the Google Sheets name box used to select an exact range.
    """
    selectors = [
        "#t-name-box",
        "input[aria-label='Name box']",
        "input[aria-label*='Name box']",
    ]

    for selector in selectors:
        locator = page.locator(
            selector
        ).first

        try:
            if locator.is_visible(
                timeout=2_000
            ):
                return locator
        except Exception:
            continue

    raise RuntimeError(
        "Could not find the Google Sheets name box."
    )


def select_range(
    page: Page,
    cell_range: str,
) -> None:
    """
    Select an exact Google Sheets range through the name box.
    """
    name_box = get_name_box(
        page
    )

    name_box.click()
    name_box.fill(
        cell_range
    )
    name_box.press(
        "Enter"
    )

    page.wait_for_timeout(
        500
    )


def clear_column_data(
    page: Page,
    column_letter: str,
    data_start_row: int,
    clear_through_row: int = 0,
) -> None:
    """
    Clear all existing values beneath one destination header using
    an open-ended range that stays within the worksheet grid.

    clear_through_row is retained for compatibility with older
    callers but is intentionally not used.
    """
    target_range = (
        f"{column_letter}{data_start_row}:"
        f"{column_letter}"
    )

    select_range(
        page,
        target_range,
    )

    page.keyboard.press(
        "Delete"
    )

    page.wait_for_timeout(
        500
    )


def paste_column_data(
    page: Page,
    column_letter: str,
    data_start_row: int,
    values: list[str],
) -> None:
    """
    Paste one complete destination column.
    """
    if not values:
        return

    starting_cell = (
        f"{column_letter}{data_start_row}"
    )

    select_range(
        page,
        starting_cell,
    )

    clipboard_text = "\n".join(
        "" if value is None else str(value)
        for value in values
    )

    page.evaluate(
        """
        async (text) => {
            await navigator.clipboard.writeText(text);
        }
        """,
        clipboard_text,
    )

    page.keyboard.press(
        "Control+V"
    )

    page.wait_for_timeout(
        1_500
    )


def verify_google_sheet_update(
    browser_context: BrowserContext,
    spreadsheet_id: str,
    worksheet_name: str,
    target_headers: list[str],
    expected_rows: list[dict[str, str]],
) -> None:
    """
    Re-export the worksheet and confirm that all replacement values
    match the source data.
    """
    worksheet_rows = get_worksheet_csv(
        browser_context=browser_context,
        spreadsheet_id=spreadsheet_id,
        worksheet_name=worksheet_name,
    )

    target_columns = find_target_columns(
        worksheet_rows=worksheet_rows,
        target_headers=target_headers,
    )

    actual_rows = []

    for worksheet_row in worksheet_rows[1:]:
        extracted_row = {}

        for header in target_headers:
            zero_based_index = (
                target_columns[header] - 1
            )

            extracted_row[header] = (
                worksheet_row[zero_based_index].strip()
                if zero_based_index < len(
                    worksheet_row
                )
                else ""
            )

        if any(
            extracted_row.values()
        ):
            actual_rows.append(
                extracted_row
            )

    if len(actual_rows) != len(
        expected_rows
    ):
        raise RuntimeError(
            "Google Sheet verification failed. "
            f"Expected {len(expected_rows)} rows but found "
            f"{len(actual_rows)} populated rows."
        )

    for row_number, (
        expected_row,
        actual_row,
    ) in enumerate(
        zip(
            expected_rows,
            actual_rows,
        ),
        start=2,
    ):
        for header in target_headers:
            expected_value = str(
                expected_row.get(
                    header,
                    "",
                )
            ).strip()

            actual_value = str(
                actual_row.get(
                    header,
                    "",
                )
            ).strip()

            if expected_value != actual_value:
                raise RuntimeError(
                    "Google Sheet verification failed at "
                    f"row {row_number}, column '{header}'. "
                    f"Expected '{expected_value}' but found "
                    f"'{actual_value}'."
                )


def update_google_sheet(
    spreadsheet_id: str,
    worksheet_name: str,
    browser_profile_path: Path,
    target_headers: list[str],
    rows: list[dict[str, str]],
    data_start_row: int = 2,
) -> dict:
    """
    Replace the data beneath the specified Google Sheet headers while
    preserving unrelated worksheet columns.
    """
    if not rows:
        raise ValueError(
            "No vehicle rows were provided for the Google Sheet."
        )

    browser_profile_path.mkdir(
        parents=True,
        exist_ok=True,
    )

    with sync_playwright() as playwright:
        browser_context = (
            playwright.chromium.launch_persistent_context(
                user_data_dir=str(browser_profile_path),
                channel="chrome",
                headless=False,
                accept_downloads=True,
                permissions=[
                    "clipboard-read",
                    "clipboard-write",
                ],
            )
        )

        try:
            # Keep one independent page alive. Google can replace or close
            # the tab used during authentication and initial navigation.
            keeper_page = browser_context.new_page()
            keeper_page.goto("about:blank")

            worksheet_rows, page = get_worksheet_csv_with_login_retry(
                browser_context=browser_context,
                spreadsheet_id=spreadsheet_id,
                worksheet_name=worksheet_name,
            )

            target_columns = find_target_columns(
                worksheet_rows=worksheet_rows,
                target_headers=target_headers,
            )

            open_target_worksheet(
                page=page,
                spreadsheet_id=spreadsheet_id,
                worksheet_name=worksheet_name,
            )

            print(
                f"Updating Google Sheet tab "
                f"'{worksheet_name}'..."
            )

            # Clear all three destination columns before inserting
            # the replacement dataset.
            for header in target_headers:
                column_letter = (
                    column_number_to_letter(
                        target_columns[header]
                    )
                )

                clear_column_data(
                    page=page,
                    column_letter=column_letter,
                    data_start_row=data_start_row,
                )

            # Paste each column separately so unrelated worksheet
            # columns remain untouched.
            for header in target_headers:
                column_letter = (
                    column_number_to_letter(
                        target_columns[header]
                    )
                )

                column_values = [
                    row.get(
                        header,
                        "",
                    )
                    for row in rows
                ]

                paste_column_data(
                    page=page,
                    column_letter=column_letter,
                    data_start_row=data_start_row,
                    values=column_values,
                )

            print(
                "Waiting for Google Sheets to save changes..."
            )

            page.wait_for_timeout(
                8_000
            )

            verify_google_sheet_update(
                browser_context=browser_context,
                spreadsheet_id=spreadsheet_id,
                worksheet_name=worksheet_name,
                target_headers=target_headers,
                expected_rows=rows,
            )

            print(
                f"Google Sheet updated and verified successfully: "
                f"{len(rows)} vehicle rows."
            )

            return {
                "rows_written": len(rows),
                "worksheet_name": worksheet_name,
                "target_columns": target_columns,
            }

        except Exception as error:
            raise RuntimeError(
                "Could not update the Google Sheet. Confirm that "
                "the Google session is active, the account has edit "
                "access, and the destination headers have not been "
                "renamed."
            ) from error

        finally:
            try:
                browser_context.close()
            except Exception:
                pass
