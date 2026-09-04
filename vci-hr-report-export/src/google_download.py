from pathlib import Path

from playwright.sync_api import sync_playwright


GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
)

GOOGLE_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "{spreadsheet_id}/export?format=xlsx"
)


def setup_google_login(
    browser_profile_path: Path,
    spreadsheet_id: str,
) -> None:
    """
    Open a dedicated Chrome profile so the user can sign in to Google.

    The authenticated session remains in browser_profile_path and is
    reused by future downloads.
    """
    browser_profile_path.mkdir(
        parents=True,
        exist_ok=True,
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
            GOOGLE_SHEET_URL.format(
                spreadsheet_id=spreadsheet_id
            ),
            wait_until="domcontentloaded",
            timeout=120_000,
        )

        print()
        print("Google Chrome is now open.")
        print(
            "Sign in with the authorized VCI Google account and "
            "confirm that the spreadsheet opens."
        )
        print(
            "Return to this terminal and press Enter only after "
            "the spreadsheet is visible."
        )

        input()

        try:
            browser_context.close()
        except Exception as error:
            if (
                "Target page, context or browser has been closed"
                not in str(error)
            ):
                raise


def download_reports(
    reports: list[dict],
    download_folder: Path,
    browser_profile_path: Path,
) -> list[dict]:
    """
    Export configured Google Sheets as Excel workbooks through the
    authenticated Chrome browser context.
    """
    download_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    browser_profile_path.mkdir(
        parents=True,
        exist_ok=True,
    )

    download_results = []

    with sync_playwright() as playwright:
        browser_context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(browser_profile_path),
            channel="chrome",
            headless=False,
            accept_downloads=True,
        )

        try:
            for report in reports:
                report_name = report["name"]
                spreadsheet_id = report["spreadsheet_id"]

                temporary_filename = (
                    report_name.lower().replace(" ", "_")
                    + "_download.xlsx"
                )

                destination_path = (
                    download_folder
                    / temporary_filename
                )

                export_url = GOOGLE_EXPORT_URL.format(
                    spreadsheet_id=spreadsheet_id
                )

                print(
                    f"Downloading {report_name}..."
                )

                try:
                    response = browser_context.request.get(
                        export_url,
                        timeout=180_000,
                    )

                except Exception as error:
                    raise RuntimeError(
                        f"Google did not return the export for "
                        f"{report_name}."
                    ) from error

                if not response.ok:
                    raise RuntimeError(
                        f"Google returned HTTP {response.status} "
                        f"while exporting {report_name}."
                    )

                response_content_type = (
                    response.headers.get(
                        "content-type",
                        "",
                    ).lower()
                )

                workbook_bytes = response.body()

                if not workbook_bytes:
                    raise ValueError(
                        f"The downloaded file for {report_name} is empty."
                    )

                # An XLSX workbook is a ZIP package and begins with PK.
                if not workbook_bytes.startswith(
                    b"PK"
                ):
                    raise RuntimeError(
                        f"Google did not return an Excel workbook for "
                        f"{report_name}. Returned content type: "
                        f"{response_content_type or 'unknown'}. "
                        "The Google session may have expired."
                    )

                destination_path.write_bytes(
                    workbook_bytes
                )

                if not destination_path.exists():
                    raise FileNotFoundError(
                        f"The downloaded file for {report_name} "
                        "was not created."
                    )

                if destination_path.stat().st_size == 0:
                    raise ValueError(
                        f"The downloaded file for {report_name} is empty."
                    )

                download_results.append(
                    {
                        "report": report,
                        "download_path": destination_path,
                    }
                )

                print(
                    f"Downloaded {report_name} successfully."
                )

        finally:
            browser_context.close()

    return download_results