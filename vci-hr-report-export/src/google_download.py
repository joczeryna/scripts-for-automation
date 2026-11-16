from __future__ import annotations

import os
from pathlib import Path
import time

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright


GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
)

GOOGLE_EXPORT_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "{spreadsheet_id}/export?format=xlsx"
)

LOGIN_TIMEOUT_SECONDS = 15 * 60
LOGIN_CHECK_INTERVAL_SECONDS = 2


class AuthenticationError(RuntimeError):
    pass


def _connect_to_open_chrome(playwright, cdp_endpoint: str | None = None):
    """Attach to an already running Chrome instance via DevToolsActivePort."""
    if cdp_endpoint and cdp_endpoint != "auto":
        return playwright.chromium.connect_over_cdp(cdp_endpoint)

    local_app_data = os.environ.get("LOCALAPPDATA")
    if not local_app_data:
        raise AuthenticationError("LOCALAPPDATA environment variable not found.")

    port_file = (
        Path(local_app_data)
        / "Google"
        / "Chrome"
        / "User Data"
        / "DevToolsActivePort"
    )
    if not port_file.exists():
        raise AuthenticationError(
            "Chrome remote debugging is not active. In Chrome, open "
            "chrome://inspect/#remote-debugging and check "
            "'Allow remote debugging for this browser instance'."
        )

    lines = port_file.read_text(encoding="utf-8").strip().splitlines()
    if len(lines) < 2:
        raise AuthenticationError("Chrome DevToolsActivePort file is incomplete.")

    port, ws_path = lines[0].strip(), lines[1].strip()
    ws_url = f"ws://127.0.0.1:{port}{ws_path}"
    return playwright.chromium.connect_over_cdp(ws_url)


def setup_google_login(
    spreadsheet_id: str,
    cdp_endpoint: str | None = None,
) -> None:
    """
    Open a single tab in the user's running Chrome so they can sign in to Google.
    Closes only that tab once access is confirmed.
    """
    with sync_playwright() as playwright:
        try:
            browser = _connect_to_open_chrome(playwright, cdp_endpoint)
        except PlaywrightError as exc:
            raise AuthenticationError(
                f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
                "debugging' is checked at chrome://inspect/#remote-debugging and "
                "click Allow if Chrome prompts you."
            ) from exc

        page = None
        try:
            if not browser.contexts:
                raise AuthenticationError("No active Chrome profile context found.")

            browser_context = browser.contexts[0]
            page = browser_context.new_page()

            page.goto(
                GOOGLE_SHEET_URL.format(
                    spreadsheet_id=spreadsheet_id
                ),
                wait_until="domcontentloaded",
                timeout=120_000,
            )

            print()
            print("A new tab is now open in your Chrome window.")
            print(
                "Sign in with the authorized VCI Google account and "
                "confirm that the spreadsheet opens."
            )
            print(
                "This tab will close automatically after Google "
                "confirms access to the spreadsheet."
            )

            export_url = GOOGLE_EXPORT_URL.format(
                spreadsheet_id=spreadsheet_id
            )
            deadline = time.monotonic() + LOGIN_TIMEOUT_SECONDS

            while time.monotonic() < deadline:
                if page.is_closed():
                    raise RuntimeError(
                        "The Chrome tab was closed before Google sign-in completed."
                    )

                try:
                    response = browser_context.request.get(
                        export_url,
                        timeout=30_000,
                    )

                    if response.ok and response.body().startswith(b"PK"):
                        print(
                            "Google sign-in confirmed. Closing setup tab..."
                        )
                        break
                except Exception:
                    # Google may interrupt requests while the interactive
                    # sign-in flow is moving between pages. Keep waiting.
                    pass

                page.wait_for_timeout(
                    LOGIN_CHECK_INTERVAL_SECONDS * 1000
                )
            else:
                raise TimeoutError(
                    "Google sign-in was not confirmed within 15 minutes. "
                    "Run --setup-login again and complete the sign-in."
                )
        finally:
            if page is not None and not page.is_closed():
                page.close()
            # Detaches Playwright only — leaves your Chrome window and open tabs untouched
            browser.close()


def download_reports(
    reports: list[dict],
    download_folder: Path,
    cdp_endpoint: str | None = None,
) -> list[dict]:
    """
    Export configured Google Sheets as Excel workbooks through a single tab
    in the user's running Chrome browser context.
    """
    download_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    download_results = []

    with sync_playwright() as playwright:
        try:
            browser = _connect_to_open_chrome(playwright, cdp_endpoint)
        except PlaywrightError as exc:
            raise AuthenticationError(
                f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
                "debugging' is checked at chrome://inspect/#remote-debugging and "
                "click Allow if Chrome prompts you."
            ) from exc

        page = None
        try:
            if not browser.contexts:
                raise AuthenticationError("No active Chrome profile context found.")

            browser_context = browser.contexts[0]
            page = browser_context.new_page()

            # Open the first spreadsheet in the tab to confirm/refresh active Google SSO
            first_sheet_id = reports[0]["spreadsheet_id"]
            sheet_response = page.goto(
                GOOGLE_SHEET_URL.format(spreadsheet_id=first_sheet_id),
                wait_until="domcontentloaded",
                timeout=120_000,
            )
            if "accounts.google.com" in page.url:
                raise AuthenticationError(
                    "Your Chrome Google session has expired. Sign in via Chrome and retry."
                )
            if sheet_response is not None and sheet_response.status >= 400:
                raise RuntimeError(
                    f"Google spreadsheet page failed with HTTP {sheet_response.status}."
                )

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

                if response.status in (401, 403):
                    raise AuthenticationError(
                        f"Google rejected the active Chrome session while exporting {report_name}."
                    )

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
                if not workbook_bytes.startswith(b"PK"):
                    raise AuthenticationError(
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
            if page is not None and not page.is_closed():
                page.close()
            browser.close()

    return download_results