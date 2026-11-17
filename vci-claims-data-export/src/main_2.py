from __future__ import annotations

import os
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


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


def setup_login(spreadsheet_id: str, cdp_endpoint: str | None = "auto") -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"

    with sync_playwright() as playwright:
        try:
            browser = _connect_to_open_chrome(playwright, cdp_endpoint)
        except PlaywrightError as exc:
            raise AuthenticationError(
                f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
                "debugging' is checked at chrome://inspect/#remote-debugging."
            ) from exc

        page = None
        try:
            if not browser.contexts:
                raise AuthenticationError("No active Chrome profile context found.")

            context = browser.contexts[0]
            page = context.new_page()
            page.goto(sheet_url, wait_until="domcontentloaded", timeout=120_000)
            
            print("Sign in to Google and confirm that the spreadsheet opens.")
            input("When the CURRENT tab is visible, return here and press Enter...")

            # Exact SSO redirect logic from your original script
            if spreadsheet_id not in page.url:
                try:
                    page.goto(
                        sheet_url,
                        wait_until="domcontentloaded",
                        timeout=120_000,
                    )
                except PlaywrightError as exc:
                    if "interrupted by another navigation" not in str(exc):
                        raise

            if spreadsheet_id not in page.url:
                try:
                    page.wait_for_url(
                        f"**/spreadsheets/d/{spreadsheet_id}/**",
                        timeout=120_000,
                    )
                except PlaywrightTimeoutError as exc:
                    raise AuthenticationError(
                        "Google/CompanyView sign-in did not return to the "
                        "authorized spreadsheet."
                    ) from exc

            if "accounts.google.com" in page.url:
                raise AuthenticationError(
                    "Google sign-in was not completed. The spreadsheet is still "
                    "redirecting to the login page."
                )
            if spreadsheet_id not in page.url:
                raise AuthenticationError(
                    "The authorized spreadsheet was not visible after sign-in."
                )

            print("Google spreadsheet access confirmed.")
        finally:
            if page is not None and not page.is_closed():
                page.close()
            browser.close()


def download_tab_values(
    spreadsheet_id: str,
    source_tab: str,
    output_path: Path,
    cdp_endpoint: str | None = "auto",
) -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq"
        f"?tqx=out:csv&sheet={source_tab}"
    )
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        try:
            browser = _connect_to_open_chrome(playwright, cdp_endpoint)
        except PlaywrightError as exc:
            raise AuthenticationError(
                f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
                "debugging' is checked at chrome://inspect/#remote-debugging."
            ) from exc

        page = None
        try:
            if not browser.contexts:
                raise AuthenticationError("No active Chrome profile context found.")

            context = browser.contexts[0]
            page = context.new_page()

            response = page.goto(
                sheet_url,
                wait_until="domcontentloaded",
                timeout=120_000,
            )

            # Give silent Okta/CompanyView SSO redirects up to 15s to finish
            if spreadsheet_id not in page.url:
                try:
                    page.wait_for_url(
                        f"**/spreadsheets/d/{spreadsheet_id}/**",
                        timeout=15_000,
                    )
                except PlaywrightTimeoutError:
                    pass

            if "accounts.google.com" in page.url:
                raise AuthenticationError(
                    "The saved Google session has expired. Open Chrome and refresh your login."
                )
            if response is not None and response.status >= 400:
                raise RuntimeError(
                    f"Google spreadsheet page failed with HTTP {response.status}."
                )

            export_response = context.request.get(export_url, timeout=120_000)
            if export_response.status in (401, 403):
                raise AuthenticationError(
                    "Google rejected the saved session. Open Chrome and refresh your login."
                )
            if not export_response.ok:
                raise RuntimeError(
                    f"Google CSV export failed with HTTP {export_response.status}."
                )
            output_path.write_bytes(export_response.body())
        finally:
            if page is not None and not page.is_closed():
                page.close()
            browser.close()

    body_prefix = output_path.read_bytes()[:2] if output_path.exists() else b""
    if body_prefix in (b"<!", b"<h"):
        if output_path.exists():
            output_path.unlink()
        raise AuthenticationError(
            "Google did not return the tab values. Open Chrome and refresh your login."
        )
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError("Google returned an empty CSV export.")