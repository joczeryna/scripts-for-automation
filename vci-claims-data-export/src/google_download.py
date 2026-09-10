from __future__ import annotations

from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


class AuthenticationError(RuntimeError):
    pass


def _get_single_sheet_page(context, spreadsheet_id: str):
    """Reuse one browser tab and remove tabs left over from earlier runs."""
    open_pages = [page for page in context.pages if not page.is_closed()]
    matching_pages = [page for page in open_pages if spreadsheet_id in page.url]

    if matching_pages:
        page = matching_pages[0]
    elif open_pages:
        # Reuse the existing browser tab instead of adding another one.
        page = open_pages[0]
    else:
        page = context.new_page()

    # This is a dedicated automation profile. Keep exactly one tab so restored
    # login or spreadsheet tabs do not accumulate on subsequent runs.
    for extra_page in open_pages:
        if extra_page != page and not extra_page.is_closed():
            extra_page.close()

    return page


def setup_login(spreadsheet_id: str, profile_dir: Path) -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            channel="chrome",
            headless=False,
        )
        page = _get_single_sheet_page(context, spreadsheet_id)
        page.goto(sheet_url, wait_until="domcontentloaded", timeout=120_000)
        print("Sign in to Google and confirm that the spreadsheet opens.")
        input("When the CURRENT tab is visible, return here and press Enter...")

        # Do not report success merely because Enter was pressed. Reuse the
        # authenticated spreadsheet tab. Avoid reloading it when it is already
        # open because the VCI CompanyView SSO handoff can interrupt a second
        # navigation even though authentication succeeded.
        page = _get_single_sheet_page(context, spreadsheet_id)

        if spreadsheet_id not in page.url:
            try:
                page.goto(
                    sheet_url,
                    wait_until="domcontentloaded",
                    timeout=120_000,
                )
            except PlaywrightError as exc:
                # CompanyView may start its token redirect while Playwright is
                # waiting for the Google navigation. That interruption is an
                # expected part of SSO; wait for the redirect chain to finish.
                if "interrupted by another navigation" not in str(exc):
                    raise

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

        print("Google spreadsheet access confirmed. Saving the browser session...")
        context.close()


def download_tab_values(
    spreadsheet_id: str,
    source_tab: str,
    profile_dir: Path,
    output_path: Path,
) -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    # GViz CSV returns the displayed values from the selected tab. This is
    # intentional: a normal XLSX export can preserve an IMPORTRANGE formula
    # instead of materializing the values it currently displays.
    export_url = (
        f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/gviz/tq"
        f"?tqx=out:csv&sheet={source_tab}"
    )
    profile_dir.mkdir(parents=True, exist_ok=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            channel="chrome",
            headless=True,
            accept_downloads=True,
        )
        try:
            page = _get_single_sheet_page(context, spreadsheet_id)
            response = page.goto(
                sheet_url,
                wait_until="domcontentloaded",
                timeout=120_000,
            )
            if "accounts.google.com" in page.url:
                raise AuthenticationError(
                    "The saved Google session has expired. Run with --setup-login."
                )
            if response is not None and response.status >= 400:
                raise RuntimeError(
                    f"Google spreadsheet page failed with HTTP {response.status}."
                )

            # BrowserContext.request shares the authenticated browser context's
            # cookie storage, while avoiding Chrome download-event lifecycle
            # failures that can close the target before Playwright saves the file.
            export_response = context.request.get(export_url, timeout=120_000)
            if export_response.status in (401, 403):
                raise AuthenticationError(
                    "Google rejected the saved session. Run with --setup-login."
                )
            if not export_response.ok:
                raise RuntimeError(
                    f"Google CSV export failed with HTTP {export_response.status}."
                )
            output_path.write_bytes(export_response.body())
        finally:
            context.close()

    body_prefix = output_path.read_bytes()[:2] if output_path.exists() else b""
    if body_prefix in (b"<!", b"<h"):
        if output_path.exists():
            output_path.unlink()
        raise AuthenticationError(
            "Google did not return the tab values. Run with --setup-login."
        )
    if not output_path.exists() or output_path.stat().st_size == 0:
        raise RuntimeError("Google returned an empty CSV export.")
