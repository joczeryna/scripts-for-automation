from __future__ import annotations

from pathlib import Path

from playwright.sync_api import sync_playwright


class AuthenticationError(RuntimeError):
    pass


def setup_login(spreadsheet_id: str, profile_dir: Path) -> None:
    sheet_url = f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/edit"
    profile_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile_dir),
            channel="chrome",
            headless=False,
        )
        page = context.pages[0] if context.pages else context.new_page()
        page.goto(sheet_url, wait_until="domcontentloaded", timeout=120_000)
        print("Sign in to Google and confirm that the spreadsheet opens.")
        input("When the CURRENT tab is visible, return here and press Enter...")
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
            page = context.pages[0] if context.pages else context.new_page()
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
