from __future__ import annotations

import argparse
import logging
import os
import re
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import BrowserContext, Error as PlaywrightError, Page, sync_playwright

# ==========================================
# CONFIGURATION & PATHS
# ==========================================
PROJECT_ROOT = Path(__file__).resolve().parent.parent
TEMP_DIR = PROJECT_ROOT / "temp"
LOG_FOLDER = PROJECT_ROOT / "logs"

DRIVE_FOLDER_URL = "https://drive.google.com/drive/folders/1HeWm86Q67OwUgFOPGyx_AtxDPflfPkYS"
DESTINATION_DIR = Path(r"W:\Dashboards VCI\Safety Data\Current\Safety Training")

FILE_PATTERN = re.compile(r"SF Safety Training Compliance Report.*\.xlsx", re.IGNORECASE)
ARTIFACT_PATTERN = re.compile(r"SF Safety Training Compliance Report.*\.xlsx(\.crdownload)?$", re.IGNORECASE)

DOWNLOAD_START_TIMEOUT_SECONDS = 240   # 4 minutes for Google Drive to start the download
DOWNLOAD_FINISH_TIMEOUT_SECONDS = 300  # 5 minutes for the file transfer to finish


class AuthenticationError(RuntimeError):
    pass


def configure_logging() -> Path:
    LOG_FOLDER.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_FOLDER / f"safety_training_export_{stamp}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )
    return log_path


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


def wait_for_drive_folder(page: Page, timeout_ms: int = 60_000) -> Page:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        if page.is_closed():
            raise RuntimeError("The Google Drive tab was closed unexpectedly.")
        if "1HeWm86Q67OwUgFOPGyx_AtxDPflfPkYS" in page.url:
            try:
                page.wait_for_load_state("networkidle", timeout=3000)
                return page
            except Exception:
                return page
        time.sleep(1)
    raise RuntimeError("The Google Drive folder page was not detected.")


def setup_login(url: str, cdp_endpoint: str | None = "auto") -> None:
    """Open the target Google Drive folder in a single tab of your running Chrome."""
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

            context = browser.contexts[0]
            page = context.new_page()
            page.goto(url, wait_until="domcontentloaded", timeout=180_000)
            print()
            print("A new tab is open in your Chrome window for Google authentication.")
            print("Sign in, complete MFA, and ensure you are in the target Drive folder:")
            print(url)
            print("Press Enter after the Drive folder is fully visible to close the tab.")
            input()
        finally:
            if page is not None and not page.is_closed():
                page.close()
            browser.close()


def current_artifacts(
    folder: Path,
    started_at: float,
    name_filter: re.Pattern[str] | None = None,
) -> list[Path]:
    if not folder.exists():
        return []
    artifacts = []
    for path in folder.iterdir():
        try:
            if not path.is_file() or path.stat().st_mtime < started_at - 2:
                continue
            lower_name = path.name.lower()
            if not (lower_name.endswith(".xlsx") or lower_name.endswith(".xlsx.crdownload")):
                continue
            if name_filter and not name_filter.search(path.name):
                continue
            artifacts.append(path)
        except OSError:
            pass
    return artifacts


def find_started_artifact(started_at: float) -> tuple[Path, list[Path]] | None:
    """Checks TEMP_DIR first, then falls back to the user's Downloads folder."""
    temp_hits = current_artifacts(TEMP_DIR, started_at)
    if temp_hits:
        return TEMP_DIR, temp_hits

    user_downloads = Path.home() / "Downloads"
    dl_hits = current_artifacts(user_downloads, started_at, name_filter=ARTIFACT_PATTERN)
    if dl_hits:
        return user_downloads, dl_hits

    return None


def wait_for_xlsx(folder: Path, started_at: float, name_filter: re.Pattern[str] | None = None) -> Path:
    deadline = time.monotonic() + DOWNLOAD_FINISH_TIMEOUT_SECONDS
    sizes: dict[Path, tuple[int, int]] = {}
    while time.monotonic() < deadline:
        for path in current_artifacts(folder, started_at, name_filter=name_filter):
            try:
                size = path.stat().st_size
                previous, stable = sizes.get(path, (-1, 0))
                stable = stable + 1 if size == previous else 1
                sizes[path] = (size, stable)

                required_stability = 10 if path.name.endswith(".crdownload") else 3
                if size <= 0 or stable < required_stability:
                    continue

                if path.name.endswith(".crdownload"):
                    final_path = path.with_name(path.name[:-len(".crdownload")])
                    try:
                        path.replace(final_path)
                    except OSError:
                        shutil.copy2(path, final_path)
                    path = final_path
                    print("Recovered Chrome's complete stable Excel download.")
                return path
            except OSError:
                continue
        time.sleep(1)
    names = [path.name for path in current_artifacts(folder, started_at, name_filter=name_filter)]
    raise TimeoutError(f"The Excel file did not finish downloading. Files: {names}")


def configure_cdp_downloads(context: BrowserContext, target_page: Page, download_dir: Path | None) -> None:
    """Point Chrome downloads to TEMP_DIR during export, or restore default behavior when None."""
    try:
        session = context.new_cdp_session(target_page)
        if download_dir is not None:
            session.send(
                "Browser.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": str(download_dir.resolve()),
                    "eventsEnabled": True,
                },
            )
        else:
            session.send("Browser.setDownloadBehavior", {"behavior": "default"})
    except Exception:
        pass


def try_click_download_anyway_once(
    context: BrowserContext,
    script_pages: list[Page],
    download_dir: Path,
) -> bool:
    """
    Scans tabs opened by this script (including any virus-scan popup tab)
    for Google Drive's 'Download anyway' prompt and clicks it if visible.
    """
    pattern = re.compile(r"Download anyway", re.IGNORECASE)
    for pg in script_pages:
        if pg.is_closed():
            continue
        configure_cdp_downloads(context, pg, download_dir)

        for frame in pg.frames:
            candidates = [
                frame.get_by_role("button", name=pattern),
                frame.get_by_role("link", name=pattern),
                frame.locator("button, a, [role='button'], input[type='submit']").filter(has_text=pattern),
            ]
            for locator in candidates:
                try:
                    count = locator.count()
                    for idx in range(count):
                        btn = locator.nth(idx)
                        if btn.is_visible():
                            btn.click(timeout=5000)
                            print("Clicked 'Download anyway' on the virus scan dialog.")
                            return True
                except Exception:
                    continue
    return False


def download_report(cdp_endpoint: str | None = "auto") -> Path:
    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        try:
            browser = _connect_to_open_chrome(playwright, cdp_endpoint)
        except PlaywrightError as exc:
            raise AuthenticationError(
                f"Could not attach to open Chrome ({exc}). Ensure 'Allow remote "
                "debugging' is checked at chrome://inspect/#remote-debugging and "
                "click Allow if Chrome prompts you."
            ) from exc

        context = None
        page = None
        existing_pages: set[Page] = set()
        try:
            if not browser.contexts:
                raise AuthenticationError("No active Chrome profile context found.")

            context = browser.contexts[0]
            # Snapshot your open tabs so we never touch or close your existing work
            existing_pages = set(context.pages)
            page = context.new_page()

            page.goto(DRIVE_FOLDER_URL, wait_until="domcontentloaded", timeout=180_000)

            if "accounts.google.com" in page.url.lower():
                raise AuthenticationError(
                    "Your Chrome Google session has expired. Sign in via Chrome and retry."
                )

            page = wait_for_drive_folder(page)

            print("Searching for the target Excel file...")
            page.wait_for_selector("[role='row']", timeout=30000)

            file_row = page.get_by_role("row").filter(has_text=FILE_PATTERN).first
            if not file_row.is_visible(timeout=10000):
                raise RuntimeError(f"Could not find a file matching pattern: {FILE_PATTERN.pattern}")

            file_row.click(button="right")
            print("Opened context menu for the file.")

            # Target the context menu item specifically to avoid hidden toolbar tooltip collisions
            download_btn = page.get_by_role("menuitem", name="Download", exact=True).first
            download_btn.wait_for(state="visible", timeout=10000)

            configure_cdp_downloads(context, page, TEMP_DIR)

            started_at = time.time()
            download_btn.click()
            print(f"Clicked Download... Waiting up to {DOWNLOAD_START_TIMEOUT_SECONDS}s for download to begin.")

            start_begin = time.monotonic()
            start_deadline = start_begin + DOWNLOAD_START_TIMEOUT_SECONDS
            last_heartbeat = start_begin
            clicked_anyway = False
            found_location: tuple[Path, list[Path]] | None = None

            while time.monotonic() < start_deadline:
                found_location = find_started_artifact(started_at)
                if found_location:
                    break

                if not clicked_anyway:
                    script_pages = [p for p in context.pages if p not in existing_pages and not p.is_closed()]
                    clicked_anyway = try_click_download_anyway_once(context, script_pages, TEMP_DIR)

                now = time.monotonic()
                if now - last_heartbeat >= 30:
                    elapsed = int(now - start_begin)
                    print(f"Still waiting for Google Drive to start download ({elapsed}s / {DOWNLOAD_START_TIMEOUT_SECONDS}s)...")
                    last_heartbeat = now

                time.sleep(1)

            if not found_location:
                raise RuntimeError(
                    f"The Download button was clicked, but no download file started within {DOWNLOAD_START_TIMEOUT_SECONDS}s."
                )

            active_dir, _ = found_location
            active_filter = None if active_dir == TEMP_DIR else ARTIFACT_PATTERN
            print(f"The Google Drive download has started in: {active_dir}")

            excel_path = wait_for_xlsx(active_dir, started_at, name_filter=active_filter)
            print(f"Downloaded {excel_path.name} successfully.")
            return excel_path

        finally:
            # 1. Restore default Chrome download behavior so your normal downloads aren't redirected to TEMP_DIR
            if context is not None and page is not None and not page.is_closed():
                configure_cdp_downloads(context, page, None)

            # 2. Close ONLY tabs created during this run (the Drive tab + any 'Download anyway' popup tab)
            if context is not None:
                for p in list(context.pages):
                    if p not in existing_pages and not p.is_closed():
                        try:
                            p.close()
                        except Exception:
                            pass
            elif page is not None and not page.is_closed():
                page.close()

            # 3. Detach Playwright without closing your Chrome window
            browser.close()


def save_report(source: Path, destination_folder: Path) -> Path:
    if not destination_folder.exists():
        raise FileNotFoundError(f"Destination folder is unavailable: {destination_folder}")

    destination = destination_folder / source.name
    temporary = destination_folder / f"{source.stem}_temp{source.suffix}"
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()
    return destination


def parse_arguments():
    parser = argparse.ArgumentParser(description="Export the Safety Training Compliance Report from Google Drive.")
    parser.add_argument("--setup-login", action="store_true")
    return parser.parse_args()


def main() -> int:
    log_path = configure_logging()
    try:
        args = parse_arguments()

        if args.setup_login:
            setup_login(DRIVE_FOLDER_URL)
            logging.info("Google Drive login setup completed.")
            print("Google Drive tab verified successfully.")
            return 0

        logging.info("Starting Google Drive export.")
        downloaded = download_report()

        destination = save_report(downloaded, DESTINATION_DIR)
        logging.info("Saved the file to %s.", destination)

        print()
        print("VCI Safety Training export completed successfully.")
        print(f"File saved to: {destination}")
        logging.info("Log saved to %s.", log_path)
        return 0
    except AuthenticationError as exc:
        logging.error("AUTHENTICATION REQUIRED: %s", exc)
        logging.info("Review the log file: %s", log_path)
        print()
        print(f"Authentication required: {exc}")
        return 2
    except Exception:
        logging.exception("The VCI Safety Training export failed.")
        logging.info("Review the log file: %s", log_path)
        print()
        print("VCI Safety Training export failed.")
        print("Review the latest file in the logs folder.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())