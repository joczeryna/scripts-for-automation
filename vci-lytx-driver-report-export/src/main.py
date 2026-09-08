import argparse
import csv
import json
import logging
import os
import re
import shutil
import sys
import tempfile
import time
from datetime import datetime
from pathlib import Path

from playwright.sync_api import BrowserContext, Locator, Page, sync_playwright


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "settings.json"
LOG_FOLDER = PROJECT_ROOT / "logs"


def load_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as config_file:
        return json.load(config_file)


def project_path(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (PROJECT_ROOT / path).resolve()


def configure_logging() -> Path:
    LOG_FOLDER.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_FOLDER / f"lytx_driver_report_export_{stamp}.log"
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


def get_lytx_page(context: BrowserContext) -> Page:
    pages = [page for page in context.pages if not page.is_closed()]
    for page in reversed(pages):
        if "driversreport" in page.url.lower():
            print("Reusing the existing Lytx Driver Report tab...")
            return page
    for page in reversed(pages):
        if "lytx.com" in page.url.lower():
            print("Reusing the existing Lytx tab...")
            return page
    return pages[0] if pages else context.new_page()


def wait_for_driver_report(context: BrowserContext, timeout_ms: int = 60_000) -> Page:
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for page in reversed(context.pages):
            if page.is_closed():
                continue
            if "driversreport" in page.url.lower():
                try:
                    if page.get_by_text("DRIVERS REPORT", exact=False).first.is_visible():
                        return page
                except Exception:
                    pass
        time.sleep(1)
    raise RuntimeError("The Lytx Drivers Report page was not detected.")


def setup_login(profile: Path, url: str) -> None:
    profile.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile), channel="chrome", headless=False,
            accept_downloads=True,
        )
        page = get_lytx_page(context)
        page.goto(url, wait_until="domcontentloaded", timeout=180_000)
        print()
        print("Chrome is open for Lytx authentication.")
        print("Sign in, complete MFA, and manually open:")
        print(url)
        print("Press Enter after the Drivers Report is fully visible.")
        input()
        try:
            context.close()
        except Exception:
            pass


def choose_last_30_days(page: Page) -> None:
    print("Selecting the Last 30 Days report period...")
    shortcut = page.get_by_text("Last 30 Days", exact=True)
    if not shortcut.first.is_visible(timeout=2_000):
        date_candidates = [
            page.get_by_text(re.compile(r"\d{1,2}/\d{1,2}/\d{2,4}\s*-\s*\d{1,2}/\d{1,2}/\d{2,4}")),
            page.locator("button, [role='button']").filter(has_text=re.compile(r"\d{1,2}/\d{1,2}/")),
        ]
        opened = False
        for candidate in date_candidates:
            try:
                if candidate.first.is_visible(timeout=2_000):
                    candidate.first.click()
                    opened = True
                    break
            except Exception:
                continue
        if not opened:
            raise RuntimeError("Could not open the Lytx date-range selector.")

    shortcut = page.get_by_text("Last 30 Days", exact=True)
    shortcut.wait_for(state="visible", timeout=30_000)
    shortcut.click()

    apply_button = page.get_by_role("button", name="Apply", exact=True)
    apply_button.wait_for(state="visible", timeout=30_000)
    apply_button.click()
    page.wait_for_timeout(5_000)
    print("Applied Last 30 Days successfully.")


def find_download_link(page: Page, label: str, timeout_ms: int = 180_000) -> Locator:
    pattern = re.compile(rf"{re.escape(label)}(?:\s*\(\d+\))?", re.IGNORECASE)
    deadline = time.monotonic() + timeout_ms / 1000
    while time.monotonic() < deadline:
        for frame in page.frames:
            try:
                frame.evaluate("""() => {
                    window.scrollTo(0, document.body.scrollHeight);
                    for (const el of document.querySelectorAll('*')) {
                        const s = getComputedStyle(el);
                        if (el.scrollHeight > el.clientHeight &&
                            (s.overflowY === 'auto' || s.overflowY === 'scroll')) {
                            el.scrollTop = el.scrollHeight;
                        }
                    }
                }""")
            except Exception:
                pass
            for locator in [
                frame.get_by_role("link", name=pattern),
                frame.get_by_role("button", name=pattern),
                frame.locator("a, button, [role='button']").filter(has_text=pattern),
            ]:
                try:
                    for index in range(locator.count() - 1, -1, -1):
                        match = locator.nth(index)
                        if match.is_visible():
                            print("Found the Lytx Download CSV control.")
                            return match
                except Exception:
                    continue
        time.sleep(1)
    raise TimeoutError("The Download CSV control did not become visible.")


def current_artifacts(folder: Path, started_at: float) -> list[Path]:
    artifacts = []
    for path in folder.iterdir():
        try:
            if path.is_file() and path.stat().st_mtime >= started_at - 2 and (
                path.name.lower().endswith(".csv")
                or path.name.lower().endswith(".csv.crdownload")
            ):
                artifacts.append(path)
        except OSError:
            pass
    return artifacts


def count_csv_rows(path: Path) -> tuple[list[str], int]:
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as csv_file:
        rows = list(csv.reader(csv_file))
    if not rows:
        return [], 0
    count = sum(1 for row in rows[1:] if any(str(value).strip() for value in row))
    return rows[0], count


def wait_for_csv(folder: Path, started_at: float, expected_rows: int | None) -> Path:
    deadline = time.monotonic() + 180
    sizes: dict[Path, tuple[int, int]] = {}
    while time.monotonic() < deadline:
        for path in current_artifacts(folder, started_at):
            try:
                size = path.stat().st_size
                previous, stable = sizes.get(path, (-1, 0))
                stable = stable + 1 if size == previous else 1
                sizes[path] = (size, stable)
                required_stability = 10 if path.name.endswith(".crdownload") else 3
                if size <= 0 or stable < required_stability:
                    continue
                headers, row_count = count_csv_rows(path)
                if not headers or row_count <= 0:
                    continue
                if expected_rows is not None and row_count != expected_rows:
                    continue
                if path.name.endswith(".crdownload"):
                    final_path = path.with_name(path.name[:-len(".crdownload")])
                    try:
                        path.replace(final_path)
                    except OSError:
                        shutil.copy2(path, final_path)
                    path = final_path
                    print("Recovered Chrome's complete stable CSV download.")
                return path
            except (OSError, csv.Error):
                continue
        time.sleep(1)
    names = [path.name for path in current_artifacts(folder, started_at)]
    raise TimeoutError(f"The Driver Report CSV did not finish downloading. Files: {names}")


def download_report(config: dict) -> Path:
    profile = project_path(config["browser_profile_folder"])
    temp = project_path(config["temporary_folder"])
    profile.mkdir(parents=True, exist_ok=True)
    temp.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile), channel="chrome", headless=False,
            accept_downloads=True, downloads_path=str(temp),
        )
        try:
            page = get_lytx_page(context)
            if "driversreport" not in page.url.lower():
                page.goto(config["drivers_report_url"], wait_until="domcontentloaded", timeout=180_000)
            if "login.lytx.com" in page.url.lower():
                print()
                print("Lytx authentication is required.")
                print("Complete sign-in and MFA, then manually open:")
                print(config["drivers_report_url"])
                print("Press Enter after the Drivers Report is fully visible.")
                input()
            page = wait_for_driver_report(context)
            choose_last_30_days(page)
            link = find_download_link(page, config["download_link_text"])
            label = link.inner_text().strip()
            match = re.search(r"\((\d+)\)", label)
            expected_rows = int(match.group(1)) if match else None
            if expected_rows:
                print(f"Lytx reports {expected_rows} drivers for this download.")

            session = context.new_cdp_session(link.page)
            session.send("Browser.setDownloadBehavior", {
                "behavior": "allow", "downloadPath": str(temp.resolve()), "eventsEnabled": True,
            })
            started_at = time.time()
            print("Downloading the Lytx Driver Report CSV...")
            link.scroll_into_view_if_needed()
            link.click(timeout=30_000)

            start_deadline = time.monotonic() + 20
            while time.monotonic() < start_deadline and not current_artifacts(temp, started_at):
                time.sleep(1)
            if not current_artifacts(temp, started_at):
                raise RuntimeError("The Download CSV control was clicked, but no download file started.")
            print("The Lytx Driver Report download has started.")
            csv_path = wait_for_csv(temp, started_at, expected_rows)
            print(f"Downloaded {csv_path.name} successfully.")
            return csv_path
        finally:
            try:
                context.close()
            except Exception:
                pass


def save_report(source: Path, destination_folder: Path) -> Path:
    if not destination_folder.exists():
        raise FileNotFoundError(f"Destination folder is unavailable: {destination_folder}")
    destination = destination_folder / source.name
    descriptor, temporary_name = tempfile.mkstemp(
        prefix="lytx_driver_report_", suffix=".tmp", dir=str(destination_folder)
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(source, temporary)
        os.replace(temporary, destination)
    finally:
        if temporary.exists():
            temporary.unlink()
    return destination


def parse_arguments():
    parser = argparse.ArgumentParser(description="Export the Lytx Last 30 Days Driver Report.")
    parser.add_argument("--setup-login", action="store_true")
    return parser.parse_args()


def main() -> int:
    log_path = configure_logging()
    try:
        config = load_config()
        args = parse_arguments()
        profile = project_path(config["browser_profile_folder"])
        if args.setup_login:
            setup_login(profile, config["drivers_report_url"])
            logging.info("Lytx login setup completed.")
            print("Lytx login profile configured successfully.")
            return 0

        logging.info("Starting Lytx Driver Report export.")
        downloaded = download_report(config)
        headers, row_count = count_csv_rows(downloaded)
        if not headers or row_count <= 0:
            raise RuntimeError("The downloaded Driver Report CSV is empty.")
        logging.info("Validated %s Driver Report rows.", row_count)
        destination = save_report(downloaded, Path(config["destination_folder"]))
        logging.info("Saved the Lytx Driver Report to %s.", destination)
        print()
        print("VCI Lytx Driver Report export completed successfully.")
        print(f"Rows downloaded: {row_count}")
        print(f"CSV saved to: {destination}")
        logging.info("Log saved to %s.", log_path)
        return 0
    except Exception:
        logging.exception("The VCI Lytx Driver Report export failed.")
        logging.info("The existing destination file was not intentionally changed after the failure.")
        logging.info("Review the log file: %s", log_path)
        print()
        print("VCI Lytx Driver Report export failed.")
        print("Review the latest file in the logs folder.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())

