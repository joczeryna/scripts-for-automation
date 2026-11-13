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

EVENTS_URL = "https://app.lytx.com/#/driver-safety/library/events"
EVENTS_DATE_SHORTCUT = "Last 180 Days"
EVENTS_DESTINATION = Path(r"W:\Dashboards VCI\Safety Data\Current\Lytx-Events")


def load_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as config_file:
        return json.load(config_file)


def project_path(value: str) -> Path:
    path = Path(value)
    return path if path.is_absolute() else (PROJECT_ROOT / path).resolve()


def configure_logging() -> Path:
    LOG_FOLDER.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = LOG_FOLDER / f"lytx_events_export_{stamp}.log"
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


def open_fresh_tab(context: BrowserContext) -> Page:
    existing_pages = [p for p in context.pages if not p.is_closed()]
    page = context.new_page()
    for old_page in existing_pages:
        try:
            old_page.close()
        except Exception:
            pass
    return page


def _try_sidebar_events_navigation(page: Page) -> None:
    try:
        events_nav = page.locator("a, button, [role='menuitem'], div, span").filter(
            has_text=re.compile(r"^Events$", re.IGNORECASE)
        )
        for i in range(events_nav.count()):
            item = events_nav.nth(i)
            if item.is_visible():
                item.click(timeout=3_000)
                return

        library_nav = page.locator("a, button, [role='menuitem'], div, span").filter(
            has_text=re.compile(r"^Library$", re.IGNORECASE)
        )
        for i in range(library_nav.count()):
            item = library_nav.nth(i)
            if item.is_visible():
                item.click(timeout=3_000)
                page.wait_for_timeout(1_000)
                break

        events_nav = page.locator("a, button, [role='menuitem'], div, span").filter(
            has_text=re.compile(r"^Events$", re.IGNORECASE)
        )
        for i in range(events_nav.count()):
            item = events_nav.nth(i)
            if item.is_visible():
                item.click(timeout=3_000)
                return
    except Exception:
        pass


def navigate_to_events_page(page: Page, target_url: str, timeout_ms: int = 90_000) -> None:
    print(f"Navigating to Lytx Events page: {target_url}")
    page.goto(target_url, wait_until="domcontentloaded", timeout=180_000)

    deadline = time.monotonic() + timeout_ms / 1000
    last_url = ""
    last_nudge = 0.0
    nudge_attempts = 0

    while time.monotonic() < deadline:
        current_url = page.url
        if current_url != last_url:
            logging.info("Current browser URL: %s", current_url)
            last_url = current_url

        if "login.lytx.com" in current_url.lower():
            print()
            print("Lytx authentication is required.")
            print("Complete sign-in and MFA, then manually open:")
            print(target_url)
            print("Press Enter after the Events page is fully visible.")
            input()
            page.goto(target_url, wait_until="domcontentloaded", timeout=180_000)
            last_nudge = time.monotonic()
            continue

        if "library/events" in current_url.lower():
            try:
                has_filter = page.get_by_text("Filter By Group", exact=False).first.is_visible()
                has_table = page.get_by_text("EVENT ID", exact=True).first.is_visible()
                if has_filter or has_table:
                    print("Lytx Events Library page loaded successfully.")
                    return
            except Exception:
                pass
        elif "app.lytx.com" in current_url.lower() and (time.monotonic() - last_nudge) > 5:
            # Lytx SPA booted onto another route (e.g., Home or Drivers Report); force Events route
            nudge_attempts += 1
            last_nudge = time.monotonic()
            logging.info("Redirecting SPA to Events Library (attempt %s)...", nudge_attempts)
            if nudge_attempts % 2 == 1:
                try:
                    page.evaluate(
                        """(url) => {
                            window.location.href = url;
                            window.dispatchEvent(new HashChangeEvent('hashchange'));
                        }""",
                        target_url,
                    )
                except Exception:
                    page.goto(target_url, wait_until="domcontentloaded", timeout=60_000)
            else:
                _try_sidebar_events_navigation(page)

        time.sleep(1)

    raise RuntimeError(f"The Lytx Events page was not detected. Last URL: {page.url}")


def setup_login(profile: Path, url: str) -> None:
    profile.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile), channel="chrome", headless=False,
            accept_downloads=True,
        )
        page = open_fresh_tab(context)
        page.goto(url, wait_until="domcontentloaded", timeout=180_000)
        print()
        print("Chrome is open for Lytx authentication.")
        print("Sign in, complete MFA, and manually open:")
        print(url)
        print("Press Enter after the Events page is fully visible.")
        input()
        try:
            context.close()
        except Exception:
            pass


def read_download_label(page: Page, label: str) -> str | None:
    pattern = re.compile(rf"{re.escape(label)}(?:\s*\(\d+\))?", re.IGNORECASE)
    for frame in page.frames:
        try:
            locator = frame.locator("a, button, [role='button'], span").filter(has_text=pattern)
            if locator.count() > 0:
                return locator.last.inner_text().strip()
        except Exception:
            continue
    return None


def choose_date_range(page: Page, shortcut_text: str, download_label: str) -> None:
    print(f"Selecting the {shortcut_text} events period...")
    shortcut = page.get_by_text(shortcut_text, exact=True)
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
            raise RuntimeError("Could not open the Lytx date-range selector on the Events page.")

    previous_label = read_download_label(page, download_label)

    shortcut = page.get_by_text(shortcut_text, exact=True)
    shortcut.wait_for(state="visible", timeout=30_000)
    shortcut.click()

    apply_button = page.get_by_role("button", name="Apply", exact=True)
    apply_button.wait_for(state="visible", timeout=30_000)
    apply_button.click()

    # Wait for the default 7-day event count to update to the 180-day count
    if previous_label:
        refresh_deadline = time.monotonic() + 25
        while time.monotonic() < refresh_deadline:
            current_label = read_download_label(page, download_label)
            if current_label and current_label != previous_label:
                break
            time.sleep(1)

    page.wait_for_timeout(3_000)
    print(f"Applied {shortcut_text} successfully.")


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
                frame.locator("a, button, [role='button'], span").filter(has_text=pattern),
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
    deadline = time.monotonic() + 240
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
    raise TimeoutError(f"The Events CSV did not finish downloading. Files: {names}")


def download_events(config: dict) -> Path:
    profile = project_path(config["browser_profile_folder"])
    temp = project_path(config["temporary_folder"])
    profile.mkdir(parents=True, exist_ok=True)
    temp.mkdir(parents=True, exist_ok=True)

    target_url = config.get("events_url", EVENTS_URL)
    shortcut_text = config.get("events_date_shortcut", EVENTS_DATE_SHORTCUT)
    link_text = config.get("download_link_text", "Download CSV")

    with sync_playwright() as playwright:
        context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(profile), channel="chrome", headless=False,
            accept_downloads=True, downloads_path=str(temp),
        )
        try:
            page = open_fresh_tab(context)
            navigate_to_events_page(page, target_url)
            choose_date_range(page, shortcut_text, link_text)

            link = find_download_link(page, link_text)
            label = link.inner_text().strip()
            match = re.search(r"\((\d+)\)", label)
            expected_rows = int(match.group(1)) if match else None
            if expected_rows:
                print(f"Lytx reports {expected_rows} events for this download.")

            session = context.new_cdp_session(link.page)
            session.send("Browser.setDownloadBehavior", {
                "behavior": "allow", "downloadPath": str(temp.resolve()), "eventsEnabled": True,
            })
            started_at = time.time()
            print("Downloading the Lytx Events CSV...")
            link.scroll_into_view_if_needed()
            link.click(timeout=30_000)

            start_deadline = time.monotonic() + 60
            while time.monotonic() < start_deadline and not current_artifacts(temp, started_at):
                time.sleep(1)
            if not current_artifacts(temp, started_at):
                raise RuntimeError("The Download CSV control was clicked, but no download file started.")
            print("The Lytx Events download has started.")
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
        prefix="lytx_events_", suffix=".tmp", dir=str(destination_folder)
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
    parser = argparse.ArgumentParser(description="Export the Lytx Last 180 Days Events CSV.")
    parser.add_argument("--setup-login", action="store_true")
    return parser.parse_args()


def main() -> int:
    log_path = configure_logging()
    try:
        config = load_config()
        args = parse_arguments()
        profile = project_path(config["browser_profile_folder"])
        target_url = config.get("events_url", EVENTS_URL)
        dest_folder = Path(config.get("events_destination_folder", str(EVENTS_DESTINATION)))

        if args.setup_login:
            setup_login(profile, target_url)
            logging.info("Lytx login setup completed.")
            print("Lytx login profile configured successfully.")
            return 0

        logging.info("Starting Lytx Events export.")
        downloaded = download_events(config)
        headers, row_count = count_csv_rows(downloaded)
        if not headers or row_count <= 0:
            raise RuntimeError("The downloaded Events CSV is empty.")
        logging.info("Validated %s Events rows.", row_count)
        destination = save_report(downloaded, dest_folder)
        logging.info("Saved the Lytx Events CSV to %s.", destination)
        print()
        print("VCI Lytx Events export completed successfully.")
        print(f"Rows downloaded: {row_count}")
        print(f"CSV saved to: {destination}")
        logging.info("Log saved to %s.", log_path)
        return 0
    except Exception:
        logging.exception("The VCI Lytx Events export failed.")
        logging.info("The existing destination file was not intentionally changed after the failure.")
        logging.info("Review the log file: %s", log_path)
        print()
        print("VCI Lytx Events export failed.")
        print("Review the latest file in the logs folder.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())