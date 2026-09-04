import re
import time
from pathlib import Path

from playwright.sync_api import (
    BrowserContext,
    Locator,
    Page,
    sync_playwright,
)


DOWNLOAD_TIMEOUT_SECONDS = 180
DOWNLOAD_STABLE_CHECKS = 3
DOWNLOAD_POLL_SECONDS = 1


def get_or_create_lytx_page(
    browser_context: BrowserContext,
) -> Page:
    """Reuse the best existing Lytx page before creating a new one."""
    available_pages = [
        page
        for page in browser_context.pages
        if not page.is_closed()
    ]

    for page in reversed(available_pages):
        if "admin/vehicles" in page.url.lower():
            print("Reusing the existing Lytx Admin Vehicles tab...")
            return page

    for page in reversed(available_pages):
        if "lytx.com" in page.url.lower():
            print("Reusing the existing Lytx tab...")
            return page

    if available_pages:
        print("Reusing the existing automation browser tab...")
        return available_pages[0]

    print("Opening a new automation browser tab...")
    return browser_context.new_page()


def setup_lytx_login(browser_profile_path: Path, vehicles_url: str) -> None:
    """Open Chrome for the authorized Lytx login session."""
    browser_profile_path.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser_context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(browser_profile_path),
            channel="chrome",
            headless=False,
            accept_downloads=True,
        )
        page = get_or_create_lytx_page(browser_context)
        page.goto(
            vehicles_url,
            wait_until="domcontentloaded",
            timeout=120_000,
        )

        print()
        print("Google Chrome is now open for Lytx.")
        print(
            "Sign in with the authorized VCI Lytx account and "
            "complete MFA if requested."
        )
        print(
            "Navigate to the Admin Vehicles page and confirm that "
            "the vehicle list is visible."
        )
        print(
            "Return to this terminal and press Enter only after "
            "the vehicle list is displayed."
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


def find_download_link(
    page: Page,
    download_link_text: str,
    timeout: int = 180_000,
) -> Locator:
    """Scroll all page containers until Download CSV renders."""
    download_pattern = re.compile(
        rf"{re.escape(download_link_text)}(?:\s*\(\d+\))?",
        re.IGNORECASE,
    )
    deadline = time.monotonic() + timeout / 1000
    last_url = ""

    while time.monotonic() < deadline:
        available_pages = [
            candidate_page
            for candidate_page in page.context.pages
            if not candidate_page.is_closed()
            and candidate_page.url != "about:blank"
        ]

        for candidate_page in reversed(available_pages):
            last_url = candidate_page.url

            try:
                candidate_page.wait_for_timeout(500)
            except Exception:
                continue

            for frame in candidate_page.frames:
                try:
                    frame.evaluate(
                        """
                        () => {
                            window.scrollTo(0, document.body.scrollHeight);

                            const elements =
                                document.querySelectorAll("*");

                            for (const element of elements) {
                                const style =
                                    window.getComputedStyle(element);
                                const canScroll =
                                    element.scrollHeight >
                                    element.clientHeight;

                                if (
                                    canScroll &&
                                    (
                                        style.overflowY === "auto" ||
                                        style.overflowY === "scroll"
                                    )
                                ) {
                                    element.scrollTop =
                                        element.scrollHeight;
                                }
                            }
                        }
                        """
                    )
                except Exception:
                    pass

                candidate_locators = [
                    frame.get_by_role(
                        "link",
                        name=download_pattern,
                    ),
                    frame.get_by_role(
                        "button",
                        name=download_pattern,
                    ),
                    frame.locator(
                        "a, button, [role='button']"
                    ).filter(has_text=download_pattern),
                    frame.get_by_text(download_pattern),
                ]

                for locator in candidate_locators:
                    try:
                        for index in range(
                            locator.count() - 1,
                            -1,
                            -1,
                        ):
                            match = locator.nth(index)

                            if match.is_visible():
                                print(
                                    "Found the Lytx Download CSV "
                                    "control."
                                )
                                return match
                    except Exception:
                        continue

        time.sleep(1)

    raise TimeoutError(
        "The Download CSV control was not found within "
        f"{timeout // 1000} seconds. Last page URL: {last_url}"
    )


def wait_for_vehicle_page(
    browser_context: BrowserContext,
    timeout: int = 60_000,
) -> Page:
    """Return the page displaying the Lytx Vehicle Management UI."""
    deadline = time.monotonic() + timeout / 1000
    last_urls = []

    while time.monotonic() < deadline:
        available_pages = [
            candidate_page
            for candidate_page in browser_context.pages
            if not candidate_page.is_closed()
            and candidate_page.url != "about:blank"
        ]
        last_urls = [
            candidate_page.url
            for candidate_page in available_pages
        ]

        for candidate_page in reversed(available_pages):
            if "admin/vehicles" in candidate_page.url.lower():
                return candidate_page

            try:
                vehicle_heading = candidate_page.get_by_text(
                    re.compile(
                        r"VEHICLE\s+MANAGEMENT",
                        re.IGNORECASE,
                    )
                ).first

                if (
                    vehicle_heading.count() > 0
                    and vehicle_heading.is_visible()
                ):
                    return candidate_page
            except Exception:
                continue

        time.sleep(1)

    raise RuntimeError(
        "The Lytx Vehicle Management page was not detected after "
        f"{timeout // 1000} seconds. Open pages: {last_urls}"
    )


def wait_for_downloaded_csv(
    temporary_folder: Path,
    expected_filename_pattern: str,
    download_started_at: float,
    timeout_seconds: int = DOWNLOAD_TIMEOUT_SECONDS,
) -> Path:
    """Wait until Chrome finishes writing the expected Lytx CSV."""
    deadline = time.monotonic() + timeout_seconds
    stable_sizes: dict[Path, tuple[int, int]] = {}

    while time.monotonic() < deadline:
        candidates = []

        for candidate in temporary_folder.glob("*.csv"):
            try:
                candidate_name = candidate.name
                candidate_stat = candidate.stat()

                if not re.fullmatch(
                    expected_filename_pattern,
                    candidate_name,
                    flags=re.IGNORECASE,
                ):
                    continue

                if candidate_stat.st_mtime < download_started_at - 2:
                    continue

                if candidate_stat.st_size <= 0:
                    continue

                candidates.append(candidate)
            except OSError:
                continue

        candidates.sort(
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )

        for candidate in candidates:
            if candidate.with_suffix(candidate.suffix + ".crdownload").exists():
                continue

            current_size = candidate.stat().st_size
            previous_size, stable_count = stable_sizes.get(
                candidate,
                (-1, 0),
            )

            if current_size == previous_size:
                stable_count += 1
            else:
                stable_count = 1

            stable_sizes[candidate] = (
                current_size,
                stable_count,
            )

            if stable_count >= DOWNLOAD_STABLE_CHECKS:
                return candidate

        time.sleep(DOWNLOAD_POLL_SECONDS)

    partial_files = [
        path.name
        for path in temporary_folder.glob("*.crdownload")
    ]
    raise TimeoutError(
        "The Lytx CSV did not finish downloading within "
        f"{timeout_seconds} seconds. Partial downloads: "
        f"{partial_files}"
    )


def download_lytx_csv(
    vehicles_url: str,
    download_link_text: str,
    browser_profile_path: Path,
    temporary_folder: Path,
    expected_filename_pattern: str,
) -> dict:
    """Open Lytx, download the vehicle CSV, and save it to temp."""
    browser_profile_path.mkdir(parents=True, exist_ok=True)
    temporary_folder.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as playwright:
        browser_context = playwright.chromium.launch_persistent_context(
            user_data_dir=str(browser_profile_path),
            channel="chrome",
            headless=False,
            accept_downloads=True,
            downloads_path=str(temporary_folder),
        )

        try:
            page = get_or_create_lytx_page(browser_context)

            if "admin/vehicles" in page.url.lower():
                print("The Lytx Vehicles page is already open.")
                page.bring_to_front()
                page.wait_for_timeout(2_000)
            else:
                print("Opening the Lytx Vehicles page...")
                page.goto(
                    vehicles_url,
                    wait_until="domcontentloaded",
                    timeout=180_000,
                )
                page.wait_for_timeout(5_000)

            # Preserve the user's confirmed working manual login and
            # navigation path whenever Lytx requires authentication.
            if "login.lytx.com" in page.url.lower():
                print()
                print("Lytx authentication is required.")
                print(
                    "Complete sign-in and MFA in Chrome, then "
                    "manually open:"
                )
                print(vehicles_url)
                print(
                    "Press Enter here after the vehicle list is "
                    "fully visible."
                )
                input()

            try:
                page = wait_for_vehicle_page(
                    browser_context=browser_context,
                    timeout=10_000,
                )
            except RuntimeError:
                print()
                print(
                    "Navigate the open Chrome window to the Admin "
                    "Vehicles page:"
                )
                print(vehicles_url)
                print(
                    "Press Enter here after the vehicle list is "
                    "fully visible."
                )
                input()
                page = wait_for_vehicle_page(
                    browser_context=browser_context,
                    timeout=60_000,
                )

            print(
                "Waiting for the Lytx vehicle list and download "
                "control..."
            )
            download_link = find_download_link(
                page=page,
                download_link_text=download_link_text,
                timeout=180_000,
            )

            print("Downloading the Lytx vehicle CSV...")
            download_page = download_link.page

            # Direct Chrome downloads to the processing folder so the
            # file survives even if Lytx closes its page after the click.
            cdp_session = browser_context.new_cdp_session(download_page)
            cdp_session.send(
                "Browser.setDownloadBehavior",
                {
                    "behavior": "allow",
                    "downloadPath": str(temporary_folder.resolve()),
                    "eventsEnabled": True,
                },
            )

            download_started_at = time.time()
            click_error = None

            try:
                download_link.click(timeout=30_000)
            except Exception as error:
                # Lytx can close the initiating page immediately after
                # starting the download. Continue checking the folder.
                click_error = error

            try:
                temporary_csv_path = wait_for_downloaded_csv(
                    temporary_folder=temporary_folder,
                    expected_filename_pattern=expected_filename_pattern,
                    download_started_at=download_started_at,
                )
            except Exception:
                if click_error is not None:
                    raise click_error
                raise

            suggested_filename = temporary_csv_path.name

            if not temporary_csv_path.exists():
                raise FileNotFoundError(
                    "The Lytx CSV was not created."
                )

            if temporary_csv_path.stat().st_size == 0:
                raise ValueError(
                    "The downloaded Lytx CSV is empty."
                )

            print(
                f"Downloaded {suggested_filename} successfully."
            )

            return {
                "download_path": temporary_csv_path,
                "suggested_filename": suggested_filename,
            }

        except Exception as error:
            try:
                diagnostic_folder = (
                    browser_profile_path.parent / "logs"
                )
                diagnostic_folder.mkdir(
                    parents=True,
                    exist_ok=True,
                )
                diagnostic_path = (
                    diagnostic_folder
                    / "lytx_download_failure.png"
                )
                active_pages = [
                    candidate_page
                    for candidate_page in browser_context.pages
                    if not candidate_page.is_closed()
                    and candidate_page.url != "about:blank"
                ]

                if active_pages:
                    active_pages[-1].screenshot(
                        path=str(diagnostic_path),
                        full_page=True,
                    )
                    print(
                        "Diagnostic screenshot saved to "
                        f"{diagnostic_path}"
                    )
            except Exception:
                pass

            raise RuntimeError(
                "Could not download the Lytx vehicle CSV. "
                "The Lytx session may have expired, the vehicle "
                "page may not have fully loaded, or the Download "
                "CSV control may have changed."
            ) from error

        finally:
            try:
                browser_context.close()
            except Exception as error:
                if (
                    "Target page, context or browser has been closed"
                    not in str(error)
                ):
                    raise
