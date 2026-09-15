from __future__ import annotations

import json
import logging
import sys
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from browser_automation import (
    AuthenticationError,
    export_dvir_details,
    export_safety_deviations,
    upload_to_drive,
)
from workbook_processor import validate_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def load_settings() -> dict:
    with (PROJECT_ROOT / "config" / "settings.json").open(encoding="utf-8") as handle:
        return json.load(handle)


def configure_logging() -> Path:
    log_dir = PROJECT_ROOT / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"vci_dvir_safety_etl_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()],
    )
    return log_path


def launch_context(playwright, browser: dict):
    profile = (PROJECT_ROOT / browser["profile_directory"]).resolve()
    profile.mkdir(parents=True, exist_ok=True)
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str(profile),
        channel=browser.get("channel", "chrome"),
        headless=browser.get("headless", False),
        accept_downloads=True,
        slow_mo=browser.get("slow_mo_ms", 0),
    )
    context.set_default_timeout(browser.get("timeout_ms", 45_000))
    pages = [page for page in context.pages if not page.is_closed()]
    page = pages[0] if pages else context.new_page()
    for extra in pages[1:]:
        try:
            extra.close()
        except Exception:
            pass
    return context, page


def close_context(context) -> None:
    try:
        context.close()
    except PlaywrightError:
        logging.info("Chrome had already closed after its completed action.")


def run_browser_action(browser: dict, action):
    """Initialize Playwright only when browser work is actually required."""
    with sync_playwright() as playwright:
        context, page = launch_context(playwright, browser)
        try:
            return action(page)
        finally:
            close_context(context)


def validate_report(path: Path, report: dict) -> tuple[int, int]:
    return validate_workbook(
        path,
        report["expected_sheet"],
        report.get("required_headers", []),
        minimum_columns=report.get("minimum_columns"),
        data_start_row=int(report.get("data_start_row", 2)),
        first_column_value=report.get("first_column_value"),
        required_data_columns=report.get("required_data_columns"),
    )


def validate_existing_report(path: Path, report: dict) -> tuple[int, int] | None:
    if not path.exists():
        return None
    try:
        return validate_report(path, report)
    except Exception as exc:
        logging.warning("Existing file will be replaced because validation failed: %s", exc)
        return None


def main() -> int:
    log_path = configure_logging()
    settings = load_settings()
    run_date = date.today()
    output_dir = PROJECT_ROOT / settings["download_directory"] / f"{run_date:%Y-%m-%d}"
    output_dir.mkdir(parents=True, exist_ok=True)
    browser = settings["browser"]
    auth_wait = int(browser.get("authentication_wait_seconds", 300))
    test_mode = settings.get("test_mode", {})

    logging.info("Starting VCI DVIR and Safety Deviations ETL for %s", run_date)
    dvir_settings = settings["dvir_details"]
    dvir_path = output_dir / dvir_settings["output_filename"]
    existing_dvir = validate_existing_report(dvir_path, dvir_settings)
    if existing_dvir is not None:
        dvir_rows, dvir_columns = existing_dvir
        logging.info("Reusing validated DVIR export: %s", dvir_path)
    else:
        dvir_path = run_browser_action(
            browser,
            lambda page: export_dvir_details(
                page, dvir_settings, run_date, output_dir, auth_wait
            ),
        )
        dvir_rows, dvir_columns = validate_report(dvir_path, dvir_settings)
    logging.info("Validated DVIR export: %d rows, %d columns", dvir_rows, dvir_columns)

    if test_mode.get("enabled") and test_mode.get("stop_after_stage") == "dvir_export":
        logging.info("Test mode complete: stopping after DVIR export")
        logging.info("Log: %s", log_path)
        return 0

    safety_settings = settings["safety_deviations"]
    date_text = run_date.strftime(safety_settings.get("output_date_format", "%m-%d-%Y"))
    safety_filename = safety_settings["output_filename_pattern"].format(run_date=date_text)
    safety_path = output_dir / safety_filename
    existing_safety = validate_existing_report(safety_path, safety_settings)
    if existing_safety is not None:
        safety_rows, safety_columns = existing_safety
        logging.info("Reusing validated Safety Deviations export: %s", safety_path)
    else:
        safety_path = run_browser_action(
            browser,
            lambda page: export_safety_deviations(
                page, safety_settings, run_date, output_dir, auth_wait
            ),
        )
        safety_rows, safety_columns = validate_report(safety_path, safety_settings)
    logging.info(
        "Validated Safety Deviations export: %d rows, %d columns",
        safety_rows,
        safety_columns,
    )

    if test_mode.get("enabled") and test_mode.get("stop_after_stage") == "safety_export":
        logging.info("Test mode complete: stopping after Safety Deviations export")
        logging.info("Log: %s", log_path)
        return 0

    upload_marker = output_dir / "drive_upload_complete.flag"
    if upload_marker.exists():
        logging.info("Skipping Drive upload because today's validated pair is already complete.")
    else:
        run_browser_action(
            browser,
            lambda page: upload_to_drive(
                page, settings["google_drive"], [dvir_path, safety_path]
            ),
        )
        upload_marker.write_text(
            f"Uploaded {datetime.now().isoformat(timespec='seconds')}\n",
            encoding="utf-8",
        )

    if test_mode.get("enabled") and test_mode.get("stop_after_stage") == "drive_upload":
        logging.info("Test mode complete: stopping after verified Drive upload")
        logging.info("Log: %s", log_path)
        return 0

    logging.info("ETL completed successfully. Apps Script was not triggered.")
    logging.info("Log: %s", log_path)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except AuthenticationError as exc:
        logging.error("AUTHENTICATION REQUIRED: %s", exc)
        raise SystemExit(2)
    except Exception:
        logging.exception("ETL failed. No Apps Script action was triggered.")
        raise SystemExit(1)
