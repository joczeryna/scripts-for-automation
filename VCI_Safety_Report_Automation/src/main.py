from __future__ import annotations

import json
import logging
import sys
from copy import deepcopy
from datetime import date, datetime
from pathlib import Path

from playwright.sync_api import Error as PlaywrightError, sync_playwright

from browser_automation import export_power_bi_report, update_executive_dashboard, upload_to_drive
from workbook_processor import clean_dvir_date_window, validate_workbook


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def load_settings() -> dict:
    with (PROJECT_ROOT / "config" / "settings.json").open(encoding="utf-8") as handle:
        return json.load(handle)


def configure_logging() -> None:
    log_dir = PROJECT_ROOT / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"vci_safety_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()],
    )


def find_last_export_date(download_root: Path, report_name: str, run_date: date) -> date:
    candidates: list[date] = []
    if download_root.exists():
        for folder in download_root.iterdir():
            if not folder.is_dir():
                continue
            try:
                folder_date = datetime.strptime(folder.name, "%Y-%m-%d").date()
            except ValueError:
                continue
            if folder_date >= run_date:
                continue
            if any(folder.glob(f"{report_name}_*.xlsx")):
                candidates.append(folder_date)

    if candidates:
        return max(candidates)

    logging.warning(
        "No prior DVIR Details export was found; using the previous day as the initial boundary"
    )
    return run_date.fromordinal(run_date.toordinal() - 1)


def choose_power_bi_period(start_date: date, run_date: date) -> str:
    gap_days = max(1, (run_date - start_date).days)
    for maximum, label in (
        (1, "Past 1 day"),
        (7, "Past 7 days"),
        (30, "Past 30 days"),
        (60, "Past 60 days"),
        (90, "Past 90 days"),
    ):
        if gap_days <= maximum:
            return label
    raise RuntimeError(
        f"The last DVIR Details export is {gap_days} days old, beyond the 90-day report range"
    )


def prepare_runtime_report(
    report: dict,
    download_root: Path,
    run_date: date,
) -> tuple[dict, date | None]:
    runtime_report = deepcopy(report)
    if report["key"] != "dvir_details":
        return runtime_report, None

    last_export_date = find_last_export_date(download_root, report["name"], run_date)
    period = choose_power_bi_period(last_export_date, run_date)
    for item in runtime_report["filters"]:
        if item["label"] == "Days":
            item.pop("value_by_day", None)
            item["value"] = period
            break

    logging.info(
        "DVIR Details window: last export %s, run date %s, Power BI period %s",
        last_export_date,
        run_date,
        period,
    )
    return runtime_report, last_export_date


def launch_browser_context(playwright, browser: dict):
    context = playwright.chromium.launch_persistent_context(
        user_data_dir=str((PROJECT_ROOT / browser["profile_directory"]).resolve()),
        channel=browser.get("channel", "chrome"),
        headless=browser.get("headless", False),
        accept_downloads=True,
        slow_mo=browser.get("slow_mo_ms", 0),
    )
    context.set_default_timeout(browser.get("timeout_ms", 45000))
    page = context.pages[0] if context.pages else context.new_page()
    return context, page


def close_browser_context(context) -> None:
    try:
        context.close()
    except PlaywrightError:
        logging.info("Browser session had already closed after its completed action")


def main() -> int:
    configure_logging()
    settings = load_settings()
    run_date = date.today()
    download_root = PROJECT_ROOT / settings["download_directory"]
    output_dir = download_root / f"{run_date:%Y-%m-%d}"
    output_dir.mkdir(parents=True, exist_ok=True)
    browser = settings["browser"]

    logging.info("Starting VCI Safety automation for %s", run_date)
    downloaded: list[Path] = []

    with sync_playwright() as playwright:
        for configured_report in settings["reports"]:
            report, cleanup_start_date = prepare_runtime_report(
                configured_report, download_root, run_date
            )
            existing_path = output_dir / f"{report['name']}_{run_date:%Y-%m-%d}.xlsx"
            reuse_existing = False
            if existing_path.exists():
                try:
                    validate_workbook(existing_path)
                    reuse_existing = True
                except Exception as exc:
                    logging.warning(
                        "Existing report is not valid and will be re-exported: %s (%s)",
                        existing_path,
                        exc,
                    )

            if reuse_existing:
                path = existing_path
                logging.info("Reusing validated report already downloaded today: %s", path)
            else:
                context, page = launch_browser_context(playwright, browser)
                try:
                    path = export_power_bi_report(page, report, run_date, output_dir)
                finally:
                    close_browser_context(context)
                validate_workbook(path)

            cleanup = report.get("gap_cleanup", {})
            if cleanup.get("enabled") and cleanup_start_date:
                removed = clean_dvir_date_window(
                    path, cleanup["date_header"], cleanup_start_date, run_date
                )
                logging.info("Gap cleanup removed %d rows from %s", removed, path.name)
                validate_workbook(path)

            downloaded.append(path)
            test_mode = settings.get("test_mode", {})
            if (
                test_mode.get("enabled")
                and test_mode.get("stop_after_report") == report["key"]
            ):
                logging.info("Test mode complete: stopping after %s", report["name"])
                return 0

        if len(downloaded) != 4:
            raise RuntimeError(f"Expected 4 reports but downloaded {len(downloaded)}")

        upload_marker = output_dir / "drive_upload_complete.flag"
        dashboard_marker = output_dir / "safety_compliance_triggered.flag"
        if not (upload_marker.exists() and dashboard_marker.exists()):
            context, page = launch_browser_context(playwright, browser)
            try:
                if upload_marker.exists():
                    logging.info("Skipping Drive upload because today's upload is complete")
                else:
                    upload_to_drive(page, settings["google_drive"], downloaded)
                    upload_marker.write_text(
                        f"Uploaded {datetime.now().isoformat(timespec='seconds')}\n",
                        encoding="utf-8",
                    )

                test_mode = settings.get("test_mode", {})
                if (
                    test_mode.get("enabled")
                    and test_mode.get("stop_after_stage") == "drive_upload"
                ):
                    logging.info("Test mode complete: stopping after Drive upload")
                    return 0

                if dashboard_marker.exists():
                    logging.info("Skipping Safety Compliance Report because it is complete")
                else:
                    update_executive_dashboard(
                        page, settings["executive_dashboard"], run_date
                    )
                    dashboard_marker.write_text(
                        f"Triggered {datetime.now().isoformat(timespec='seconds')}\n",
                        encoding="utf-8",
                    )
            finally:
                close_browser_context(context)
        else:
            logging.info("Drive upload and Safety Compliance Report are already complete")

    logging.info("Automation completed successfully")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception:
        logging.exception("Automation failed")
        raise SystemExit(1)
