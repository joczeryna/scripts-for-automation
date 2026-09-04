import argparse
import json
import logging
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from uuid import uuid4

from google_download import download_reports, setup_google_login
from workbook_processor import clean_workbook


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_PATH = PROJECT_ROOT / "config" / "reports.json"
LOG_FOLDER = PROJECT_ROOT / "logs"


def load_config() -> dict:
    """Load and validate the project configuration."""
    if not CONFIG_PATH.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {CONFIG_PATH}"
        )

    with CONFIG_PATH.open(
        "r",
        encoding="utf-8",
    ) as config_file:
        config = json.load(
            config_file
        )

    required_settings = {
        "destination_folder",
        "browser_profile_folder",
        "temporary_folder",
        "worksheet_date_format",
        "required_headers",
        "reports",
    }

    missing_settings = required_settings.difference(
        config
    )

    if missing_settings:
        raise ValueError(
            "Missing configuration settings: "
            + ", ".join(
                sorted(
                    missing_settings
                )
            )
        )

    if not config["reports"]:
        raise ValueError(
            "No reports are configured."
        )

    return config


def configure_logging() -> Path:
    """Write run details to both the terminal and a log file."""
    LOG_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    log_path = (
        LOG_FOLDER
        / (
            "hr_report_export_"
            + datetime.now().strftime(
                "%Y%m%d_%H%M%S"
            )
            + ".log"
        )
    )

    logging.basicConfig(
        level=logging.INFO,
        format=(
            "%(asctime)s | %(levelname)s | %(message)s"
        ),
        handlers=[
            logging.FileHandler(
                log_path,
                encoding="utf-8",
            ),
            logging.StreamHandler(
                sys.stdout
            ),
        ],
    )

    return log_path


def replace_destination_file(
    cleaned_file: Path,
    destination_file: Path,
) -> None:
    """
    Stage the validated workbook in the destination folder and then
    replace the production file atomically.
    """
    destination_file.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    staged_file = destination_file.with_name(
        "."
        + destination_file.name
        + "."
        + uuid4().hex
        + ".tmp"
    )

    try:
        shutil.copy2(
            cleaned_file,
            staged_file,
        )

        if not staged_file.exists():
            raise FileNotFoundError(
                f"Could not stage {destination_file.name}."
            )

        if staged_file.stat().st_size == 0:
            raise ValueError(
                f"The staged file {destination_file.name} is empty."
            )

        os.replace(
            staged_file,
            destination_file,
        )

    except PermissionError as error:
        raise PermissionError(
            f"Could not replace {destination_file}. "
            "Close the workbook in Excel and confirm the W drive "
            "is available."
        ) from error

    finally:
        if staged_file.exists():
            staged_file.unlink()


def run_setup(
    config: dict,
) -> None:
    """Create the authenticated browser profile."""
    browser_profile_path = (
        PROJECT_ROOT
        / config["browser_profile_folder"]
    )

    first_spreadsheet_id = (
        config["reports"][0]["spreadsheet_id"]
    )

    setup_google_login(
        browser_profile_path=browser_profile_path,
        spreadsheet_id=first_spreadsheet_id,
    )


def run_export(
    config: dict,
) -> None:
    """Download, validate, clean, and replace both HR reports."""
    destination_folder = Path(
        config["destination_folder"]
    )

    if not destination_folder.exists():
        raise FileNotFoundError(
            "The destination folder is unavailable: "
            f"{destination_folder}. Confirm that the W drive "
            "is connected."
        )

    browser_profile_path = (
        PROJECT_ROOT
        / config["browser_profile_folder"]
    )

    temporary_root = (
        PROJECT_ROOT
        / config["temporary_folder"]
    )

    temporary_root.mkdir(
        parents=True,
        exist_ok=True,
    )

    logging.info(
        "Starting VCI HR report export."
    )

    with tempfile.TemporaryDirectory(
        prefix="vci_hr_reports_",
        dir=temporary_root,
    ) as temporary_directory:
        run_folder = Path(
            temporary_directory
        )

        download_folder = (
            run_folder
            / "downloads"
        )

        cleaned_folder = (
            run_folder
            / "cleaned"
        )

        downloads = download_reports(
            reports=config["reports"],
            download_folder=download_folder,
            browser_profile_path=browser_profile_path,
        )

        cleaned_reports = []

        for download_result in downloads:
            report = download_result["report"]
            downloaded_file = download_result["download_path"]

            cleaned_file = (
                cleaned_folder
                / report["output_filename"]
            )

            logging.info(
                "Cleaning %s.",
                report["name"],
            )

            cleanup_result = clean_workbook(
                source_path=downloaded_file,
                output_path=cleaned_file,
                date_format=config["worksheet_date_format"],
                required_headers=config["required_headers"],
            )

            logging.info(
                "%s retained worksheet %s with headers on row %s.",
                report["name"],
                cleanup_result["latest_worksheet"],
                cleanup_result["header_row"],
            )

            cleaned_reports.append(
                {
                    "report": report,
                    "cleaned_file": cleaned_file,
                    "cleanup_result": cleanup_result,
                }
            )

        # Files are replaced only after both reports download,
        # clean, and validate successfully.
        for cleaned_report in cleaned_reports:
            report = cleaned_report["report"]

            destination_file = (
                destination_folder
                / report["output_filename"]
            )

            logging.info(
                "Replacing %s.",
                destination_file,
            )

            replace_destination_file(
                cleaned_file=cleaned_report["cleaned_file"],
                destination_file=destination_file,
            )

            logging.info(
                "Updated %s successfully.",
                destination_file,
            )

    logging.info(
        "Both VCI HR reports were updated successfully."
    )


def parse_arguments():
    parser = argparse.ArgumentParser(
        description=(
            "Download and prepare the VCI Active and Terminated "
            "Employee reports."
        )
    )

    parser.add_argument(
        "--setup-login",
        action="store_true",
        help=(
            "Open Chrome and create the authorized Google session."
        ),
    )

    return parser.parse_args()


def main() -> int:
    log_path = configure_logging()

    try:
        config = load_config()
        arguments = parse_arguments()

        if arguments.setup_login:
            run_setup(
                config
            )

            logging.info(
                "Google login setup completed."
            )
        else:
            run_export(
                config
            )

        logging.info(
            "Log saved to %s.",
            log_path,
        )

        return 0

    except Exception:
        logging.exception(
            "The VCI HR report automation failed."
        )

        logging.info(
            "The existing production files were not intentionally "
            "changed after the failure."
        )

        logging.info(
            "Review the log file: %s",
            log_path,
        )

        return 1


if __name__ == "__main__":
    raise SystemExit(
        main()
    )