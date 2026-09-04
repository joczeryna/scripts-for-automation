import argparse
import json
import logging
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path

from csv_processor import process_vehicle_csv
from google_sheet_update import (
    setup_google_login,
    update_google_sheet,
)
from lytx_download import (
    download_lytx_csv,
    setup_lytx_login,
)


PROJECT_ROOT = Path(
    __file__
).resolve().parent.parent

CONFIG_PATH = (
    PROJECT_ROOT
    / "config"
    / "settings.json"
)

LOG_FOLDER = (
    PROJECT_ROOT
    / "logs"
)


def configure_logging() -> Path:
    """
    Configure terminal and file logging for the current run.
    """
    LOG_FOLDER.mkdir(
        parents=True,
        exist_ok=True,
    )

    timestamp = datetime.now().strftime(
        "%Y%m%d_%H%M%S"
    )

    log_path = (
        LOG_FOLDER
        / f"lytx_vehicle_sync_{timestamp}.log"
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
        force=True,
    )

    return log_path


def load_config() -> dict:
    """
    Read and validate the project configuration.
    """
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

    required_sections = [
        "lytx",
        "google_sheet",
        "column_mapping",
    ]

    missing_sections = [
        section
        for section in required_sections
        if section not in config
    ]

    if missing_sections:
        raise ValueError(
            "The configuration file is missing these sections: "
            + ", ".join(
                missing_sections
            )
        )

    return config


def resolve_project_path(
    path_value: str,
) -> Path:
    """
    Resolve a project-relative path into an absolute path.
    Absolute Windows paths remain unchanged.
    """
    configured_path = Path(
        path_value
    )

    if configured_path.is_absolute():
        return configured_path

    return (
        PROJECT_ROOT
        / configured_path
    ).resolve()


def copy_csv_to_destination(
    source_csv_path: Path,
    destination_folder: Path,
    destination_filename: str,
) -> Path:
    """
    Copy the validated CSV into its final destination using a
    temporary file and atomic replacement.
    """
    destination_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    destination_path = (
        destination_folder
        / destination_filename
    )

    temporary_file_descriptor = None
    temporary_destination_path = None

    try:
        (
            temporary_file_descriptor,
            temporary_filename,
        ) = tempfile.mkstemp(
            prefix="lytx_vehicle_",
            suffix=".tmp",
            dir=str(
                destination_folder
            ),
        )

        os.close(
            temporary_file_descriptor
        )
        temporary_file_descriptor = None

        temporary_destination_path = Path(
            temporary_filename
        )

        shutil.copy2(
            source_csv_path,
            temporary_destination_path,
        )

        if (
            not temporary_destination_path.exists()
            or temporary_destination_path.stat().st_size == 0
        ):
            raise RuntimeError(
                "The temporary destination CSV was not created "
                "successfully."
            )

        os.replace(
            temporary_destination_path,
            destination_path,
        )

        logging.info(
            "Saved the Lytx CSV to %s.",
            destination_path,
        )

        return destination_path

    finally:
        if temporary_file_descriptor is not None:
            os.close(
                temporary_file_descriptor
            )

        if (
            temporary_destination_path is not None
            and temporary_destination_path.exists()
        ):
            temporary_destination_path.unlink()


def run_lytx_login_setup(
    config: dict,
) -> None:
    """
    Create or refresh the dedicated authenticated Lytx profile.
    """
    lytx_config = config["lytx"]

    browser_profile_path = resolve_project_path(
        lytx_config["browser_profile_folder"]
    )

    setup_lytx_login(
        browser_profile_path=browser_profile_path,
        vehicles_url=lytx_config["vehicles_url"],
    )

    logging.info(
        "Lytx login setup completed."
    )


def run_google_login_setup(
    config: dict,
) -> None:
    """
    Create or refresh the dedicated authenticated Google profile.
    """
    google_config = config["google_sheet"]

    browser_profile_path = resolve_project_path(
        google_config[
            "browser_profile_folder"
        ]
    )

    setup_google_login(
        browser_profile_path=browser_profile_path,
        spreadsheet_id=google_config[
            "spreadsheet_id"
        ],
        worksheet_name=google_config[
            "worksheet_name"
        ],
    )

    logging.info(
        "Google Sheets login setup completed."
    )


def run_sync(
    config: dict,
) -> dict:
    """
    Run the full Lytx CSV to Google Sheets workflow.
    """
    lytx_config = config["lytx"]
    google_config = config["google_sheet"]
    column_mapping = config["column_mapping"]

    lytx_browser_profile = resolve_project_path(
        lytx_config["browser_profile_folder"]
    )

    google_browser_profile = resolve_project_path(
        google_config[
            "browser_profile_folder"
        ]
    )

    temporary_folder = resolve_project_path(
        config.get(
            "temporary_folder",
            "temp",
        )
    )

    destination_folder = Path(
        lytx_config["destination_folder"]
    )

    target_headers = list(
        column_mapping.values()
    )

    logging.info(
        "Starting Lytx vehicle synchronization."
    )

    download_result = download_lytx_csv(
        vehicles_url=lytx_config[
            "vehicles_url"
        ],
        download_link_text=lytx_config[
            "download_link_text"
        ],
        browser_profile_path=(
            lytx_browser_profile
        ),
        temporary_folder=temporary_folder,
        expected_filename_pattern=(
            lytx_config[
                "expected_filename_pattern"
            ]
        ),
    )

    downloaded_csv_path = Path(
        download_result["download_path"]
    )

    downloaded_filename = download_result[
        "suggested_filename"
    ]

    logging.info(
        "Validating and transforming %s.",
        downloaded_filename,
    )

    processed_result = process_vehicle_csv(
        csv_path=downloaded_csv_path,
        column_mapping=column_mapping,
    )

    vehicle_rows = processed_result["rows"]

    if not vehicle_rows:
        raise RuntimeError(
            "The downloaded CSV did not contain any vehicle rows. "
            "The Google Sheet was not changed."
        )

    logging.info(
        "Validated %s vehicle rows.",
        len(vehicle_rows),
    )

    google_result = update_google_sheet(
        spreadsheet_id=google_config[
            "spreadsheet_id"
        ],
        worksheet_name=google_config[
            "worksheet_name"
        ],
        browser_profile_path=(
            google_browser_profile
        ),
        target_headers=target_headers,
        rows=vehicle_rows,
        data_start_row=google_config.get(
            "data_start_row",
            2,
        ),
    )

    final_csv_path = copy_csv_to_destination(
        source_csv_path=downloaded_csv_path,
        destination_folder=destination_folder,
        destination_filename=downloaded_filename,
    )

    logging.info(
        "Lytx vehicle synchronization completed successfully."
    )

    return {
        "source_rows": len(vehicle_rows),
        "google_rows_written": google_result[
            "rows_written"
        ],
        "final_csv_path": final_csv_path,
        "worksheet_name": google_result[
            "worksheet_name"
        ],
    }


def parse_arguments():
    """
    Read supported command-line modes.
    """
    parser = argparse.ArgumentParser(
        description=(
            "Download the Lytx vehicle CSV and replace the "
            "configured Google Sheet vehicle data."
        )
    )

    parser.add_argument(
        "--setup-lytx",
        action="store_true",
        help=(
            "Open Chrome to establish or refresh the authorized "
            "Lytx login session."
        ),
    )

    parser.add_argument(
        "--setup-google",
        action="store_true",
        help=(
            "Open Chrome to establish or refresh the authorized "
            "Google Sheets login session."
        ),
    )

    return parser.parse_args()


def main() -> int:
    """
    Project entry point.
    """
    log_path = configure_logging()

    try:
        config = load_config()
        arguments = parse_arguments()

        if (
            arguments.setup_lytx
            and arguments.setup_google
        ):
            run_lytx_login_setup(
                config
            )
            run_google_login_setup(
                config
            )

            print()
            print(
                "Both login profiles were configured successfully."
            )

        elif arguments.setup_lytx:
            run_lytx_login_setup(
                config
            )

            print()
            print(
                "Lytx login profile configured successfully."
            )

        elif arguments.setup_google:
            run_google_login_setup(
                config
            )

            print()
            print(
                "Google Sheets login profile configured "
                "successfully."
            )

        else:
            result = run_sync(
                config
            )

            print()
            print(
                "VCI Lytx vehicle synchronization completed "
                "successfully."
            )
            print(
                f"Vehicles updated: "
                f"{result['google_rows_written']}"
            )
            print(
                f"Google Sheet tab: "
                f"{result['worksheet_name']}"
            )
            print(
                f"CSV saved to: "
                f"{result['final_csv_path']}"
            )

        logging.info(
            "Log saved to %s.",
            log_path,
        )

        return 0

    except Exception:
        logging.exception(
            "The VCI Lytx vehicle synchronization failed."
        )

        logging.info(
            "The existing destination data was not intentionally "
            "changed after the failure."
        )

        logging.info(
            "Review the log file: %s",
            log_path,
        )

        print()
        print(
            "VCI Lytx vehicle synchronization failed."
        )
        print(
            "Review the latest file in the logs folder."
        )

        return 1


if __name__ == "__main__":
    raise SystemExit(
        main()
    )