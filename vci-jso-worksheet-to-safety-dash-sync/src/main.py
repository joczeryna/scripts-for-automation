import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

from google_sheet_audit_sync import (
    run_google_sheet_sync as run_audit_google_sheet_sync,
    setup_google_login as setup_audit_google_login,
)
from google_sheet_sync import (
    run_google_sheet_sync as run_shift_google_sheet_sync,
    setup_google_login as setup_shift_google_login,
)


PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_FOLDER = PROJECT_ROOT / "config"
LOG_FOLDER = PROJECT_ROOT / "logs"

JOB_CONFIGS = {
    "audit": CONFIG_FOLDER / "audit_sync.json",
    "shift": CONFIG_FOLDER / "shift_allocation_sync.json",
}

JOB_FUNCTIONS = {
    "audit": {
        "run": run_audit_google_sheet_sync,
        "setup_login": setup_audit_google_login,
    },
    "shift": {
        "run": run_shift_google_sheet_sync,
        "setup_login": setup_shift_google_login,
    },
}


def parse_arguments():
    parser = argparse.ArgumentParser(
        description="Copy governed Google Sheet ranges as values only."
    )
    parser.add_argument(
        "--job",
        choices=["audit", "shift"],
        help="Synchronization job to run.",
    )
    parser.add_argument(
        "--setup-login",
        action="store_true",
        help=(
            "Open Chrome to establish or refresh the authorized "
            "Google Sheets login session."
        ),
    )
    return parser.parse_args()


def configure_logging(job_name: str) -> Path:
    LOG_FOLDER.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_job_name = job_name.lower().replace(" ", "_")
    log_path = LOG_FOLDER / f"{safe_job_name}_{timestamp}.log"

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


def load_config(job_key: str) -> dict:
    config_path = JOB_CONFIGS[job_key]

    if not config_path.exists():
        raise FileNotFoundError(
            f"Configuration file not found: {config_path}"
        )

    with config_path.open("r", encoding="utf-8") as config_file:
        config = json.load(config_file)

    required_fields = [
        "job_name",
        "source_spreadsheet_id",
        "destination_spreadsheet_id",
        "browser_profile_folder",
        "data_start_row",
        "mappings",
    ]
    missing_fields = [
        field for field in required_fields if field not in config
    ]

    if missing_fields:
        raise ValueError(
            "The configuration is missing these fields: "
            + ", ".join(missing_fields)
        )

    if not isinstance(config["mappings"], list) or not config["mappings"]:
        raise ValueError(
            "The configuration must contain at least one worksheet mapping."
        )

    required_mapping_fields = [
        "source_worksheet",
        "destination_worksheet",
        "start_column",
        "end_column",
    ]

    for mapping_number, mapping in enumerate(config["mappings"], start=1):
        missing_mapping_fields = [
            field
            for field in required_mapping_fields
            if field not in mapping or str(mapping[field]).strip() == ""
        ]

        if missing_mapping_fields:
            raise ValueError(
                f"Mapping {mapping_number} is missing: "
                + ", ".join(missing_mapping_fields)
            )

    if config["source_spreadsheet_id"] == config["destination_spreadsheet_id"]:
        raise ValueError(
            "The source and destination spreadsheet IDs must be different."
        )

    return config


def resolve_project_path(configured_path: str) -> Path:
    path = Path(configured_path)
    if path.is_absolute():
        return path
    return (PROJECT_ROOT / path).resolve()


def run_login_setup(job_key: str, config: dict) -> None:
    browser_profile_path = resolve_project_path(
        config["browser_profile_folder"]
    )
    setup_login_function = JOB_FUNCTIONS[job_key]["setup_login"]
    setup_login_function(
        browser_profile_path=browser_profile_path,
        spreadsheet_id=config["destination_spreadsheet_id"],
    )
    logging.info("Google Sheets login setup completed.")


def run_job(job_key: str, config: dict) -> list[dict]:
    browser_profile_path = resolve_project_path(
        config["browser_profile_folder"]
    )
    logging.info("Starting %s.", config["job_name"])

    run_function = JOB_FUNCTIONS[job_key]["run"]
    results = run_function(
        job_config=config,
        browser_profile_path=browser_profile_path,
    )

    total_rows = sum(result["rows_written"] for result in results)
    logging.info("%s completed successfully.", config["job_name"])
    logging.info(
        "Completed %s worksheet mappings and wrote %s total rows.",
        len(results),
        total_rows,
    )
    return results


def main() -> int:
    arguments = parse_arguments()

    if arguments.job is None:
        print("ERROR: A job must be selected.")
        print()
        print("Use --job audit or --job shift.")
        return 1

    config = None
    log_path = None

    try:
        config = load_config(arguments.job)
        log_path = configure_logging(arguments.job)

        if arguments.setup_login:
            run_login_setup(arguments.job, config)
            print()
            print("Google Sheets login setup completed successfully.")
        else:
            results = run_job(arguments.job, config)
            total_rows = sum(result["rows_written"] for result in results)

            print()
            print(f"{config['job_name']} completed successfully.")
            print(f"Worksheet mappings completed: {len(results)}")
            print(f"Total rows written: {total_rows}")
            print()
            print("Worksheet results:")

            for result in results:
                print(
                    f"- {result['source_worksheet']} "
                    f"to {result['destination_worksheet']}: "
                    f"{result['rows_written']} rows "
                    f"({result['range']})"
                )

        logging.info("Log saved to %s.", log_path)
        return 0

    except Exception:
        if log_path is None:
            log_path = configure_logging(arguments.job or "unknown_job")

        logging.exception(
            "The Google Sheet values synchronization failed."
        )
        logging.info("Review the log file: %s", log_path)
        print()
        print("Google Sheet values synchronization failed.")
        print("Review the latest file in the logs folder.")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
