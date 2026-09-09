from __future__ import annotations

import argparse
import json
import logging
import sys
from datetime import datetime
from pathlib import Path

from .google_download import AuthenticationError, download_tab_values, setup_login
from .workbook_processor import prepare_current_tab, replace_destination


PROJECT_ROOT = Path(__file__).resolve().parent.parent


def configure_logging() -> Path:
    log_dir = PROJECT_ROOT / "logs"
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"claims_export_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(log_path, encoding="utf-8"), logging.StreamHandler()],
    )
    return log_path


def load_config() -> dict[str, str]:
    config_path = PROJECT_ROOT / "config.json"
    with config_path.open(encoding="utf-8") as handle:
        config = json.load(handle)

    required = ("spreadsheet_id", "source_tab", "destination")
    missing = [key for key in required if not str(config.get(key, "")).strip()]
    if missing:
        raise ValueError(f"Missing configuration values: {', '.join(missing)}")
    return config


def run() -> int:
    parser = argparse.ArgumentParser(description="Export CURRENT to CLAIMS-DATA.xlsx")
    parser.add_argument("--setup-login", action="store_true")
    args = parser.parse_args()

    log_path = configure_logging()
    try:
        config = load_config()
        profile_dir = PROJECT_ROOT / "browser-profile"

        if args.setup_login:
            setup_login(config["spreadsheet_id"], profile_dir)
            logging.info("Google sign-in setup completed.")
            return 0

        run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        temp_dir = PROJECT_ROOT / "temp"
        downloaded = temp_dir / f"source_{run_stamp}.csv"
        prepared = temp_dir / f"claims_data_{run_stamp}.xlsx"
        destination = Path(config["destination"])

        logging.info("Downloading static values from the authorized %s tab.", config["source_tab"])
        download_tab_values(
            config["spreadsheet_id"], config["source_tab"], profile_dir, downloaded
        )

        logging.info("Validating and isolating the %s tab.", config["source_tab"])
        rows, columns = prepare_current_tab(downloaded, prepared, config["source_tab"])

        logging.info("Replacing %s", destination)
        replace_destination(prepared, destination)
        logging.info(
            "SUCCESS: exported %s with %d rows and %d columns to %s",
            config["source_tab"], rows, columns, destination,
        )
        return 0
    except AuthenticationError as exc:
        logging.error("AUTHENTICATION REQUIRED: %s", exc)
        return 2
    except PermissionError:
        logging.exception(
            "Permission denied. Close CLAIMS-DATA.xlsx and confirm write access to W:."
        )
        return 3
    except Exception:
        logging.exception("Export failed; the existing destination file was not replaced.")
        return 1
    finally:
        logging.info("Log: %s", log_path)


if __name__ == "__main__":
    sys.exit(run())
