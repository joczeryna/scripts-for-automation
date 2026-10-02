from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import time
from datetime import date
from pathlib import Path

import update_power_bi_month_filters_v2 as power_bi
from google_sheets_jso_load_preview_v2 import preview


LOG = logging.getLogger("leadership_jso_production")


def run_step(project_dir: Path, script: str, *arguments: Path | str) -> None:
    command = [sys.executable, str(project_dir / script)] + [str(value) for value in arguments]
    LOG.info("Starting controlled step: %s", script)
    completed = subprocess.run(command, cwd=str(project_dir), check=False)
    if completed.returncode != 0:
        raise RuntimeError(
            f"Controlled step {script} failed with exit code {completed.returncode}"
        )


def open_and_wait_for_power_bi(pbix_path: Path, timeout_seconds: int = 180) -> tuple[object, bool]:
    opened_by_runner = False
    try:
        window = power_bi.power_bi_window()
    except Exception:
        if not pbix_path.exists():
            raise RuntimeError(f"Power BI file was not found: {pbix_path}")
        if os.name != "nt":
            raise RuntimeError("This production runner must be executed on the Windows workstation")
        os.startfile(str(pbix_path))  # type: ignore[attr-defined]
        opened_by_runner = True
        deadline = time.time() + timeout_seconds
        last_error: Exception | None = None
        while time.time() < deadline:
            time.sleep(3)
            try:
                window = power_bi.power_bi_window()
                break
            except Exception as error:
                last_error = error
        else:
            raise RuntimeError(
                f"Safety Reports did not become ready within {timeout_seconds} seconds"
            ) from last_error

    window.set_focus()
    power_bi.exact(window, "Report view", "TabItem").click_input()
    time.sleep(3)
    return window, opened_by_runner


def main() -> int:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    project_dir = Path(__file__).resolve().parent
    power_bi_settings_path = Path(
        sys.argv[1] if len(sys.argv) > 1 else project_dir / "settings.json"
    ).resolve()
    google_settings_path = Path(
        sys.argv[2]
        if len(sys.argv) > 2
        else project_dir / "google_sheets_preview_settings.json"
    ).resolve()

    try:
        power_bi_settings = json.loads(
            power_bi_settings_path.read_text(encoding="utf-8")
        )
        pbix_path = Path(
            power_bi_settings.get(
                "power_bi_file",
                r"C:\Users\jd0855\Downloads\Safety Reports.pbix",
            )
        )
        _, opened_by_runner = open_and_wait_for_power_bi(
            pbix_path,
            int(power_bi_settings.get("power_bi_open_timeout_seconds", 180)),
        )

        # Refreshes first, saves after refresh, enforces the current-year
        # Leadership policy, enforces the rolling three-month policy on the
        # other two visuals, and saves any filter changes.
        run_step(
            project_dir,
            "update_power_bi_month_filters_v2.py",
            power_bi_settings_path,
        )
        run_step(
            project_dir,
            "export_power_bi_jso_tables_v5.py",
            power_bi_settings_path,
        )

        # Validate Google destinations once. The two append/overwrite writers
        # consume this same gate and also recheck their live overlap rows just
        # before writing, avoiding repeated expensive full-workbook previews.
        gate = preview(google_settings_path)
        gate_dir = project_dir / "logs"
        gate_dir.mkdir(parents=True, exist_ok=True)
        stamp = time.strftime("%Y-%m-%d_%H%M%S")
        gate_path = gate_dir / f"google_validation_gate_{stamp}.json"
        gate_path.write_text(json.dumps(gate, indent=2), encoding="utf-8")

        run_step(
            project_dir,
            "apply_leadership_google_update_v1.py",
            google_settings_path,
        )
        run_step(
            project_dir,
            "apply_deviation_statuses_google_update_v1.py",
            google_settings_path,
            "--gate-file",
            gate_path,
        )
        run_step(
            project_dir,
            "apply_all_jsos_google_update_v1.py",
            google_settings_path,
            "--gate-file",
            gate_path,
        )

        result = {
            "mode": "LEADERSHIP JSO CONTROLLED PRODUCTION WORKFLOW",
            "runDate": date.today().isoformat(),
            "powerBiFile": str(pbix_path),
            "powerBiOpenedByRunner": opened_by_runner,
            "powerBiRefreshedAndSaved": True,
            "exportsValidated": 3,
            "googleValidationGate": str(gate_path),
            "leadershipRows": gate["leadership"]["proposedImportRows"],
            "deviationRowsFromOverlap": gate["deviationStatuses"][
                "proposedRowsFromOverlap"
            ],
            "allJsosRowsToAppend": gate["allJsos"]["proposedAppendRows"],
            "controlledGoogleWritesCompleted": [
                "Leadership Data Import A2:G and B1",
                "Deviation Statuses A:N from verified overlap",
                "All JSOs A:O append after verified overlap",
            ],
            "valuesOnly": True,
            "completed": True,
        }
        LOG.info("LEADERSHIP_JSO_PRODUCTION_RESULT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Leadership JSO production workflow failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
