from __future__ import annotations

import json
import logging
import subprocess
import sys
import time
from datetime import date
from pathlib import Path

import export_power_bi_jso_tables_v5 as exporter
import update_power_bi_month_filters_v2 as power_bi


LOG = logging.getLogger("leadership_jso_end_to_end")


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
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

        window = power_bi.power_bi_window()
        window.set_focus()
        power_bi.exact(window, "Report view", "TabItem").click_input()
        time.sleep(3)

        refresh = power_bi.refresh_report(
            window,
            int(power_bi_settings.get("refresh_timeout_seconds", 1800)),
        )
        save = power_bi.save_report(window, pbix_path)
        leadership_filter = power_bi.validate_leadership(window, date.today().year)

        stamp = time.strftime("%Y-%m-%d_%H%M%S")
        export_dir = project_dir / "exports" / stamp
        export_dir.mkdir(parents=True, exist_ok=False)
        output = export_dir / f"Leadership_JSOs_{stamp}.csv"
        exporter.export_page(window, "Leadership JSOs", output)
        export_validation = exporter.validate_export(output, "Leadership JSOs")

        writer_path = project_dir / "apply_leadership_google_update_v1.py"
        completed = subprocess.run(
            [sys.executable, str(writer_path), str(google_settings_path)],
            cwd=str(project_dir),
            check=False,
        )
        if completed.returncode != 0:
            raise RuntimeError(
                f"Leadership Google Sheets writer failed with exit code {completed.returncode}"
            )

        result = {
            "mode": "LEADERSHIP JSO END-TO-END WORKFLOW",
            "runDate": date.today().isoformat(),
            "powerBiRefresh": refresh,
            "powerBiSave": save,
            "leadershipFilter": leadership_filter,
            "export": export_validation,
            "googleSheetsWriterCompleted": True,
            "allowedGoogleWriteRanges": [
                "Data Import!A2:G",
                "Data Import!B1",
            ],
            "outOfScope": [
                "Summary",
                "All JSOs",
                "Deviation Statuses",
            ],
        }
        LOG.info("LEADERSHIP_JSO_END_TO_END_RESULT=%s", json.dumps(result))
        return 0
    except Exception as error:
        LOG.exception("Leadership JSO end-to-end workflow failed: %s", error)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
