# VCI Safety Report Automation

Windows batch + Python/Playwright automation modeled on the VCI Lytx workflow.

## Scope

The automation:

1. Opens four authenticated Power BI paginated reports in Chrome.
2. Applies the required report filters and exports each report as `.xlsx`.
3. On Mondays, keeps the DVIR Detail records from the most recent Friday through the run Monday and removes older dates from the seven-day export.
4. Validates all four workbooks before continuing.
5. Uploads the files only to **DYCOM Reports V2**.
6. Opens only the **Executive Safety Dashboard**.
7. Runs **Updates > Safety Compliance Report**.
8. Unhides **Updates**, sets **Report Date** to today, and hides the sheet again.

The VCI Safety Dashboard, original DYCOM Reports folder, and manual division-formula review are excluded.

## First-time setup

1. Install Python 3.10 or later and Google Chrome.
2. Open this folder in VS Code.
3. Double-click `setup_vci_safety_automation.bat`.
4. Double-click `run_vci_safety_reports.bat`.
5. In the Chrome window opened by the automation, complete Power BI and Google sign-in/MFA when prompted.
6. Close Chrome after the run. Later runs reuse the dedicated signed-in browser profile.

Do not open the automation's Chrome profile in another Chrome process while the script is running.

## Routine use

Double-click `run_vci_safety_reports.bat`. Keep the Chrome window available until the completion message appears. Downloaded files are retained by date under `downloads`, and execution details are written to `logs`.

## Configuration

Edit `config/settings.json` to change URLs, report filters, timeouts, menu labels, or the browser channel. User-specific paths are not hard-coded.

## Important first-run validation

Power BI and Google periodically change their page markup. The automation uses visible labels rather than brittle coordinates, but the first run should be watched. If a control label differs from the written procedure, the run stops safely and records the missing label in the newest log rather than continuing with incomplete data.

The Monday cleanup interprets the existing procedure as retaining the most recent Friday through Monday and deleting older Monday-through-Thursday records from the seven-day export. This rule is isolated in `src/workbook_processor.py`.

## Troubleshooting

- **Sign-in timeout:** finish login/MFA in the automation Chrome window, close it, and run again.
- **Filter label not found:** compare the on-screen Power BI label with `config/settings.json`.
- **Google menu not found:** wait for the custom `Updates` menu to load, then rerun.
- **Upload does not complete:** confirm that the signed-in Google account has edit access to DYCOM Reports V2.
- **Run fails:** open the newest timestamped file in `logs`.
