# VCI Claims Data Export

Exports the resolved, static values displayed on the `CURRENT` tab from the
configured Google Sheet into:

`W:\Dashboards VCI\Safety Data\Current\GL\CLAIMS-DATA.xlsx`

The current destination is replaced only after the download and prepared workbook pass validation.

## One-time setup

Open **Command Prompt** in this project folder and run:

```bat
py -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m playwright install chromium
run_claims_data_export.bat --setup-login
```

Sign in using the authorized Google account. Confirm the spreadsheet opens and the `CURRENT` tab is accessible, then return to Command Prompt and press Enter.

The login window uses the installed Google Chrome browser because Google may
stall or reject sign-in from Playwright's bundled Chromium browser. Google
Chrome must be installed on the computer.

Do not copy, share, or commit `browser-profile`. It contains the saved Google session.

## Routine execution

Before running:

- Confirm the `W:` drive is connected.
- Close `CLAIMS-DATA.xlsx` if it is open in Excel.
- Confirm no other operator is running the export.

Run:

```bat
run_claims_data_export.bat
```

A successful run ends with a `SUCCESS` message showing the exported row and column counts. Logs are written to `logs`.

## Safety behavior

- Requires a worksheet named exactly `CURRENT`.
- Requires at least a header and one populated data row.
- Produces an Excel workbook containing only `CURRENT` as static values.
- Does not carry the source `IMPORTRANGE` formula into the Excel output.
- Validates the prepared workbook before publication.
- Writes a staged file beside the destination and replaces the destination only after validation.
- Leaves the existing destination unchanged when download, authentication, validation, or write access fails.

## Troubleshooting

- **Authentication required:** run `run_claims_data_export.bat --setup-login` again.
- **W: unavailable:** reconnect the mapped drive before running.
- **Permission denied:** close the destination workbook in Excel.
- **CURRENT missing:** verify the source worksheet name has not changed.
- **Details:** review the newest file in `logs`.

## Project contents

- `config.json` — source spreadsheet, worksheet, and destination configuration
- `run_claims_data_export.bat` — Windows entry point
- `src/google_download.py` — authenticated Google workbook download
- `src/workbook_processor.py` — worksheet isolation and validation
- `src/main.py` — orchestration, logs, and exit codes
