# VCI Lytx Driver Report Export

## Purpose

This project exports the Lytx Drivers Report for `Last 30 Days` and saves the populated CSV using Lytx's generated filename.

## Source

`https://app.lytx.com/#/driver-safety/insights/driversreport`

## Destination

`W:\Dashboards VCI\Safety Data\Current\Lytx`

If the same generated filename already exists, it is replaced only after the new CSV downloads and validates successfully.

## First-Time Setup

From Command Prompt in the project folder:

```bat
py -m venv .venv
.venv\Scripts\python.exe -m pip install --upgrade pip
.venv\Scripts\python.exe -m pip install -r requirements.txt
run_lytx_driver_report_export.bat --setup-login
```

Complete Lytx sign-in and MFA, manually open the Drivers Report, and press Enter after it is fully visible.

## Normal Use

1. Confirm the W drive is connected.
2. Double-click `run_lytx_driver_report_export.bat`.
3. Complete Lytx authentication if requested.
4. Keep Chrome open while the report runs.
5. Wait for the success message.

## Automation Rules

The automation:

1. Reuses an existing Lytx Driver Report tab when available.
2. Selects `Last 30 Days` and clicks `Apply`.
3. Scrolls until `Download CSV` renders.
4. Downloads every driver in the report.
5. Confirms the CSV is populated and matches the displayed driver count when available.
6. Preserves Lytx's generated filename.
7. Atomically replaces a same-name destination file.

## Logs

Logs are saved in `logs`.

## Security

Do not copy, share, email, or commit the `browser-profile` folder. Each authorized user must establish their own Lytx session.
