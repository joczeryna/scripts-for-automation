VCI JSO Worksheet to Safety Dashboard Sync

Important File Path Notice

Windows file paths in this guide use <username> as a placeholder.

Each user must replace <username> with their own Windows username. For example:

C:\Users\<username>\Desktop\VCI_Automation\vci-jso-worksheet-to-safety-dash-sync

Purpose

This project copies governed worksheet data between authorized Google Sheets used by the VCI Executive Safety Dashboard.

It contains two independent synchronization jobs:

JSO Shift Allocation Sync

JSO Audit Data Sync

Both jobs replace destination data from row 2 onward and use paste values only. Header rows are preserved.

Project Folder

C:\Users\<username>\Desktop\VCI_Automation\vci-jso-worksheet-to-safety-dash-sync

JSO Shift Allocation Sync

Source Spreadsheet

Google Sheet ID:

1tA4rLMzhQGo66d2205TaHPKO3dKZJYiwgY15XaJelz4

Destination Spreadsheet

Google Sheet ID:

10AhgLWuwcOXWGT_Z_ARKDiATn_5GI5hVn-9yzJS8TiE

Worksheet Mappings

Source Worksheet

Destination Worksheet

Range

Shift Allocation

Shift Allocation

A2

Shift Allocation PW

Shift Allocation PW

A2

JTSA Details Auditor

JTSA Details Auditor

A2

Run-Date Update

The Shift Allocation job updates cell F1 on the JTSA Summary worksheet in both the source and destination spreadsheets to the date the script is run.

JSO Audit Data Sync

Primary Source Spreadsheet

Google Sheet ID:

11ScapYS8Ac84Yvq6SNMLnfi9g-uEcSQJYwsLNnHxMhk

Primary Destination Spreadsheet

Google Sheet ID:

1ZcpnQn6ZqXFhqaEShqMrWGIpEWKtH-Z-0z9ZGw-OXdw

Worksheet Mappings

Source Worksheet

Source GID

Destination Worksheet

Range

Division Audit Totals

855789445

Division Audit Totals

A2

Deviation Statuses

1972900369

Deviation by Status

A2

Foremen

184113978

Foremen Data

A2

Deviations

1407739991

Deviations

A2

Reversed Final Mapping

The final mapping runs in the opposite spreadsheet direction:

Source Spreadsheet

Source Worksheet

Source GID

Destination Spreadsheet

Destination Worksheet

Range

1ZcpnQn6ZqXFhqaEShqMrWGIpEWKtH-Z-0z9ZGw-OXdw

Monthly Audit Totals

1542493577

11ScapYS8Ac84Yvq6SNMLnfi9g-uEcSQJYwsLNnHxMhk

Audit Totals

A2

Normal Use

Run Shift Allocation Sync

Confirm the computer is connected to the internet.

Close any Chrome window using the project's dedicated Google profile.

Double-click run_shift_allocation_sync.bat.

Allow the script to finish all three worksheet mappings.

Confirm that the terminal displays a successful completion message.

Review the latest log if an error occurs.

Run Audit Data Sync

Confirm the computer is connected to the internet.

Close any Chrome window using the project's dedicated Google profile.

Double-click run_audit_sync.bat.

Allow the script to finish all five worksheet mappings.

Confirm that the terminal displays a successful completion message.

Review the latest log if an error occurs.

Run only one synchronization job at a time.

Google Login Setup

The first-time login setup is run from Command Prompt in the project folder.

Shift Allocation Login Setup

.venv\Scripts\python.exe src\main.py --job shift --setup-login

Audit Login Setup

.venv\Scripts\python.exe src\main.py --job audit --setup-login

Sign in with the authorized VCI Google account and confirm that the spreadsheet opens and is editable before returning to the terminal.

The project uses a dedicated local Chrome profile stored in:

google-browser-profile

Do not copy, share, email, or commit this folder. It contains the user's authenticated Google session.

Synchronization Rules

For every configured worksheet mapping, the automation:

Opens the authorized source spreadsheet.

Selects the configured source worksheet.

Reads or copies the configured source columns beginning at row 2.

Determines the current populated source rows.

Stops before clearing the destination if the source cannot be read completely.

Opens the configured destination worksheet.

Clears the destination range from row 2 onward.

Pastes the replacement dataset using values only.

Processes large datasets in controlled batches.

Allows Google Sheets time to save before continuing.

Values-Only Rule

The automation transfers displayed values only.

It does not intentionally transfer:

Formulas

Formatting

Conditional formatting

Data validation

Comments or notes

Worksheet structure

Existing destination headers and formatting are preserved because replacement begins at row 2.

Range-Limit Workaround

Large destination datasets are pasted in batches. If Google Sheets displays a row-expansion or range-limit confirmation, the automation attempts to select OK and continue.

Do not manually click inside the worksheet while a job is running, because this can change the selected destination cell.

Safety Controls

The automation is designed to prevent partial source reads from replacing governed destination data.

The destination is not intentionally cleared unless:

The correct source spreadsheet opens.

The configured source worksheet is identified.

The source range contains populated rows.

The source data is successfully captured.

Required mapping settings are present.

The Audit job uses exact source GIDs to prevent similarly named or previously active worksheets from being used.

Logs

Run logs are saved in:

logs

The latest log contains the worksheet being processed, source row counts, destination actions, and complete error details.

Log files should be excluded from source control.

Configuration Files

Audit Configuration

config\audit_sync.json

Shift Allocation Configuration

config\shift_allocation_sync.json

Do not change spreadsheet IDs, worksheet names, GIDs, column ranges, or run-date settings without validating the source and destination structure.

Troubleshooting

Python virtual environment not found

Confirm this file exists:

.venv\Scripts\python.exe

If it does not exist, complete the new-user setup process.

PowerShell blocks virtual-environment activation

Activation is not required. Run Python directly through:

.venv\Scripts\python.exe

Google session expired

Run the appropriate login setup command again:

.venv\Scripts\python.exe src\main.py --job shift --setup-login

or:

.venv\Scripts\python.exe src\main.py --job audit --setup-login

Spreadsheet opens but the script does not continue

Confirm the Google account has access to both source and destination spreadsheets.

Confirm the spreadsheet is fully loaded and editable.

Do not close the automated Chrome window.

Review the latest log for the worksheet where processing stopped.

Incorrect source worksheet

Confirm the source worksheet name and GID in config\audit_sync.json match the intended source tab.

Incorrect row count

Confirm column A is populated through the final source data row.

Confirm the source contains no unexpected blank records below the header.

Confirm the worksheet structure and configured range have not changed.

Review the logged source row count before rerunning.

Data pastes into the header row

Confirm data_start_row is set to 2 and the destination worksheet mapping is correct.

Google Sheets row-expansion dialog remains open

Select OK only if the automation cannot dismiss the dialog automatically, then allow the script to continue without clicking elsewhere.

Browser profile is already in use

Close Chrome windows opened by this project and rerun the job. Do not run both batch files simultaneously.

Worksheet renamed or removed

Restore the expected worksheet name or update the applicable configuration only after validating the new mapping.

Executive Safety Dashboard

After a successful synchronization, refresh the VCI Executive Safety Dashboard or its Power BI semantic model according to the established refresh process.

The synchronized data supports Shift Allocation, JTSA, audit totals, deviation, foreman, and related operational safety reporting.