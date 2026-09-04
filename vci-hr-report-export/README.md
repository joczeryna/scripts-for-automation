VCI HR Report Export Automation

Setup and User Guide

Purpose

This automation downloads two authorized VCI Google Sheets as Microsoft Excel workbooks:

Dashboard Report – Active Employees

Dashboard Report – Term

The automation then:

Identifies the most recent worksheet using the MMDDYYYY tab-name format.

Keeps only the most recent worksheet.

Validates the required column headers.

Removes all other worksheets.

Replaces the governed Excel files used by Power BI.

Creates a timestamped processing log.

The automation uses the employee’s authorized Google account through a dedicated local Chrome session. It does not require a Google Cloud project or API credentials.

Prerequisites

The user must have:

Editor access to both authorized Google Sheets

Access to W:\Dashboards VCI\HR Data

Google Chrome installed

Visual Studio Code installed

Python 3.12 or 3.13 installed

A Windows computer

Only one person should run the automation at a time.

Source Reports

Active Employees

Google Sheet ID:
1t6ONs-qNpVkks_-qkWZIbCNcvKLZGtWKeGLuS4T0VbI
Destination file:
W:\Dashboards VCI\HR Data\Dashboard Report - Active Employees.xlsx

Terminated Employees

Google Sheet ID:
1fjDBCaFJHke4OTcwpn4gts89H2v3UENstcqdWODd-AU
Destination file:
W:\Dashboards VCI\HR Data\Dashboard Report - Term.xlsx

Project Contents

A clean project copy should contain:
vci-hr-report-export
 config
  reports.json
 src
  google_download.py
  main.py
  workbook_processor.py
 .gitignore
 README.md
 requirements.txt
 run_hr_report_export.bat
The project copy must not include another user’s:

.venv

browser-profile

logs

temp

The browser-profile folder contains an authenticated Google session and must never be copied or shared.

First-Time Installation

Step 1 — Copy the Project

Create this folder on the user’s computer:
C:\Users\YOUR_USERNAME\Desktop\VCI_Automation
Copy or extract the project into that folder.
The completed location should resemble:
C:\Users\YOUR_USERNAME\Desktop\VCI_Automation\vci-hr-report-export

Step 2 — Install Python

Check VCI’s approved Software Center or application portal first.
If Python is available, install Python 3.13.x.
If it is not available, use the approved Windows 64-bit installer from:
https://www.python.org/downloads/windows/
On the first installer screen, select:
Add python.exe to PATH
Then select:
Install Now
A per-user installation is sufficient unless VCI directs otherwise.
After installation:

Close Visual Studio Code.

Reopen Visual Studio Code.

Step 3 — Open the Project in VS Code

Open Visual Studio Code.

Select File → Open Folder.

Open the vci-hr-report-export folder.

If prompted to trust the folder, select Yes, I trust the authors.

Step 4 — Open a Command Prompt Terminal

Use Command Prompt instead of PowerShell.

Select Terminal → New Terminal.

Select the dropdown arrow beside the terminal’s plus button.

Select Command Prompt.

If Command Prompt is not listed:

Select Select Default Profile.

Choose Command Prompt.

Open a new terminal.

The terminal path should end with:
VCI_Automation\vci-hr-report-export
If the terminal opens one folder above the project, run:
cd vci-hr-report-export

Step 5 — Confirm Python

Run:
python --version
The result should resemble:
Python 3.13.x
If Python is not recognized, try:
py --version

Step 6 — Create the Virtual Environment

If the python command works, run:
python -m venv .venv
If only the py command works, run:
py -m venv .venv
PowerShell activation is not required. The project runs the virtual-environment Python directly.

Step 7 — Install the Required Packages

Run:
.venv\Scripts\python.exe -m pip install --upgrade pip
Then run:
.venv\Scripts\python.exe -m pip install -r requirements.txt
Verify the installation:
.venv\Scripts\python.exe -c "import playwright, openpyxl; print('Dependencies installed successfully')"
Expected result:
Dependencies installed successfully

Step 8 — Validate the Project

Run:
.venv\Scripts\python.exe -m py_compile src\main.py src\google_download.py src\workbook_processor.py
Then run:
.venv\Scripts\python.exe -m json.tool config\reports.json > nul
If both commands return to the prompt without an error, the project files are valid.

Step 9 — Confirm the Destination Folder

Open:
W:\Dashboards VCI\HR Data
The user must be able to open and write to this folder.
If the user’s mapped drive uses a different letter, update the destination_folder setting in:
config\reports.json
The standard setting is:
"destination_folder": "W:\Dashboards VCI\HR Data"

Step 10 — Create the Google Login Session

Each user must create their own authenticated browser session.
Run:
run_hr_report_export.bat --setup-login
A dedicated Chrome window will open.
In that window:

Sign in with the user’s authorized VCI Google account.

Complete multifactor authentication if required.

Confirm the Active Employees Google Sheet opens.

Leave the automation Chrome window open.

Return to the terminal.

Press Enter only after the spreadsheet is fully visible.

Allow the automation to close Chrome. Do not close the automation Chrome window manually.

Expected result:
Google login setup completed.
The authenticated session is stored locally in:
browser-profile
Do not copy or share this folder.

First Automation Run

Before running the automation:

Confirm the W: drive is connected.

Close both destination Excel workbooks.

Do not refresh Power BI during file replacement.

Confirm no other user is running the automation.

Run:
run_hr_report_export.bat

If the stored Google session has expired, the batch file will report that the export failed and ask:

Refresh the Google login and retry? [Y/N]

Press Y. When the dedicated Chrome window opens:

Sign in with the authorized VCI Google account.

Complete multifactor authentication if requested.

Confirm the Active Employees spreadsheet is fully visible.

Leave Chrome open.

Return to the terminal and press Enter.

Allow the script to close Chrome and retry the export automatically.

The process should report:

Starting VCI HR report export

Downloading Active Employees

Downloading Terminated Employees

Cleaning Active Employees

Cleaning Terminated Employees

Replacing both destination files

Successful completion displays:
Both VCI HR reports were updated successfully.

First-Run Verification

Open:
W:\Dashboards VCI\HR Data
Confirm these files exist:

Dashboard Report - Active Employees.xlsx

Dashboard Report - Term.xlsx

Open each workbook and confirm:

Only one worksheet remains.

The worksheet has the most recent valid MMDDYYYY name.

The worksheet includes Employee Number.

The worksheet includes Revenue/Non Revenue.

The rows appear complete.

Excel does not display a repair or corruption warning.

Close both workbooks after verification.

Power BI Validation

After the automation completes successfully:

Open the applicable Power BI file.

Run the refresh.

Confirm the active and termination staging queries load successfully.

Confirm the queries return:

Employee Number

RevenueClassification

Workforce Group

Confirm COER and G&A values reconcile.

Confirm the Replacement Cost KPI and Executive Alert remain aligned.

Routine Operating Procedure

After first-time setup, the user only needs to:

Confirm the W: drive is connected.

Close the two destination Excel workbooks.

Confirm no other user is running the automation.

Double-click run_hr_report_export.bat.

If prompted to refresh the Google login, press Y, sign in, confirm the Active Employees spreadsheet is visible, leave Chrome open, and press Enter in the terminal.

Wait while the batch file retries the export automatically.

Wait for the success message.

Press any key to close the window.

Refresh Power BI.

The user does not need to open Visual Studio Code for routine updates.

Workbook Processing Rules

For each workbook, the automation:

Downloads the authorized Google Sheet in Microsoft Excel format.

Identifies worksheets whose names follow MMDDYYYY.

Converts valid worksheet names into dates.

Selects the most recent date.

Confirms the required headers exist.

Deletes every other worksheet.

Saves a cleaned temporary workbook.

Replaces the governed destination file.

Worksheet position does not determine which tab is retained. The date in the worksheet name determines the latest tab.
Column sequence does not matter.
Required business headers:

Employee Number

Revenue/Non Revenue

Troubleshooting

Python Is Not Recognized

Close and reopen Visual Studio Code, then run:
python --version
If that fails, try:
py --version
If py works, create the virtual environment with:
py -m venv .venv

PowerShell Says Running Scripts Is Disabled

Do not change the company execution policy.
Use a Command Prompt terminal. The project does not require PowerShell activation.

Required Python Packages Are Missing

Run:
.venv\Scripts\python.exe -m pip install -r requirements.txt
Then verify:
.venv\Scripts\python.exe -c "import playwright, openpyxl; print('Dependencies installed successfully')"

Google Session Expired

The normal batch file detects a failed export and offers to refresh the Google login.

When prompted with Refresh the Google login and retry? [Y/N], press Y.

Sign in using the authorized VCI Google account.

Complete multifactor authentication if requested.

Confirm the Active Employees spreadsheet is fully visible.

Leave the automation Chrome window open.

Return to the terminal and press Enter.

Allow the automation to close Chrome and retry the export automatically.

If the automatic prompt is unavailable, run:

run_hr_report_export.bat --setup-login

After setup succeeds, run:

run_hr_report_export.bat

Google Sheet Does Not Download

Confirm:

The user has access to both Google Sheets.

The correct Google account is signed in.

The Google session has not expired.

Google Chrome is installed.

The computer has internet access.

If necessary, rerun the Google login setup.

W: Drive Is Unavailable

Open:
W:\Dashboards VCI\HR Data
If it does not open, reconnect the mapped drive or contact IT.
The automation will not intentionally replace the production files if the destination folder is unavailable.

Destination Workbook Is Locked

Close:

Dashboard Report - Active Employees.xlsx

Dashboard Report - Term.xlsx

Then rerun the automation.

No Dated Worksheet Is Found

Confirm the latest Google Sheet tab uses exactly eight digits in MMDDYYYY format.
Example:
08312026

Required Header Is Missing

Confirm the latest worksheet contains:

Employee Number

Revenue/Non Revenue

Column spacing and sequence may vary, but both business fields must be present.

Batch File Does Not Open

Confirm the filename is exactly:
run_hr_report_export.bat
It must not be:
run_hr_report_export.bat.txt
In File Explorer, enable:
View → Show → File name extensions
The file type should be Windows Batch File.
The user can also run it from Command Prompt:
run_hr_report_export.bat

Review the Logs

Logs are stored in:
logs
Open the most recent file named similarly to:
hr_report_export_YYYYMMDD_HHMMSS.log
The log identifies the report, worksheet, file, and processing step that failed.

Security Requirements

Never share the browser-profile folder.

Never place Google passwords in the source code.

Never copy another employee’s authenticated browser profile.

Do not commit downloaded HR workbooks to Git.

Do not email downloaded HR workbooks unless separately authorized.

Keep the project on an approved VCI computer.

Run the automation only under an authorized Windows and Google account.

Multi-User Operating Rule

Multiple authorized employees may install and run the project, but each user must have:

Their own local project installation

Their own Python virtual environment

Their own authenticated browser-profile

Their own Google Sheet permissions

Access to the shared W: destination

Only one user should run the automation at a time because every installation replaces the same two production Excel files.