@echo off
setlocal
cd /d "%~dp0"

py -3 -m venv .venv
if errorlevel 1 goto :error

".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :error

".venv\Scripts\python.exe" -m pip install -r requirements.txt
if errorlevel 1 goto :error

if not exist downloads mkdir downloads
if not exist logs mkdir logs
if not exist browser_profiles mkdir browser_profiles

echo.
echo Setup complete. Run run_vci_safety_reports.bat.
pause
exit /b 0

:error
echo.
echo Setup failed. Review the messages above.
pause
exit /b 1

