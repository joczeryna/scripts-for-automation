@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERROR: .venv\Scripts\python.exe was not found.
  echo Run this file from the leadership_jso_powerbi_etl project folder.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" "run_leadership_jso_production.py" "settings.json" "google_sheets_preview_settings.json"
set "RUN_EXIT=%ERRORLEVEL%"

if not "%RUN_EXIT%"=="0" (
  echo.
  echo Production workflow failed. Review the final error above.
) else (
  echo.
  echo Production workflow completed successfully.
)

pause
exit /b %RUN_EXIT%
