@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERROR: Python environment not found.
  echo Run setup_vci_safety_automation.bat first.
  pause
  exit /b 1
)

".venv\Scripts\python.exe" "src\main.py"
set "VCI_EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%VCI_EXIT_CODE%"=="0" (
  echo Automation stopped with an error. Review the newest file in the logs folder.
) else (
  echo VCI Safety report automation completed successfully.
)
pause
exit /b %VCI_EXIT_CODE%

