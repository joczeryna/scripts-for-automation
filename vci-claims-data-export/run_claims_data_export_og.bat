@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERROR: Virtual environment not found.
  echo Run setup first using the instructions in README.md.
  exit /b 1
)

".venv\Scripts\python.exe" -m src.main %*
set "exit_code=%ERRORLEVEL%"

if not "%exit_code%"=="0" (
  echo.
  echo Export failed. Review the newest file in the logs folder.
)

exit /b %exit_code%
