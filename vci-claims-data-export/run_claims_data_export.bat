@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE="
if exist ".venv\Scripts\python.exe" (
  set "PYTHON_EXE=.venv\Scripts\python.exe"
) else if exist "..\.venv\Scripts\python.exe" (
  set "PYTHON_EXE=..\.venv\Scripts\python.exe"
)

if not defined PYTHON_EXE (
  echo ERROR: Virtual environment not found.
  echo Run setup first using the instructions in README.md.
  exit /b 1
)

"%PYTHON_EXE%" -m src.main %*
set "exit_code=%ERRORLEVEL%"

if "%exit_code%"=="2" (
  echo.
  echo AUTHENTICATION REQUIRED:
  echo 1. Ensure Google Chrome is open.
  echo 2. Check "Allow remote debugging for this browser instance" at chrome://inspect/#remote-debugging
  echo 3. Confirm you are signed into Google in Chrome, then retry.
) else if "%exit_code%"=="3" (
  echo.
  echo PERMISSION DENIED:
  echo Close the destination Excel workbook and confirm the W: drive is connected.
) else if not "%exit_code%"=="0" (
  echo.
  echo Export failed. Review the newest file in the logs folder.
)

exit /b %exit_code%