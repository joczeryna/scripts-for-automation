@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo ERROR: Virtual environment not found.
  echo Run setup first using the instructions in README.md.
  exit /b 1
)

set "python_exe=.venv\Scripts\python.exe"

"%python_exe%" -m src.main %*
set "exit_code=%ERRORLEVEL%"

if "%exit_code%"=="0" exit /b 0
if not "%exit_code%"=="2" goto export_failed

echo.
echo The saved Google login has expired or is no longer authorized.
set "login_choice="
set /P "login_choice=Refresh the Google login and retry? [Y/N]: "
if /I not "%login_choice%"=="Y" (
  echo.
  echo Export cancelled. The existing CLAIMS-DATA.xlsx was not replaced.
  exit /b 2
)

echo.
echo Google Chrome will open for authentication.
echo Sign in with the authorized Google account and complete MFA if prompted.
echo Wait until the spreadsheet and CURRENT tab are visible, then return here.
echo.

"%python_exe%" -m src.main --setup-login
set "login_exit_code=%ERRORLEVEL%"

if not "%login_exit_code%"=="0" (
  echo.
  echo Google login refresh failed. Review the newest file in the logs folder.
  exit /b %login_exit_code%
)

echo.
echo Saving the refreshed Google session...
timeout /T 5 /NOBREAK >NUL

echo.
echo Google login was refreshed. Retrying the claims data export...
echo.

"%python_exe%" -m src.main
set "exit_code=%ERRORLEVEL%"

if "%exit_code%"=="0" exit /b 0

if "%exit_code%"=="2" (
  echo.
  echo Authentication is still required. Confirm that the authorized Google
  echo account can open the spreadsheet and CURRENT tab, then run this batch file again.
  exit /b 2
)

:export_failed
echo.
echo Export failed. Review the newest file in the logs folder.
exit /b %exit_code%
