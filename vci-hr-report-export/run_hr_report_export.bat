@echo off
setlocal

cd /d "%~dp0"

echo.
echo VCI HR Report Automation
echo ========================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The Python virtual environment was not found.
    echo Expected:
    echo %CD%\.venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

rem Allow manual login setup.
if /I "%~1"=="--setup-login" (
    ".venv\Scripts\python.exe" "src\main.py" --setup-login
    set "EXIT_CODE=%ERRORLEVEL%"
    goto FINISH
)

rem Run the normal export.
".venv\Scripts\python.exe" "src\main.py"
set "EXIT_CODE=%ERRORLEVEL%"

if "%EXIT_CODE%"=="0" (
    goto FINISH
)

echo.
echo The export failed.
echo The Google login session may have expired.
echo.
choice /C YN /N /M "Refresh the Google login and retry? [Y/N]: "

if errorlevel 2 goto FINISH

echo.
echo Opening Google login setup...
".venv\Scripts\python.exe" "src\main.py" --setup-login

if errorlevel 1 (
    echo.
    echo Google login setup failed.
    set "EXIT_CODE=1"
    goto FINISH
)

echo.
echo Login setup completed. Retrying the HR report export...
echo.

".venv\Scripts\python.exe" "src\main.py"
set "EXIT_CODE=%ERRORLEVEL%"

:FINISH
echo.

if "%EXIT_CODE%"=="0" (
    echo VCI HR report automation completed successfully.
) else (
    echo VCI HR report automation failed.
    echo Review the latest file in the logs folder.
)

echo.
pause
exit /b %EXIT_CODE%