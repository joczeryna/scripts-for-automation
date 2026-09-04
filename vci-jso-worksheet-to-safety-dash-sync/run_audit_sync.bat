@echo off
setlocal

cd /d "%~dp0"

echo.
echo VCI JSO Audit Data Sync
echo =======================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The Python virtual environment was not found.
    echo Expected:
    echo %CD%\.venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" "src\main.py" --job audit %*

set "EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXIT_CODE%"=="0" (
    echo VCI JSO Audit Data Sync failed.
    echo Review the latest file in the logs folder.
) else (
    echo VCI JSO Audit Data Sync completed successfully.
)

echo.
pause
exit /b %EXIT_CODE%