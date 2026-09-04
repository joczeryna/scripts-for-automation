@echo off
setlocal

cd /d "%~dp0"

title VCI JSO Shift Allocation Sync

echo.
echo VCI JSO Shift Allocation Sync
echo =============================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The Python virtual environment was not found.
    echo Expected:
    echo %CD%\.venv\Scripts\python.exe
    echo.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" "src\main.py" --job shift %*

set "SYNC_EXIT_CODE=%ERRORLEVEL%"

echo.

if not "%SYNC_EXIT_CODE%"=="0" (
    echo VCI JSO Shift Allocation Sync failed.
    echo Review the latest file in the logs folder.
    echo.
    pause
    exit /b %SYNC_EXIT_CODE%
)

echo VCI JSO Shift Allocation Sync completed successfully.
echo.
pause

endlocal