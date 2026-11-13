@echo off
setlocal
cd /d "%~dp0"
title VCI Lytx Driver Report and Events Export

echo.
echo VCI Lytx Driver Report and Events Export
echo ==============================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The Python virtual environment was not found.
    echo Expected: %CD%\.venv\Scripts\python.exe
    echo.
    echo Complete the project setup before running this file.
    echo.
    pause
    exit /b 1
)

echo --- [1/2] Running Driver Report Export ---
".venv\Scripts\python.exe" "src\main.py" %*
set "MAIN_EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%MAIN_EXIT_CODE%"=="0" (
    echo [ERROR] VCI Lytx Driver Report export failed.
    echo Review the latest file in the logs folder.
    echo.
    pause
    exit /b %MAIN_EXIT_CODE%
)

echo --- [2/2] Running Events Export ---
".venv\Scripts\python.exe" "src\events_export.py" %*
set "EVENTS_EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EVENTS_EXIT_CODE%"=="0" (
    echo [ERROR] VCI Lytx Events export failed.
    echo Review the latest file in the logs folder.
    echo.
    pause
    exit /b %EVENTS_EXIT_CODE%
)

echo [SUCCESS] Both VCI Lytx exports completed successfully.
echo.
pause
endlocal