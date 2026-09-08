@echo off
setlocal
cd /d "%~dp0"
title VCI Lytx Driver Report Export

echo.
echo VCI Lytx Driver Report Export
echo ==============================
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

".venv\Scripts\python.exe" "src\main.py" %*
set "EXPORT_EXIT_CODE=%ERRORLEVEL%"

echo.
if not "%EXPORT_EXIT_CODE%"=="0" (
    echo VCI Lytx Driver Report export failed.
    echo Review the latest file in the logs folder.
    echo.
    pause
    exit /b %EXPORT_EXIT_CODE%
)

echo VCI Lytx Driver Report export completed successfully.
echo.
pause
endlocal

