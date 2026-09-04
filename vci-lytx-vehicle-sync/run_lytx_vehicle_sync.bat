@echo off
setlocal

cd /d "%~dp0"

title VCI Lytx Vehicle Synchronization

echo.
echo VCI Lytx Vehicle Synchronization
echo =================================
echo.

if not exist ".venv\Scripts\python.exe" (
    echo ERROR: The Python virtual environment was not found.
    echo Expected:
    echo %CD%\.venv\Scripts\python.exe
    echo.
    echo Complete the project setup before running this file.
    echo.
    pause
    exit /b 1
)

".venv\Scripts\python.exe" "src\main.py"

set "SYNC_EXIT_CODE=%ERRORLEVEL%"

echo.

if not "%SYNC_EXIT_CODE%"=="0" (
    echo VCI Lytx vehicle synchronization failed.
    echo Review the latest file in the logs folder.
    echo.
    pause
    exit /b %SYNC_EXIT_CODE%
)

echo Synchronization completed successfully.
echo.
pause

endlocal