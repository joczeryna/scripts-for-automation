@echo off
cd /d "%~dp0"

:: If you use a virtual environment, uncomment the line below and comment out the plain python line:
:: call ".venv\Scripts\activate.bat"

python -m src.main_2
set EXIT_CODE=%ERRORLEVEL%

if %EXIT_CODE% NEQ 0 (
    echo.
    echo Export failed with exit code %EXIT_CODE%. Check the logs folder for details.
    pause
)

exit /b %EXIT_CODE%