@echo off
setlocal
cd /d "C:\Users\jd0855\Desktop\VCI_Automation\ljso-dash-refresh-and-pdf-export"

echo ========================================================
echo Starting Leadership JSO Refresh and PDF Export Workflow
echo ========================================================

python refresh_and_export_leadership_jso.py settings.json

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Workflow failed with exit code %ERRORLEVEL%.
    pause
    exit /b %ERRORLEVEL%
)

echo [SUCCESS] Workflow finished cleanly.
endlocal