@echo off
setlocal
cd /d "%~dp0"
title Run VASU PM Workflow
echo.
echo ================================
echo      VASU PM Workflow
echo ================================
echo.
echo Source folder:
echo D:\VASU_PM_WORKFLOW\01_Source_Files
echo.
echo Required files:
echo - VASU's Tracker 2026.xlsx
echo - project_master.xlsx
echo.
set /p REPORT_MONTH=Enter report month in YYYY-MM format or press Enter for current month: 
echo.
python -m pip install -r requirements.txt
if errorlevel 1 goto failed
python publish_guides.py
if errorlevel 1 goto failed
if "%REPORT_MONTH%"=="" (
  python pm_workflow.py
) else (
  python pm_workflow.py --report-month %REPORT_MONTH%
)
if errorlevel 1 goto failed
echo.
echo Completed. Open D:\VASU_PM_WORKFLOW\02_Output for your month folder.
echo.
pause
endlocal
exit /b 0

:failed
echo.
echo Workflow failed. Please check the message above.
echo.
pause
endlocal
exit /b 1
