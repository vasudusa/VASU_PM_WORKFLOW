@echo off
setlocal
cd /d "%~dp0"
set /p REPORT_MONTH=Enter report month in YYYY-MM format, for example 2026-06:
if "%REPORT_MONTH%"=="" (
  echo No month entered. Running current month.
  python -m pip install -r requirements.txt
  python publish_guides.py
  python client_report_generator.py
  python main.py
) else (
  python -m pip install -r requirements.txt
  python publish_guides.py
  python client_report_generator.py --report-month %REPORT_MONTH%
  python main.py --report-month %REPORT_MONTH%
)
endlocal
