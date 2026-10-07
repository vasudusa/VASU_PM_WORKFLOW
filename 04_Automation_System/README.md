# Monthly Reporting Automation

This package auto-detects the latest CSV/XLSX input, cleans and validates the data, builds KPI and trend analysis, creates a formatted Excel workbook, and creates a 10-slide executive PowerPoint review.

## Folder Structure

```text
Monthly_Reporting_Automation/
  input/       Place monthly CSV/XLSX files here when not using an external source folder
  output/      Generated Excel, PowerPoint, quality reports, and insight files
  logs/        Processing logs and error reports
  templates/   Optional future Excel/PPT templates
  config.json
  main.py
  reporting_engine.py
  excel_generator.py
  ppt_generator.py
```

## Run Manually

```powershell
cd "C:\Users\Vasu Dusa\OneDrive - Spine Technologies India Private Limited\Documents\PM Monthly Dashboard\Monthly_Reporting_Automation"
python -m pip install -r requirements.txt
python client_report_generator.py
```

To regenerate a specific month, for example June 2026:

```powershell
python client_report_generator.py --report-month 2026-06
python main.py --report-month 2026-06
```

For the generic business-report engine, run:

```powershell
python main.py
```

## Outputs

VASU monthly dashboard files are saved by the current run month:

```text
D:\Reports\<Current Month>\
  VASU_Tracker_Latest_<Month><Year>.xlsx
  VASU_Tracker_Latest_<Month><Year>_Presentation.pptx
```

The generic report engine saves here:

```text
D:\Reports\Monthly Business Reports\<Year>\<Month>\
  Monthly_Business_Report_<Month>_<Year>.xlsx
  Monthly_Executive_Review_<Month>_<Year>.pptx
  Data_Quality_Report_<Month>_<Year>.csv
  Strategic_Insights_<Month>_<Year>.csv
```

## Excel Workbook

Generated sheets:

1. Raw_Data
2. Cleaned_Data
3. Dashboard
4. KPI_Summary
5. Trend_Analysis
6. Charts
7. Executive_Summary

The workbook includes filters, frozen headers, professional corporate formatting, formulas, trend charts, top/bottom performer summaries, conditional formatting, and data quality checks.

Status colors are configurable in `config.json` and `pm_config.json`:

- Active: green
- Completed: light blue
- Yet To Start: yellow
- Held Up: orange
- Lost: red

The VASU workbook includes a `Quarterly Followup` sheet that spreads completed-client follow-ups across the quarter.

## PowerPoint Deck

Generated slide flow:

1. Title Slide
2. Executive Summary
3. KPI Overview
4. Trend Analysis
5. Top Performers
6. Risk & Issues
7. Financial Analysis
8. Forecast & Outlook
9. Recommendations
10. Closing Summary

## Windows Task Scheduler

1. Open **Task Scheduler**.
2. Choose **Create Basic Task**.
3. Name it `Monthly Business Report Automation`.
4. Trigger: **Monthly**, select your preferred day and time.
5. Action: **Start a program**.
6. Program/script:

```text
C:\Users\Vasu Dusa\OneDrive - Spine Technologies India Private Limited\Documents\PM Monthly Dashboard\Monthly_Reporting_Automation\run_monthly_report.bat
```

7. Start in:

```text
C:\Users\Vasu Dusa\OneDrive - Spine Technologies India Private Limited\Documents\PM Monthly Dashboard\Monthly_Reporting_Automation
```

8. Finish, then open task properties and enable **Run whether user is logged on or not** if required by your IT policy.

## PM Cockpit Workflow

Use this for PM management after downloading the SCH `project_master.xlsx` export into `D:\Reports\Current Status`. The tracker remains the primary file; SCH Project Master enriches missing details and milestone context.

```powershell
cd "C:\Users\Vasu Dusa\OneDrive - Spine Technologies India Private Limited\Documents\PM Monthly Dashboard\Monthly_Reporting_Automation"
python pm_cockpit.py --mode all
```

Or run:

```text
D:\Reports\RUN_PM_COCKPIT.bat
```

Inputs:

- `D:\Reports\Current Status\VASU's Tracker 2026.xlsx`
- `D:\Reports\Current Status\project_master.xlsx`
- `D:\Reports\Current Status\Current Status.xlsx`
- `D:\Reports\Current Status\Book1.xlsx`

Outputs:

- `D:\Reports\PM Cockpit\PM_Execution_Master_Latest.xlsx`
- `D:\Reports\PM Cockpit\PM_Cockpit_Latest.xlsx`
- `D:\Reports\PM Cockpit\config\Alias_Matching_Review.xlsx`
- `D:\Reports\PM Cockpit\<Year>\<Month>\PM_Cockpit_<YYYY-MM-DD>.xlsx`
- `D:\Reports\PM Cockpit\<Year>\<Month>\project_summaries\Project_Summary_Index_<YYYY-MM-DD>.csv`
- One HTML project summary per current project under `project_summaries`
- `D:\Reports\PM Cockpit\config\client_alias_map.csv`
- `D:\Reports\PM Cockpit\config\project_outlook_group_map.csv`

Recommended operating rhythm:

1. Download the latest SCH `project_master.xlsx`.
2. Update tracker rows and PM notes as needed.
3. Run `run_pm_cockpit.bat`.
4. Open `PM_Cockpit_Latest.xlsx`.
5. Review `Execution_Master`, `Project_Summary_View`, `Action_List`, `Completed_Followup_Plan`, `Status_Mismatch`, and `Alias_Matching_Tracker`.
6. Send/use the individual project summary files for current projects.
7. Update `client_alias_map.csv` for valid unmatched SCH client-name pairs.
8. When Outlook groups are ready, fill `project_outlook_group_map.csv`; project-wise Outlook draft generation can then be enabled in `pm_config.json`.

See `PM_Cockpit_Guide.md` for the full PM workflow, milestone definitions, alias-mapping steps, and Outlook group setup.
The latest guides are also published to `D:\Reports\PM Automation Guides`.

Optional Task Scheduler setup:

- Program/script:

```text
C:\Users\Vasu Dusa\OneDrive - Spine Technologies India Private Limited\Documents\PM Monthly Dashboard\Monthly_Reporting_Automation\run_pm_cockpit.bat
```

- Start in:

```text
C:\Users\Vasu Dusa\OneDrive - Spine Technologies India Private Limited\Documents\PM Monthly Dashboard\Monthly_Reporting_Automation
```

## Configuration Notes

- Update `input_directories` in `config.json` if monthly files are stored elsewhere.
- `report_month_mode` is set to `current_date`, so output month/year comes from the day you run the automation unless you pass `--report-month YYYY-MM`.
- `monthly_report.selected_report_month` can be used with `monthly_report.period_mode = manual` if you want config-only month selection.
- Existing monthly output files are replaced automatically on rerun. Close the Excel/PPT first if Windows has the file locked.
- Add revenue, cost, budget, and target columns in future source files to activate financial KPI and variance reporting automatically.
- Update `pm_config.json` for PM Cockpit paths, Outlook recipients, and risk thresholds.
