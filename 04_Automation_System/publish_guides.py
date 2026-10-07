from __future__ import annotations

import shutil
from datetime import datetime
from pathlib import Path


AUTOMATION_DIR = Path(__file__).resolve().parent
FINAL_ROOT = Path(r"D:\VASU_PM_WORKFLOW")
SOURCE_DIR = FINAL_ROOT / "01_Source_Files"
OUTPUT_ROOT = FINAL_ROOT / "02_Output"
GUIDE_DIR = FINAL_ROOT / "03_Guides"


def write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content.strip() + "\n", encoding="utf-8")


def write_launcher(path: Path, target_bat: Path) -> None:
    content = f"""@echo off
call "{target_bat}"
"""
    write_text(path, content)


def copy_if_exists(source: Path, destination: Path) -> None:
    if source.exists():
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, destination)


def start_here_guide() -> str:
    return f"""
# Start Here - VASU PM Workflow

Generated: {datetime.now():%d-%b-%Y %H:%M}

This is the only active folder you need:

`D:\\VASU_PM_WORKFLOW`

## What Goes Where

1. Put the latest source files here:

   `D:\\VASU_PM_WORKFLOW\\01_Source_Files`

   Required files:

   - `VASU's Tracker 2026.xlsx`
   - `project_master.xlsx`

2. Run the workflow from either:

   - `D:\\VASU_PM_WORKFLOW\\RUN_PM_WORKFLOW.bat`
   - Desktop shortcut: `Run VASU PM Workflow`

3. Use monthly outputs from:

   `D:\\VASU_PM_WORKFLOW\\02_Output\\<Year>\\<MM-Month>`

   Inside every month:

   - `00_Source_Files_Used`: exact source files used for that run.
   - `01_Execution`: your final PM execution workbook.
   - `02_Monthly_To_Sir`: Excel and PPT to send to Sir.
   - `03_Weekly_Mail_Merge`: one weekly update workbook for mail merge.

4. Use guides from:

   `D:\\VASU_PM_WORKFLOW\\03_Guides`

## Daily Working Rule

Tracker is the PM source of truth. SCH Project Master fills missing details, latest project status, implementation remarks, owner context, product, and employee count.

For PM notes:

- If you update `PM Notes`, that note is used.
- If `PM Notes` is blank, SCH `Project Update` is used.
- If that is blank, SCH `Project Remark` is used.
- If both are blank, the workflow drafts a conservative update from status, milestone, risk, owner, and next action.

No separate weekly HTML files are created. Weekly updates are kept in one Excel mail-merge workbook.

Old folders under `D:\\Reports` are history only unless you intentionally need an old file.
"""


def monthly_report_guide() -> str:
    return f"""
# Monthly Output Guide

Generated: {datetime.now():%d-%b-%Y %H:%M}

## Run

Run:

`D:\\VASU_PM_WORKFLOW\\RUN_PM_WORKFLOW.bat`

Enter the report month as `YYYY-MM`, for example:

`2026-09`

Press Enter to use the current month.

## Final Monthly Folder

The workflow creates:

`D:\\VASU_PM_WORKFLOW\\02_Output\\<Year>\\<MM-Month>`

Use:

- `01_Execution`: one execution master for your PM work.
- `02_Monthly_To_Sir`: final Excel dashboard and PowerPoint for Sir.
- `03_Weekly_Mail_Merge`: one Excel file with project-wise weekly update rows.

If you rerun the same month, generated outputs are replaced. Close Excel and PowerPoint before rerunning.

## Monthly Excel Sheets

- Dashboard
- Overall Summary
- Year on Year Analysis
- Monthly Metrics
- Latest Month
- Projects
- Quarterly Followup

The `Quarterly Followup` sheet splits completed clients across the quarter so you can stay in touch without trying to call every completed client in one month.
"""


def source_file_guide() -> str:
    return f"""
# Source File Guide

Generated: {datetime.now():%d-%b-%Y %H:%M}

Use only this source folder:

`D:\\VASU_PM_WORKFLOW\\01_Source_Files`

Keep these names exactly:

- `VASU's Tracker 2026.xlsx`
- `project_master.xlsx`

Optional files:

- `Book1.xlsx`: used only if you still keep notes there.
- `Alias_Matching_Review.xlsx`: created by the workflow for tracker/SCH name matching.
- `Project_Outlook_Group_Map.csv`: created by the workflow for future mail automation mapping.

Before running the batch file, close the source Excel files.
"""


def batch_file_guide() -> str:
    return f"""
# Batch File Guide

Generated: {datetime.now():%d-%b-%Y %H:%M}

## Recommended Run

1. Copy/download latest source files into `D:\\VASU_PM_WORKFLOW\\01_Source_Files`.
2. Double-click the Desktop shortcut `Run VASU PM Workflow`.
3. Enter the month as `YYYY-MM`, or press Enter for the current month.
4. Open the month folder shown at the end of the run.

The root launcher is:

`D:\\VASU_PM_WORKFLOW\\RUN_PM_WORKFLOW.bat`

The automation system files are in:

`D:\\VASU_PM_WORKFLOW\\04_Automation_System`

You normally do not need to open the automation system folder.
"""


def main() -> None:
    FINAL_ROOT.mkdir(parents=True, exist_ok=True)
    SOURCE_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)
    GUIDE_DIR.mkdir(parents=True, exist_ok=True)

    write_text(FINAL_ROOT / "START_HERE.md", start_here_guide())
    write_text(GUIDE_DIR / "START_HERE.md", start_here_guide())
    write_text(GUIDE_DIR / "Monthly_Output_Guide.md", monthly_report_guide())
    write_text(GUIDE_DIR / "Source_File_Guide.md", source_file_guide())
    write_text(GUIDE_DIR / "Batch_File_Guide.md", batch_file_guide())

    write_launcher(FINAL_ROOT / "RUN_PM_WORKFLOW.bat", AUTOMATION_DIR / "run_pm_workflow.bat")

    print(f"FINAL_ROOT={FINAL_ROOT}")
    print(f"GUIDES={GUIDE_DIR}")
    print(f"START_HERE={FINAL_ROOT / 'START_HERE.md'}")
    print(f"RUN_FILE={FINAL_ROOT / 'RUN_PM_WORKFLOW.bat'}")


if __name__ == "__main__":
    main()
