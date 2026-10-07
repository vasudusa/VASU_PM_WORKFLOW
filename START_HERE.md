# Start Here - VASU PM Workflow

Generated: 06-Oct-2026 12:25

This is the only active folder you need:

`D:\VASU_PM_WORKFLOW`

## What Goes Where

1. Put the latest source files here:

   `D:\VASU_PM_WORKFLOW\01_Source_Files`

   Required files:

   - `VASU's Tracker 2026.xlsx`
   - `project_master.xlsx`

2. Run the workflow from either:

   - `D:\VASU_PM_WORKFLOW\RUN_PM_WORKFLOW.bat`
   - Desktop shortcut: `Run VASU PM Workflow`

3. Use monthly outputs from:

   `D:\VASU_PM_WORKFLOW\02_Output\<Year>\<MM-Month>`

   Inside every month:

   - `00_Source_Files_Used`: exact source files used for that run.
   - `01_Execution`: your final PM execution workbook.
   - `02_Monthly_To_Sir`: Excel and PPT to send to Sir.
   - `03_Weekly_Mail_Merge`: one weekly update workbook for mail merge.

4. Use guides from:

   `D:\VASU_PM_WORKFLOW\03_Guides`

## Daily Working Rule

Tracker is the PM source of truth. SCH Project Master fills missing details, latest project status, implementation remarks, owner context, product, and employee count.

For PM notes:

- If you update `PM Notes`, that note is used.
- If `PM Notes` is blank, SCH `Project Update` is used.
- If that is blank, SCH `Project Remark` is used.
- If both are blank, the workflow drafts a conservative update from status, milestone, risk, owner, and next action.

No separate weekly HTML files are created. Weekly updates are kept in one Excel mail-merge workbook.

Old folders under `D:\Reports` are history only unless you intentionally need an old file.
