# PM Cockpit Guide

## Operating Flow

1. Update `D:\Reports\Current Status\VASU's Tracker 2026.xlsx`.
2. Download the latest SCH `project_master.xlsx` into `D:\Reports\Current Status`.
3. Run `run_pm_cockpit.bat`.
4. Open `D:\Reports\PM Cockpit\PM_Execution_Master_Latest.xlsx`.
5. Review `Execution_Master` and `Project_Summary_View` for current projects.
6. Review `Stakeholder PM Update`, `PM Note Source`, `Current Milestone`, `Next Action`, `Owner`, and `Due Date`.
7. If a project needs a specific human update, write it in `PM Notes`; otherwise keep it blank and let SCH update/remark become the default PM note.
8. Use the HTML files in `project_summaries` for separate project-wise stakeholder updates.
9. At month-end, run the monthly report/PPT generator.

Tracker is the primary file. SCH Project Master is used to enrich details such as official project remark, SCH status, implementation engineer, employee count, product, and client context. `Current Status.xlsx` is not required for the merge.

## PM Note Automation Logic

Every current project gets one ready-to-review stakeholder update.

1. If `PM Notes` has text, the system uses that as the main update.
2. If `PM Notes` is blank, the system uses `SCH Project Update`.
3. If `SCH Project Update` is blank, the system uses `SCH Project Remark`.
4. If SCH text is also blank, the system creates a conservative update from Tracker status, SCH status, milestone, risk, owner, and next action.

The output columns are:

- `PM Note Source`: tells you whether the note came from your manual PM note, SCH project update, SCH project remark, or a system default.
- `Stakeholder PM Update`: the final update text used in the project-wise HTML file.
- `Stakeholder Update Subject`: the email subject to use for that project.

Weekly consolidated mail is disabled. The intended weekly flow is one project-wise update per current project.

## Milestone Flow

Use these milestones for PM tracking and project-wise updates:

1. Kick Off Meeting
2. Datafile Explanation
3. Data & Policies Received
4. Payroll Configuration & Training
5. Salary Tallied
6. Parallel Salary Processed
7. Attendance Integration
8. Leave & Attendance Configuration
9. Employee Training & Login Roll Out
10. Live Salary Processing
11. Phase 2
12. Completed

## SCH Status Meaning

- `Pending For Milestone`: Yet to start; no progress or movement so far.
- `Active`: Project has started and is moving through implementation milestones.
- `AC`: Active Complete; implementation is complete and client is on watch list.
- `FC`: Fully Complete; live salary processing is complete and client is on watch list.
- `Go LIVE`: Handover from implementation team to support team.
- `Dropped`: Implementation dropped and license/client discontinued.
- `Held`: Implementation on hold while license/client is still running.
- `Lost`: Project/client lost and license deactivated.
- `FC Lost`: Project was fully completed but later lost.

## Alias Mapping: Step By Step

Use alias mapping when the tracker project name and SCH client name refer to the same project but do not match exactly.

1. Open `D:\Reports\PM Cockpit\config\Alias_Matching_Review.xlsx`.
2. Go to `Tracker_to_SCH`.
3. Review `Tracker Project Name`, `Suggested SCH Match 1/2/3`, and `Current Matched SCH Client`.
4. Put the correct SCH name in `Final SCH Client Name`.
5. Set `Approved` to `Yes` only when you are confident.
6. Use `SCH_to_Tracker` when you want the reverse view from SCH names back to tracker names.
7. Save the workbook.
8. Run `run_pm_cockpit.bat` again.
9. Confirm the project match method is updated in `Execution_Master`.

## Reviewing Existing Mappings

1. Open `Alias_Map`.
2. Review rows where `Match Type` is `Manual Alias` or `Legacy Current Status`.
3. Open `Match_Review` for duplicate SCH candidates.
4. Open `Status_Mismatch` to find tracker/SCH status conflicts.
5. If a mapping is wrong, update `client_alias_map.csv`:
   - Correct `SCH Client Name`, or
   - set `Approved` to `False`.
6. Rerun the cockpit and recheck the project.

Rule: do not approve uncertain matches. A missing alias is safer than a wrong alias because wrong mapping can pull the wrong SCH remarks into project updates.

## Completed Client Followup

Completed clients are now split across the quarter in `Completed_Followup_Plan`. Use this sheet to contact a manageable set each month and capture:

- current usage status
- challenges or pain points
- support required
- upsell scope
- PM follow-up notes

This replaces the generic go-live follow-up approach.

## Colors

Colors are controlled in `pm_config.json` under `theme`.

- Active: green
- Completed: light blue
- Yet To Start: yellow
- Held Up: orange
- Lost: red

## Outlook Group Mapping

When project-wise Outlook groups are ready:

1. Open `D:\Reports\PM Cockpit\config\project_outlook_group_map.csv`.
2. Fill `Outlook Group / To` for each project.
3. Add `CC` if needed.
4. Set `Enabled` to `Yes`.
5. In `pm_config.json`, set `project_summaries.create_outlook_drafts` to `true`.
6. Run `run_pm_cockpit.bat`.

The system will then create one reviewable Outlook draft per mapped current project.
