from __future__ import annotations

import argparse
import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd

import pm_cockpit
from client_report_generator import build_report_data, create_presentation, create_workbook


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_CONFIG = BASE_DIR / "config.json"
DEFAULT_PM_CONFIG = BASE_DIR / "pm_config.json"
DEFAULT_FINAL_ROOT = Path(r"D:\VASU_PM_WORKFLOW")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run the simplified PM monthly workflow.")
    parser.add_argument("--report-month", default=None, help="Optional report month in YYYY-MM format, e.g. 2026-06.")
    parser.add_argument("--config", default=str(DEFAULT_CONFIG), help="Monthly report config path.")
    parser.add_argument("--pm-config", default=str(DEFAULT_PM_CONFIG), help="PM cockpit config path.")
    return parser.parse_args()


def load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def resolve_report_date(report_month: str | None, config: dict[str, Any]) -> pd.Timestamp:
    if report_month:
        period = pd.Period(report_month, freq="M")
        return pd.Timestamp(period.end_time.date())
    monthly = config.get("monthly_report", {})
    if monthly.get("period_mode") == "manual" and monthly.get("selected_report_month"):
        period = pd.Period(str(monthly["selected_report_month"]), freq="M")
        return pd.Timestamp(period.end_time.date())
    return pd.Timestamp(datetime.now().date())


def month_folder(root: Path, report_date: pd.Timestamp) -> Path:
    return root / str(report_date.year) / f"{report_date:%m-%B}"


def replace_copy(source: Path, destination: Path) -> None:
    destination.parent.mkdir(parents=True, exist_ok=True)
    if destination.exists():
        destination.unlink()
    shutil.copy2(source, destination)


def require_file(path: Path, label: str) -> None:
    if not path.exists() or path.name.startswith("~$"):
        raise FileNotFoundError(
            f"{label} not found: {path}\n"
            "Put the latest Tracker and project_master files in the final source folder, then run again."
        )


def archive_source_files(paths: list[Path], destination_dir: Path) -> None:
    destination_dir.mkdir(parents=True, exist_ok=True)
    for path in paths:
        if path.exists() and not path.name.startswith("~$"):
            replace_copy(path, destination_dir / path.name)


def build_mail_merge_rows(data: dict[str, pd.DataFrame], config: dict[str, Any]) -> pd.DataFrame:
    project_view = data.get("Project_Summary_View", pd.DataFrame()).copy()
    if project_view.empty:
        return pd.DataFrame()
    group_map = pm_cockpit.load_project_group_map(config, project_view)
    group_by_key = {
        str(row["Tracker Name Key"]): row
        for _, row in group_map.iterrows()
        if str(row.get("Tracker Name Key", "")).strip()
    }
    rows: list[dict[str, Any]] = []
    for _, row in project_view.iterrows():
        map_row = group_by_key.get(str(row.get("Tracker Name Key", "")), {})
        to_value = pm_cockpit.value_to_text(pm_cockpit.get_map_value(map_row, "Outlook Group / To"))
        cc_value = pm_cockpit.value_to_text(pm_cockpit.get_map_value(map_row, "CC"))
        project_name = pm_cockpit.value_to_text(row.get("Project Name"))
        subject = pm_cockpit.value_to_text(row.get("Stakeholder Update Subject")) or pm_cockpit.build_stakeholder_subject(row, config)
        pm_update = pm_cockpit.value_to_text(row.get("Stakeholder PM Update"))
        next_action = pm_cockpit.value_to_text(row.get("Next Action"))
        milestone = pm_cockpit.value_to_text(row.get("Current Milestone"))
        body = (
            "Dear Team,\n\n"
            "Please find the latest PM update for the below project.\n\n"
            f"Project: {project_name}\n"
            f"Current Milestone: {milestone}\n"
            f"Latest PM Update: {pm_update}\n"
            f"Next Action: {next_action}\n\n"
            f"Tracker Status: {pm_cockpit.value_to_text(row.get('Tracker Status'))}\n"
            f"SCH Status: {pm_cockpit.value_to_text(row.get('SCH Status'))}\n"
            f"Owner: {pm_cockpit.value_to_text(row.get('Owner'))}\n\n"
            f"Regards,\n{config.get('owner_name', 'Vasu')}"
        )
        rows.append(
            {
                "Send Flag": "Ready" if to_value else "Need Recipient",
                "To": to_value,
                "CC": cc_value,
                "Subject": subject,
                "Project Code": pm_cockpit.value_to_text(row.get("Project Code")),
                "Project Name": project_name,
                "Tracker Status": pm_cockpit.value_to_text(row.get("Tracker Status")),
                "SCH Status": pm_cockpit.value_to_text(row.get("SCH Status")),
                "Risk Level": pm_cockpit.value_to_text(row.get("Risk Level")),
                "Current Milestone": milestone,
                "Owner": pm_cockpit.value_to_text(row.get("Owner")),
                "PM Note Source": pm_cockpit.value_to_text(row.get("PM Note Source")),
                "Stakeholder PM Update": pm_update,
                "Next Action": next_action,
                "Due Date": pm_cockpit.value_to_text(row.get("Due Date")),
                "Mail Body": body,
                "Review Notes": "",
            }
        )
    return pd.DataFrame(rows)


def create_weekly_mail_merge_workbook(rows: pd.DataFrame, output_path: Path, config: dict[str, Any]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    theme = pm_cockpit.resolved_theme(config)
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        rows.to_excel(writer, sheet_name="Weekly_Mail_Merge", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Weekly_Mail_Merge"]
        header = workbook.add_format({"bold": True, "font_color": "white", "bg_color": theme["primary"], "border": 1})
        wrap = workbook.add_format({"text_wrap": True, "valign": "top", "border": 1})
        ready = workbook.add_format({"bg_color": theme["status_colors"]["Active"], "font_color": "white"})
        need = workbook.add_format({"bg_color": theme["status_colors"]["Yet To Start"], "font_color": "#1F1F1F"})
        for col_idx, column in enumerate(rows.columns):
            worksheet.write(0, col_idx, column, header)
            width = 18
            if column in {"Stakeholder PM Update", "Mail Body", "Next Action", "Review Notes"}:
                width = 55
            elif column in {"Project Name", "Subject"}:
                width = 36
            worksheet.set_column(col_idx, col_idx, width, wrap)
        if not rows.empty:
            worksheet.autofilter(0, 0, len(rows), len(rows.columns) - 1)
            worksheet.freeze_panes(1, 0)
            flag_col = rows.columns.get_loc("Send Flag")
            worksheet.conditional_format(1, flag_col, len(rows), flag_col, {"type": "text", "criteria": "containing", "value": "Ready", "format": ready})
            worksheet.conditional_format(1, flag_col, len(rows), flag_col, {"type": "text", "criteria": "containing", "value": "Need Recipient", "format": need})


def write_month_readme(path: Path, outputs: dict[str, Path], report_date: pd.Timestamp) -> None:
    text = f"""PM Monthly Workspace - {report_date:%B %Y}

Use this folder for the month. The active PM workflow lives in D:\\VASU_PM_WORKFLOW.

1. Source files used:
   {outputs['source_snapshot']}

2. Execution file:
   {outputs['execution']}

3. Monthly pack to send to Sir:
   {outputs['monthly_excel']}
   {outputs['monthly_ppt']}

4. Weekly mail-merge file:
   {outputs['weekly_mail_merge']}

Recommended flow:
- Download or copy latest Tracker and project_master files into D:\\VASU_PM_WORKFLOW\\01_Source_Files.
- Run D:\\VASU_PM_WORKFLOW\\RUN_PM_WORKFLOW.bat or the desktop shortcut.
- Review the execution file first.
- Use the monthly pack for Sir at month end.
- Use Weekly_Mail_Merge for project-wise stakeholder updates.
"""
    path.write_text(text, encoding="utf-8")


def main() -> int:
    args = parse_args()
    config = load_json(Path(args.config))
    pm_config = load_json(Path(args.pm_config))
    report_date = resolve_report_date(args.report_month, config)
    report_label = f"{report_date:%B}{report_date.year}"

    final_root = Path(pm_config.get("final_root", DEFAULT_FINAL_ROOT))
    source_dir = Path(pm_config.get("source_dir", final_root / "01_Source_Files"))
    workspace_root = Path(pm_config.get("workspace_root", final_root / "02_Output"))
    month_root = month_folder(workspace_root, report_date)
    source_snapshot_dir = month_root / "00_Source_Files_Used"
    execution_dir = month_root / "01_Execution"
    monthly_dir = month_root / "02_Monthly_To_Sir"
    weekly_dir = month_root / "03_Weekly_Mail_Merge"
    latest_dir = workspace_root / "Latest"
    for folder in [source_dir, source_snapshot_dir, execution_dir, monthly_dir, weekly_dir, latest_dir]:
        folder.mkdir(parents=True, exist_ok=True)

    pm_config["run_date"] = str(report_date.date())
    pm_config["source_dir"] = str(source_dir)
    pm_config["workspace_root"] = str(workspace_root)
    pm_config["output_root"] = str(workspace_root)
    pm_config["alias_review_file"] = str(Path(pm_config.get("alias_review_file", source_dir / "Alias_Matching_Review.xlsx")))

    tracker_file = Path(pm_config.get("tracker_file", source_dir / "VASU's Tracker 2026.xlsx"))
    sch_file = Path(pm_config.get("sch_file", source_dir / "project_master.xlsx"))
    require_file(tracker_file, "Tracker file")
    require_file(sch_file, "SCH project master file")
    archive_source_files([tracker_file, sch_file], source_snapshot_dir)

    pm_data = pm_cockpit.build_pm_dataset(pm_config)
    execution_path = execution_dir / f"PM_Execution_Master_{report_label}.xlsx"
    latest_execution_path = latest_dir / "PM_Execution_Master_Latest.xlsx"
    pm_cockpit.create_pm_cockpit_workbook(pm_data, execution_path, latest_execution_path, pm_config)
    replace_copy(execution_path, workspace_root / "PM_Cockpit_Latest.xlsx")
    replace_copy(execution_path, workspace_root / "PM_Execution_Master_Latest.xlsx")
    pm_cockpit.create_alias_matching_review_workbook(pm_data, Path(pm_config["alias_review_file"]), pm_config)

    weekly_rows = build_mail_merge_rows(pm_data, pm_config)
    weekly_path = weekly_dir / f"Weekly_Project_Updates_{report_label}.xlsx"
    create_weekly_mail_merge_workbook(weekly_rows, weekly_path, pm_config)
    replace_copy(weekly_path, latest_dir / "Weekly_Project_Updates_Latest.xlsx")

    report_data = build_report_data(tracker_file, report_date)
    monthly_excel = monthly_dir / f"VASU_Tracker_YoY_{report_label}.xlsx"
    monthly_ppt = monthly_dir / f"VASU_Tracker_YoY_{report_label}_Presentation.pptx"
    create_workbook(report_data, monthly_excel)
    create_presentation(report_data, monthly_ppt)
    replace_copy(monthly_excel, latest_dir / monthly_excel.name)
    replace_copy(monthly_ppt, latest_dir / monthly_ppt.name)

    if config.get("monthly_report", {}).get("legacy_month_copy_enabled", False):
        compatibility_dir = Path(config.get("monthly_report", {}).get("output_root", str(workspace_root))) / report_data.report_month
        compatibility_dir.mkdir(parents=True, exist_ok=True)
        replace_copy(monthly_excel, compatibility_dir / monthly_excel.name)
        replace_copy(monthly_ppt, compatibility_dir / monthly_ppt.name)

    outputs = {
        "source_snapshot": source_snapshot_dir,
        "execution": execution_path,
        "monthly_excel": monthly_excel,
        "monthly_ppt": monthly_ppt,
        "weekly_mail_merge": weekly_path,
    }
    write_month_readme(month_root / "README_THIS_MONTH.txt", outputs, report_date)

    print("PM_WORKFLOW_COMPLETE")
    print(f"MONTH_WORKSPACE={month_root}")
    print(f"EXECUTION_MASTER={execution_path}")
    print(f"MONTHLY_EXCEL={monthly_excel}")
    print(f"MONTHLY_PPT={monthly_ppt}")
    print(f"WEEKLY_MAIL_MERGE={weekly_path}")
    print(f"ALIAS_REVIEW={pm_config['alias_review_file']}")
    print(f"WEEKLY_ROWS={len(weekly_rows)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
