from __future__ import annotations

import argparse
import html
import json
import logging
import re
import sys
import traceback
from dataclasses import dataclass
from datetime import datetime, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


PM_FIELD_COLUMNS = [
    "Current Milestone",
    "PM Notes",
    "Next Action",
    "Owner",
    "Due Date",
    "Action Date",
    "Manual Risk Reason",
    "Mail Include",
]


TRACKER_COLUMNS = [
    "Project Code",
    "Project Type",
    "Project Name",
    "Received From",
    "Start Date",
    "End Date",
    "Priority",
    "Status",
]


@dataclass
class PMRunResult:
    cockpit_path: Path
    latest_path: Path
    weekly_html_path: Path
    project_summary_index_path: Path
    outlook_status: str
    summary: dict[str, Any]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create PM Cockpit and project-wise summaries from tracker + SCH export.")
    parser.add_argument("--config", default="pm_config.json", help="Path to PM config JSON.")
    parser.add_argument(
        "--mode",
        default="all",
        choices=["all", "cockpit", "project-summaries", "weekly-draft"],
        help="Run cockpit, project-wise summaries, or the legacy consolidated weekly draft.",
    )
    parser.add_argument("--no-outlook", action="store_true", help="Create HTML draft but skip Outlook draft creation.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = base_dir / config_path

    try:
        config = load_config(config_path)
        output_root = Path(config["output_root"])
        log_path = setup_logging(output_root / "logs")
        logging.info("PM Cockpit automation started")
        logging.info("Config: %s", config_path)

        data = build_pm_dataset(config)
        output_paths = prepare_output_paths(output_root)
        cockpit_path = output_paths["cockpit"]
        latest_path = output_paths["latest"]
        execution_path = output_paths["execution"]
        latest_execution_path = output_paths["latest_execution"]
        html_path = output_paths["html"]
        project_index_path = output_paths["project_index"]

        if args.mode in {"all", "cockpit"}:
            create_pm_cockpit_workbook(data, cockpit_path, latest_path, config)
            execution_path.write_bytes(cockpit_path.read_bytes())
            latest_execution_path.write_bytes(cockpit_path.read_bytes())
            create_alias_matching_review_workbook(data, Path(config["alias_review_file"]), config)
        elif not latest_path.exists():
            create_pm_cockpit_workbook(data, cockpit_path, latest_path, config)
            execution_path.write_bytes(cockpit_path.read_bytes())
            latest_execution_path.write_bytes(cockpit_path.read_bytes())
            create_alias_matching_review_workbook(data, Path(config["alias_review_file"]), config)

        project_summary_status = "Skipped by mode"
        if args.mode in {"all", "project-summaries"} and config.get("project_summaries", {}).get("enabled", True):
            project_index = create_project_summaries(data, output_paths["project_dir"], project_index_path, config, skip_outlook=args.no_outlook)
            data["Project_Summary_Index"] = project_index
            project_summary_status = f"Created {len(project_index)} project summaries"

        outlook_status = project_summary_status
        weekly_created = False
        if args.mode == "weekly-draft" or (
            args.mode == "all" and config.get("weekly_consolidated", {}).get("enabled", False)
        ):
            body = build_weekly_email_html(data, config)
            html_path.parent.mkdir(parents=True, exist_ok=True)
            html_path.write_text(body, encoding="utf-8")
            weekly_created = True
            if args.no_outlook:
                outlook_status = "Skipped by --no-outlook"
            else:
                outlook_status = create_outlook_draft(body, latest_path, config)

        summary = build_run_summary(data)
        write_run_summary(
            output_root / "latest_run_summary.json",
            summary,
            cockpit_path,
            html_path if weekly_created else None,
            project_index_path,
            outlook_status,
            log_path,
        )

        print("PM_COCKPIT_COMPLETE")
        print(f"COCKPIT={cockpit_path}")
        print(f"LATEST_COCKPIT={latest_path}")
        print(f"EXECUTION_MASTER={execution_path}")
        print(f"LATEST_EXECUTION_MASTER={latest_execution_path}")
        print(f"ALIAS_REVIEW_FILE={config['alias_review_file']}")
        if weekly_created:
            print(f"WEEKLY_DRAFT_HTML={html_path}")
        else:
            print("CONSOLIDATED_WEEKLY_DRAFT=Disabled")
        print(f"PROJECT_SUMMARY_INDEX={project_index_path}")
        print(f"OUTLOOK_STATUS={outlook_status}")
        for key, value in summary.items():
            print(f"{key.upper()}={value}")
        print(f"LOG_FILE={log_path}")
        return 0
    except Exception as exc:
        try:
            output_root = Path(load_config(config_path).get("output_root", Path.cwd()))
            error_dir = output_root / "logs"
            error_dir.mkdir(parents=True, exist_ok=True)
            error_path = error_dir / f"pm_cockpit_error_{datetime.now():%Y%m%d_%H%M%S}.txt"
            error_path.write_text(
                f"PM Cockpit automation failed\n\nError: {exc}\n\n{traceback.format_exc()}",
                encoding="utf-8",
            )
            print("PM_COCKPIT_FAILED", file=sys.stderr)
            print(f"ERROR_REPORT={error_path}", file=sys.stderr)
        except Exception:
            pass
        logging.exception("PM Cockpit automation failed")
        print(str(exc), file=sys.stderr)
        return 1


def load_config(config_path: Path) -> dict[str, Any]:
    with config_path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)
    return config


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"pm_cockpit_{datetime.now():%Y%m%d_%H%M%S}.log"
    root = logging.getLogger()
    for handler in list(root.handlers):
        root.removeHandler(handler)
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(levelname)s | %(message)s"))
    root.addHandler(console)
    return log_path


def prepare_output_paths(output_root: Path) -> dict[str, Path]:
    now = datetime.now()
    dated_dir = output_root / str(now.year) / f"{now:%m-%B}"
    return {
        "cockpit": dated_dir / f"PM_Cockpit_{now:%Y-%m-%d}.xlsx",
        "latest": output_root / "PM_Cockpit_Latest.xlsx",
        "execution": dated_dir / f"PM_Execution_Master_{now:%Y-%m-%d}.xlsx",
        "latest_execution": output_root / "PM_Execution_Master_Latest.xlsx",
        "html": dated_dir / "weekly_drafts" / f"Weekly_Update_Draft_{now:%Y-%m-%d}.html",
        "project_dir": dated_dir / "project_summaries",
        "project_index": dated_dir / "project_summaries" / f"Project_Summary_Index_{now:%Y-%m-%d}.csv",
    }


def config_run_timestamp(config: dict[str, Any]) -> pd.Timestamp:
    run_date = config.get("run_date") or config.get("report_date")
    if run_date:
        return pd.Timestamp(run_date)
    return pd.Timestamp(datetime.now().date())


def build_pm_dataset(config: dict[str, Any]) -> dict[str, pd.DataFrame]:
    tracker = load_tracker(config)
    sch = load_sch(config)
    alias_map = load_alias_map(config, tracker, sch)
    previous_notes = load_previous_notes(config)
    merged, unmatched, match_review = merge_tracker_sch(tracker, sch, alias_map, config)
    merged = enrich_with_pm_notes(merged, previous_notes)
    merged = add_pm_intelligence(merged, config)

    execution_master = build_execution_master(merged)
    alias_matching_tracker = build_alias_matching_tracker(tracker, sch, alias_map, merged, config)
    alias_matching_sch = build_alias_matching_sch(tracker, sch, alias_map, merged, config)
    completed_followup_plan = build_completed_followup_plan(merged, config)
    action_list = build_action_list(merged)
    stale_items = merged[merged["Risk Level"].isin(["Critical", "High", "Watch"])].copy()
    status_mismatch = merged[merged["Status Mismatch"].eq(True)].copy()
    weekly_mail_view = build_weekly_mail_view(merged, config)
    project_summary_view = build_project_summary_view(merged, config)
    monthly_summary = build_monthly_summary(merged)
    pm_playbook = build_pm_playbook(config)
    alias_mapping_guide = build_alias_mapping_guide()
    theme_config = build_theme_config(config)
    run_log = build_validation_log(tracker, sch, merged, unmatched, previous_notes)

    return {
        "Execution_Master": execution_master,
        "PM_Cockpit": merged,
        "Action_List": action_list,
        "Completed_Followup_Plan": completed_followup_plan,
        "Stale_Items": stale_items,
        "Status_Mismatch": status_mismatch,
        "Alias_Matching_Tracker": alias_matching_tracker,
        "Alias_Matching_SCH": alias_matching_sch,
        "Unmatched_Clients": unmatched,
        "Match_Review": match_review,
        "Alias_Map": alias_map,
        "Theme_Config": theme_config,
        "Weekly_Mail_View": weekly_mail_view,
        "Project_Summary_View": project_summary_view,
        "Monthly_Summary": monthly_summary,
        "PM_Playbook": pm_playbook,
        "Alias_Mapping_Guide": alias_mapping_guide,
        "TrackerData": tracker,
        "SCH_Project_Master": sch,
        "Preserved_Notes": previous_notes,
        "Run_Log": run_log,
    }


def load_tracker(config: dict[str, Any]) -> pd.DataFrame:
    path = Path(config["tracker_file"])
    if not path.exists():
        raise FileNotFoundError(f"Tracker file not found: {path}")
    df = pd.read_excel(path, sheet_name=config.get("tracker_sheet", "Sheet1"))
    df = normalize_columns(df)
    if "Start date" in df.columns and "Start Date" not in df.columns:
        df = df.rename(columns={"Start date": "Start Date"})
    if "Assigned Engineer " in df.columns and "Assigned Engineer" not in df.columns:
        df = df.rename(columns={"Assigned Engineer ": "Assigned Engineer"})
    df = repair_wrapped_project_names(df)

    missing = [column for column in TRACKER_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"Tracker is missing required columns: {', '.join(missing)}")

    df = df[[column for column in df.columns if not str(column).startswith("Unnamed")]].copy()
    for column in df.select_dtypes(include=["object"]).columns:
        df[column] = clean_text_series(df[column])
    df["Start Date"] = pd.to_datetime(df["Start Date"], errors="coerce")
    df["End Date"] = pd.to_datetime(df["End Date"], errors="coerce")

    df = df[
        df["Project Code"].notna()
        & df["Project Name"].notna()
        & df["Status"].notna()
        & df["Start Date"].notna()
    ].copy()
    df["Project Code"] = pd.to_numeric(df["Project Code"], errors="coerce").astype("Int64")
    df["Tracker Name Key"] = df["Project Name"].apply(normalize_name)
    df["Tracker Row Number"] = np.arange(2, len(df) + 2)
    logging.info("Loaded tracker rows: %s", len(df))
    return df.reset_index(drop=True)


def load_sch(config: dict[str, Any]) -> pd.DataFrame:
    path = Path(config["sch_file"])
    if not path.exists():
        raise FileNotFoundError(f"SCH export not found: {path}")
    df = pd.read_excel(path)
    df = normalize_columns(df)
    if "Client Name" not in df.columns:
        raise ValueError("SCH export is missing required column: Client Name")

    for column in df.select_dtypes(include=["object"]).columns:
        df[column] = clean_text_series(df[column])
    date_cols = [
        "Project AC Date",
        "Amc Date",
        "Project Start Date",
        "Project Installation Date",
        "Project Implementation Date",
        "Project Close Date",
    ]
    for column in date_cols:
        if column in df.columns:
            df[column] = pd.to_datetime(df[column], errors="coerce")
    df["SCH Name Key"] = df["Client Name"].apply(normalize_name)
    df["SCH Freshness Date"] = df[[column for column in date_cols if column in df.columns]].max(axis=1, skipna=True)
    df["SCH Row Number"] = np.arange(2, len(df) + 2)
    logging.info("Loaded SCH rows: %s", len(df))
    return df.reset_index(drop=True)


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    cleaned.columns = [re.sub(r"\s+", " ", str(column).replace("\xa0", " ")).strip() for column in cleaned.columns]
    return cleaned


def clean_text_series(series: pd.Series) -> pd.Series:
    return (
        series.astype("string")
        .str.replace("\xa0", " ", regex=False)
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    )


def repair_wrapped_project_names(df: pd.DataFrame) -> pd.DataFrame:
    if "Project Code" not in df.columns or "Project Name" not in df.columns:
        return df
    fixed = df.copy()
    for idx in range(1, len(fixed)):
        current_code = fixed.at[idx, "Project Code"]
        current_name = fixed.at[idx, "Project Name"]
        if pd.isna(current_code) and pd.notna(current_name):
            previous_idx = idx - 1
            while previous_idx >= 0 and pd.isna(fixed.at[previous_idx, "Project Code"]):
                previous_idx -= 1
            if previous_idx >= 0 and pd.notna(fixed.at[previous_idx, "Project Name"]):
                fixed.at[previous_idx, "Project Name"] = f"{fixed.at[previous_idx, 'Project Name']} {current_name}".strip()
            fixed.at[idx, "Project Name"] = pd.NA
    return fixed


def normalize_name(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).lower()
    replacements = {
        "&": "and",
        "private": "pvt",
        "limited": "ltd",
        "company": "co",
        "technologies": "tech",
        "international": "intl",
    }
    for old, new in replacements.items():
        text = re.sub(rf"\b{old}\b", new, text)
    return re.sub(r"[^a-z0-9]+", "", text)


def load_alias_map(config: dict[str, Any], tracker: pd.DataFrame, sch: pd.DataFrame) -> pd.DataFrame:
    alias_path = alias_map_path(config)
    alias_path.parent.mkdir(parents=True, exist_ok=True)
    columns = [
        "Tracker Project Name",
        "Tracker Name Key",
        "SCH Client Name",
        "SCH Name Key",
        "Approved",
        "Match Type",
        "Last Reviewed",
        "Notes",
    ]
    if alias_path.exists():
        alias_map = pd.read_csv(alias_path)
        for column in columns:
            if column not in alias_map.columns:
                alias_map[column] = ""
    else:
        alias_map = pd.DataFrame(columns=columns)
    review_aliases = build_review_aliases(config)
    legacy_aliases = build_legacy_aliases(config) if config.get("use_current_status_file", False) else pd.DataFrame()
    if not review_aliases.empty or not legacy_aliases.empty:
        alias_inputs = [df for df in [alias_map, review_aliases, legacy_aliases] if not df.empty]
        alias_map = pd.concat(alias_inputs, ignore_index=True) if alias_inputs else pd.DataFrame(columns=columns)
    alias_map["Tracker Name Key"] = alias_map.apply(
        lambda row: row["Tracker Name Key"] if pd.notna(row.get("Tracker Name Key")) and row.get("Tracker Name Key") else normalize_name(row.get("Tracker Project Name")),
        axis=1,
    )
    alias_map["SCH Name Key"] = alias_map.apply(
        lambda row: row["SCH Name Key"] if pd.notna(row.get("SCH Name Key")) and row.get("SCH Name Key") else normalize_name(row.get("SCH Client Name")),
        axis=1,
    )
    alias_map["Approved"] = alias_map["Approved"].astype("string").str.lower().isin(["yes", "true", "1", "approved"])
    alias_map = alias_map[alias_map["Tracker Name Key"].astype(str).ne("") & alias_map["SCH Name Key"].astype(str).ne("")]
    alias_map = alias_map.sort_values(["Tracker Name Key", "Match Type"]).drop_duplicates(
        subset=["Tracker Name Key", "SCH Name Key"], keep="first"
    )
    alias_map.to_csv(alias_path, index=False)
    logging.info("Loaded alias map rows: %s", len(alias_map))
    return alias_map[columns]


def alias_map_path(config: dict[str, Any]) -> Path:
    return Path(config["output_root"]) / "config" / "client_alias_map.csv"


def build_review_aliases(config: dict[str, Any]) -> pd.DataFrame:
    review_path = Path(config.get("alias_review_file", ""))
    if not review_path.exists():
        return pd.DataFrame()
    alias_rows: list[dict[str, Any]] = []
    sheet_specs = [
        ("Tracker_to_SCH", "Tracker Project Name", "Final SCH Client Name"),
        ("SCH_to_Tracker", "Final Tracker Project Name", "SCH Client Name"),
    ]
    for sheet_name, tracker_col, sch_col in sheet_specs:
        try:
            df = pd.read_excel(review_path, sheet_name=sheet_name)
        except Exception:
            continue
        df = normalize_columns(df)
        if tracker_col not in df.columns or sch_col not in df.columns:
            continue
        approved = (
            df["Approved"].astype("string").str.lower().isin(["yes", "true", "1", "approved"])
            if "Approved" in df.columns
            else pd.Series(False, index=df.index)
        )
        df = df[approved].copy()
        for _, row in df.iterrows():
            tracker_name = value_to_text(row.get(tracker_col))
            sch_name = value_to_text(row.get(sch_col))
            if not tracker_name or not sch_name:
                continue
            alias_rows.append(
                {
                    "Tracker Project Name": tracker_name,
                    "Tracker Name Key": normalize_name(tracker_name),
                    "SCH Client Name": sch_name,
                    "SCH Name Key": normalize_name(sch_name),
                    "Approved": True,
                    "Match Type": f"Alias Review {sheet_name}",
                    "Last Reviewed": datetime.now().strftime("%d-%b-%Y"),
                    "Notes": "Approved from Alias_Matching_Review workbook.",
                }
            )
    if not alias_rows:
        return pd.DataFrame()
    return pd.DataFrame(alias_rows).drop_duplicates(subset=["Tracker Name Key", "SCH Name Key"])


def build_legacy_aliases(config: dict[str, Any]) -> pd.DataFrame:
    current_status = Path(config.get("current_status_file", ""))
    if not current_status.exists():
        return pd.DataFrame()
    try:
        df = pd.read_excel(current_status, sheet_name="Raw Data")
    except Exception as exc:
        logging.warning("Could not seed aliases from Current Status: %s", exc)
        return pd.DataFrame()
    df = normalize_columns(df)
    if "Project Name" not in df.columns or "project_master xlsx.Client Name" not in df.columns:
        return pd.DataFrame()
    df["Project Name"] = clean_text_series(df["Project Name"])
    df["project_master xlsx.Client Name"] = clean_text_series(df["project_master xlsx.Client Name"])
    df = df[df["Project Name"].notna() & df["project_master xlsx.Client Name"].notna()].copy()
    if df.empty:
        return pd.DataFrame()
    aliases = pd.DataFrame(
        {
            "Tracker Project Name": df["Project Name"],
            "Tracker Name Key": df["Project Name"].apply(normalize_name),
            "SCH Client Name": df["project_master xlsx.Client Name"],
            "SCH Name Key": df["project_master xlsx.Client Name"].apply(normalize_name),
            "Approved": True,
            "Match Type": "Legacy Current Status",
            "Last Reviewed": datetime.now().strftime("%d-%b-%Y"),
            "Notes": "Seeded from existing Current Status workbook.",
        }
    )
    return aliases.drop_duplicates(subset=["Tracker Name Key", "SCH Name Key"])


def load_previous_notes(config: dict[str, Any]) -> pd.DataFrame:
    sources: list[pd.DataFrame] = []
    output_root = Path(config["output_root"])
    latest_cockpit = output_root / "PM_Cockpit_Latest.xlsx"
    if latest_cockpit.exists():
        sources.append(read_notes_source(latest_cockpit, "PM_Cockpit", "Previous PM Cockpit", priority=30))
    current_status = Path(config.get("current_status_file", ""))
    if config.get("use_current_status_file", False) and current_status.exists():
        sources.append(read_notes_source(current_status, "Raw Data", "Current Status", priority=20))
    book1 = Path(config.get("book1_notes_file", ""))
    if book1.exists():
        sources.append(read_notes_source(book1, "Sheet1", "Book1", priority=25))

    sources = [source for source in sources if not source.empty]
    if not sources:
        return pd.DataFrame(columns=["Project Name", "Notes Key", "Source", "Priority", *PM_FIELD_COLUMNS])

    notes = pd.concat(sources, ignore_index=True)
    notes = notes[notes["Notes Key"].ne("")]
    notes = notes.sort_values(["Notes Key", "Priority"], ascending=[True, False])
    combined_rows: list[dict[str, Any]] = []
    for key, group in notes.groupby("Notes Key", dropna=False):
        row: dict[str, Any] = {
            "Notes Key": key,
            "Project Name": first_nonblank(group["Project Name"]),
            "Source": ", ".join(group["Source"].dropna().astype(str).unique()),
            "Priority": int(group["Priority"].max()),
        }
        for column in PM_FIELD_COLUMNS:
            row[column] = first_nonblank(group[column]) if column in group.columns else ""
        combined_rows.append(row)
    combined = pd.DataFrame(combined_rows)
    logging.info("Loaded preserved note keys: %s", len(combined))
    return combined


def read_notes_source(path: Path, sheet_name: str, source_name: str, priority: int) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except Exception as exc:
        logging.warning("Could not read notes from %s [%s]: %s", path, sheet_name, exc)
        return pd.DataFrame()
    df = normalize_columns(df)
    if "Project Name" not in df.columns:
        return pd.DataFrame()
    for column in PM_FIELD_COLUMNS:
        if column not in df.columns:
            df[column] = ""
    df["Project Name"] = clean_text_series(df["Project Name"])
    df["Notes Key"] = df["Project Name"].apply(normalize_name)
    df["Source"] = source_name
    df["Priority"] = priority
    keep = ["Project Name", "Notes Key", "Source", "Priority", *PM_FIELD_COLUMNS]
    has_note = df[PM_FIELD_COLUMNS].notna().any(axis=1)
    return df.loc[has_note, keep].copy()


def first_nonblank(series: pd.Series) -> Any:
    for value in series:
        if pd.notna(value) and str(value).strip():
            return value
    return ""


def merge_tracker_sch(
    tracker: pd.DataFrame,
    sch: pd.DataFrame,
    alias_map: pd.DataFrame,
    config: dict[str, Any],
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    sch_groups = {key: group.copy() for key, group in sch.groupby("SCH Name Key") if key}
    approved_aliases = alias_map[alias_map["Approved"].eq(True)].copy()
    alias_by_tracker = {
        row["Tracker Name Key"]: row["SCH Name Key"]
        for _, row in approved_aliases.iterrows()
        if row.get("Tracker Name Key") and row.get("SCH Name Key")
    }
    sch_names = sch[["Client Name", "SCH Name Key"]].dropna().drop_duplicates()

    merged_rows: list[dict[str, Any]] = []
    unmatched_rows: list[dict[str, Any]] = []
    review_rows: list[dict[str, Any]] = []

    for _, tracker_row in tracker.iterrows():
        tracker_key = tracker_row["Tracker Name Key"]
        alias_key = alias_by_tracker.get(tracker_key)
        if alias_key and alias_key in sch_groups:
            candidates = sch_groups[alias_key]
            match_method = "Approved Alias"
            match_score = 1.0
        elif tracker_key in sch_groups:
            candidates = sch_groups[tracker_key]
            match_method = "Exact Name"
            match_score = 1.0
        else:
            candidates = pd.DataFrame()
            match_method = "No Match"
            match_score = 0.0

        if candidates.empty:
            fuzzy = fuzzy_candidates(tracker_row["Project Name"], sch_names, config)
            unmatched_rows.append(
                {
                    "Tracker Project Code": tracker_row.get("Project Code"),
                    "Tracker Project Name": tracker_row.get("Project Name"),
                    "Tracker Status": tracker_row.get("Status"),
                    "Best SCH Candidate 1": fuzzy[0][0] if len(fuzzy) > 0 else "",
                    "Score 1": fuzzy[0][1] if len(fuzzy) > 0 else "",
                    "Best SCH Candidate 2": fuzzy[1][0] if len(fuzzy) > 1 else "",
                    "Score 2": fuzzy[1][1] if len(fuzzy) > 1 else "",
                    "Best SCH Candidate 3": fuzzy[2][0] if len(fuzzy) > 2 else "",
                    "Score 3": fuzzy[2][1] if len(fuzzy) > 2 else "",
                    "Recommended Action": "Add approved alias if this project exists in SCH.",
                }
            )
            merged_rows.append(build_merged_row(tracker_row, None, match_method, match_score, 0))
        else:
            selected = select_sch_candidate(candidates, tracker_row.get("Status"))
            match_count = len(candidates)
            if match_count > 1:
                review_rows.append(
                    {
                        "Tracker Project Code": tracker_row.get("Project Code"),
                        "Tracker Project Name": tracker_row.get("Project Name"),
                        "Tracker Status": tracker_row.get("Status"),
                        "Match Method": match_method,
                        "Candidate Count": match_count,
                        "Selected SCH Client": selected.get("Client Name"),
                        "Selected SCH Status": selected.get("Status"),
                        "Recommended Action": "Review duplicate SCH candidates and add alias if selection is not correct.",
                    }
                )
            merged_rows.append(build_merged_row(tracker_row, selected, match_method, match_score, match_count))

    unmatched = pd.DataFrame(unmatched_rows)
    match_review = pd.DataFrame(review_rows)
    merged = pd.DataFrame(merged_rows)
    logging.info("Merged rows: %s | unmatched: %s | match review: %s", len(merged), len(unmatched), len(match_review))
    return merged, unmatched, match_review


def fuzzy_candidates(project_name: Any, sch_names: pd.DataFrame, config: dict[str, Any]) -> list[tuple[str, float]]:
    limit = int(config.get("fuzzy_candidate_limit", 3))
    min_score = float(config.get("fuzzy_candidate_min_score", 0.58))
    key = normalize_name(project_name)
    candidates: list[tuple[str, float]] = []
    if not key:
        return candidates
    for _, row in sch_names.iterrows():
        sch_key = row["SCH Name Key"]
        if not sch_key:
            continue
        score = SequenceMatcher(None, key, sch_key).ratio()
        if score >= min_score:
            candidates.append((row["Client Name"], round(score, 3)))
    return sorted(candidates, key=lambda item: item[1], reverse=True)[:limit]


def select_sch_candidate(candidates: pd.DataFrame, tracker_status: Any) -> pd.Series:
    status = str(tracker_status).strip().lower()
    preferred = {
        "completed": ["go live", "fc", "ac", "pending for milestone", "active"],
        "lost": ["lost", "fc lost", "dropped", "held"],
        "active": ["active", "pending for milestone", "ac", "fc", "held"],
        "yet to start": ["pending for milestone", "active", "ac", "fc", "held"],
        "held up": ["held", "active", "pending for milestone", "ac", "fc"],
    }
    order = preferred.get(status, ["active", "pending for milestone", "ac", "fc", "go live", "held", "lost"])

    def score(row: pd.Series) -> tuple[int, pd.Timestamp]:
        sch_status = str(row.get("Status", "")).strip().lower()
        status_rank = order.index(sch_status) if sch_status in order else len(order)
        freshness = row.get("SCH Freshness Date")
        if pd.isna(freshness):
            freshness = pd.Timestamp("1900-01-01")
        return (status_rank, -pd.Timestamp(freshness).timestamp())

    ranked = candidates.copy()
    ranked["_rank"] = ranked.apply(score, axis=1)
    ranked = ranked.sort_values("_rank")
    return ranked.iloc[0]


def build_merged_row(
    tracker_row: pd.Series,
    sch_row: pd.Series | None,
    match_method: str,
    match_score: float,
    match_count: int,
) -> dict[str, Any]:
    row: dict[str, Any] = {
        "Project Code": tracker_row.get("Project Code"),
        "Project Type": tracker_row.get("Project Type"),
        "Project Name": tracker_row.get("Project Name"),
        "Received From": tracker_row.get("Received From"),
        "Start Date": tracker_row.get("Start Date"),
        "End Date": tracker_row.get("End Date"),
        "Tracker Priority": tracker_row.get("Priority"),
        "Tracker Status": tracker_row.get("Status"),
        "Employee Count": tracker_row.get("Employee Count"),
        "Product": tracker_row.get("Product"),
        "Assigned Engineer": tracker_row.get("Assigned Engineer"),
        "Team Leader": tracker_row.get("Team Leader"),
        "Marketing Executive": tracker_row.get("Marketing Executive"),
        "Tracker Client Type": tracker_row.get("Client Type"),
        "Tracker Name Key": tracker_row.get("Tracker Name Key"),
        "Match Method": match_method,
        "Match Score": match_score,
        "SCH Match Count": match_count,
        "Match Review Needed": bool(match_count > 1 or match_method == "No Match"),
    }
    sch_columns = {
        "Client Code": "SCH Client Code",
        "Project Code": "SCH Project Code",
        "Client Status": "SCH Client Status",
        "Client Name": "SCH Client Name",
        "Priority": "SCH Priority",
        "Class": "SCH Class",
        "Co-ordinator": "SCH Co-ordinator",
        "Team Leader": "SCH Team Leader",
        "Implementation Engg": "SCH Implementation Engg",
        "Project MKT": "SCH Project MKT",
        "Project Team Manager": "SCH Project Team Manager",
        "Project Manager": "SCH Project Manager",
        "Client Type": "SCH Client Type",
        "No. Emp": "SCH No. Emp",
        "Project Start Date": "SCH Project Start Date",
        "Project Implementation Date": "SCH Project Implementation Date",
        "Project Close Date": "SCH Project Close Date",
        "Type of Product": "SCH Product Type",
        "Achived Percent": "SCH Achieved Percent",
        "Project Product": "SCH Project Product",
        "Status": "SCH Status",
        "Project Mode": "SCH Project Mode",
        "Project Remark": "SCH Project Remark",
        "Project Update": "SCH Project Update",
        "Last Processed Year Month": "SCH Last Processed Year Month",
        "SCH Name Key": "SCH Name Key",
        "SCH Row Number": "SCH Row Number",
    }
    for source_column, output_column in sch_columns.items():
        row[output_column] = sch_row.get(source_column) if sch_row is not None and source_column in sch_row.index else ""
    return row


def enrich_with_pm_notes(merged: pd.DataFrame, previous_notes: pd.DataFrame) -> pd.DataFrame:
    df = merged.copy()
    notes = previous_notes.copy()
    if notes.empty:
        for column in PM_FIELD_COLUMNS:
            df[column] = ""
        df["Notes Source"] = ""
        return df

    note_columns = ["Notes Key", "Source", *PM_FIELD_COLUMNS]
    notes = notes[[column for column in note_columns if column in notes.columns]].copy()
    df = df.merge(notes, how="left", left_on="Tracker Name Key", right_on="Notes Key")
    df = df.rename(columns={"Source": "Notes Source"})
    for column in PM_FIELD_COLUMNS:
        if column not in df.columns:
            df[column] = ""
        df[column] = df[column].fillna("")
    df = df.drop(columns=[column for column in ["Notes Key"] if column in df.columns])
    return df


def add_pm_intelligence(df: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    enriched = df.copy()
    today = config_run_timestamp(config)
    start_dates = pd.to_datetime(enriched["Start Date"], errors="coerce")
    end_dates = pd.to_datetime(enriched["End Date"], errors="coerce")
    enriched["Age Days"] = (today - start_dates).dt.days
    enriched["Cycle Days"] = (end_dates - start_dates).dt.days
    definitions = config.get("sch_status_definitions", {})
    enriched["SCH Status Meaning"] = enriched["SCH Status"].apply(lambda value: status_definition(value, definitions, "meaning"))
    enriched["SCH Stage Group"] = enriched["SCH Status"].apply(lambda value: status_definition(value, definitions, "stage"))
    enriched["SCH Suggested Milestone"] = enriched["SCH Status"].apply(lambda value: status_definition(value, definitions, "suggested_milestone"))
    enriched["SCH Milestone"] = enriched["SCH Suggested Milestone"]
    enriched["Status Mismatch"] = enriched.apply(is_status_mismatch, axis=1)

    risk = enriched.apply(lambda row: classify_risk(row, config), axis=1, result_type="expand")
    enriched["Risk Level"] = risk[0]
    enriched["Risk Reason"] = risk[1]

    enriched["Owner"] = enriched["Owner"].replace("", pd.NA)
    enriched["Owner"] = enriched["Owner"].fillna(enriched["Assigned Engineer"]).fillna(enriched["SCH Implementation Engg"]).fillna("")
    enriched["Current Milestone"] = enriched["Current Milestone"].replace("", pd.NA)
    enriched["Current Milestone"] = enriched["Current Milestone"].fillna(enriched.apply(default_current_milestone, axis=1)).fillna("")
    enriched["Next Action"] = enriched.apply(default_next_action, axis=1)
    enriched["Manual Risk Reason"] = enriched["Manual Risk Reason"].fillna("")
    enriched["Due Date"] = enriched["Due Date"].fillna("")
    enriched["Action Date"] = enriched["Action Date"].fillna("")
    enriched["PM Notes"] = enriched["PM Notes"].fillna("")
    enriched["Mail Include"] = enriched.apply(lambda row: default_mail_include(row, config), axis=1)
    pm_updates = enriched.apply(
        lambda row: pd.Series(
            build_stakeholder_pm_update(row, config),
            index=["PM Note Source", "Stakeholder PM Update"],
        ),
        axis=1,
    )
    enriched[["PM Note Source", "Stakeholder PM Update"]] = pm_updates
    enriched["Stakeholder Update Subject"] = enriched.apply(lambda row: build_stakeholder_subject(row, config), axis=1)
    enriched["Effective Employee Count"] = enriched["Employee Count"].replace("", pd.NA).fillna(enriched["SCH No. Emp"]).fillna("")
    enriched["Effective Product"] = (
        enriched["Product"]
        .replace("", pd.NA)
        .fillna(enriched["SCH Project Product"])
        .fillna(enriched["SCH Product Type"])
        .fillna("")
    )
    enriched["Effective Client Type"] = enriched["Tracker Client Type"].replace("", pd.NA).fillna(enriched["SCH Client Type"]).fillna("")

    display_order = [
        "Project Code",
        "Project Type",
        "Project Name",
        "Received From",
        "Start Date",
        "End Date",
        "Tracker Priority",
        "Tracker Status",
        "Age Days",
        "Cycle Days",
        "Risk Level",
        "Risk Reason",
        "Status Mismatch",
        "SCH Status",
        "SCH Status Meaning",
        "SCH Stage Group",
        "SCH Milestone",
        "Current Milestone",
        "SCH Client Name",
        "SCH Project Code",
        "SCH Client Status",
        "SCH Project Remark",
        "SCH Project Update",
        "SCH Achieved Percent",
        "SCH Last Processed Year Month",
        "Owner",
        "Effective Employee Count",
        "Effective Product",
        "Effective Client Type",
        "PM Note Source",
        "Stakeholder PM Update",
        "Stakeholder Update Subject",
        "Next Action",
        "Due Date",
        "Action Date",
        "PM Notes",
        "Manual Risk Reason",
        "Mail Include",
        "Match Method",
        "Match Score",
        "SCH Match Count",
        "Match Review Needed",
        "Assigned Engineer",
        "Team Leader",
        "Marketing Executive",
        "Product",
        "Employee Count",
        "SCH Implementation Engg",
        "SCH Project Manager",
        "SCH Project MKT",
        "SCH No. Emp",
        "Tracker Name Key",
        "SCH Name Key",
        "Notes Source",
    ]
    remaining = [column for column in enriched.columns if column not in display_order]
    return enriched[[column for column in display_order if column in enriched.columns] + remaining]


def is_status_mismatch(row: pd.Series) -> bool:
    tracker = str(row.get("Tracker Status", "")).strip().lower()
    sch = str(row.get("SCH Status", "")).strip().lower()
    if not sch or sch == "nan":
        return False
    if tracker == "completed" and sch not in {"go live", "fc", "fc lost"}:
        return True
    if tracker in {"active", "yet to start", "held up"} and sch in {"go live", "lost", "fc lost", "dropped"}:
        return True
    if tracker == "lost" and sch not in {"lost", "fc lost", "dropped", "held"}:
        return True
    return False


def status_definition(value: Any, definitions: dict[str, Any], field: str) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    definition = definitions.get(text, {})
    if not definition:
        for key, candidate in definitions.items():
            if key.lower() == text.lower():
                definition = candidate
                break
    return str(definition.get(field, "")) if isinstance(definition, dict) else ""


def build_stakeholder_pm_update(row: pd.Series, config: dict[str, Any]) -> tuple[str, str]:
    settings = config.get("auto_pm_note", {})
    max_note_chars = int(settings.get("max_note_characters", 1400))
    max_sch_chars = int(settings.get("max_sch_context_characters", 900))
    manual_note = clean_display_text(row.get("PM Notes"), max_note_chars)
    if manual_note:
        return "Manual PM Note", ensure_sentence(manual_note)

    fallback_order = settings.get("fallback_order", ["SCH Project Update", "SCH Project Remark"])
    source = "Tracker/SCH Status"
    sch_context = ""
    for column in fallback_order:
        sch_context = clean_display_text(row.get(column), max_sch_chars)
        if sch_context:
            source = str(column)
            break

    project_name = clean_display_text(row.get("Project Name"), 160)
    tracker_status = clean_display_text(row.get("Tracker Status"), 80)
    sch_status = clean_display_text(row.get("SCH Status"), 80)
    sch_meaning = clean_display_text(row.get("SCH Status Meaning"), 220)
    milestone = clean_display_text(row.get("Current Milestone"), 160)
    owner = clean_display_text(row.get("Owner"), 120)
    next_action = clean_display_text(row.get("Next Action"), 320)
    due_date = clean_display_text(row.get("Due Date"), 80)
    risk_level = clean_display_text(row.get("Risk Level"), 80)
    risk_reason = clean_display_text(row.get("Risk Reason"), 320)

    sentences: list[str] = []
    if sch_context:
        label = "SCH project update" if source == "SCH Project Update" else "SCH project remark"
        sentences.append(f"As per the latest {label}: {ensure_sentence(sch_context)}")
    else:
        subject = f" for {project_name}" if project_name else ""
        sentences.append(f"No manual PM note is available{subject}; this update is generated from Tracker and SCH status.")

    status_parts = []
    if tracker_status:
        status_parts.append(f"Tracker status is {tracker_status}")
    if sch_status:
        sch_phrase = f"SCH status is {sch_status}"
        if sch_meaning:
            sch_phrase += f" ({sch_meaning})"
        status_parts.append(sch_phrase)
    if status_parts:
        sentences.append(ensure_sentence("; ".join(status_parts)))

    if milestone:
        sentences.append(f"Current milestone: {ensure_sentence(milestone)}")
    if owner:
        sentences.append(f"Current owner/implementation contact: {ensure_sentence(owner)}")
    if next_action:
        action_sentence = f"Next action: {next_action}"
        if due_date:
            action_sentence += f" Target date: {due_date}"
        sentences.append(ensure_sentence(action_sentence))
    if risk_level in {"Critical", "High"} and risk_reason:
        sentences.append(f"Attention required: {ensure_sentence(risk_reason)}")

    update = " ".join(sentence for sentence in sentences if sentence).strip()
    return source, update or "Project is pending review. Please validate Tracker and SCH details before sending this update."


def build_stakeholder_subject(row: pd.Series, config: dict[str, Any]) -> str:
    settings = config.get("project_summaries", {})
    prefix = clean_display_text(settings.get("subject_prefix", "Project Status Update"), 80) or "Project Status Update"
    project_name = clean_display_text(row.get("Project Name"), 140) or "Project"
    milestone = clean_display_text(row.get("Current Milestone"), 100)
    if milestone:
        return f"{prefix} - {project_name} - {milestone}"
    return f"{prefix} - {project_name}"


def clean_display_text(value: Any, max_chars: int | None = None) -> str:
    text = value_to_text(value)
    text = re.sub(r"\s+", " ", text).strip()
    if max_chars and len(text) > max_chars:
        return text[: max_chars - 3].rstrip() + "..."
    return text


def ensure_sentence(value: Any) -> str:
    text = clean_display_text(value)
    if not text:
        return ""
    return text if text[-1] in ".!?" else f"{text}."


def classify_risk(row: pd.Series, config: dict[str, Any]) -> tuple[str, str]:
    thresholds = config.get("risk_thresholds_days", {"watch": 30, "high": 90, "critical": 180})
    watch_days = int(thresholds.get("watch", 30))
    high_days = int(thresholds.get("high", 90))
    critical_days = int(thresholds.get("critical", 180))
    tracker = str(row.get("Tracker Status", "")).strip()
    tracker_l = tracker.lower()
    sch_status = str(row.get("SCH Status", "")).strip()
    sch_l = sch_status.lower()
    age = row.get("Age Days")
    age = int(age) if pd.notna(age) else 0
    notes = str(row.get("PM Notes", "") or "").strip()
    reasons: list[str] = []

    if tracker_l == "held up" or sch_l == "held":
        reasons.append("Held-up status needs management intervention")
        return "Critical", "; ".join(reasons)
    if row.get("Status Mismatch") is True:
        reasons.append(f"Tracker status '{tracker}' conflicts with SCH status '{sch_status}'")
    if row.get("Match Method") == "No Match":
        reasons.append("No approved SCH match")

    if tracker_l in {"active", "yet to start"}:
        if age >= critical_days:
            reasons.append(f"Ageing {age} days from start date")
            return "Critical", "; ".join(reasons)
        if age >= high_days:
            reasons.append(f"Ageing {age} days from start date")
            return "High", "; ".join(reasons)
        if sch_l in {"pending for milestone", "ac", "fc"} and age >= 60:
            reasons.append(f"SCH milestone '{sch_status}' has not progressed for {age} days")
            return "High", "; ".join(reasons)
        if not notes and age >= watch_days:
            reasons.append("PM note missing for ageing open item")
            return "Watch", "; ".join(reasons)
        if age >= watch_days:
            reasons.append(f"Ageing {age} days from start date")
            return "Watch", "; ".join(reasons)

    if reasons:
        return "High" if row.get("Status Mismatch") is True else "Watch", "; ".join(reasons)
    if tracker_l in {"completed", "lost"}:
        return "Closed", "Closed/lost item, monitor only"
    return "Healthy", "No immediate exception detected"


def default_next_action(row: pd.Series) -> str:
    existing = str(row.get("Next Action", "") or "").strip()
    if existing:
        return existing
    notes = str(row.get("PM Notes", "") or "").strip()
    if notes:
        return notes
    milestone = str(row.get("Current Milestone", "") or "").strip()
    tracker = str(row.get("Tracker Status", "")).strip().lower()
    risk = str(row.get("Risk Level", "")).strip()
    if risk in {"Critical", "High"}:
        return "Review blocker, confirm owner, and update next closure plan."
    if tracker == "yet to start":
        return "Confirm kickoff date, data readiness, and owner."
    if tracker == "active":
        return f"Confirm progress for milestone: {milestone}." if milestone else "Confirm milestone progress and next client action."
    if tracker == "held up":
        return "Escalate blocker and confirm revised action plan."
    return ""


def default_current_milestone(row: pd.Series) -> str:
    tracker = str(row.get("Tracker Status", "") or "").strip().lower()
    sch_stage = str(row.get("SCH Stage Group", "") or "").strip().lower()
    suggested = str(row.get("SCH Suggested Milestone", "") or "").strip()
    if tracker == "completed":
        return "Completed"
    if tracker == "yet to start":
        return "Kick Off Meeting"
    if tracker == "held up":
        return suggested or "Datafile Explanation"
    if tracker == "active":
        if sch_stage in {"active complete", "fully complete", "go live"}:
            return suggested or "Live Salary Processing"
        return suggested or "Datafile Explanation"
    return suggested


def default_mail_include(row: pd.Series, config: dict[str, Any]) -> str:
    if (
        str(row.get("Risk Level", "")).strip() == "Closed"
        and str(row.get("Tracker Status", "")).strip().lower() in {"completed", "lost"}
        and row.get("Status Mismatch") is not True
    ):
        return "No"
    existing = str(row.get("Mail Include", "") or "").strip()
    if existing:
        return existing
    risk_levels = set(config.get("mail_include_risk_levels", ["Critical", "High"]))
    statuses = {status.lower() for status in config.get("mail_include_tracker_statuses", ["Active", "Yet To Start", "Held Up"])}
    if row.get("Risk Level") in risk_levels:
        return "Yes"
    if str(row.get("Tracker Status", "")).strip().lower() in statuses:
        return "Yes"
    return "No"


def build_action_list(merged: pd.DataFrame) -> pd.DataFrame:
    mask = (
        merged["Tracker Status"].astype(str).str.lower().isin(["active", "yet to start", "held up"])
        | merged["Risk Level"].isin(["Critical", "High"])
        | merged["Mail Include"].astype(str).str.lower().eq("yes")
    )
    columns = [
        "Risk Level",
        "Project Name",
        "Tracker Status",
        "SCH Status",
        "SCH Status Meaning",
        "SCH Stage Group",
        "SCH Milestone",
        "Current Milestone",
        "Age Days",
        "Owner",
        "Next Action",
        "Due Date",
        "PM Notes",
        "PM Note Source",
        "Stakeholder PM Update",
        "Risk Reason",
    ]
    action = merged.loc[mask, [column for column in columns if column in merged.columns]].copy()
    risk_order = {"Critical": 1, "High": 2, "Watch": 3, "Healthy": 4, "Closed": 5}
    action["_risk_order"] = action["Risk Level"].map(risk_order).fillna(9)
    return action.sort_values(["_risk_order", "Age Days"], ascending=[True, False]).drop(columns=["_risk_order"])


def build_weekly_mail_view(merged: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    mask = merged["Mail Include"].astype(str).str.lower().eq("yes")
    columns = [
        "Project Name",
        "Tracker Status",
        "SCH Status",
        "SCH Milestone",
        "Current Milestone",
        "Age Days",
        "Risk Level",
        "Owner",
        "Next Action",
        "PM Notes",
        "PM Note Source",
        "Stakeholder PM Update",
    ]
    weekly = merged.loc[mask, [column for column in columns if column in merged.columns]].copy()
    risk_order = {"Critical": 1, "High": 2, "Watch": 3, "Healthy": 4, "Closed": 5}
    weekly["_risk_order"] = weekly["Risk Level"].map(risk_order).fillna(9)
    return weekly.sort_values(["_risk_order", "Age Days"], ascending=[True, False]).drop(columns=["_risk_order"])


def build_project_summary_view(merged: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    statuses = {
        status.lower()
        for status in config.get("project_summaries", {}).get("current_statuses", ["Active", "Yet To Start", "Held Up"])
    }
    current = merged[merged["Tracker Status"].astype(str).str.lower().isin(statuses)].copy()
    columns = [
        "Project Code",
        "Project Name",
        "Project Type",
        "Tracker Status",
        "SCH Status",
        "SCH Status Meaning",
        "SCH Stage Group",
        "SCH Milestone",
        "Current Milestone",
        "SCH Suggested Milestone",
        "Age Days",
        "Risk Level",
        "Risk Reason",
        "Owner",
        "Effective Employee Count",
        "Effective Product",
        "Effective Client Type",
        "PM Note Source",
        "Stakeholder PM Update",
        "Stakeholder Update Subject",
        "Next Action",
        "Due Date",
        "Action Date",
        "PM Notes",
        "SCH Project Remark",
        "SCH Project Update",
        "SCH Achieved Percent",
        "Match Method",
        "Match Review Needed",
        "Tracker Name Key",
    ]
    current = current[[column for column in columns if column in current.columns]]
    risk_order = {"Critical": 1, "High": 2, "Watch": 3, "Healthy": 4, "Closed": 5}
    current["_risk_order"] = current["Risk Level"].map(risk_order).fillna(9)
    return current.sort_values(["_risk_order", "Age Days", "Project Name"], ascending=[True, False, True]).drop(columns=["_risk_order"])


def build_execution_master(merged: pd.DataFrame) -> pd.DataFrame:
    columns = [
        "Project Code",
        "Project Name",
        "Project Type",
        "Tracker Status",
        "SCH Client Name",
        "SCH Status",
        "SCH Status Meaning",
        "Current Milestone",
        "Owner",
        "Risk Level",
        "Risk Reason",
        "Next Action",
        "PM Notes",
        "Stakeholder PM Update",
        "PM Note Source",
        "Match Method",
        "Match Review Needed",
        "Start Date",
        "End Date",
        "Age Days",
        "Effective Employee Count",
        "Effective Product",
        "Effective Client Type",
        "SCH Implementation Engg",
        "SCH Project Remark",
        "SCH Project Update",
    ]
    return merged[[column for column in columns if column in merged.columns]].copy()


def build_alias_matching_tracker(
    tracker: pd.DataFrame,
    sch: pd.DataFrame,
    alias_map: pd.DataFrame,
    merged: pd.DataFrame,
    config: dict[str, Any],
) -> pd.DataFrame:
    sch_names = sch[["Client Name", "SCH Name Key"]].dropna().drop_duplicates()
    alias_by_tracker = {
        row["Tracker Name Key"]: row["SCH Client Name"]
        for _, row in alias_map[alias_map["Approved"].eq(True)].iterrows()
        if row.get("Tracker Name Key") and row.get("SCH Client Name")
    }
    merged_by_tracker = merged.set_index("Tracker Name Key", drop=False) if "Tracker Name Key" in merged.columns else pd.DataFrame()
    rows: list[dict[str, Any]] = []
    for _, tracker_row in tracker.iterrows():
        tracker_name = value_to_text(tracker_row.get("Project Name"))
        tracker_key = value_to_text(tracker_row.get("Tracker Name Key"))
        match_row = merged_by_tracker.loc[tracker_key] if not merged_by_tracker.empty and tracker_key in merged_by_tracker.index else {}
        if isinstance(match_row, pd.DataFrame):
            match_row = match_row.iloc[0]
        suggestions = fuzzy_candidates(tracker_name, sch_names, config)
        current_match = value_to_text(get_row_value(match_row, "SCH Client Name"))
        approved_final = alias_by_tracker.get(tracker_key, "")
        final_name = approved_final or current_match
        rows.append(
            {
                "Tracker Project Code": tracker_row.get("Project Code"),
                "Tracker Project Name": tracker_name,
                "Tracker Status": tracker_row.get("Status"),
                "Current Matched SCH Client": current_match,
                "Suggested SCH Match 1": suggestions[0][0] if len(suggestions) > 0 else "",
                "Score 1": suggestions[0][1] if len(suggestions) > 0 else "",
                "Suggested SCH Match 2": suggestions[1][0] if len(suggestions) > 1 else "",
                "Score 2": suggestions[1][1] if len(suggestions) > 1 else "",
                "Suggested SCH Match 3": suggestions[2][0] if len(suggestions) > 2 else "",
                "Score 3": suggestions[2][1] if len(suggestions) > 2 else "",
                "Final SCH Client Name": final_name,
                "Approved": "Yes" if final_name else "No",
                "Match Method": get_row_value(match_row, "Match Method"),
                "Review Notes": "Update Final SCH Client Name and Approved=Yes if the suggested match is correct.",
            }
        )
    return pd.DataFrame(rows)


def build_alias_matching_sch(
    tracker: pd.DataFrame,
    sch: pd.DataFrame,
    alias_map: pd.DataFrame,
    merged: pd.DataFrame,
    config: dict[str, Any],
) -> pd.DataFrame:
    tracker_names = tracker[["Project Name", "Tracker Name Key"]].dropna().drop_duplicates()
    alias_by_sch = {
        row["SCH Name Key"]: row["Tracker Project Name"]
        for _, row in alias_map[alias_map["Approved"].eq(True)].iterrows()
        if row.get("SCH Name Key") and row.get("Tracker Project Name")
    }
    matched_tracker_by_sch = (
        merged.groupby("SCH Name Key")["Project Name"].apply(lambda values: ", ".join(sorted(set(value_to_text(v) for v in values if value_to_text(v)))))
        if "SCH Name Key" in merged.columns
        else pd.Series(dtype="object")
    )
    rows: list[dict[str, Any]] = []
    for _, sch_row in sch.iterrows():
        sch_name = value_to_text(sch_row.get("Client Name"))
        sch_key = value_to_text(sch_row.get("SCH Name Key"))
        suggestions = tracker_fuzzy_candidates(sch_name, tracker_names, config)
        current_match = value_to_text(matched_tracker_by_sch.get(sch_key, ""))
        approved_final = alias_by_sch.get(sch_key, "")
        final_name = approved_final or current_match
        rows.append(
            {
                "SCH Client Name": sch_name,
                "SCH Status": sch_row.get("Status"),
                "Current Matched Tracker Project": current_match,
                "Suggested Tracker Match 1": suggestions[0][0] if len(suggestions) > 0 else "",
                "Score 1": suggestions[0][1] if len(suggestions) > 0 else "",
                "Suggested Tracker Match 2": suggestions[1][0] if len(suggestions) > 1 else "",
                "Score 2": suggestions[1][1] if len(suggestions) > 1 else "",
                "Suggested Tracker Match 3": suggestions[2][0] if len(suggestions) > 2 else "",
                "Score 3": suggestions[2][1] if len(suggestions) > 2 else "",
                "Final Tracker Project Name": final_name,
                "Approved": "Yes" if final_name else "No",
                "Review Notes": "Use this reverse view when SCH has a client name that needs to be mapped back to Tracker.",
            }
        )
    return pd.DataFrame(rows)


def tracker_fuzzy_candidates(sch_name: Any, tracker_names: pd.DataFrame, config: dict[str, Any]) -> list[tuple[str, float]]:
    limit = int(config.get("fuzzy_candidate_limit", 3))
    min_score = float(config.get("fuzzy_candidate_min_score", 0.58))
    key = normalize_name(sch_name)
    candidates: list[tuple[str, float]] = []
    if not key:
        return candidates
    for _, row in tracker_names.iterrows():
        tracker_key = row["Tracker Name Key"]
        if not tracker_key:
            continue
        score = SequenceMatcher(None, key, tracker_key).ratio()
        if score >= min_score:
            candidates.append((row["Project Name"], round(score, 3)))
    return sorted(candidates, key=lambda item: item[1], reverse=True)[:limit]


def get_row_value(row: Any, column: str) -> Any:
    if isinstance(row, pd.Series):
        return row.get(column, "")
    if isinstance(row, dict):
        return row.get(column, "")
    return ""


def build_completed_followup_plan(merged: pd.DataFrame, config: dict[str, Any]) -> pd.DataFrame:
    completed = merged[merged["Tracker Status"].astype(str).str.lower().eq("completed")].copy()
    if completed.empty:
        return pd.DataFrame(
            columns=[
                "Followup Due This Month",
                "Followup Quarter",
                "Followup Month",
                "Suggested Followup Date",
                "Project Code",
                "Project Name",
                "Completed Date",
                "SCH Status",
                "Owner",
                "Last Processed Year Month",
                "Client Usage Status",
                "Challenges / Pain Points",
                "Support Required",
                "Upsell Scope",
                "PM Followup Notes",
                "Next Followup Date",
            ]
        )
    today = config_run_timestamp(config)
    quarter = pd.Period(today, freq="Q")
    quarter_months = list(range(quarter.start_time.month, quarter.end_time.month + 1))
    completed["_followup_slot"] = completed["Project Code"].apply(stable_followup_slot)
    rows: list[dict[str, Any]] = []
    for _, row in completed.sort_values(["_followup_slot", "Project Name"]).iterrows():
        slot = int(row["_followup_slot"])
        followup_month_number = quarter_months[slot]
        followup_date = pd.Timestamp(year=today.year, month=followup_month_number, day=15)
        rows.append(
            {
                "Followup Due This Month": "Yes" if followup_month_number == today.month else "No",
                "Followup Quarter": f"Q{quarter.quarter} {today.year}",
                "Followup Month": followup_date.strftime("%B"),
                "Suggested Followup Date": followup_date,
                "Project Code": row.get("Project Code"),
                "Project Name": row.get("Project Name"),
                "Completed Date": row.get("End Date"),
                "SCH Status": row.get("SCH Status"),
                "Owner": row.get("Owner"),
                "Last Processed Year Month": row.get("SCH Last Processed Year Month"),
                "Client Usage Status": "",
                "Challenges / Pain Points": "",
                "Support Required": "",
                "Upsell Scope": "",
                "PM Followup Notes": "Check if client is using the system properly, whether any challenges are open, and if optimization/support is needed.",
                "Next Followup Date": "",
            }
        )
    return pd.DataFrame(rows)


def stable_followup_slot(value: Any) -> int:
    text = value_to_text(value)
    digits = re.sub(r"\D+", "", text)
    if digits:
        return int(digits) % 3
    return sum(ord(char) for char in text) % 3 if text else 0


def create_project_summaries(
    data: dict[str, pd.DataFrame],
    output_dir: Path,
    index_path: Path,
    config: dict[str, Any],
    skip_outlook: bool = False,
) -> pd.DataFrame:
    project_view = data.get("Project_Summary_View", pd.DataFrame()).copy()
    output_dir.mkdir(parents=True, exist_ok=True)
    group_map = load_project_group_map(config, project_view)
    group_by_key = {
        str(row["Tracker Name Key"]): row
        for _, row in group_map.iterrows()
        if str(row.get("Tracker Name Key", "")).strip()
    }
    settings = config.get("project_summaries", {})
    create_drafts = bool(settings.get("create_outlook_drafts", False)) and not skip_outlook

    index_rows: list[dict[str, Any]] = []
    for _, row in project_view.iterrows():
        project_code = value_to_text(row.get("Project Code"))
        project_name = value_to_text(row.get("Project Name"))
        filename = f"{project_code}_{safe_filename(project_name)}.html" if project_code else f"{safe_filename(project_name)}.html"
        summary_path = output_dir / filename
        map_row = group_by_key.get(str(row.get("Tracker Name Key", "")), {})
        to_value = value_to_text(get_map_value(map_row, "Outlook Group / To"))
        cc_value = value_to_text(get_map_value(map_row, "CC"))
        enabled = value_to_text(get_map_value(map_row, "Enabled")).lower() in {"yes", "true", "1"}
        body = build_project_summary_html(row, config)
        summary_path.write_text(body, encoding="utf-8")
        subject = value_to_text(row.get("Stakeholder Update Subject")) or build_stakeholder_subject(row, config)

        draft_status = "HTML only"
        if create_drafts:
            if enabled and to_value:
                draft_status = create_project_outlook_draft(row, body, to_value, cc_value, settings)
            else:
                draft_status = "No Outlook group mapped"

        index_rows.append(
            {
                "Project Code": project_code,
                "Project Name": project_name,
                "Tracker Status": value_to_text(row.get("Tracker Status")),
                "SCH Status": value_to_text(row.get("SCH Status")),
                "Risk Level": value_to_text(row.get("Risk Level")),
                "Owner": value_to_text(row.get("Owner")),
                "PM Note Source": value_to_text(row.get("PM Note Source")),
                "Stakeholder Update Subject": subject,
                "HTML Path": str(summary_path),
                "Outlook Group / To": to_value,
                "CC": cc_value,
                "Draft Status": draft_status,
            }
        )

    index = pd.DataFrame(index_rows)
    index_path.parent.mkdir(parents=True, exist_ok=True)
    index.to_csv(index_path, index=False)
    logging.info("Project summaries created: %s", len(index))
    return index


def load_project_group_map(config: dict[str, Any], project_view: pd.DataFrame) -> pd.DataFrame:
    settings = config.get("project_summaries", {})
    path = Path(settings.get("group_map_file", Path(config["output_root"]) / "config" / "project_outlook_group_map.csv"))
    path.parent.mkdir(parents=True, exist_ok=True)
    columns = ["Project Name", "Tracker Name Key", "Outlook Group / To", "CC", "Enabled", "Notes"]
    if path.exists():
        group_map = pd.read_csv(path)
        for column in columns:
            if column not in group_map.columns:
                group_map[column] = ""
    else:
        group_map = pd.DataFrame(columns=columns)

    existing = set(group_map["Tracker Name Key"].dropna().astype(str))
    additions = []
    for _, row in project_view.iterrows():
        key = str(row.get("Tracker Name Key", ""))
        if key and key not in existing:
            additions.append(
                {
                    "Project Name": value_to_text(row.get("Project Name")),
                    "Tracker Name Key": key,
                    "Outlook Group / To": "",
                    "CC": "",
                    "Enabled": "No",
                    "Notes": "Add Outlook group/email here when ready.",
                }
            )
            existing.add(key)
    if additions:
        group_map = pd.concat([group_map, pd.DataFrame(additions)], ignore_index=True)
    group_map = group_map[columns]
    group_map.to_csv(path, index=False)
    return group_map


def get_map_value(map_row: Any, column: str) -> Any:
    if isinstance(map_row, pd.Series):
        return map_row.get(column, "")
    if isinstance(map_row, dict):
        return map_row.get(column, "")
    return ""


def build_project_summary_html(row: pd.Series, config: dict[str, Any]) -> str:
    theme = resolved_theme(config)
    project_name = value_to_text(row.get("Project Name"))
    tracker_status = value_to_text(row.get("Tracker Status"))
    sch_status = value_to_text(row.get("SCH Status"))
    risk_level = value_to_text(row.get("Risk Level"))
    pm_update = clean_display_text(row.get("Stakeholder PM Update"), 2200)
    if not pm_update:
        _, pm_update = build_stakeholder_pm_update(row, config)
    subject = clean_display_text(row.get("Stakeholder Update Subject"), 220) or build_stakeholder_subject(row, config)
    note_source = clean_display_text(row.get("PM Note Source"), 120)
    milestone = clean_display_text(row.get("Current Milestone"), 160)
    next_action = clean_display_text(row.get("Next Action"), 600) or "Confirm latest milestone progress and update next closure plan."
    due_date = clean_display_text(row.get("Due Date"), 80)
    action_date = clean_display_text(row.get("Action Date"), 80)
    risk_color = theme["risk_colors"].get(risk_level, theme["primary"])
    risk_text_color = "#1f1f1f" if risk_level == "Watch" else "#ffffff"
    status_text = sch_status
    meaning = clean_display_text(row.get("SCH Status Meaning"), 260)
    if meaning:
        status_text = f"{sch_status} - {meaning}" if sch_status else meaning

    snapshot_fields = [
        ("Project Code", row.get("Project Code")),
        ("Project Type", row.get("Project Type")),
        ("Tracker Status", tracker_status),
        ("SCH Status", status_text),
        ("Current Milestone", milestone),
        ("Owner / Engineer", row.get("Owner")),
        ("Age Days", row.get("Age Days")),
        ("Employee Count", row.get("Effective Employee Count")),
        ("Product", row.get("Effective Product")),
        ("Client Type", row.get("Effective Client Type")),
        ("Risk Level", risk_level),
        ("Risk Reason", row.get("Risk Reason")),
    ]
    sch_context_fields = [
        ("SCH Project Update", clean_display_text(row.get("SCH Project Update"), 1200)),
        ("SCH Project Remark", clean_display_text(row.get("SCH Project Remark"), 1200)),
        ("SCH Achieved Percent", row.get("SCH Achieved Percent")),
        ("SCH Suggested Milestone", row.get("SCH Suggested Milestone")),
        ("Match Method", row.get("Match Method")),
    ]
    label_bg = theme["status_colors"].get("Completed", "#D9EAF7")

    def table_rows(fields: list[tuple[str, Any]]) -> str:
        rows = []
        for label, value in fields:
            text = value_to_text(value)
            if not text:
                text = "-"
            rows.append(
                f"<tr><td style='font-weight:bold;background:{label_bg};width:190px;border:1px solid #B7C9D9;padding:7px;'>"
                f"{html.escape(label)}</td><td style='border:1px solid #B7C9D9;padding:7px;'>"
                f"{html.escape(text)}</td></tr>"
            )
        return "\n".join(rows)

    target_text = []
    if due_date:
        target_text.append(f"Due Date: {due_date}")
    if action_date:
        target_text.append(f"Action Date: {action_date}")
    target_line = f"<p style='margin:6px 0 0;color:#666;font-size:9pt;'>{html.escape(' | '.join(target_text))}</p>" if target_text else ""

    return f"""
<html>
<body style="font-family:Arial, sans-serif;color:{theme['text']};background:{theme['background']};margin:0;padding:18px;">
  <div style="max-width:980px;background:#FFFFFF;border:1px solid {theme['border']};padding:22px;margin:0 auto;">
    <p style="margin:0 0 12px;color:#666;font-size:9pt;"><b>Subject:</b> {html.escape(subject)}</p>
    <h2 style="color:{theme['primary']};margin:0 0 4px;">{html.escape(project_name)}</h2>
    <p style="margin:0 0 14px;color:#666;">Project-wise PM update generated on {datetime.now():%d-%b-%Y %H:%M}</p>

    <div style="padding:10px 12px;background:{risk_color};color:{risk_text_color};font-weight:bold;display:inline-block;margin-bottom:14px;">
      {html.escape(risk_level or 'Status Review')}
    </div>

    <p style="margin:8px 0 14px;">Dear Team,</p>
    <p style="margin:0 0 14px;">Please find the latest PM update for the below project.</p>

    <h3 style="color:{theme['primary']};margin:18px 0 8px;">Latest PM Update</h3>
    <div style="border-left:5px solid {theme['primary']};background:#EEF4FA;padding:12px 14px;line-height:1.45;">
      {html.escape(pm_update)}
      <p style="margin:8px 0 0;color:#666;font-size:9pt;">PM note source: {html.escape(note_source or 'System generated')}</p>
    </div>

    <h3 style="color:{theme['primary']};margin:18px 0 8px;">Next Action</h3>
    <div style="background:#FFF8E5;border:1px solid #E7C65A;padding:11px 13px;line-height:1.45;">
      <b>{html.escape(next_action)}</b>
      {target_line}
    </div>

    <h3 style="color:{theme['primary']};margin:18px 0 8px;">Project Snapshot</h3>
    <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:10pt;">
      {table_rows(snapshot_fields)}
    </table>

    <h3 style="color:{theme['primary']};margin:18px 0 8px;">SCH Context Used</h3>
    <table cellpadding="0" cellspacing="0" style="border-collapse:collapse;width:100%;font-size:10pt;">
      {table_rows(sch_context_fields)}
    </table>

    <p style="margin:18px 0 4px;">Regards,<br>{html.escape(config.get('owner_name', 'Vasu'))}</p>
    <p style="color:#666;font-size:9pt;margin-top:12px;">
      Tracker is the primary PM source. SCH Project Master is used for latest status, implementation remark, project update, and operational context.
    </p>
  </div>
</body>
</html>
"""


def create_project_outlook_draft(row: pd.Series, body: str, to_value: str, cc_value: str, settings: dict[str, Any]) -> str:
    try:
        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_value
        mail.CC = cc_value
        mail.Subject = value_to_text(row.get("Stakeholder Update Subject")) or f"{settings.get('subject_prefix', 'Project Status Update')} - {value_to_text(row.get('Project Name'))}"
        mail.HTMLBody = body
        mail.Save()
        if settings.get("display_drafts", False):
            mail.Display(False)
        return "Outlook draft saved"
    except Exception as exc:
        logging.warning("Project Outlook draft failed for %s: %s", row.get("Project Name"), exc)
        return f"Outlook draft failed ({exc})"


def safe_filename(value: Any) -> str:
    text = re.sub(r"[^A-Za-z0-9._ -]+", "", value_to_text(value))
    text = re.sub(r"\s+", "_", text).strip("._ ")
    return text[:120] or "project"


def value_to_text(value: Any) -> str:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.strftime("%d-%b-%Y")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def build_monthly_summary(merged: pd.DataFrame) -> pd.DataFrame:
    df = merged.copy()
    df["Start Month"] = pd.to_datetime(df["Start Date"], errors="coerce").dt.to_period("M").astype("string")
    df["End Month"] = pd.to_datetime(df["End Date"], errors="coerce").dt.to_period("M").astype("string")
    months = sorted(set(df["Start Month"].dropna()) | set(df["End Month"].dropna()))
    opening = 0
    rows = []
    for month in months:
        starts = df[df["Start Month"].eq(month)]
        ends = df[df["End Month"].eq(month)]
        add_alloc = int(starts["Project Type"].eq("Allocation").sum())
        add_escal = int(starts["Project Type"].eq("Escalation").sum())
        completed = int(ends["Tracker Status"].eq("Completed").sum())
        lost = int(ends["Tracker Status"].eq("Lost").sum())
        held = int(ends["Tracker Status"].eq("Held Up").sum())
        total_add = add_alloc + add_escal
        net = total_add - completed - lost - held
        closing = opening + net
        rows.append(
            {
                "Month": month,
                "Opening": opening,
                "Add Alloc": add_alloc,
                "Add Escal": add_escal,
                "Total Add": total_add,
                "Completed": completed,
                "Lost": lost,
                "Held": held,
                "Net Change": net,
                "Closing": closing,
            }
        )
        opening = closing
    return pd.DataFrame(rows)


def build_validation_log(
    tracker: pd.DataFrame,
    sch: pd.DataFrame,
    merged: pd.DataFrame,
    unmatched: pd.DataFrame,
    previous_notes: pd.DataFrame,
) -> pd.DataFrame:
    rows = [
        ("Run Timestamp", datetime.now().strftime("%d-%b-%Y %H:%M:%S")),
        ("Tracker Rows Used", len(tracker)),
        ("SCH Rows Used", len(sch)),
        ("Merged Cockpit Rows", len(merged)),
        ("Unmatched Tracker Rows", len(unmatched)),
        ("Rows Needing Match Review", int(merged["Match Review Needed"].sum())),
        ("Preserved Note Keys", len(previous_notes)),
        ("Critical Risks", int(merged["Risk Level"].eq("Critical").sum())),
        ("High Risks", int(merged["Risk Level"].eq("High").sum())),
        ("Open Pipeline", int(merged["Tracker Status"].isin(["Active", "Yet To Start", "Held Up"]).sum())),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def build_pm_playbook(config: dict[str, Any]) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    rows.extend(
        [
            {
                "Section": "Source Of Truth",
                "Item": "Tracker",
                "Description": "VASU's Tracker is the primary PM file for project list, allocation/escalation type, start/end date, PM-owned status, owner context, and notes.",
            },
            {
                "Section": "Source Of Truth",
                "Item": "SCH Project Master",
                "Description": "SCH Project Master enriches missing details, official project remarks, implementation engineer, client information, and SCH status.",
            },
            {
                "Section": "Project Type",
                "Item": "Allocation",
                "Description": "PM is involved from the start: kickoff, engineer assignment, datafile explanation, setup, training, integration, salary tally, handholding, then go-live.",
            },
            {
                "Section": "Project Type",
                "Item": "Escalation",
                "Description": "PM enters midway or when stuck: understand current state, coordinate internally and with client/partner, resolve, track feedback for a month, then close.",
            },
        ]
    )
    for index, milestone in enumerate(config.get("milestone_sequence", []), start=1):
        rows.append({"Section": "Milestone Sequence", "Item": f"{index:02d}", "Description": milestone})
    for status, definition in config.get("sch_status_definitions", {}).items():
        rows.append(
            {
                "Section": "SCH Status Meaning",
                "Item": status,
                "Description": f"{definition.get('stage', '')}: {definition.get('meaning', '')} Suggested milestone: {definition.get('suggested_milestone', '')}",
            }
        )
    rows.extend(
        [
            {
                "Section": "Best PM Routine",
                "Item": "Daily",
                "Description": "Review Action_List for Critical/High items, update Current Milestone, PM Notes, Next Action, Owner, and Due Date.",
            },
            {
                "Section": "Best PM Routine",
                "Item": "Weekly",
                "Description": "Refresh tracker and SCH export, run the cockpit, review project summaries, and send project-wise updates to mapped stakeholders.",
            },
            {
                "Section": "Best PM Routine",
                "Item": "Monthly",
                "Description": "Generate the monthly dashboard workbook and executive PPT, then review trends, risks, completions, and next-month focus.",
            },
        ]
    )
    return pd.DataFrame(rows)


def build_alias_mapping_guide() -> pd.DataFrame:
    rows = [
        ("Purpose", "Why mapping exists", "Tracker project codes and SCH project codes are different. Matching is done by project/client name."),
        ("Purpose", "What alias means", "An alias links one tracker Project Name to one SCH Client Name when the names differ but represent the same client/project."),
        ("Step 1", "Open Unmatched_Clients", "Review rows where Match Method is No Match. Start with current Active, Yet To Start, and Held Up projects."),
        ("Step 2", "Check candidates", "Compare Best SCH Candidate 1/2/3 with the tracker project name and SCH Project Remark/context."),
        ("Step 3", "Open client_alias_map.csv", "File path: D:\\Reports\\PM Cockpit\\config\\client_alias_map.csv."),
        ("Step 4", "Add or edit row", "Fill Tracker Project Name, SCH Client Name, Approved=True, Match Type=Manual Alias, Last Reviewed=today, Notes=reason."),
        ("Step 5", "Use exact names", "Copy names exactly from Unmatched_Clients and SCH Project Master to avoid spelling mistakes."),
        ("Step 6", "Rerun cockpit", "Run run_pm_cockpit.bat. The project should move from Unmatched_Clients into PM_Cockpit with Match Method Approved Alias."),
        ("Review Existing", "Open Alias_Map", "Review existing approved mappings. If a mapping is wrong, correct SCH Client Name or set Approved=False in client_alias_map.csv."),
        ("Review Existing", "Open Match_Review", "Rows with multiple SCH candidates need human confirmation. Add a manual alias for the correct SCH client."),
        ("Review Existing", "Open Status_Mismatch", "If tracker says Completed/Lost but SCH says Active/AC/etc., verify whether tracker or SCH needs updating."),
        ("Rule", "Do not force uncertain matches", "If not sure, leave unmapped. A wrong alias is worse than a missing alias because it can send the wrong project remark."),
    ]
    return pd.DataFrame(rows, columns=["Section", "Action", "Details"])


def build_theme_config(config: dict[str, Any]) -> pd.DataFrame:
    theme = resolved_theme(config)
    rows = []
    for key in ["primary", "secondary", "background", "panel", "text", "muted", "border"]:
        rows.append({"Section": "Base", "Item": key, "Color": theme.get(key, "")})
    for status, color in theme.get("status_colors", {}).items():
        rows.append({"Section": "Status", "Item": status, "Color": color})
    for risk, color in theme.get("risk_colors", {}).items():
        rows.append({"Section": "Risk", "Item": risk, "Color": color})
    return pd.DataFrame(rows)


def create_alias_matching_review_workbook(data: dict[str, pd.DataFrame], output_path: Path, config: dict[str, Any]) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy") as writer:
        workbook = writer.book
        formats = workbook_formats(workbook, config)
        write_dataframe_sheet(writer, "Tracker_to_SCH", data.get("Alias_Matching_Tracker", pd.DataFrame()), formats)
        write_dataframe_sheet(writer, "SCH_to_Tracker", data.get("Alias_Matching_SCH", pd.DataFrame()), formats)
        guide = pd.DataFrame(
            [
                ("How to use", "Tracker_to_SCH", "Review tracker projects, choose/correct Final SCH Client Name, and set Approved=Yes."),
                ("How to use", "SCH_to_Tracker", "Use the reverse SCH view when SCH client names need mapping back to tracker names."),
                ("Next run", "Automation", "Approved final names in this workbook are read before the next merge and are added to client_alias_map.csv."),
                ("Rule", "Safety", "Only approve matches you are confident about. A blank match is safer than a wrong match."),
            ],
            columns=["Section", "Sheet", "Instruction"],
        )
        write_dataframe_sheet(writer, "Guide", guide, formats)
    logging.info("Alias matching review workbook created: %s", output_path)


def create_pm_cockpit_workbook(
    data: dict[str, pd.DataFrame],
    output_path: Path,
    latest_path: Path,
    config: dict[str, Any],
) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    latest_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="dd-mmm-yyyy", date_format="dd-mmm-yyyy") as writer:
        workbook = writer.book
        formats = workbook_formats(workbook, config)
        write_dashboard(writer, data, config, formats)
        sheet_order = [
            "Execution_Master",
            "PM_Cockpit",
            "Action_List",
            "Completed_Followup_Plan",
            "Stale_Items",
            "Status_Mismatch",
            "Alias_Matching_Tracker",
            "Alias_Matching_SCH",
            "Unmatched_Clients",
            "Match_Review",
            "Alias_Map",
            "Theme_Config",
            "Weekly_Mail_View",
            "Project_Summary_View",
            "Monthly_Summary",
            "PM_Playbook",
            "Alias_Mapping_Guide",
            "TrackerData",
            "SCH_Project_Master",
            "Preserved_Notes",
            "Run_Log",
        ]
        for sheet in sheet_order:
            write_dataframe_sheet(writer, sheet, data.get(sheet, pd.DataFrame()), formats)
        add_dashboard_charts(writer, data, formats)

    latest_path.write_bytes(output_path.read_bytes())
    logging.info("PM Cockpit created: %s", output_path)


def workbook_formats(workbook: Any, config: dict[str, Any] | None = None) -> dict[str, Any]:
    theme = resolved_theme(config or {})
    primary = theme["primary"]
    secondary = theme["secondary"]
    status_colors = theme["status_colors"]
    risk_colors = theme["risk_colors"]
    formats = {
        "title": workbook.add_format({"bold": True, "font_size": 18, "font_color": primary}),
        "subtitle": workbook.add_format({"font_size": 10, "font_color": theme["muted"]}),
        "header": workbook.add_format(
            {"bold": True, "font_color": "white", "bg_color": primary, "border": 1, "align": "center", "valign": "vcenter"}
        ),
        "section": workbook.add_format({"bold": True, "font_color": "white", "bg_color": secondary, "border": 1}),
        "label": workbook.add_format({"bold": True, "bg_color": status_colors.get("Completed", "#D9EAF7"), "border": 1}),
        "text": workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"}),
        "date": workbook.add_format({"border": 1, "num_format": "dd-mmm-yyyy"}),
        "integer": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "number": workbook.add_format({"border": 1, "num_format": "#,##0.00"}),
        "percent": workbook.add_format({"border": 1, "num_format": "0.0%"}),
        "card": workbook.add_format(
            {"bold": True, "font_size": 18, "font_color": primary, "bg_color": theme["panel"], "border": 1, "align": "center", "valign": "vcenter"}
        ),
        "card_title": workbook.add_format({"bold": True, "font_color": "white", "bg_color": primary, "border": 1, "align": "center"}),
        "critical": workbook.add_format({"bg_color": risk_colors.get("Critical", "#C00000"), "font_color": "#FFFFFF"}),
        "high": workbook.add_format({"bg_color": risk_colors.get("High", "#ED7D31"), "font_color": "#1F1F1F"}),
        "watch": workbook.add_format({"bg_color": risk_colors.get("Watch", "#FFC000"), "font_color": "#1F1F1F"}),
        "healthy": workbook.add_format({"bg_color": risk_colors.get("Healthy", "#70AD47"), "font_color": "#FFFFFF"}),
    }
    status_formats = {}
    for status, color in status_colors.items():
        font_color = "#1F1F1F" if status in {"Completed", "Yet To Start", "Held Up"} else "#FFFFFF"
        status_formats[status] = workbook.add_format({"bg_color": color, "font_color": font_color})
    formats["status_formats"] = status_formats
    formats["theme"] = theme
    return formats


def resolved_theme(config: dict[str, Any]) -> dict[str, Any]:
    default = {
        "primary": "#1F4E78",
        "secondary": "#305496",
        "background": "#F7F9FC",
        "panel": "#F3F6FA",
        "text": "#1F1F1F",
        "muted": "#606060",
        "border": "#D7DEE8",
        "status_colors": {
            "Active": "#70AD47",
            "Completed": "#D9EAF7",
            "Yet To Start": "#FFC000",
            "Held Up": "#ED7D31",
            "Lost": "#C00000",
        },
        "risk_colors": {
            "Critical": "#C00000",
            "High": "#ED7D31",
            "Watch": "#FFC000",
            "Healthy": "#70AD47",
            "Closed": "#D9EAF7",
        },
    }
    incoming = config.get("theme", {}) if isinstance(config, dict) else {}
    merged = {**default, **{key: value for key, value in incoming.items() if key not in {"status_colors", "risk_colors"}}}
    merged["status_colors"] = {**default["status_colors"], **incoming.get("status_colors", {})}
    merged["risk_colors"] = {**default["risk_colors"], **incoming.get("risk_colors", {})}
    return merged


def write_dashboard(writer: pd.ExcelWriter, data: dict[str, pd.DataFrame], config: dict[str, Any], formats: dict[str, Any]) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Dashboard")
    writer.sheets["Dashboard"] = worksheet
    cockpit = data["PM_Cockpit"]
    run_log = data["Run_Log"]
    summary = build_run_summary(data)
    worksheet.hide_gridlines(2)
    worksheet.set_column("A:A", 2)
    worksheet.set_column("B:M", 14)
    worksheet.merge_range("B1:M1", "PM Cockpit Dashboard", formats["title"])
    worksheet.write("B2", f"{config.get('company_name', '')} | Generated {datetime.now():%d-%b-%Y %H:%M}", formats["subtitle"])

    cards = [
        ("B4:C4", "B5:C6", "Total Projects", summary["total_projects"]),
        ("D4:E4", "D5:E6", "Open Pipeline", summary["open_pipeline"]),
        ("F4:G4", "F5:G6", "Critical Risks", summary["critical_risks"]),
        ("H4:I4", "H5:I6", "High Risks", summary["high_risks"]),
        ("J4:K4", "J5:K6", "Unmatched", summary["unmatched"]),
        ("L4:M4", "L5:M6", "Project Summaries", summary["project_summaries_due"]),
    ]
    for title_range, value_range, title, value in cards:
        worksheet.merge_range(title_range, title, formats["card_title"])
        worksheet.merge_range(value_range, value, formats["card"])

    worksheet.write("B8", "Recommended Focus", formats["section"])
    focus = [
        f"Critical risk items: {summary['critical_risks']}",
        f"High risk items: {summary['high_risks']}",
        f"Open pipeline: {summary['open_pipeline']}",
        f"Unmatched tracker rows: {summary['unmatched']}",
        "Review Action_List first, then update PM Notes / Next Action / Owner.",
    ]
    for idx, item in enumerate(focus, start=9):
        worksheet.write(idx - 1, 1, item, formats["text"])

    setup_table_block(worksheet, run_log, 16, 1, formats, "Run Health")
    worksheet.freeze_panes(8, 0)


def write_dataframe_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    safe_name = sheet_name[:31]
    df_to_write = prepare_excel_df(df)
    df_to_write.to_excel(writer, sheet_name=safe_name, index=False)
    worksheet = writer.sheets[safe_name]
    worksheet.hide_gridlines(2)
    worksheet.freeze_panes(1, 0)
    if df_to_write.empty:
        worksheet.write(0, 0, "No records", formats["header"])
        return
    for col_idx, column in enumerate(df_to_write.columns):
        worksheet.write(0, col_idx, column, formats["header"])
    worksheet.autofilter(0, 0, len(df_to_write), len(df_to_write.columns) - 1)
    autosize_columns(worksheet, df_to_write)
    apply_formats(worksheet, df_to_write, formats)
    apply_risk_conditioning(worksheet, df_to_write, formats)
    apply_status_conditioning(worksheet, df_to_write, formats)


def prepare_excel_df(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    for column in cleaned.columns:
        if pd.api.types.is_datetime64_any_dtype(cleaned[column]):
            continue
        cleaned[column] = cleaned[column].apply(lambda value: "" if pd.isna(value) else value)
    return cleaned


def autosize_columns(worksheet: Any, df: pd.DataFrame) -> None:
    for col_idx, column in enumerate(df.columns):
        sample = df[column].astype("string").fillna("").head(200)
        width = int(max([len(str(column)), *(sample.str.len().fillna(0).tolist())]) + 2)
        if column in {"SCH Project Remark", "SCH Project Update", "PM Notes", "Stakeholder PM Update", "Next Action", "Risk Reason"}:
            width = min(max(width, 28), 65)
        else:
            width = min(max(width, 10), 28)
        worksheet.set_column(col_idx, col_idx, width)


def apply_formats(worksheet: Any, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    for col_idx, column in enumerate(df.columns):
        lower = column.lower()
        if "date" in lower:
            worksheet.set_column(col_idx, col_idx, None, formats["date"])
        elif "days" in lower or "count" in lower or "code" in lower:
            worksheet.set_column(col_idx, col_idx, None, formats["integer"])
        elif "score" in lower or "percent" in lower:
            worksheet.set_column(col_idx, col_idx, None, formats["number"])
        else:
            worksheet.set_column(col_idx, col_idx, None, formats["text"])


def apply_risk_conditioning(worksheet: Any, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    if "Risk Level" not in df.columns or df.empty:
        return
    col = df.columns.get_loc("Risk Level")
    worksheet.conditional_format(1, col, len(df), col, {"type": "text", "criteria": "containing", "value": "Critical", "format": formats["critical"]})
    worksheet.conditional_format(1, col, len(df), col, {"type": "text", "criteria": "containing", "value": "High", "format": formats["high"]})
    worksheet.conditional_format(1, col, len(df), col, {"type": "text", "criteria": "containing", "value": "Watch", "format": formats["watch"]})
    worksheet.conditional_format(1, col, len(df), col, {"type": "text", "criteria": "containing", "value": "Healthy", "format": formats["healthy"]})


def apply_status_conditioning(worksheet: Any, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    if df.empty:
        return
    status_formats = formats.get("status_formats", {})
    status_columns = [column for column in df.columns if "status" in column.lower()]
    for column in status_columns:
        col = df.columns.get_loc(column)
        for status, fmt in status_formats.items():
            worksheet.conditional_format(
                1,
                col,
                len(df),
                col,
                {"type": "text", "criteria": "containing", "value": status, "format": fmt},
            )


def setup_table_block(worksheet: Any, df: pd.DataFrame, start_row: int, start_col: int, formats: dict[str, Any], title: str) -> None:
    worksheet.write(start_row, start_col, title, formats["section"])
    if df.empty:
        worksheet.write(start_row + 1, start_col, "No records", formats["text"])
        return
    for col_idx, column in enumerate(df.columns):
        worksheet.write(start_row + 1, start_col + col_idx, column, formats["header"])
    for row_idx, row in enumerate(df.itertuples(index=False), start=start_row + 2):
        for col_idx, value in enumerate(row):
            worksheet.write(row_idx, start_col + col_idx, "" if pd.isna(value) else value, formats["text"])


def add_dashboard_charts(writer: pd.ExcelWriter, data: dict[str, pd.DataFrame], formats: dict[str, Any]) -> None:
    workbook = writer.book
    worksheet = writer.sheets["Dashboard"]
    cockpit = data["PM_Cockpit"]
    theme = formats.get("theme", resolved_theme({}))
    status_colors = theme["status_colors"]
    risk_colors = theme["risk_colors"]
    risk_counts = cockpit["Risk Level"].value_counts().reindex(["Critical", "High", "Watch", "Healthy", "Closed"], fill_value=0).reset_index()
    risk_counts.columns = ["Risk Level", "Count"]
    status_order = ["Completed", "Active", "Yet To Start", "Held Up", "Lost"]
    status_raw = cockpit["Tracker Status"].value_counts()
    ordered_status = [status for status in status_order if status in status_raw.index]
    ordered_status.extend([status for status in status_raw.index if status not in ordered_status])
    status_counts = status_raw.reindex(ordered_status, fill_value=0).reset_index()
    status_counts.columns = ["Tracker Status", "Count"]

    risk_start = 8
    status_start = 8
    setup_table_block(worksheet, risk_counts, risk_start, 4, formats, "Risk Mix")
    setup_table_block(worksheet, status_counts, status_start, 8, formats, "Status Mix")

    if not risk_counts.empty:
        chart = workbook.add_chart({"type": "doughnut"})
        chart.add_series(
            {
                "name": "Risk Mix",
                "categories": ["Dashboard", risk_start + 2, 4, risk_start + 1 + len(risk_counts), 4],
                "values": ["Dashboard", risk_start + 2, 5, risk_start + 1 + len(risk_counts), 5],
                "points": [
                    {"fill": {"color": risk_colors.get(value, "#1F4E78")}}
                    for value in risk_counts["Risk Level"].astype(str)
                ],
            }
        )
        chart.set_title({"name": "Risk Mix"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 360, "height": 245})
        worksheet.insert_chart("E17", chart)

    if not status_counts.empty:
        chart = workbook.add_chart({"type": "doughnut"})
        chart.add_series(
            {
                "name": "Status Mix",
                "categories": ["Dashboard", status_start + 2, 8, status_start + 1 + len(status_counts), 8],
                "values": ["Dashboard", status_start + 2, 9, status_start + 1 + len(status_counts), 9],
                "points": [
                    {"fill": {"color": status_colors.get(value, "#1F4E78")}}
                    for value in status_counts["Tracker Status"].astype(str)
                ],
            }
        )
        chart.set_title({"name": "Status Mix"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 360, "height": 245})
        worksheet.insert_chart("I17", chart)

    monthly = data.get("Monthly_Summary", pd.DataFrame())
    if not monthly.empty:
        chart = workbook.add_chart({"type": "line"})
        rows = len(monthly) + 1
        chart.add_series(
            {
                "name": "Closing",
                "categories": ["Monthly_Summary", 1, 0, rows - 1, 0],
                "values": ["Monthly_Summary", 1, 9, rows - 1, 9],
                "line": {"color": theme["primary"], "width": 2.25},
            }
        )
        chart.add_series(
            {
                "name": "Completed",
                "categories": ["Monthly_Summary", 1, 0, rows - 1, 0],
                "values": ["Monthly_Summary", 1, 5, rows - 1, 5],
                "line": {"color": status_colors.get("Completed", "#D9EAF7"), "width": 2.25},
            }
        )
        chart.set_title({"name": "Monthly Pipeline"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 560, "height": 260})
        worksheet.insert_chart("B28", chart)


def build_weekly_email_html(data: dict[str, pd.DataFrame], config: dict[str, Any]) -> str:
    cockpit = data["PM_Cockpit"]
    weekly = data["Weekly_Mail_View"]
    today = datetime.now().date()
    week_start = today - timedelta(days=today.weekday())
    completed_this_week = cockpit[
        cockpit["Tracker Status"].eq("Completed")
        & pd.to_datetime(cockpit["End Date"], errors="coerce").dt.date.ge(week_start)
    ]
    new_this_week = cockpit[pd.to_datetime(cockpit["Start Date"], errors="coerce").dt.date.ge(week_start)]
    summary = build_run_summary(data)

    top_rows = weekly.head(15)
    body = f"""
<html>
<body style="font-family:Arial, sans-serif; color:#1f1f1f;">
  <p>Dear Team,</p>
  <p>Please find the weekly PM update as of <b>{today:%d-%b-%Y}</b>.</p>

  <h3 style="color:#1F4E78;">Executive Snapshot</h3>
  <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;">
    <tr style="background:#1F4E78;color:white;">
      <th>Total Projects</th><th>Open Pipeline</th><th>Critical</th><th>High</th><th>Unmatched</th><th>Project Summaries</th>
    </tr>
    <tr>
      <td>{summary['total_projects']}</td>
      <td>{summary['open_pipeline']}</td>
      <td>{summary['critical_risks']}</td>
      <td>{summary['high_risks']}</td>
      <td>{summary['unmatched']}</td>
      <td>{summary['project_summaries_due']}</td>
    </tr>
  </table>

  <h3 style="color:#1F4E78;">This Week</h3>
  <ul>
    <li>New projects started this week: <b>{len(new_this_week)}</b></li>
    <li>Projects completed this week: <b>{len(completed_this_week)}</b></li>
    <li>Critical/high items requiring attention: <b>{summary['critical_risks'] + summary['high_risks']}</b></li>
  </ul>

  <h3 style="color:#1F4E78;">Priority Action Items</h3>
  {df_to_html_table(top_rows, ['Project Name','Tracker Status','SCH Status','SCH Milestone','Age Days','Risk Level','Owner','Next Action'])}

  <p>The detailed PM cockpit is attached for review.</p>
  <p>Regards,<br>{html.escape(config.get('owner_name', 'Vasu'))}</p>
</body>
</html>
"""
    return body


def df_to_html_table(df: pd.DataFrame, columns: list[str]) -> str:
    if df.empty:
        return "<p>No priority action items found.</p>"
    visible = df[[column for column in columns if column in df.columns]].copy()
    rows = []
    header = "".join(f"<th>{html.escape(str(column))}</th>" for column in visible.columns)
    rows.append(f"<tr style='background:#1F4E78;color:white;'>{header}</tr>")
    for _, row in visible.iterrows():
        cells = "".join(f"<td>{html.escape('' if pd.isna(value) else str(value))}</td>" for value in row)
        rows.append(f"<tr>{cells}</tr>")
    return "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;font-size:10pt;'>" + "".join(rows) + "</table>"


def create_outlook_draft(body: str, attachment_path: Path, config: dict[str, Any]) -> str:
    email_config = config.get("email", {})
    if not email_config.get("create_outlook_draft", True):
        return "Skipped by config"
    try:
        import win32com.client  # type: ignore

        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = email_config.get("to", "")
        mail.CC = email_config.get("cc", "")
        mail.Subject = f"{email_config.get('subject_prefix', 'PM Weekly Update')} - {datetime.now():%d-%b-%Y}"
        mail.HTMLBody = body
        if attachment_path.exists():
            mail.Attachments.Add(str(attachment_path))
        mail.Save()
        if email_config.get("display_draft", False):
            mail.Display(False)
        return "Outlook draft saved"
    except Exception as exc:
        logging.warning("Outlook draft creation failed: %s", exc)
        return f"Outlook draft failed; HTML draft created ({exc})"


def build_run_summary(data: dict[str, pd.DataFrame]) -> dict[str, int]:
    cockpit = data["PM_Cockpit"]
    project_view = data.get("Project_Summary_View", pd.DataFrame())
    project_index = data.get("Project_Summary_Index", pd.DataFrame())
    return {
        "total_projects": int(len(cockpit)),
        "open_pipeline": int(cockpit["Tracker Status"].isin(["Active", "Yet To Start", "Held Up"]).sum()),
        "critical_risks": int(cockpit["Risk Level"].eq("Critical").sum()),
        "high_risks": int(cockpit["Risk Level"].eq("High").sum()),
        "watch_items": int(cockpit["Risk Level"].eq("Watch").sum()),
        "unmatched": int(cockpit["Match Method"].eq("No Match").sum()),
        "match_review_needed": int(cockpit["Match Review Needed"].sum()),
        "project_summaries_due": int(len(project_view)),
        "project_summaries_created": int(len(project_index)) if not project_index.empty else 0,
        "legacy_weekly_mail_items": int(cockpit["Mail Include"].astype(str).str.lower().eq("yes").sum()),
        "preserved_note_keys": int(len(data["Preserved_Notes"])),
    }


def write_run_summary(
    path: Path,
    summary: dict[str, int],
    cockpit_path: Path,
    html_path: Path | None,
    project_index_path: Path,
    outlook_status: str,
    log_path: Path,
) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "cockpit_path": str(cockpit_path),
        "weekly_html_path": str(html_path) if html_path else "",
        "project_summary_index_path": str(project_index_path),
        "outlook_status": outlook_status,
        "log_path": str(log_path),
        "summary": summary,
    }
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


if __name__ == "__main__":
    raise SystemExit(main())
