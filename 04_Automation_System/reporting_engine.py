from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd


DATE_FORMAT = "%d-%b-%Y"


@dataclass
class SourceSnapshot:
    file_path: Path
    file_modified: datetime
    sheets: dict[str, pd.DataFrame]
    primary_sheet: str
    raw_df: pd.DataFrame


@dataclass
class AnalysisResult:
    source: SourceSnapshot
    cleaned_df: pd.DataFrame
    overview: pd.DataFrame
    missing_values: pd.DataFrame
    dtype_summary: pd.DataFrame
    kpi_summary: pd.DataFrame
    monthly_trend: pd.DataFrame
    weekly_trend: pd.DataFrame
    top_performers: dict[str, pd.DataFrame]
    bottom_performers: dict[str, pd.DataFrame]
    outliers: pd.DataFrame
    data_quality: pd.DataFrame
    insights: dict[str, list[str]]
    report_month: str
    report_year: int
    report_label: str
    financial_columns: dict[str, str | None]


def load_config(config_path: Path) -> dict[str, Any]:
    with config_path.open("r", encoding="utf-8") as handle:
        config = json.load(handle)

    base_dir = config_path.parent
    for key in ("output_directory", "logs_directory", "templates_directory"):
        config[key] = str(resolve_config_path(base_dir, config[key]))

    config["input_directories"] = [
        str(resolve_config_path(base_dir, value)) for value in config.get("input_directories", [])
    ]
    config["explicit_input_files"] = [
        str(resolve_config_path(base_dir, value)) for value in config.get("explicit_input_files", [])
    ]
    return config


def resolve_config_path(base_dir: Path, value: str) -> Path:
    path = Path(value)
    if path.is_absolute():
        return path
    return (base_dir / path).resolve()


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / f"monthly_reporting_{datetime.now():%Y%m%d_%H%M%S}.log"
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
    )
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    console.setFormatter(logging.Formatter("%(levelname)s | %(message)s"))
    logging.getLogger().addHandler(console)
    return log_path


def detect_latest_input_file(config: dict[str, Any]) -> Path:
    patterns = config.get("file_patterns", ["*.xlsx", "*.xls", "*.csv"])
    exclusions = [value.lower() for value in config.get("exclude_filename_contains", [])]
    candidates: list[Path] = []
    explicit_candidates: list[Path] = []

    for explicit in config.get("explicit_input_files", []):
        path = Path(explicit)
        if path.exists() and path.is_file():
            explicit_candidates.append(path.resolve())

    if config.get("prefer_explicit_input_files", False) and explicit_candidates:
        latest_explicit = max(explicit_candidates, key=lambda path: path.stat().st_mtime)
        logging.info("Selected explicit input file: %s", latest_explicit)
        return latest_explicit

    candidates.extend(explicit_candidates)

    for directory in config.get("input_directories", []):
        root = Path(directory)
        if not root.exists() or not root.is_dir():
            continue
        for pattern in patterns:
            candidates.extend(root.glob(pattern))

    unique_candidates: dict[str, Path] = {}
    for candidate in candidates:
        name_lower = candidate.name.lower()
        if any(exclusion in name_lower for exclusion in exclusions):
            continue
        if candidate.suffix.lower() not in {".xlsx", ".xls", ".csv"}:
            continue
        unique_candidates[str(candidate.resolve()).lower()] = candidate.resolve()

    if not unique_candidates:
        raise FileNotFoundError(
            "No CSV/XLSX input files were found. Add a file to input/ or update config.json."
        )

    latest = max(unique_candidates.values(), key=lambda path: path.stat().st_mtime)
    logging.info("Selected latest input file: %s", latest)
    return latest


def load_source_file(file_path: Path) -> SourceSnapshot:
    suffix = file_path.suffix.lower()
    if suffix == ".csv":
        raw = pd.read_csv(file_path)
        raw = standardize_columns(raw)
        sheets = {"CSV_Data": raw.copy()}
        primary_sheet = "CSV_Data"
    elif suffix in {".xlsx", ".xls"}:
        sheets = pd.read_excel(file_path, sheet_name=None)
        sheets = {name: standardize_columns(df) for name, df in sheets.items()}
        primary_sheet = choose_primary_sheet(sheets)
        raw = sheets[primary_sheet].copy()
    else:
        raise ValueError(f"Unsupported input file type: {suffix}")

    modified = datetime.fromtimestamp(file_path.stat().st_mtime)
    logging.info("Loaded %s with primary sheet '%s'", file_path.name, primary_sheet)
    return SourceSnapshot(file_path, modified, sheets, primary_sheet, raw)


def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    cleaned = df.copy()
    seen: dict[str, int] = {}
    new_columns: list[str] = []
    for column in cleaned.columns:
        name = normalize_column_name(column)
        seen[name] = seen.get(name, 0) + 1
        if seen[name] > 1:
            name = f"{name}_{seen[name]}"
        new_columns.append(name)
    cleaned.columns = new_columns
    return cleaned


def normalize_column_name(value: Any) -> str:
    if value is None or (isinstance(value, float) and np.isnan(value)):
        return "Unnamed"
    text = str(value).replace("\xa0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    text = text.replace("\n", " ")
    return text if text else "Unnamed"


def choose_primary_sheet(sheets: dict[str, pd.DataFrame]) -> str:
    def score(item: tuple[str, pd.DataFrame]) -> float:
        sheet_name, df = item
        non_empty_cells = int(df.notna().sum().sum())
        useful_headers = sum(
            1
            for col in df.columns
            if normalize_token(col) in {"projectcode", "projectname", "status", "startdate", "enddate"}
        )
        followup_penalty = 500 if "follow" in sheet_name.lower() else 0
        return non_empty_cells + useful_headers * 100 - followup_penalty

    return max(sheets.items(), key=score)[0]


def normalize_token(value: Any) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).lower())


def analyze_source(source: SourceSnapshot, config: dict[str, Any]) -> AnalysisResult:
    raw_df = source.raw_df.copy()
    cleaned_df = clean_and_enrich_data(raw_df, source.sheets, config)
    date_columns = detect_date_columns(cleaned_df)
    report_date = choose_report_date(cleaned_df, date_columns, source.file_modified, config)
    report_month = report_date.strftime("%B")
    report_year = int(report_date.year)
    report_label = f"{report_month}_{report_year}"

    overview = build_dataset_overview(raw_df, cleaned_df, date_columns, source)
    missing_values = build_missing_values(cleaned_df)
    dtype_summary = build_dtype_summary(cleaned_df)
    financial_columns = detect_financial_columns(cleaned_df, config)
    monthly_trend = build_monthly_trend(cleaned_df)
    weekly_trend = build_weekly_trend(cleaned_df)
    kpi_summary = build_kpi_summary(cleaned_df, monthly_trend, financial_columns)
    top_performers, bottom_performers = build_performer_tables(cleaned_df, config)
    outliers = detect_outliers(cleaned_df)
    data_quality = build_data_quality_report(cleaned_df)
    insights = generate_insights(
        cleaned_df=cleaned_df,
        kpi_summary=kpi_summary,
        monthly_trend=monthly_trend,
        top_performers=top_performers,
        bottom_performers=bottom_performers,
        outliers=outliers,
        data_quality=data_quality,
        financial_columns=financial_columns,
    )

    return AnalysisResult(
        source=source,
        cleaned_df=cleaned_df,
        overview=overview,
        missing_values=missing_values,
        dtype_summary=dtype_summary,
        kpi_summary=kpi_summary,
        monthly_trend=monthly_trend,
        weekly_trend=weekly_trend,
        top_performers=top_performers,
        bottom_performers=bottom_performers,
        outliers=outliers,
        data_quality=data_quality,
        insights=insights,
        report_month=report_month,
        report_year=report_year,
        report_label=report_label,
        financial_columns=financial_columns,
    )


def clean_and_enrich_data(
    raw_df: pd.DataFrame,
    sheets: dict[str, pd.DataFrame],
    config: dict[str, Any],
) -> pd.DataFrame:
    df = raw_df.copy()
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    df = df.loc[:, ~df.columns.duplicated()]

    for column in df.select_dtypes(include=["object"]).columns:
        df[column] = (
            df[column]
            .astype("string")
            .str.replace("\xa0", " ", regex=False)
            .str.strip()
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        )

    columns = infer_business_columns(df)
    for column in {columns.get("start_date"), columns.get("end_date")}:
        if column and column in df.columns:
            df[column] = pd.to_datetime(df[column], errors="coerce")

    employee_col = columns.get("employee_count")
    if employee_col and employee_col in df.columns:
        df[employee_col] = pd.to_numeric(df[employee_col], errors="coerce")

    status_col = columns.get("status")
    priority_col = columns.get("priority")
    start_col = columns.get("start_date")
    end_col = columns.get("end_date")

    df["Status Group"] = df[status_col].apply(lambda value: classify_status(value, config)) if status_col else "Unknown"
    df["Is Completed"] = df["Status Group"].eq("Completed")
    df["Is Lost"] = df["Status Group"].eq("Lost")
    df["Is Open"] = df["Status Group"].isin(["Active", "Not Started", "Held Up", "Other"])
    df["Is Held Up"] = df["Status Group"].eq("Held Up")

    today = pd.Timestamp(datetime.now().date())
    if start_col:
        start_dates = pd.to_datetime(df[start_col], errors="coerce")
    else:
        start_dates = pd.Series(pd.NaT, index=df.index)
    if end_col:
        end_dates = pd.to_datetime(df[end_col], errors="coerce")
    else:
        end_dates = pd.Series(pd.NaT, index=df.index)

    effective_end_dates = end_dates.fillna(today)
    df["Cycle Days"] = (effective_end_dates - start_dates).dt.days
    df.loc[start_dates.isna(), "Cycle Days"] = np.nan
    df.loc[df["Cycle Days"] < 0, "Cycle Days"] = np.nan

    if priority_col:
        df["SLA Target Days"] = df[priority_col].apply(lambda value: sla_target_for_priority(value, config))
    else:
        df["SLA Target Days"] = config.get("sla_targets_days", {}).get("Default", 90)
    df["SLA Met"] = np.where(df["Cycle Days"].notna(), df["Cycle Days"] <= df["SLA Target Days"], pd.NA)
    df["Ageing Risk"] = df["Is Open"] & df["Cycle Days"].gt(df["SLA Target Days"])

    activity_date = end_dates.fillna(start_dates)
    df["Activity Date"] = activity_date
    df["Activity Month"] = activity_date.dt.to_period("M").astype("string")
    df["Activity Week"] = activity_date.dt.to_period("W").astype("string")

    if employee_col:
        df["Employee Size Band"] = pd.cut(
            df[employee_col],
            bins=[-np.inf, 50, 100, 250, 500, 1000, np.inf],
            labels=["<=50", "51-100", "101-250", "251-500", "501-1000", ">1000"],
        ).astype("string")
    else:
        df["Employee Size Band"] = pd.NA

    risk_col = columns.get("risk")
    if risk_col and risk_col in df.columns:
        df["Risk Flag"] = df[risk_col].notna() | df["Is Held Up"] | df["Ageing Risk"]
    else:
        df["Risk Flag"] = df["Is Held Up"] | df["Ageing Risk"]

    df = merge_followup_data(df, sheets)
    return df.reset_index(drop=True)


def infer_business_columns(df: pd.DataFrame) -> dict[str, str | None]:
    aliases = {
        "project_code": ["Project Code", "Project ID", "Code", "ID"],
        "project_type": ["Project Type", "Type"],
        "project_name": ["Project Name", "Project", "Client", "Customer"],
        "start_date": ["Start date", "Start Date", "Received Date", "Created Date", "Open Date"],
        "end_date": ["End Date", "Completed Date", "Close Date", "Closed Date", "Finish Date"],
        "priority": ["Priority"],
        "status": ["Status", "Project Status"],
        "risk": ["Risk Identified", "Risk", "Issue"],
        "employee_count": ["Employee Count", "Employees", "Headcount"],
        "product": ["Product", "Module", "Category"],
        "assigned_engineer": ["Assigned Engineer", "Engineer", "Owner", "Assignee"],
        "team_leader": ["Team Leader", "Lead", "Manager"],
        "marketing_executive": ["Marketing Executive", "Sales Executive", "Executive"],
        "client_type": ["Client Type", "Customer Type", "Segment"],
    }
    return {key: find_column(df, values) for key, values in aliases.items()}


def find_column(df: pd.DataFrame, aliases: list[str]) -> str | None:
    normalized_columns = {normalize_token(column): column for column in df.columns}
    for alias in aliases:
        token = normalize_token(alias)
        if token in normalized_columns:
            return normalized_columns[token]
    for alias in aliases:
        token = normalize_token(alias)
        for column in df.columns:
            if token and token in normalize_token(column):
                return column
    return None


def classify_status(value: Any, config: dict[str, Any]) -> str:
    if pd.isna(value):
        return "Unknown"
    text = str(value).strip().lower()
    groups = config.get("status_groups", {})
    if text in groups.get("completed", []):
        return "Completed"
    if text in groups.get("lost", []):
        return "Lost"
    if text in groups.get("active", []):
        return "Active"
    if text in groups.get("not_started", []):
        return "Not Started"
    if text in groups.get("held_up", []):
        return "Held Up"
    return "Other"


def sla_target_for_priority(value: Any, config: dict[str, Any]) -> int:
    targets = config.get("sla_targets_days", {})
    if pd.isna(value):
        return int(targets.get("Default", 90))
    text = str(value).strip().lower()
    for priority, days in targets.items():
        if priority.lower() == text:
            return int(days)
    return int(targets.get("Default", 90))


def merge_followup_data(df: pd.DataFrame, sheets: dict[str, pd.DataFrame]) -> pd.DataFrame:
    project_code = infer_business_columns(df).get("project_code")
    if not project_code:
        return df

    followup_sheet = None
    for sheet_name, sheet_df in sheets.items():
        if "follow" in sheet_name.lower() and sheet_df is not df:
            followup_sheet = sheet_df.copy()
            break

    if followup_sheet is None:
        df["Followup Available"] = False
        return df

    followup_sheet = followup_sheet.dropna(axis=1, how="all").dropna(axis=0, how="all")
    followup_code = infer_business_columns(followup_sheet).get("project_code")
    if not followup_code:
        df["Followup Available"] = False
        return df

    followup_cols = [
        column
        for column in followup_sheet.columns
        if column != followup_code and normalize_token(column) not in {"projectname", "status"}
    ]
    renamed = followup_sheet[[followup_code, *followup_cols]].copy()
    renamed = renamed.rename(columns={column: f"Followup {column}" for column in followup_cols})
    merged = df.merge(
        renamed.drop_duplicates(subset=[followup_code]),
        how="left",
        left_on=project_code,
        right_on=followup_code,
        suffixes=("", "_FollowupKey"),
    )
    if followup_code != project_code and followup_code in merged.columns:
        merged = merged.drop(columns=[followup_code])
    merged["Followup Available"] = merged[[column for column in merged.columns if column.startswith("Followup ")]].notna().any(axis=1)
    return merged


def detect_date_columns(df: pd.DataFrame) -> list[str]:
    date_columns: list[str] = []
    for column in df.columns:
        if pd.api.types.is_datetime64_any_dtype(df[column]):
            date_columns.append(column)
            continue
        if "date" in column.lower():
            converted = pd.to_datetime(df[column], errors="coerce")
            if converted.notna().sum() > 0:
                date_columns.append(column)
    return date_columns


def choose_report_date(
    df: pd.DataFrame,
    date_columns: list[str],
    file_modified: datetime,
    config: dict[str, Any],
) -> pd.Timestamp:
    mode = config.get("report_month_mode", "max_data_date")
    if mode == "manual":
        manual_month = config.get("manual_report_month")
        if not manual_month:
            raise ValueError("manual_report_month is required when report_month_mode is manual")
        return pd.Timestamp(datetime.strptime(str(manual_month), "%Y-%m").date())
    if mode == "current_date":
        return pd.Timestamp(datetime.now().date())
    if mode == "file_modified" or not date_columns:
        return pd.Timestamp(file_modified)

    max_dates = []
    for column in date_columns:
        values = pd.to_datetime(df[column], errors="coerce")
        if values.notna().any():
            max_dates.append(values.max())
    return pd.Timestamp(max(max_dates)) if max_dates else pd.Timestamp(file_modified)


def build_dataset_overview(
    raw_df: pd.DataFrame,
    cleaned_df: pd.DataFrame,
    date_columns: list[str],
    source: SourceSnapshot,
) -> pd.DataFrame:
    min_date = None
    max_date = None
    if date_columns:
        dates = pd.concat([pd.to_datetime(cleaned_df[col], errors="coerce") for col in date_columns])
        dates = dates.dropna()
        if not dates.empty:
            min_date = dates.min().strftime(DATE_FORMAT)
            max_date = dates.max().strftime(DATE_FORMAT)

    rows = [
        ("Source File", str(source.file_path)),
        ("Primary Sheet", source.primary_sheet),
        ("File Modified", source.file_modified.strftime("%d-%b-%Y %H:%M")),
        ("Total Rows", len(raw_df)),
        ("Total Columns", raw_df.shape[1]),
        ("Cleaned Rows", len(cleaned_df)),
        ("Cleaned Columns", cleaned_df.shape[1]),
        ("Missing Values", int(raw_df.isna().sum().sum())),
        ("Duplicate Records", int(raw_df.duplicated().sum())),
        ("Date Columns", ", ".join(date_columns) if date_columns else "Not detected"),
        ("Date Range Covered", f"{min_date} to {max_date}" if min_date and max_date else "Not available"),
    ]
    return pd.DataFrame(rows, columns=["Metric", "Value"])


def build_missing_values(df: pd.DataFrame) -> pd.DataFrame:
    missing = df.isna().sum().reset_index()
    missing.columns = ["Column", "Missing Values"]
    missing["Missing %"] = np.where(len(df) > 0, missing["Missing Values"] / len(df), 0.0)
    return missing.sort_values(["Missing Values", "Column"], ascending=[False, True])


def build_dtype_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for column in df.columns:
        rows.append(
            {
                "Column": column,
                "Data Type": str(df[column].dtype),
                "Non Null": int(df[column].notna().sum()),
                "Unique Values": int(df[column].nunique(dropna=True)),
            }
        )
    return pd.DataFrame(rows)


def detect_financial_columns(df: pd.DataFrame, config: dict[str, Any]) -> dict[str, str | None]:
    aliases = config.get("business_kpi_aliases", {})
    return {
        key: find_column(df, alias_list)
        for key, alias_list in aliases.items()
        if key in {"revenue", "cost", "profit", "conversion", "productivity", "utilization", "sla"}
    }


def build_monthly_trend(df: pd.DataFrame) -> pd.DataFrame:
    if "Activity Date" not in df.columns or df["Activity Date"].dropna().empty:
        return pd.DataFrame(
            columns=[
                "Month",
                "Total Projects",
                "Completed",
                "Lost",
                "Active/Open",
                "Held Up",
                "Completion Rate",
                "Loss Rate",
                "Avg Cycle Days",
                "MoM Growth %",
            ]
        )

    trend_df = df[df["Activity Date"].notna()].copy()
    trend_df["Month"] = trend_df["Activity Date"].dt.to_period("M").dt.to_timestamp()
    grouped = (
        trend_df.groupby("Month")
        .agg(
            **{
                "Total Projects": ("Status Group", "size"),
                "Completed": ("Is Completed", "sum"),
                "Lost": ("Is Lost", "sum"),
                "Active/Open": ("Is Open", "sum"),
                "Held Up": ("Is Held Up", "sum"),
                "Avg Cycle Days": ("Cycle Days", "mean"),
            }
        )
        .reset_index()
        .sort_values("Month")
    )
    grouped["Completion Rate"] = np.where(
        grouped["Total Projects"] > 0, grouped["Completed"] / grouped["Total Projects"], 0.0
    )
    grouped["Loss Rate"] = np.where(grouped["Total Projects"] > 0, grouped["Lost"] / grouped["Total Projects"], 0.0)
    grouped["MoM Growth %"] = grouped["Total Projects"].pct_change().replace([np.inf, -np.inf], np.nan).fillna(0.0)
    grouped["Month"] = grouped["Month"].dt.strftime("%b-%Y")
    return grouped


def build_weekly_trend(df: pd.DataFrame) -> pd.DataFrame:
    if "Activity Date" not in df.columns or df["Activity Date"].dropna().empty:
        return pd.DataFrame(columns=["Week", "Total Projects", "Completed", "Lost", "Completion Rate"])

    trend_df = df[df["Activity Date"].notna()].copy()
    trend_df["Week"] = trend_df["Activity Date"].dt.to_period("W").astype(str)
    grouped = (
        trend_df.groupby("Week")
        .agg(
            **{
                "Total Projects": ("Status Group", "size"),
                "Completed": ("Is Completed", "sum"),
                "Lost": ("Is Lost", "sum"),
            }
        )
        .reset_index()
        .sort_values("Week")
    )
    grouped["Completion Rate"] = np.where(
        grouped["Total Projects"] > 0, grouped["Completed"] / grouped["Total Projects"], 0.0
    )
    return grouped


def build_kpi_summary(
    df: pd.DataFrame,
    monthly_trend: pd.DataFrame,
    financial_columns: dict[str, str | None],
) -> pd.DataFrame:
    total = len(df)
    completed = int(df["Is Completed"].sum()) if "Is Completed" in df else 0
    lost = int(df["Is Lost"].sum()) if "Is Lost" in df else 0
    open_items = int(df["Is Open"].sum()) if "Is Open" in df else 0
    held_up = int(df["Is Held Up"].sum()) if "Is Held Up" in df else 0
    age_risk = int(df["Ageing Risk"].sum()) if "Ageing Risk" in df else 0
    sla_known = df["SLA Met"].dropna() if "SLA Met" in df else pd.Series(dtype=bool)
    cycle_days = pd.to_numeric(df.get("Cycle Days", pd.Series(dtype=float)), errors="coerce").dropna()
    employee_count = pd.to_numeric(df.get("Employee Count", pd.Series(dtype=float)), errors="coerce")

    previous_volume = None
    latest_volume = None
    mom_growth = 0.0
    if len(monthly_trend) >= 2:
        previous_volume = float(monthly_trend["Total Projects"].iloc[-2])
        latest_volume = float(monthly_trend["Total Projects"].iloc[-1])
        mom_growth = (latest_volume - previous_volume) / previous_volume if previous_volume else 0.0
    elif len(monthly_trend) == 1:
        latest_volume = float(monthly_trend["Total Projects"].iloc[-1])

    rows = [
        ("Total Projects", total, "Count", "Overall project volume in the selected tracker"),
        ("Completed Projects", completed, "Count", "Projects marked completed"),
        ("Active/Open Projects", open_items, "Count", "Projects not yet completed or lost"),
        ("Lost Projects", lost, "Count", "Projects marked lost/cancelled"),
        ("Held Up Projects", held_up, "Count", "Blocked/on-hold work requiring intervention"),
        ("Completion Rate", safe_div(completed, total), "Percent", "Completed projects divided by total projects"),
        ("Loss Rate", safe_div(lost, total), "Percent", "Lost projects divided by total projects"),
        ("Conversion Rate", safe_div(completed, completed + lost), "Percent", "Completed divided by completed plus lost"),
        ("SLA Met Rate", float(sla_known.mean()) if len(sla_known) else np.nan, "Percent", "Projects within configured cycle-time SLA"),
        ("Ageing Risk Count", age_risk, "Count", "Open projects older than their SLA target"),
        ("Average Cycle Days", float(cycle_days.mean()) if len(cycle_days) else np.nan, "Number", "Average days from start to end/today"),
        ("Median Cycle Days", float(cycle_days.median()) if len(cycle_days) else np.nan, "Number", "Median days from start to end/today"),
        ("Total Employee Count", float(employee_count.sum(skipna=True)) if employee_count.notna().any() else np.nan, "Number", "Total employee count linked to projects"),
        ("Average Employee Count", float(employee_count.mean(skipna=True)) if employee_count.notna().any() else np.nan, "Number", "Average client employee count per populated row"),
        ("Latest Month Volume", latest_volume if latest_volume is not None else np.nan, "Count", "Projects in latest detected activity month"),
        ("Previous Month Volume", previous_volume if previous_volume is not None else np.nan, "Count", "Projects in prior activity month"),
        ("MoM Volume Growth %", mom_growth, "Percent", "Month-over-month project volume growth"),
    ]

    for metric, column in financial_columns.items():
        if metric in {"revenue", "cost", "profit"}:
            if column:
                values = pd.to_numeric(df[column], errors="coerce")
                rows.append((metric.title(), float(values.sum(skipna=True)), "Currency", f"Detected from column '{column}'"))
            else:
                rows.append((metric.title(), np.nan, "Currency", "Not available in current source file"))

    if financial_columns.get("revenue") and financial_columns.get("cost") and not financial_columns.get("profit"):
        revenue = pd.to_numeric(df[financial_columns["revenue"]], errors="coerce").sum(skipna=True)
        cost = pd.to_numeric(df[financial_columns["cost"]], errors="coerce").sum(skipna=True)
        rows.append(("Computed Profit", revenue - cost, "Currency", "Revenue minus cost"))
        rows.append(("Profit Margin", safe_div(revenue - cost, revenue), "Percent", "Computed profit divided by revenue"))

    return pd.DataFrame(rows, columns=["KPI", "Value", "Format", "Definition"])


def safe_div(numerator: float, denominator: float) -> float:
    return float(numerator / denominator) if denominator else 0.0


def build_performer_tables(
    df: pd.DataFrame,
    config: dict[str, Any],
) -> tuple[dict[str, pd.DataFrame], dict[str, pd.DataFrame]]:
    top_n = int(config.get("top_n", 10))
    columns = infer_business_columns(df)
    dimensions = {
        "Assigned Engineer": columns.get("assigned_engineer"),
        "Team Leader": columns.get("team_leader"),
        "Product": columns.get("product"),
        "Project Type": columns.get("project_type"),
        "Client Type": columns.get("client_type"),
    }
    top: dict[str, pd.DataFrame] = {}
    bottom: dict[str, pd.DataFrame] = {}

    for label, column in dimensions.items():
        if not column or column not in df.columns:
            continue
        subset = df[df[column].notna()].copy()
        if subset.empty:
            continue
        grouped = (
            subset.groupby(column)
            .agg(
                **{
                    "Total Projects": ("Status Group", "size"),
                    "Completed": ("Is Completed", "sum"),
                    "Lost": ("Is Lost", "sum"),
                    "Open": ("Is Open", "sum"),
                    "Ageing Risk": ("Ageing Risk", "sum"),
                    "Avg Cycle Days": ("Cycle Days", "mean"),
                }
            )
            .reset_index()
            .rename(columns={column: label})
        )
        grouped["Completion Rate"] = np.where(
            grouped["Total Projects"] > 0, grouped["Completed"] / grouped["Total Projects"], 0.0
        )
        grouped["Risk Rate"] = np.where(
            grouped["Total Projects"] > 0, grouped["Ageing Risk"] / grouped["Total Projects"], 0.0
        )
        grouped = grouped.sort_values(["Completed", "Completion Rate", "Total Projects"], ascending=False)
        top[label] = grouped.head(top_n).reset_index(drop=True)
        bottom[label] = (
            grouped[grouped["Total Projects"] >= 1]
            .sort_values(["Completion Rate", "Ageing Risk", "Total Projects"], ascending=[True, False, False])
            .head(top_n)
            .reset_index(drop=True)
        )
    return top, bottom


def detect_outliers(df: pd.DataFrame) -> pd.DataFrame:
    rows: list[dict[str, Any]] = []
    for column in ["Cycle Days", "Employee Count"]:
        if column not in df.columns:
            continue
        values = pd.to_numeric(df[column], errors="coerce")
        values = values.dropna()
        if len(values) < 4:
            continue
        q1 = values.quantile(0.25)
        q3 = values.quantile(0.75)
        iqr = q3 - q1
        upper = q3 + 1.5 * iqr
        lower = q1 - 1.5 * iqr
        outlier_index = df[(pd.to_numeric(df[column], errors="coerce") > upper) | (pd.to_numeric(df[column], errors="coerce") < lower)].index
        for idx in outlier_index:
            rows.append(
                {
                    "Row Number": int(idx) + 2,
                    "Column": column,
                    "Value": df.loc[idx, column],
                    "Lower Bound": lower,
                    "Upper Bound": upper,
                    "Reason": "IQR outlier",
                }
            )
    return pd.DataFrame(rows)


def build_data_quality_report(df: pd.DataFrame) -> pd.DataFrame:
    columns = infer_business_columns(df)
    checks: list[dict[str, Any]] = []

    def add_check(severity: str, issue: str, mask: pd.Series | None, column: str | None, recommendation: str) -> None:
        count = int(mask.sum()) if mask is not None else 0
        if count > 0:
            checks.append(
                {
                    "Severity": severity,
                    "Issue": issue,
                    "Row Count": count,
                    "Column": column or "",
                    "Recommendation": recommendation,
                }
            )

    for key, friendly in [
        ("project_code", "Project code missing"),
        ("project_name", "Project name missing"),
        ("status", "Status missing"),
        ("start_date", "Start date missing"),
    ]:
        column = columns.get(key)
        if column:
            add_check("High", friendly, df[column].isna(), column, f"Populate {column} for reporting completeness.")

    project_code = columns.get("project_code")
    if project_code:
        duplicates = df[project_code].duplicated(keep=False) & df[project_code].notna()
        add_check("Medium", "Duplicate project codes", duplicates, project_code, "Confirm whether duplicate project rows are intentional.")

    start_col = columns.get("start_date")
    end_col = columns.get("end_date")
    if start_col and end_col:
        start_dates = pd.to_datetime(df[start_col], errors="coerce")
        end_dates = pd.to_datetime(df[end_col], errors="coerce")
        add_check("High", "End date before start date", end_dates.lt(start_dates), end_col, "Correct date sequence.")
        completed_without_end = df["Is Completed"] & end_dates.isna()
        add_check("Medium", "Completed project missing end date", completed_without_end, end_col, "Add completion date for cycle-time accuracy.")

    if "Ageing Risk" in df.columns:
        add_check(
            "High",
            "Open projects older than SLA target",
            df["Ageing Risk"].fillna(False),
            "Cycle Days",
            "Review blockers and reset delivery plan for overdue open items.",
        )

    assigned = columns.get("assigned_engineer")
    if assigned:
        add_check(
            "Medium",
            "Open project missing assigned engineer",
            df["Is Open"] & df[assigned].isna(),
            assigned,
            "Assign an owner to each open project.",
        )

    product = columns.get("product")
    if product:
        add_check("Low", "Product/category missing", df[product].isna(), product, "Populate product/category to improve portfolio reporting.")

    if not checks:
        checks.append(
            {
                "Severity": "Info",
                "Issue": "No major data quality issues detected",
                "Row Count": 0,
                "Column": "",
                "Recommendation": "Continue monthly validation.",
            }
        )
    return pd.DataFrame(checks)


def generate_insights(
    cleaned_df: pd.DataFrame,
    kpi_summary: pd.DataFrame,
    monthly_trend: pd.DataFrame,
    top_performers: dict[str, pd.DataFrame],
    bottom_performers: dict[str, pd.DataFrame],
    outliers: pd.DataFrame,
    data_quality: pd.DataFrame,
    financial_columns: dict[str, str | None],
) -> dict[str, list[str]]:
    kpis = dict(zip(kpi_summary["KPI"], kpi_summary["Value"]))
    insights: dict[str, list[str]] = {
        "executive_summary": [],
        "best_performance_areas": [],
        "risk_areas": [],
        "significant_changes": [],
        "root_cause_analysis": [],
        "growth_opportunities": [],
        "efficiency_recommendations": [],
        "cost_optimization": [],
        "forecast_commentary": [],
        "recommendations": [],
    }

    total = int(kpis.get("Total Projects", 0) or 0)
    completed = int(kpis.get("Completed Projects", 0) or 0)
    completion_rate = kpis.get("Completion Rate", 0) or 0
    loss_rate = kpis.get("Loss Rate", 0) or 0
    sla_rate = kpis.get("SLA Met Rate", np.nan)
    age_risk = int(kpis.get("Ageing Risk Count", 0) or 0)
    avg_cycle = kpis.get("Average Cycle Days", np.nan)

    insights["executive_summary"].append(
        f"{completed} of {total} projects are completed, giving a completion rate of {completion_rate:.1%}."
    )
    insights["executive_summary"].append(
        f"Loss rate is {loss_rate:.1%}; ageing-risk items stand at {age_risk} based on configured SLA targets."
    )
    if not pd.isna(avg_cycle):
        insights["executive_summary"].append(f"Average cycle time is {avg_cycle:.1f} days across dated projects.")
    if not pd.isna(sla_rate):
        insights["executive_summary"].append(f"SLA adherence is {sla_rate:.1%} using the priority-based SLA model.")

    for label, table in top_performers.items():
        if table.empty:
            continue
        first = table.iloc[0]
        name = first[label]
        completed_count = int(first.get("Completed", 0))
        rate = first.get("Completion Rate", 0)
        insights["best_performance_areas"].append(
            f"{label} '{name}' leads delivery with {completed_count} completions and {rate:.1%} completion rate."
        )
        break

    if "Project Type" in top_performers and not top_performers["Project Type"].empty:
        row = top_performers["Project Type"].iloc[0]
        insights["growth_opportunities"].append(
            f"Scale repeatable delivery practices from the strongest project type, '{row['Project Type']}', which has {int(row['Total Projects'])} tracked projects."
        )

    high_quality = data_quality[data_quality["Severity"].isin(["High", "Medium"])]
    if not high_quality.empty:
        top_issue = high_quality.sort_values("Row Count", ascending=False).iloc[0]
        insights["risk_areas"].append(
            f"Primary data/operational risk: {top_issue['Issue']} impacts {int(top_issue['Row Count'])} rows."
        )

    if age_risk > 0:
        insights["risk_areas"].append(
            f"{age_risk} open projects are older than target SLA and should be reviewed in the next delivery standup."
        )

    if len(monthly_trend) >= 2:
        latest = monthly_trend.iloc[-1]
        previous = monthly_trend.iloc[-2]
        delta = int(latest["Total Projects"] - previous["Total Projects"])
        insights["significant_changes"].append(
            f"Latest month volume changed by {delta:+d} projects versus the previous month."
        )
        insights["forecast_commentary"].append(
            f"Short-term project volume indicator is {latest['Total Projects']} in the latest month versus {previous['Total Projects']} previously."
        )
    elif len(monthly_trend) == 1:
        latest = monthly_trend.iloc[-1]
        insights["forecast_commentary"].append(
            f"Only one activity month is available; baseline monthly volume is {latest['Total Projects']} projects."
        )

    if not outliers.empty:
        cycle_outliers = outliers[outliers["Column"].eq("Cycle Days")]
        if not cycle_outliers.empty:
            insights["root_cause_analysis"].append(
                f"{len(cycle_outliers)} cycle-time outliers were detected; review scope, dependency, and assignment notes for these projects."
            )

    if "Assigned Engineer" in bottom_performers and not bottom_performers["Assigned Engineer"].empty:
        weak = bottom_performers["Assigned Engineer"].iloc[0]
        insights["efficiency_recommendations"].append(
            f"Coach or rebalance work for '{weak['Assigned Engineer']}', where completion rate is {weak['Completion Rate']:.1%}."
        )

    if financial_columns.get("revenue") is None and financial_columns.get("cost") is None:
        insights["cost_optimization"].append(
            "Revenue and cost columns are not present; add commercial fields to quantify margin, cost-to-serve, and budget variance next month."
        )
    else:
        insights["cost_optimization"].append(
            "Use the detected financial columns to compare delivery cycle time with revenue and cost concentration."
        )

    insights["recommendations"].extend(
        [
            "Review ageing-risk projects first, then reassign owners where open items do not have clear accountability.",
            "Standardize mandatory fields for status, start/end date, owner, product, and client type before monthly close.",
            "Add revenue, cost, budget, and target SLA columns to unlock financial variance and margin reporting.",
        ]
    )

    for key, values in insights.items():
        if not values:
            insights[key].append("No material exception detected from the current dataset.")
    return insights


def flatten_insights(insights: dict[str, list[str]]) -> pd.DataFrame:
    rows = []
    for section, values in insights.items():
        title = section.replace("_", " ").title()
        for item in values:
            rows.append({"Section": title, "Insight": item})
    return pd.DataFrame(rows)
