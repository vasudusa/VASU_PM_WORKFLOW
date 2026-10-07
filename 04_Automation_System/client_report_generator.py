from __future__ import annotations

import argparse
import json
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.marker import DataPoint
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.dml.color import RGBColor
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Inches, Pt


SOURCE_FILE = Path(r"D:\VASU_PM_WORKFLOW\01_Source_Files\VASU's Tracker 2026.xlsx")
REPORTS_ROOT = Path(r"D:\VASU_PM_WORKFLOW\02_Output")
CONFIG_FILE = Path(__file__).with_name("config.json")

DEFAULT_THEME = {
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


def load_theme() -> dict[str, Any]:
    config_path = CONFIG_FILE
    if not config_path.exists():
        return DEFAULT_THEME.copy()
    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception:
        return DEFAULT_THEME.copy()
    incoming = config.get("theme", {})
    theme = {**DEFAULT_THEME, **{key: value for key, value in incoming.items() if key not in {"status_colors", "risk_colors"}}}
    theme["status_colors"] = {**DEFAULT_THEME["status_colors"], **incoming.get("status_colors", {})}
    theme["risk_colors"] = {**DEFAULT_THEME["risk_colors"], **incoming.get("risk_colors", {})}
    return theme


THEME = load_theme()


def clean_hex(value: str) -> str:
    return str(value).replace("#", "").upper()


def ppt_rgb(value: str) -> RGBColor:
    text = clean_hex(value)
    return RGBColor(int(text[0:2], 16), int(text[2:4], 16), int(text[4:6], 16))


BLUE = clean_hex(THEME["primary"])
LIGHT_BLUE = clean_hex(THEME["status_colors"]["Completed"])
LIGHT_YELLOW = clean_hex(THEME["status_colors"]["Yet To Start"])
LIGHT_GREEN = clean_hex(THEME["status_colors"]["Active"])
LIGHT_RED = clean_hex(THEME["status_colors"]["Lost"])
LIGHT_ORANGE = clean_hex(THEME["status_colors"]["Held Up"])
GREY = "E7E6E6"
WHITE = "FFFFFF"
DARK_TEXT = clean_hex(THEME["text"])

PPT_BLUE = ppt_rgb(THEME["primary"])
PPT_DARK = ppt_rgb(THEME["text"])
PPT_MUTED = ppt_rgb(THEME["muted"])
PPT_GREEN = ppt_rgb(THEME["status_colors"]["Active"])
PPT_COMPLETED = ppt_rgb(THEME["status_colors"]["Completed"])
PPT_RED = ppt_rgb(THEME["status_colors"]["Lost"])
PPT_ORANGE = ppt_rgb(THEME["status_colors"]["Held Up"])
PPT_YELLOW = ppt_rgb(THEME["status_colors"]["Yet To Start"])
PPT_LIGHT = ppt_rgb(THEME["panel"])
PPT_WHITE = RGBColor(255, 255, 255)


@dataclass
class ClientReportData:
    projects: pd.DataFrame
    followup: pd.DataFrame
    monthly_metrics: pd.DataFrame
    yoy_summary: pd.DataFrame
    yoy_comparison: pd.DataFrame
    yoy_insights: list[str]
    latest: dict[str, Any]
    summary: dict[str, Any]
    report_month: str
    report_period: str
    report_year: int


def load_app_config() -> dict[str, Any]:
    if not CONFIG_FILE.exists():
        return {}
    try:
        return json.loads(CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def build_report_data(source_file: Path = SOURCE_FILE, as_of_date: pd.Timestamp | None = None) -> ClientReportData:
    raw = pd.read_excel(source_file, sheet_name="Sheet1")
    raw = raw.rename(columns={"Start date": "Start Date", "Assigned Engineer ": "Assigned Engineer"})
    raw = repair_wrapped_project_names(raw)

    core_cols = [
        "Project Code",
        "Project Type",
        "Project Name",
        "Received From",
        "Start Date",
        "End Date",
        "Priority",
        "Status",
    ]
    projects = raw[core_cols].copy()
    for col in ["Project Type", "Project Name", "Received From", "Priority", "Status"]:
        projects[col] = projects[col].astype("string").str.replace("\xa0", " ", regex=False).str.strip()
    projects["Start Date"] = pd.to_datetime(projects["Start Date"], errors="coerce")
    projects["End Date"] = pd.to_datetime(projects["End Date"], errors="coerce")
    projects = projects[
        projects["Project Code"].notna()
        & projects["Project Name"].notna()
        & projects["Status"].notna()
        & projects["Start Date"].notna()
    ].copy()
    projects["Project Code"] = projects["Project Code"].astype(int)
    projects = projects.sort_values("Project Code").reset_index(drop=True)

    report_date = pd.Timestamp(as_of_date if as_of_date is not None else datetime.now().date())
    report_month = report_date.strftime("%B")
    report_period = report_date.strftime("%Y-%m")
    report_year = int(report_date.year)

    monthly_metrics = calculate_monthly_metrics(projects, through_period=pd.Period(report_period, freq="M"))
    latest_row = monthly_metrics[monthly_metrics["Month"].eq(report_period)].iloc[0].to_dict()
    previous_period = (pd.Period(report_period, freq="M") - 1).strftime("%Y-%m")
    previous_match = monthly_metrics[monthly_metrics["Month"].eq(previous_period)]
    previous_closing = int(previous_match["Closing"].iloc[0]) if not previous_match.empty else 0

    status_counts = projects["Status"].value_counts().to_dict()
    type_counts = projects["Project Type"].value_counts().to_dict()
    total = len(projects)
    completed = int(status_counts.get("Completed", 0))
    lost = int(status_counts.get("Lost", 0))
    active = int(status_counts.get("Active", 0))
    yet_to_start = int(status_counts.get("Yet To Start", 0))
    held_up = int(status_counts.get("Held Up", 0))
    open_pipeline = active + yet_to_start

    durations = (projects["End Date"] - projects["Start Date"]).dt.days
    completed_durations = durations[projects["Status"].eq("Completed") & durations.notna()]
    summary = {
        "total": total,
        "completed": completed,
        "lost": lost,
        "active": active,
        "yet_to_start": yet_to_start,
        "held_up": held_up,
        "open_pipeline": open_pipeline,
        "completion_rate": completed / total if total else 0,
        "success_rate": (total - lost) / total if total else 0,
        "loss_rate": lost / total if total else 0,
        "allocation": int(type_counts.get("Allocation", 0)),
        "escalation": int(type_counts.get("Escalation", 0)),
        "avg_duration": float(completed_durations.mean()) if not completed_durations.empty else 0,
        "latest_month": report_period,
        "latest_opening": int(latest_row["Opening"]),
        "latest_additions": int(latest_row["Total Add"]),
        "latest_completed": int(latest_row["Completed"]),
        "latest_lost": int(latest_row["Lost"]),
        "latest_held": int(latest_row["Held"]),
        "latest_closing": int(latest_row["Closing"]),
        "previous_month": pd.Period(previous_period, freq="M").strftime("%B"),
        "previous_closing": previous_closing,
        "next_month": (pd.Period(report_period, freq="M") + 1).strftime("%B"),
    }

    followup = build_quarterly_followup_plan(projects, report_date)
    summary["quarterly_followups_due"] = int(followup["Followup Due This Month"].eq("Yes").sum()) if not followup.empty else 0
    yoy_summary = build_yoy_summary(projects, monthly_metrics, report_date)
    yoy_comparison = build_yoy_comparison(projects, report_date)
    yoy_insights = build_yoy_insights(yoy_comparison, summary, report_date)

    return ClientReportData(
        projects=projects,
        followup=followup,
        monthly_metrics=monthly_metrics,
        yoy_summary=yoy_summary,
        yoy_comparison=yoy_comparison,
        yoy_insights=yoy_insights,
        latest=latest_row,
        summary=summary,
        report_month=report_month,
        report_period=report_period,
        report_year=report_year,
    )


def repair_wrapped_project_names(df: pd.DataFrame) -> pd.DataFrame:
    fixed = df.copy()
    for idx in range(1, len(fixed)):
        current_code = fixed.at[idx, "Project Code"] if "Project Code" in fixed.columns else None
        current_name = fixed.at[idx, "Project Name"] if "Project Name" in fixed.columns else None
        if pd.isna(current_code) and pd.notna(current_name):
            previous_idx = idx - 1
            while previous_idx >= 0 and pd.isna(fixed.at[previous_idx, "Project Code"]):
                previous_idx -= 1
            if previous_idx >= 0 and pd.notna(fixed.at[previous_idx, "Project Name"]):
                fixed.at[previous_idx, "Project Name"] = f"{fixed.at[previous_idx, 'Project Name']} {current_name}".strip()
            fixed.at[idx, "Project Name"] = pd.NA
    return fixed


def calculate_monthly_metrics(projects: pd.DataFrame, through_period: pd.Period | None = None) -> pd.DataFrame:
    start_period = projects["Start Date"].min().to_period("M")
    date_maxes = [projects["Start Date"].max()]
    if projects["End Date"].notna().any():
        date_maxes.append(projects["End Date"].max())
    end_period = max(date_maxes).to_period("M")
    if through_period is not None:
        end_period = max(end_period, through_period)
        start_period = min(start_period, through_period)
    rows = []
    opening = 0
    for period in pd.period_range(start_period, end_period, freq="M"):
        starts = projects[projects["Start Date"].dt.to_period("M").eq(period)]
        ends = projects[projects["End Date"].dt.to_period("M").eq(period)]
        add_alloc = int(starts["Project Type"].eq("Allocation").sum())
        add_escal = int(starts["Project Type"].eq("Escalation").sum())
        total_add = add_alloc + add_escal
        completed = int(ends["Status"].eq("Completed").sum())
        lost = int(ends["Status"].eq("Lost").sum())
        held = int(ends["Status"].eq("Held Up").sum())
        net_change = total_add - completed - lost - held
        closing = opening + net_change
        rows.append(
            {
                "Month": str(period),
                "Opening": opening,
                "Add Alloc": add_alloc,
                "Add Escal": add_escal,
                "Total Add": total_add,
                "Completed": completed,
                "Lost": lost,
                "Held": held,
                "Net Change": net_change,
                "Closing": closing,
            }
        )
        opening = closing
    return pd.DataFrame(rows)


def build_yoy_summary(projects: pd.DataFrame, monthly_metrics: pd.DataFrame, report_date: pd.Timestamp) -> pd.DataFrame:
    start_year = int(projects["Start Date"].dt.year.min())
    end_year = int(report_date.year)
    rows = []
    for year in range(start_year, end_year + 1):
        year_starts = projects[projects["Start Date"].dt.year.eq(year)]
        year_ends = projects[projects["End Date"].dt.year.eq(year)]
        month_limit = 12 if year < end_year else int(report_date.month)
        ytd_starts = year_starts[year_starts["Start Date"].dt.month.le(month_limit)]
        ytd_ends = year_ends[year_ends["End Date"].dt.month.le(month_limit)]
        closing_period = f"{year}-{month_limit:02d}"
        closing_row = monthly_metrics[monthly_metrics["Month"].eq(closing_period)]
        closing = int(closing_row["Closing"].iloc[0]) if not closing_row.empty else 0
        additions = int(ytd_starts.shape[0])
        completed = int(ytd_ends["Status"].eq("Completed").sum())
        lost = int(ytd_ends["Status"].eq("Lost").sum())
        held = int(ytd_ends["Status"].eq("Held Up").sum())
        rows.append(
            {
                "Year": year,
                "Period Covered": f"Jan-{pd.Timestamp(year=year, month=month_limit, day=1):%b}",
                "Add Alloc": int(ytd_starts["Project Type"].eq("Allocation").sum()),
                "Add Escal": int(ytd_starts["Project Type"].eq("Escalation").sum()),
                "Total Additions": additions,
                "Completed": completed,
                "Lost": lost,
                "Held Up": held,
                "Closing Balance": closing,
                "Completion Rate": completed / additions if additions else 0,
                "Success Rate": (additions - lost) / additions if additions else 0,
            }
        )
    return pd.DataFrame(rows)


def build_yoy_comparison(projects: pd.DataFrame, report_date: pd.Timestamp) -> pd.DataFrame:
    current_year = int(report_date.year)
    previous_year = current_year - 1
    month_limit = int(report_date.month)

    def metrics_for(year: int) -> dict[str, float]:
        starts = projects[projects["Start Date"].dt.year.eq(year) & projects["Start Date"].dt.month.le(month_limit)]
        ends = projects[projects["End Date"].dt.year.eq(year) & projects["End Date"].dt.month.le(month_limit)]
        total_add = int(starts.shape[0])
        completed = int(ends["Status"].eq("Completed").sum())
        lost = int(ends["Status"].eq("Lost").sum())
        held = int(ends["Status"].eq("Held Up").sum())
        active_open = int(projects[projects["Start Date"].dt.year.le(year) & projects["Status"].isin(["Active", "Yet To Start"])].shape[0])
        return {
            "Total Additions": total_add,
            "Completed": completed,
            "Lost": lost,
            "Held Up": held,
            "Completion Rate": completed / total_add if total_add else 0,
            "Open Pipeline": active_open,
        }

    previous = metrics_for(previous_year)
    current = metrics_for(current_year)
    rows = []
    for metric in ["Total Additions", "Completed", "Lost", "Held Up", "Completion Rate", "Open Pipeline"]:
        prev = previous[metric]
        curr = current[metric]
        change = curr - prev
        change_pct = change / prev if prev else 0
        rows.append(
            {
                "Metric": metric,
                f"{previous_year} Same Period": prev,
                f"{current_year} Same Period": curr,
                "Change": change,
                "Change %": change_pct,
            }
        )
    return pd.DataFrame(rows)


def build_yoy_insights(yoy_comparison: pd.DataFrame, summary: dict[str, Any], report_date: pd.Timestamp) -> list[str]:
    insights = []
    current_year = report_date.year
    month_name = report_date.strftime("%B")
    lookup = yoy_comparison.set_index("Metric") if not yoy_comparison.empty else pd.DataFrame()
    additions_change = float(lookup.loc["Total Additions", "Change %"]) if "Total Additions" in lookup.index else 0
    completed_change = float(lookup.loc["Completed", "Change %"]) if "Completed" in lookup.index else 0
    completion_rate = summary.get("completion_rate", 0)
    insights.append(f"{month_name} {current_year} portfolio has {summary['total']} tracked projects with {summary['completed']} completed.")
    if additions_change < 0:
        insights.append(f"Same-period additions are down {abs(additions_change):.1%} versus last year; pipeline creation needs management focus.")
    elif additions_change > 0:
        insights.append(f"Same-period additions are up {additions_change:.1%} versus last year; maintain intake discipline and delivery capacity.")
    else:
        insights.append("Same-period additions are flat versus last year; delivery throughput will decide pipeline health.")
    if completed_change > 0:
        insights.append(f"Same-period completions are up {completed_change:.1%}; delivery performance is improving.")
    elif completed_change < 0:
        insights.append(f"Same-period completions are down {abs(completed_change):.1%}; review ageing active projects and owner assignment.")
    insights.append(f"Current completion rate is {completion_rate:.1%}; open pipeline is {summary['open_pipeline']} projects.")
    insights.append(f"{summary['quarterly_followups_due']} completed-client follow-ups are scheduled for this month to validate usage and capture issues.")
    return insights[:6]


def build_quarterly_followup_plan(projects: pd.DataFrame, report_date: pd.Timestamp) -> pd.DataFrame:
    completed = projects[projects["Status"].eq("Completed")].copy()
    columns = [
        "Followup Due This Month",
        "Followup Quarter",
        "Followup Month",
        "Suggested Followup Date",
        "Project Code",
        "Project Name",
        "Status",
        "Completed Date",
        "Client Interaction Date",
        "Current Usage Status",
        "Challenges / Pain Points",
        "Support Required",
        "Scope of UpSelling",
        "PM Followup Notes",
        "Next Followup Date",
    ]
    if completed.empty:
        return pd.DataFrame(columns=columns)

    quarter = pd.Period(report_date, freq="Q")
    quarter_months = list(range(quarter.start_time.month, quarter.end_time.month + 1))
    rows: list[dict[str, Any]] = []
    for _, row in completed.sort_values(["Project Code", "Project Name"]).iterrows():
        slot = stable_followup_slot(row.get("Project Code"))
        followup_month = quarter_months[slot]
        followup_date = pd.Timestamp(year=report_date.year, month=followup_month, day=15)
        rows.append(
            {
                "Followup Due This Month": "Yes" if followup_month == report_date.month else "No",
                "Followup Quarter": f"Q{quarter.quarter} {report_date.year}",
                "Followup Month": followup_date.strftime("%B"),
                "Suggested Followup Date": followup_date,
                "Project Code": row.get("Project Code"),
                "Project Name": row.get("Project Name"),
                "Status": row.get("Status"),
                "Completed Date": row.get("End Date"),
                "Client Interaction Date": "",
                "Current Usage Status": "",
                "Challenges / Pain Points": "",
                "Support Required": "",
                "Scope of UpSelling": "",
                "PM Followup Notes": "Check if the client is using the system properly, whether any challenge is open, and whether support/optimization is needed.",
                "Next Followup Date": "",
            }
        )
    return pd.DataFrame(rows, columns=columns)


def stable_followup_slot(value: Any) -> int:
    text = str(value)
    digits = "".join(char for char in text if char.isdigit())
    if digits:
        return int(digits) % 3
    return sum(ord(char) for char in text) % 3 if text else 0


def replace_existing_output(path: Path) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        raise PermissionError(f"Please close the existing output file before rerunning: {path}") from exc


def create_workbook(data: ClientReportData, output_path: Path) -> Path:
    wb = Workbook()
    wb.remove(wb.active)
    for sheet in ["Dashboard", "Overall Summary", "Year on Year Analysis", "Monthly Metrics", "Latest Month", "Projects", "Quarterly Followup"]:
        wb.create_sheet(sheet)

    styles = make_styles()
    write_dashboard(wb["Dashboard"], data, styles)
    write_overall_summary(wb["Overall Summary"], data, styles)
    write_yoy_analysis(wb["Year on Year Analysis"], data, styles)
    write_monthly_metrics(wb["Monthly Metrics"], data, styles)
    write_latest_month(wb["Latest Month"], data, styles)
    write_projects(wb["Projects"], data, styles)
    write_followup(wb["Quarterly Followup"], data, styles)

    add_workbook_charts(wb, data)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    replace_existing_output(output_path)
    wb.save(output_path)
    return output_path


def make_styles() -> dict[str, Any]:
    thin = Side(style="thin", color="BFBFBF")
    return {
        "title_font": Font(name="Arial", size=16, bold=True, color=DARK_TEXT),
        "subtitle_font": Font(name="Arial", size=10, color=DARK_TEXT),
        "section_font": Font(name="Arial", size=12, bold=True, color=DARK_TEXT),
        "header_font": Font(name="Arial", size=11, bold=True, color=WHITE),
        "label_font": Font(name="Arial", size=10, bold=True, color=DARK_TEXT),
        "body_font": Font(name="Calibri", size=11, color=DARK_TEXT),
        "header_fill": PatternFill("solid", fgColor=BLUE),
        "section_fill": PatternFill("solid", fgColor=GREY),
        "yellow_fill": PatternFill("solid", fgColor=LIGHT_YELLOW),
        "green_fill": PatternFill("solid", fgColor=LIGHT_GREEN),
        "red_fill": PatternFill("solid", fgColor=LIGHT_RED),
        "orange_fill": PatternFill("solid", fgColor=LIGHT_ORANGE),
        "blue_fill": PatternFill("solid", fgColor=LIGHT_BLUE),
        "border": Border(left=thin, right=thin, top=thin, bottom=thin),
        "center": Alignment(horizontal="center", vertical="center"),
        "wrap": Alignment(wrap_text=True, vertical="top"),
    }


def title(ws: Any, cell: str, text: str, styles: dict[str, Any]) -> None:
    ws[cell] = text
    ws[cell].font = styles["title_font"]


def subtitle(ws: Any, cell: str, text: str, styles: dict[str, Any]) -> None:
    ws[cell] = text
    ws[cell].font = styles["subtitle_font"]


def section(ws: Any, row: int, text: str, styles: dict[str, Any], cols: int = 1) -> None:
    ws.cell(row, 1).value = text
    for col in range(1, cols + 1):
        cell = ws.cell(row, col)
        cell.fill = styles["section_fill"]
        cell.font = styles["section_font"]


def header_row(ws: Any, row: int, headers: list[str], styles: dict[str, Any]) -> None:
    for col, value in enumerate(headers, start=1):
        cell = ws.cell(row, col)
        cell.value = value
        cell.fill = styles["header_fill"]
        cell.font = styles["header_font"]
        cell.alignment = styles["center"]
        cell.border = styles["border"]


def write_dashboard(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    project_last = len(data.projects) + 1
    title(ws, "A1", f"VASU's Project Tracker - {data.report_month} {data.report_year}", styles)
    subtitle(ws, "A2", f"Generated: {datetime.now():%B %d, %Y}", styles)
    section(ws, 4, "KEY METRICS", styles, cols=2)
    metrics = [
        ("Total Projects", f"=COUNTA(Projects!A2:A{project_last})"),
        ("Completed", f'=COUNTIF(Projects!H2:H{project_last},"Completed")'),
        ("Active", f'=COUNTIF(Projects!H2:H{project_last},"Active")'),
        ("Completion Rate", "=B6/B5"),
        ("Success Rate", f'=(B5-COUNTIF(Projects!H2:H{project_last},"Lost"))/B5'),
        ("Avg Duration (Days)", f"=AVERAGE(Projects!I2:I{project_last})"),
        ("Followups Due This Month", data.summary["quarterly_followups_due"]),
    ]
    for row, (label, formula) in enumerate(metrics, start=5):
        ws.cell(row, 1).value = label
        ws.cell(row, 1).font = styles["label_font"]
        ws.cell(row, 1).fill = styles["section_fill"]
        ws.cell(row, 2).value = formula
        ws.cell(row, 2).number_format = "0.0%" if "Rate" in label else "0"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 18


def write_overall_summary(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    project_last = len(data.projects) + 1
    title(ws, "A1", "VASU's Tracker 2026 - Overall Summary", styles)
    subtitle(ws, "A2", f"As of: {data.report_month} {data.report_year}", styles)
    section(ws, 4, "OVERALL PROJECT COUNTS", styles, cols=4)
    header_row(ws, 5, ["Metric", "Count", "Formula Reference", "Percentage"], styles)
    rows = [
        ("Total Projects", f"=COUNTA(Projects!A2:A{project_last})", None, "=B6/B6"),
        ("Total Allocations", f'=COUNTIF(Projects!B2:B{project_last},"Allocation")', "By Type", "=B7/B6"),
        ("Total Escalations", f'=COUNTIF(Projects!B2:B{project_last},"Escalation")', "By Type", "=B8/B6"),
        (None, None, None, None),
        ("Total Completed", f'=COUNTIF(Projects!H2:H{project_last},"Completed")', "By Status", "=B10/B6"),
        ("Total Active", f'=COUNTIF(Projects!H2:H{project_last},"Active")', "By Status", "=B11/B6"),
        ("Total Yet to Start", f'=COUNTIF(Projects!H2:H{project_last},"Yet To Start")', "By Status", "=B12/B6"),
        ("Total Held Up", f'=COUNTIF(Projects!H2:H{project_last},"Held Up")', "By Status", "=B13/B6"),
        ("Total Lost", f'=COUNTIF(Projects!H2:H{project_last},"Lost")', "By Status", "=B14/B6"),
        (None, None, None, None),
    ]
    for row_idx, values in enumerate(rows, start=6):
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row_idx, col_idx).value = value
        if values[0]:
            ws.cell(row_idx, 1).font = styles["label_font"]
            if row_idx in (7, 8):
                for col in range(1, 5):
                    ws.cell(row_idx, col).fill = styles["yellow_fill"]
            status_fill_by_label = {
                "Total Completed": styles["blue_fill"],
                "Total Active": styles["green_fill"],
                "Total Yet to Start": styles["yellow_fill"],
                "Total Held Up": styles["orange_fill"],
                "Total Lost": styles["red_fill"],
            }
            if values[0] in status_fill_by_label:
                for col in range(1, 5):
                    ws.cell(row_idx, col).fill = status_fill_by_label[values[0]]
    section(ws, 16, "KEY PERFORMANCE INDICATORS", styles, cols=4)
    header_row(ws, 17, ["KPI", "Value", "Description", ""], styles)
    kpis = [
        ("Completion Rate", "=B10/B6", "Projects Completed / Total"),
        ("Success Rate", "=(B6-B14)/B6", "(Total - Lost) / Total"),
        ("Active Project %", "=B11/B6", "Active / Total Projects"),
        ("Avg Duration (Days)", f"=AVERAGE(Projects!I2:I{project_last})", "Avg completion time"),
    ]
    for row_idx, values in enumerate(kpis, start=18):
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row_idx, col_idx).value = value
        ws.cell(row_idx, 1).font = styles["label_font"]

    for row in range(6, 22):
        ws.cell(row, 4).number_format = "0.0%"
    for col, width in zip("ABCD", [28, 18, 28, 16]):
        ws.column_dimensions[col].width = width


def write_yoy_analysis(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    title(ws, "A1", "YEAR ON YEAR ANALYSIS & COMPARISON", styles)
    subtitle(ws, "A2", f"Same-period comparison through {data.report_month} {data.report_year}", styles)

    section(ws, 4, "ANNUAL / YTD SUMMARY", styles, cols=11)
    headers = list(data.yoy_summary.columns)
    header_row(ws, 5, headers, styles)
    for row_idx, row in enumerate(data.yoy_summary.itertuples(index=False), start=6):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = value
            cell.border = styles["border"]
            cell.alignment = styles["center"]
            if headers[col_idx - 1] in {"Completion Rate", "Success Rate"}:
                cell.number_format = "0.0%"

    compare_start = len(data.yoy_summary) + 8
    section(ws, compare_start, "SAME-PERIOD YOY COMPARISON", styles, cols=5)
    compare_headers = list(data.yoy_comparison.columns)
    header_row(ws, compare_start + 1, compare_headers, styles)
    for row_idx, row in enumerate(data.yoy_comparison.itertuples(index=False), start=compare_start + 2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = value
            cell.border = styles["border"]
            cell.alignment = styles["center"] if col_idx > 1 else styles["wrap"]
            if compare_headers[col_idx - 1] == "Change %":
                cell.number_format = "0.0%"
                cell.fill = styles["green_fill"] if isinstance(value, (int, float)) and value >= 0 else styles["red_fill"]
            if compare_headers[col_idx - 1] == "Change":
                cell.fill = styles["green_fill"] if isinstance(value, (int, float)) and value >= 0 else styles["red_fill"]

    insight_start = compare_start + len(data.yoy_comparison) + 5
    section(ws, insight_start, "MANAGEMENT INSIGHTS", styles, cols=2)
    for idx, insight in enumerate(data.yoy_insights, start=insight_start + 1):
        ws.cell(idx, 1).value = idx - insight_start
        ws.cell(idx, 2).value = insight
        ws.cell(idx, 1).border = styles["border"]
        ws.cell(idx, 2).border = styles["border"]
        ws.cell(idx, 2).alignment = styles["wrap"]

    widths = [12, 18, 14, 14, 16, 14, 12, 12, 16, 16, 14]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_monthly_metrics(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    title(ws, "A1", "Monthly Project Tracking - Opening & Closing Balance", styles)
    subtitle(ws, "A2", "Logic: Closing = Opening + Additions - Completed - Lost - Held", styles)
    headers = ["Month", "Opening", "Add Alloc", "Add Escal", "Total Add", "Completed", "Lost", "Held", "Net Change", "Closing"]
    header_row(ws, 4, headers, styles)
    for row_idx, row in enumerate(data.monthly_metrics.itertuples(index=False), start=5):
        values = list(row)
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row_idx, col_idx).value = value
            ws.cell(row_idx, col_idx).alignment = styles["center"]
            ws.cell(row_idx, col_idx).border = styles["border"]
        ws.cell(row_idx, 9).value = f"=E{row_idx}-F{row_idx}-G{row_idx}-H{row_idx}"
    total_row = len(data.monthly_metrics) + 5
    ws.cell(total_row, 1).value = "TOTAL"
    ws.cell(total_row, 1).font = styles["label_font"]
    for col_idx in range(3, 10):
        letter = get_column_letter(col_idx)
        ws.cell(total_row, col_idx).value = f"=SUM({letter}5:{letter}{total_row-1})"
        ws.cell(total_row, col_idx).font = styles["label_font"]
    for col in range(1, 11):
        ws.cell(total_row, col).fill = styles["section_fill"]
        ws.cell(total_row, col).border = styles["border"]
        ws.column_dimensions[get_column_letter(col)].width = 14


def write_latest_month(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    latest = data.summary
    title(ws, "A1", f"Latest Month Performance - {data.report_period}", styles)
    section(ws, 3, "MONTH SUMMARY", styles, cols=3)
    header_row(ws, 4, ["Metric", "Value", "Details"], styles)
    rows = [
        ("Month", f"'{data.report_period}", ""),
        ("Opening Balance", latest["latest_opening"], "Projects at start of month"),
        ("New Allocations", int(data.latest["Add Alloc"]), "New allocation projects"),
        ("New Escalations", int(data.latest["Add Escal"]), "New escalation projects"),
        ("Total Additions", latest["latest_additions"], "Total new projects"),
        ("Completed", latest["latest_completed"], "Projects completed this month"),
        ("Lost", latest["latest_lost"], "Projects lost this month"),
        ("Held Up", latest["latest_held"], "Projects held this month"),
        ("Net Change", int(data.latest["Net Change"]), "Add - Comp - Lost - Held"),
        ("Closing Balance", latest["latest_closing"], "Projects at end of month"),
    ]
    for row_idx, values in enumerate(rows, start=5):
        for col_idx, value in enumerate(values, start=1):
            ws.cell(row_idx, col_idx).value = value
            ws.cell(row_idx, col_idx).border = styles["border"]
        ws.cell(row_idx, 1).font = styles["label_font"]
        if values[0] == "Completed":
            fill = styles["blue_fill"]
        elif values[0] == "Lost":
            fill = styles["red_fill"]
        elif values[0] == "Held Up":
            fill = styles["orange_fill"]
        else:
            fill = None
        if fill:
            for col_idx in range(1, 4):
                ws.cell(row_idx, col_idx).fill = fill
    for col, width in zip("ABC", [24, 16, 36]):
        ws.column_dimensions[col].width = width


def write_projects(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    headers = ["Project Code", "Project Type", "Project Name", "Received From", "Start Date", "End Date", "Priority", "Status", "Duration"]
    header_row(ws, 1, headers, styles)
    for row_idx, row in enumerate(data.projects.itertuples(index=False), start=2):
        values = list(row)
        for col_idx, value in enumerate(values, start=1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = None if pd.isna(value) else value
            cell.border = styles["border"]
            cell.alignment = styles["wrap"] if col_idx == 3 else styles["center"]
        ws.cell(row_idx, 9).value = f'=IF(F{row_idx}="","",F{row_idx}-E{row_idx})'
        ws.cell(row_idx, 9).number_format = "0"
        status_fill = {
            "Completed": styles["blue_fill"],
            "Active": styles["green_fill"],
            "Yet To Start": styles["yellow_fill"],
            "Held Up": styles["orange_fill"],
            "Lost": styles["red_fill"],
        }.get(str(row.Status))
        if status_fill:
            ws.cell(row_idx, 8).fill = status_fill
    for col_idx in [5, 6]:
        for row_idx in range(2, len(data.projects) + 2):
            ws.cell(row_idx, col_idx).number_format = "DD-MMM-YYYY"
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:I{len(data.projects)+1}"
    widths = [14, 16, 55, 18, 14, 14, 12, 14, 12]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def write_followup(ws: Any, data: ClientReportData, styles: dict[str, Any]) -> None:
    headers = list(data.followup.columns)
    header_row(ws, 1, headers, styles)
    for row_idx, row in enumerate(data.followup.itertuples(index=False), start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row_idx, col_idx)
            cell.value = value
            cell.border = styles["border"]
            cell.alignment = styles["wrap"] if col_idx in {6, 11, 12, 14} else styles["center"]
            if col_idx in {4, 8, 9, 15}:
                cell.number_format = "DD-MMM-YYYY"
        if row[0] == "Yes":
            for col_idx in range(1, len(headers) + 1):
                ws.cell(row_idx, col_idx).fill = styles["yellow_fill"]
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(data.followup)+1}"
    widths = [18, 16, 16, 18, 14, 55, 14, 18, 22, 22, 26, 20, 20, 55, 18]
    for col_idx, width in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def add_workbook_charts(wb: Workbook, data: ClientReportData) -> None:
    ws = wb["Dashboard"]
    summary = wb["Overall Summary"]
    metrics = wb["Monthly Metrics"]

    doughnut = DoughnutChart()
    labels = Reference(summary, min_col=1, min_row=10, max_row=14)
    data_ref = Reference(summary, min_col=2, min_row=10, max_row=14)
    doughnut.add_data(data_ref, titles_from_data=False)
    doughnut.set_categories(labels)
    apply_openpyxl_series_colors(doughnut, ["Completed", "Active", "Yet To Start", "Held Up", "Lost"])
    doughnut.title = "Status Breakdown"
    doughnut.height = 7
    doughnut.width = 9
    ws.add_chart(doughnut, "D4")

    line = LineChart()
    line.title = "Monthly Closing Balance"
    line.y_axis.title = "Projects"
    line.x_axis.title = "Month"
    rows = len(data.monthly_metrics) + 4
    line.add_data(Reference(metrics, min_col=10, min_row=4, max_row=rows), titles_from_data=True)
    line.set_categories(Reference(metrics, min_col=1, min_row=5, max_row=rows))
    if line.series:
        line.series[0].graphicalProperties.line.solidFill = BLUE
    line.height = 7
    line.width = 14
    ws.add_chart(line, "D19")

    bar = BarChart()
    bar.title = "Project Type Mix"
    bar.add_data(Reference(summary, min_col=2, min_row=7, max_row=8), titles_from_data=False)
    bar.set_categories(Reference(summary, min_col=1, min_row=7, max_row=8))
    if bar.series:
        bar.series[0].graphicalProperties.solidFill = BLUE
    bar.height = 7
    bar.width = 9
    ws.add_chart(bar, "D11")


def apply_openpyxl_series_colors(chart: Any, categories: list[str]) -> None:
    if not chart.series:
        return
    points = []
    for idx, category in enumerate(categories):
        point = DataPoint(idx=idx)
        point.graphicalProperties.solidFill = clean_hex(THEME["status_colors"].get(category, THEME["primary"]))
        points.append(point)
    chart.series[0].data_points = points


def create_presentation(data: ClientReportData, output_path: Path) -> Path:
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    add_ppt_title(prs, data)
    add_ppt_executive_summary(prs, data)
    add_ppt_yoy_analysis(prs, data)
    add_ppt_status_breakdown(prs, data)
    add_ppt_monthly_trends(prs, data)
    add_ppt_project_type(prs, data)
    add_ppt_latest_month(prs, data)
    add_ppt_management_insights(prs, data)
    add_ppt_recommendations(prs, data)
    add_ppt_thank_you(prs, data)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    replace_existing_output(output_path)
    prs.save(output_path)
    return output_path


def slide(prs: Presentation, title_text: str | None = None) -> Any:
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.background.fill.solid()
    s.background.fill.fore_color.rgb = PPT_WHITE
    accent = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(10), Inches(0.12))
    accent.fill.solid()
    accent.fill.fore_color.rgb = PPT_BLUE
    accent.line.color.rgb = PPT_BLUE
    if title_text:
        box = s.shapes.add_textbox(Inches(0.45), Inches(0.35), Inches(9), Inches(0.45))
        p = box.text_frame.paragraphs[0]
        p.text = title_text
        p.font.size = Pt(24)
        p.font.bold = True
        p.font.color.rgb = PPT_BLUE
    return s


def add_text(s: Any, text: str, x: float, y: float, w: float, h: float, size: int = 12, color: RGBColor = PPT_DARK, bold: bool = False, align: PP_ALIGN | None = None) -> Any:
    box = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.color.rgb = color
    p.font.bold = bold
    if align:
        p.alignment = align
    return box


def add_stat_card(s: Any, x: float, y: float, w: float, h: float, value: str, label: str, color: RGBColor = PPT_BLUE) -> None:
    shape = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    shape.fill.solid()
    shape.fill.fore_color.rgb = PPT_LIGHT
    shape.line.color.rgb = RGBColor(215, 222, 232)
    marker = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.08), Inches(h))
    marker.fill.solid()
    marker.fill.fore_color.rgb = color
    marker.line.color.rgb = color
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = value
    p.alignment = PP_ALIGN.CENTER
    p.font.size = Pt(22)
    p.font.bold = True
    p.font.color.rgb = PPT_BLUE
    p2 = tf.add_paragraph()
    p2.text = label
    p2.alignment = PP_ALIGN.CENTER
    p2.font.size = Pt(9)
    p2.font.color.rgb = PPT_MUTED


def apply_ppt_series_colors(chart: Any, categories: list[str]) -> None:
    if not chart.series:
        return
    series = chart.series[0]
    for idx, category in enumerate(categories):
        try:
            point = series.points[idx]
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = ppt_rgb(THEME["status_colors"].get(category, THEME["primary"]))
        except Exception:
            continue


def apply_ppt_line_colors(chart: Any, colors: list[RGBColor]) -> None:
    for idx, color in enumerate(colors):
        if idx >= len(chart.series):
            break
        try:
            chart.series[idx].format.line.color.rgb = color
        except Exception:
            continue


def add_bullets(s: Any, title_text: str, bullets: list[str], x: float, y: float, w: float, h: float, color: RGBColor = PPT_BLUE) -> None:
    panel = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    panel.fill.solid()
    panel.fill.fore_color.rgb = PPT_LIGHT
    panel.line.color.rgb = RGBColor(220, 226, 235)
    bar = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(0.08), Inches(h))
    bar.fill.solid()
    bar.fill.fore_color.rgb = color
    bar.line.color.rgb = color
    add_text(s, title_text, x + 0.18, y + 0.1, w - 0.3, 0.25, 12, PPT_BLUE, True)
    body = s.shapes.add_textbox(Inches(x + 0.22), Inches(y + 0.45), Inches(w - 0.35), Inches(h - 0.55))
    tf = body.text_frame
    tf.clear()
    for idx, bullet in enumerate(bullets):
        p = tf.paragraphs[0] if idx == 0 else tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(9.5)
        p.font.color.rgb = PPT_DARK
        p.space_after = Pt(4)


def add_ppt_title(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs)
    band = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0.12), Inches(10), Inches(1.3))
    band.fill.solid()
    band.fill.fore_color.rgb = PPT_BLUE
    band.line.color.rgb = PPT_BLUE
    add_text(s, "PROJECT TRACKER 2026", 0.6, 1.95, 8.8, 0.6, 34, PPT_BLUE, True, PP_ALIGN.CENTER)
    add_text(s, "Management Review & Performance Analysis", 0.8, 2.65, 8.4, 0.35, 16, PPT_MUTED, False, PP_ALIGN.CENTER)
    add_text(s, f"{data.report_month} {data.report_year}", 0.8, 3.1, 8.4, 0.35, 18, PPT_BLUE, True, PP_ALIGN.CENTER)
    add_text(s, "Prepared for Monthly Business Review", 0.8, 4.35, 8.4, 0.28, 10, PPT_MUTED, False, PP_ALIGN.CENTER)


def add_ppt_executive_summary(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "EXECUTIVE SUMMARY")
    summary = data.summary
    cards = [
        (summary["total"], "Total Projects\nSince Jan 2025", PPT_BLUE),
        (f"{summary['completion_rate']:.1%}", "Completion Rate\n{} Completed".format(summary["completed"]), PPT_COMPLETED),
        (summary["open_pipeline"], "Current Pipeline\nActive + Yet to Start", PPT_GREEN),
        (summary["lost"], "Lost Projects\n{:.1%} Loss Rate".format(summary["loss_rate"]), PPT_RED),
    ]
    for idx, (value, label, color) in enumerate(cards):
        add_stat_card(s, 0.65 + idx * 2.35, 1.05, 1.9, 1.1, str(value), label, color)
    bullets = [
        f"Total tracked project base is {summary['total']} as of {data.report_month} {data.report_year}.",
        f"Completion rate is {summary['completion_rate']:.1%}, with {summary['completed']} completed projects.",
        f"Current active pipeline is {summary['open_pipeline']} projects; {summary['held_up']} projects are held up.",
        f"{data.report_month} closing balance is {summary['latest_closing']} after {summary['latest_additions']} additions and {summary['latest_completed']} completions.",
    ]
    add_bullets(s, "KEY HIGHLIGHTS", bullets, 0.75, 2.8, 8.5, 1.8, PPT_BLUE)


def add_ppt_yoy_analysis(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "YEAR ON YEAR ANALYSIS")
    comparison = data.yoy_comparison.set_index("Metric") if not data.yoy_comparison.empty else pd.DataFrame()
    prev_col = next((col for col in data.yoy_comparison.columns if "Same Period" in col and not str(col).startswith(str(data.report_year))), "")
    curr_col = next((col for col in data.yoy_comparison.columns if str(data.report_year) in str(col)), "")

    def metric_value(metric: str, column: str) -> float:
        if comparison.empty or metric not in comparison.index or not column:
            return 0
        value = comparison.loc[metric, column]
        return float(value) if pd.notna(value) else 0

    cards = [
        ("Additions", int(metric_value("Total Additions", curr_col)), metric_value("Total Additions", "Change %"), PPT_BLUE),
        ("Completed", int(metric_value("Completed", curr_col)), metric_value("Completed", "Change %"), PPT_COMPLETED),
        ("Lost", int(metric_value("Lost", curr_col)), metric_value("Lost", "Change %"), PPT_RED),
        ("Open Pipeline", int(metric_value("Open Pipeline", curr_col)), metric_value("Open Pipeline", "Change %"), PPT_ORANGE),
    ]
    for idx, (label, value, change, color) in enumerate(cards):
        add_stat_card(s, 0.45 + idx * 2.35, 0.95, 1.85, 1.05, str(value), f"{label}\nYoY {change:+.1%}", color)

    chart_data = CategoryChartData()
    chart_data.categories = ["Additions", "Completed", "Lost", "Held Up"]
    chart_data.add_series(str(prev_col).replace(" Same Period", ""), [
        metric_value("Total Additions", prev_col),
        metric_value("Completed", prev_col),
        metric_value("Lost", prev_col),
        metric_value("Held Up", prev_col),
    ])
    chart_data.add_series(str(curr_col).replace(" Same Period", ""), [
        metric_value("Total Additions", curr_col),
        metric_value("Completed", curr_col),
        metric_value("Lost", curr_col),
        metric_value("Held Up", curr_col),
    ])
    chart = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(0.65), Inches(2.25), Inches(5.6), Inches(2.65), chart_data).chart
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.chart_title.text_frame.text = "Same-Period Performance"
    if len(chart.series) >= 2:
        chart.series[0].format.fill.solid()
        chart.series[0].format.fill.fore_color.rgb = PPT_MUTED
        chart.series[1].format.fill.solid()
        chart.series[1].format.fill.fore_color.rgb = PPT_BLUE

    add_bullets(s, "MANAGEMENT READING", data.yoy_insights[:5], 6.55, 2.1, 3.0, 2.9, PPT_BLUE)


def add_ppt_status_breakdown(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "PROJECT STATUS BREAKDOWN")
    summary = data.summary
    chart_data = CategoryChartData()
    chart_data.categories = ["Completed", "Active", "Yet To Start", "Held Up", "Lost"]
    chart_data.add_series(
        "Status",
        [summary["completed"], summary["active"], summary["yet_to_start"], summary["held_up"], summary["lost"]],
    )
    chart = s.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, Inches(0.75), Inches(1.05), Inches(4.2), Inches(3.7), chart_data).chart
    apply_ppt_series_colors(chart, ["Completed", "Active", "Yet To Start", "Held Up", "Lost"])
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.RIGHT
    chart.chart_title.text_frame.text = "Current Status Mix"
    add_bullets(
        s,
        "STATUS READING",
        [
            f"{summary['completed']} projects completed and {summary['lost']} lost.",
            f"{summary['active']} active projects require delivery follow-through.",
            f"{summary['yet_to_start']} projects are yet to start and should be planned for {summary['next_month']}.",
            f"{summary['held_up']} held-up projects need management intervention.",
        ],
        5.35,
        1.2,
        3.95,
        3.1,
        PPT_ORANGE,
    )


def add_ppt_monthly_trends(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "MONTHLY PERFORMANCE TRENDS")
    trend = data.monthly_metrics.tail(12)
    chart_data = CategoryChartData()
    chart_data.categories = list(trend["Month"])
    chart_data.add_series("Additions", list(trend["Total Add"]))
    chart_data.add_series("Completed", list(trend["Completed"]))
    chart_data.add_series("Closing", list(trend["Closing"]))
    chart = s.shapes.add_chart(XL_CHART_TYPE.LINE_MARKERS, Inches(0.55), Inches(1.0), Inches(6.0), Inches(3.65), chart_data).chart
    apply_ppt_line_colors(chart, [PPT_ORANGE, PPT_COMPLETED, PPT_BLUE])
    chart.has_legend = True
    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
    chart.chart_title.text_frame.text = "Additions, Completions & Closing Pipeline"
    add_bullets(
        s,
        "KEY INSIGHTS",
        [
            f"{data.report_month} added {data.summary['latest_additions']} projects against {data.summary['latest_completed']} completions.",
            f"Closing balance moved from {data.summary['latest_opening']} to {data.summary['latest_closing']}.",
            f"{data.summary['next_month']} delivery planning should focus on quick starts, owners, and ageing items.",
            "Sustained completions from older projects continue to reduce long-running backlog.",
        ],
        6.85,
        1.05,
        2.75,
        3.6,
        PPT_BLUE,
    )


def add_ppt_project_type(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "PROJECT TYPE ANALYSIS")
    summary = data.summary
    total = summary["total"]
    add_type_block(s, "ALLOCATION", summary["allocation"], summary["completed"], total, 0.9, PPT_GREEN)
    add_type_block(s, "ESCALATION", summary["escalation"], summary["completed"], total, 5.25, PPT_BLUE)
    chart_data = CategoryChartData()
    chart_data.categories = ["Allocation", "Escalation"]
    chart_data.add_series("Projects", [summary["allocation"], summary["escalation"]])
    chart = s.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, Inches(2.25), Inches(3.55), Inches(5.4), Inches(1.45), chart_data).chart
    chart.has_legend = False
    chart.chart_title.text_frame.text = "Portfolio Mix"


def add_type_block(s: Any, title_text: str, count: int, completed: int, total: int, x: float, color: RGBColor) -> None:
    panel = s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(1.25), Inches(3.8), Inches(1.8))
    panel.fill.solid()
    panel.fill.fore_color.rgb = PPT_LIGHT
    panel.line.color.rgb = RGBColor(220, 226, 235)
    add_text(s, title_text, x + 0.25, 1.45, 3.3, 0.3, 16, color, True, PP_ALIGN.CENTER)
    pct = count / total if total else 0
    add_text(s, f"Total: {count} ({pct:.1%})", x + 0.25, 1.95, 3.3, 0.28, 13, PPT_DARK, True, PP_ALIGN.CENTER)
    add_text(s, f"Portfolio contribution in current project base", x + 0.25, 2.35, 3.3, 0.28, 9.5, PPT_MUTED, False, PP_ALIGN.CENTER)


def add_ppt_latest_month(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, f"{data.report_month.upper()} {data.report_year} PERFORMANCE SUMMARY")
    summary = data.summary
    cards = [
        ("Total Projects", summary["total"], PPT_BLUE),
        ("Completed", summary["completed"], PPT_COMPLETED),
        ("Active", summary["active"], PPT_GREEN),
        ("Yet To Start", summary["yet_to_start"], PPT_YELLOW),
        ("Held Up", summary["held_up"], PPT_ORANGE),
    ]
    for idx, (label, value, color) in enumerate(cards):
        add_stat_card(s, 0.45 + idx * 1.9, 1.0, 1.55, 0.95, str(value), label, color)
    add_bullets(
        s,
        "CURRENT STATUS",
        [
            f"Opening balance: {summary['latest_opening']}",
            f"New additions: {summary['latest_additions']} ({int(data.latest['Add Alloc'])} allocation, {int(data.latest['Add Escal'])} escalation)",
            f"Completed in {data.report_month}: {summary['latest_completed']}",
            f"Held up in {data.report_month}: {summary['latest_held']}",
            f"Closing balance: {summary['latest_closing']}",
        ],
        0.8,
        2.45,
        8.4,
        2.1,
        PPT_BLUE,
    )


def add_ppt_management_insights(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "CRITICAL INSIGHTS & OUTLOOK")
    summary = data.summary
    pipeline_notes = [
        f"Open pipeline is {summary['open_pipeline']} projects: {summary['active']} active and {summary['yet_to_start']} yet to start.",
        f"{summary['held_up']} held-up project(s) need blocker review and revised action plan.",
        f"{summary['quarterly_followups_due']} completed-client follow-ups are planned for {data.report_month}.",
    ]
    delivery_notes = [
        f"Completion rate is {summary['completion_rate']:.1%}; completed base is {summary['completed']} projects.",
        "Use quarterly follow-up outcomes to identify usage issues, support gaps, and upsell opportunities.",
        "Weekly project updates should be sent from the mail-merge workbook after PM review.",
    ]
    action_notes = [
        "Prioritize ageing active projects with clear owner/date/action.",
        "Keep alias review updated so SCH remarks map to the correct tracker project.",
        f"Prepare {summary['next_month']} start plan for yet-to-start projects.",
    ]
    add_bullets(s, "Pipeline Health", pipeline_notes, 0.55, 1.05, 2.85, 3.8, PPT_ORANGE)
    add_bullets(s, "Client Success", delivery_notes, 3.6, 1.05, 2.85, 3.8, PPT_COMPLETED)
    add_bullets(s, "PM Actions", action_notes, 6.65, 1.05, 2.85, 3.8, PPT_BLUE)


def add_ppt_recommendations(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs, "RECOMMENDATIONS & NEXT STEPS")
    summary = data.summary
    next_month = summary["next_month"]
    add_bullets(
        s,
        "Maintain Delivery Excellence",
        [
            f"{summary['completed']} completed projects show strong delivery base.",
            f"{summary['quarterly_followups_due']} completed-client follow-ups are planned for {data.report_month}.",
            "Continue weekly review of ageing active projects.",
            "Use quarterly follow-ups to confirm usage, challenges, support needs, and upsell scope.",
        ],
        0.55,
        1.0,
        2.85,
        3.6,
        PPT_GREEN,
    )
    add_bullets(
        s,
        "Pipeline Development Focus",
        [
            f"Current open pipeline is {summary['open_pipeline']} projects.",
            f"Plan {next_month} starts for yet-to-start projects.",
            "Assign owners and dates before month-end review.",
        ],
        3.6,
        1.0,
        2.85,
        3.6,
        PPT_ORANGE,
    )
    add_bullets(
        s,
        "Resolve Held-Up Items",
        [
            f"{summary['held_up']} projects are held up and need escalation.",
            "Confirm blockers, client dependency, and revised closure plan.",
            "Track held-up projects separately until closure.",
        ],
        6.65,
        1.0,
        2.85,
        3.6,
        PPT_RED,
    )


def add_ppt_thank_you(prs: Presentation, data: ClientReportData) -> None:
    s = slide(prs)
    add_text(s, "THANK YOU", 0.8, 1.8, 8.4, 0.55, 34, PPT_BLUE, True, PP_ALIGN.CENTER)
    add_text(s, "Questions & Discussion", 0.8, 2.55, 8.4, 0.35, 18, PPT_MUTED, False, PP_ALIGN.CENTER)
    add_text(s, f"{data.report_month} {data.report_year} Project Tracker Review", 0.8, 3.35, 8.4, 0.3, 11, PPT_MUTED, False, PP_ALIGN.CENTER)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Create VASU tracker monthly Excel dashboard and PPT.")
    parser.add_argument("--source-file", default=str(SOURCE_FILE), help="Tracker workbook path.")
    parser.add_argument("--as-of-date", default=None, help="Optional report date in YYYY-MM-DD format. Defaults to today.")
    parser.add_argument("--report-month", default=None, help="Optional report month in YYYY-MM format, for example 2026-06.")
    return parser.parse_args()


def parse_as_of_date(value: str | None) -> pd.Timestamp | None:
    if not value:
        return None
    return pd.Timestamp(datetime.strptime(value, "%Y-%m-%d").date())


def resolve_report_date(args: argparse.Namespace, config: dict[str, Any]) -> pd.Timestamp:
    if args.report_month:
        return parse_report_month(args.report_month)
    if args.as_of_date:
        return parse_as_of_date(args.as_of_date) or pd.Timestamp(datetime.now().date())
    monthly = config.get("monthly_report", {})
    if monthly.get("period_mode") == "manual" and monthly.get("selected_report_month"):
        return parse_report_month(str(monthly["selected_report_month"]))
    return pd.Timestamp(datetime.now().date())


def parse_report_month(value: str) -> pd.Timestamp:
    return pd.Timestamp(datetime.strptime(value, "%Y-%m").date())


def main() -> None:
    args = parse_args()
    config = load_app_config()
    data = build_report_data(Path(args.source_file), resolve_report_date(args, config))
    output_root = Path(config.get("monthly_report", {}).get("output_root", str(REPORTS_ROOT)))
    output_dir = output_root / data.report_month
    output_dir.mkdir(parents=True, exist_ok=True)
    workbook_path = output_dir / f"VASU_Tracker_YoY_{data.report_month}{data.report_year}.xlsx"
    ppt_path = output_dir / f"VASU_Tracker_YoY_{data.report_month}{data.report_year}_Presentation.pptx"
    create_workbook(data, workbook_path)
    create_presentation(data, ppt_path)
    print(f"WORKBOOK={workbook_path}")
    print(f"PRESENTATION={ppt_path}")
    print(f"TOTAL={data.summary['total']}")
    print(f"COMPLETED={data.summary['completed']}")
    print(f"ACTIVE={data.summary['active']}")
    print(f"YET_TO_START={data.summary['yet_to_start']}")
    print(f"HELD_UP={data.summary['held_up']}")
    print(f"LOST={data.summary['lost']}")
    print(f"REPORT_MONTH={data.report_month} {data.report_year}")
    print(f"LATEST_ADDITIONS={data.summary['latest_additions']}")
    print(f"LATEST_COMPLETED={data.summary['latest_completed']}")
    print(f"LATEST_CLOSING={data.summary['latest_closing']}")


if __name__ == "__main__":
    main()
