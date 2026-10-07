from __future__ import annotations

from pathlib import Path
from typing import Any

import pandas as pd
from xlsxwriter.utility import xl_col_to_name

from reporting_engine import AnalysisResult, flatten_insights, infer_business_columns


SHEET_RAW = "Raw_Data"
SHEET_CLEAN = "Cleaned_Data"
SHEET_DASHBOARD = "Dashboard"
SHEET_KPI = "KPI_Summary"
SHEET_TREND = "Trend_Analysis"
SHEET_CHARTS = "Charts"
SHEET_EXEC = "Executive_Summary"


def create_excel_report(result: AnalysisResult, output_path: Path, config: dict[str, Any]) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        datetime_format="dd-mmm-yyyy",
        date_format="dd-mmm-yyyy",
    ) as writer:
        result.source.raw_df.to_excel(writer, sheet_name=SHEET_RAW, index=False)
        result.cleaned_df.to_excel(writer, sheet_name=SHEET_CLEAN, index=False)

        workbook = writer.book
        formats = build_formats(workbook, config)

        for sheet_name in [SHEET_DASHBOARD, SHEET_KPI, SHEET_TREND, SHEET_CHARTS, SHEET_EXEC]:
            writer.sheets[sheet_name] = workbook.add_worksheet(sheet_name)

        setup_data_sheet(writer, SHEET_RAW, result.source.raw_df, formats)
        setup_data_sheet(writer, SHEET_CLEAN, result.cleaned_df, formats)
        write_kpi_summary(writer, result, formats)
        write_trend_analysis(writer, result, formats)
        write_charts_sheet(writer, result, formats)
        write_dashboard(writer, result, formats, config)
        write_executive_summary(writer, result, formats, config)

        for worksheet in writer.sheets.values():
            worksheet.hide_gridlines(2)

    return output_path


def build_formats(workbook: Any, config: dict[str, Any]) -> dict[str, Any]:
    currency_symbol = config.get("currency_symbol", "INR")
    theme = resolved_theme(config)
    status_colors = theme["status_colors"]
    risk_colors = theme["risk_colors"]
    return {
        "title": workbook.add_format(
            {"bold": True, "font_size": 18, "font_color": theme["primary"], "align": "left", "valign": "vcenter"}
        ),
        "subtitle": workbook.add_format({"font_size": 10, "font_color": theme["muted"]}),
        "header": workbook.add_format(
            {
                "bold": True,
                "font_color": "white",
                "bg_color": theme["primary"],
                "border": 1,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "section": workbook.add_format(
            {"bold": True, "font_color": "white", "bg_color": theme["secondary"], "border": 1, "align": "left"}
        ),
        "label": workbook.add_format({"bold": True, "font_color": theme["primary"], "border": 1, "bg_color": status_colors["Completed"]}),
        "text": workbook.add_format({"border": 1, "text_wrap": True, "valign": "top"}),
        "number": workbook.add_format({"border": 1, "num_format": "#,##0.00"}),
        "integer": workbook.add_format({"border": 1, "num_format": "#,##0"}),
        "currency": workbook.add_format({"border": 1, "num_format": f'"{currency_symbol}" #,##0'}),
        "percent": workbook.add_format({"border": 1, "num_format": "0.00%"}),
        "date": workbook.add_format({"border": 1, "num_format": "dd-mmm-yyyy"}),
        "good": workbook.add_format({"bg_color": status_colors["Active"], "font_color": "#FFFFFF"}),
        "bad": workbook.add_format({"bg_color": status_colors["Lost"], "font_color": "#FFFFFF"}),
        "warning": workbook.add_format({"bg_color": status_colors["Yet To Start"], "font_color": "#1F1F1F"}),
        "card_title": workbook.add_format(
            {"bold": True, "font_color": "white", "bg_color": theme["primary"], "align": "center", "border": 1}
        ),
        "card_value": workbook.add_format(
            {
                "bold": True,
                "font_size": 18,
                "font_color": theme["primary"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bg_color": theme["panel"],
            }
        ),
        "card_percent": workbook.add_format(
            {
                "bold": True,
                "font_size": 18,
                "font_color": theme["primary"],
                "align": "center",
                "valign": "vcenter",
                "border": 1,
                "bg_color": theme["panel"],
                "num_format": "0.00%",
            }
        ),
        "formula": workbook.add_format({"border": 1, "font_color": theme["muted"], "text_wrap": True}),
        "theme": theme,
    }


def resolved_theme(config: dict[str, Any]) -> dict[str, Any]:
    default = {
        "primary": "#1F4E78",
        "secondary": "#305496",
        "panel": "#F3F6FA",
        "muted": "#606060",
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
    incoming = config.get("theme", {})
    theme = {**default, **{key: value for key, value in incoming.items() if key not in {"status_colors", "risk_colors"}}}
    theme["status_colors"] = {**default["status_colors"], **incoming.get("status_colors", {})}
    theme["risk_colors"] = {**default["risk_colors"], **incoming.get("risk_colors", {})}
    return theme


def setup_data_sheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    worksheet = writer.sheets[sheet_name]
    worksheet.freeze_panes(1, 0)
    if df.empty:
        return

    for col_num, column_name in enumerate(df.columns):
        worksheet.write(0, col_num, column_name, formats["header"])
    columns = [{"header": str(column)} for column in df.columns]
    worksheet.add_table(0, 0, len(df), len(df.columns) - 1, {"columns": columns, "style": "Table Style Medium 2"})
    autosize_columns(worksheet, df)
    apply_column_formats(worksheet, df, formats)


def autosize_columns(worksheet: Any, df: pd.DataFrame, min_width: int = 10, max_width: int = 35) -> None:
    for col_num, column in enumerate(df.columns):
        values = df[column].astype("string").fillna("")
        width = min(max(max(values.map(len).max() if len(values) else 0, len(str(column))) + 2, min_width), max_width)
        worksheet.set_column(col_num, col_num, width)


def apply_column_formats(worksheet: Any, df: pd.DataFrame, formats: dict[str, Any]) -> None:
    for col_num, column in enumerate(df.columns):
        lowered = column.lower()
        if pd.api.types.is_datetime64_any_dtype(df[column]) or "date" in lowered:
            worksheet.set_column(col_num, col_num, 13, formats["date"])
        elif "rate" in lowered or "%" in lowered or lowered.startswith("is "):
            worksheet.set_column(col_num, col_num, 12, formats["percent"] if "rate" in lowered or "%" in lowered else None)
        elif any(word in lowered for word in ["revenue", "cost", "profit", "amount", "billing", "value"]):
            worksheet.set_column(col_num, col_num, 14, formats["currency"])
        elif pd.api.types.is_numeric_dtype(df[column]):
            worksheet.set_column(col_num, col_num, 12, formats["number"])


def write_kpi_summary(writer: pd.ExcelWriter, result: AnalysisResult, formats: dict[str, Any]) -> None:
    workbook = writer.book
    worksheet = writer.sheets.get(SHEET_KPI) or workbook.add_worksheet(SHEET_KPI)
    writer.sheets[SHEET_KPI] = worksheet
    worksheet.freeze_panes(1, 0)
    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 18)
    worksheet.set_column("C:C", 42)
    worksheet.set_column("D:D", 14)
    worksheet.set_column("E:E", 60)

    headers = ["KPI", "Value", "Formula / Source", "Format", "Definition"]
    for col, header in enumerate(headers):
        worksheet.write(0, col, header, formats["header"])

    df = result.cleaned_df
    kpi_values = dict(zip(result.kpi_summary["KPI"], result.kpi_summary["Value"]))
    columns = infer_business_columns(df)
    status_range = sheet_col_range(df, "Status Group")
    project_range = sheet_col_range(df, columns.get("project_code") or df.columns[0])
    cycle_range = sheet_col_range(df, "Cycle Days")
    sla_range = sheet_col_range(df, "SLA Met")
    age_range = sheet_col_range(df, "Ageing Risk")
    employee_col = columns.get("employee_count")
    employee_range = sheet_col_range(df, employee_col) if employee_col else None
    priority_range = sheet_col_range(df, columns.get("priority")) if columns.get("priority") else None

    trend_last_row = len(result.monthly_trend) + 1
    trend_prev_row = max(trend_last_row - 1, 2)

    formulas: list[tuple[str, str | None, Any, str, str]] = [
        ("Total Projects", f"=COUNTA({project_range})", kpi_values.get("Total Projects"), "Count", "Overall project volume"),
        ("Completed Projects", f'=COUNTIFS({status_range},"Completed")', kpi_values.get("Completed Projects"), "Count", "Projects marked completed"),
        ("Active/Open Projects", f'=COUNTIFS({status_range},"Active")+COUNTIFS({status_range},"Not Started")+COUNTIFS({status_range},"Held Up")+COUNTIFS({status_range},"Other")', kpi_values.get("Active/Open Projects"), "Count", "Projects not completed or lost"),
        ("Lost Projects", f'=COUNTIFS({status_range},"Lost")', kpi_values.get("Lost Projects"), "Count", "Projects marked lost"),
        ("Held Up Projects", f'=COUNTIFS({status_range},"Held Up")', kpi_values.get("Held Up Projects"), "Count", "Blocked/on-hold projects"),
        ("Completion Rate", "=IFERROR(B3/B2,0)", kpi_values.get("Completion Rate"), "Percent", "Completed projects divided by total projects"),
        ("Loss Rate", "=IFERROR(B5/B2,0)", kpi_values.get("Loss Rate"), "Percent", "Lost projects divided by total projects"),
        ("Conversion Rate", "=IFERROR(B3/(B3+B5),0)", kpi_values.get("Conversion Rate"), "Percent", "Completed divided by completed plus lost"),
        ("SLA Met Rate", f'=IFERROR(COUNTIFS({sla_range},TRUE)/COUNTIFS({cycle_range},">=0"),0)', kpi_values.get("SLA Met Rate"), "Percent", "Projects within configured SLA"),
        ("Ageing Risk Count", f'=COUNTIFS({age_range},TRUE)', kpi_values.get("Ageing Risk Count"), "Count", "Open projects older than target SLA"),
        ("Average Cycle Days", f"=AVERAGE({cycle_range})", kpi_values.get("Average Cycle Days"), "Number", "Average project cycle days"),
        ("Median Cycle Days", f"=MEDIAN({cycle_range})", kpi_values.get("Median Cycle Days"), "Number", "Median project cycle days"),
    ]

    if employee_range:
        formulas.append(
            (
                "Completed Employee Base",
                f'=SUMIFS({employee_range},{status_range},"Completed")',
                None,
                "Number",
                "Total employee count linked to completed projects",
            )
        )
        formulas.append(
            (
                "Average Employee Count",
                f"=AVERAGE({employee_range})",
                kpi_values.get("Average Employee Count"),
                "Number",
                "Average employee count per populated project row",
            )
        )

    if priority_range:
        formulas.append(
            (
                "High Priority Completed",
                f'=COUNTIFS({priority_range},"High",{status_range},"Completed")',
                None,
                "Count",
                "Completed high-priority projects",
            )
        )

    if len(result.monthly_trend) >= 2:
        formulas.extend(
            [
                (
                    "Latest Month Volume",
                    f"='{SHEET_TREND}'!$B${trend_last_row}",
                    kpi_values.get("Latest Month Volume"),
                    "Count",
                    "Latest activity month project count",
                ),
                (
                    "Previous Month Volume",
                    f"='{SHEET_TREND}'!$B${trend_prev_row}",
                    kpi_values.get("Previous Month Volume"),
                    "Count",
                    "Previous activity month project count",
                ),
                (
                    "MoM Volume Growth %",
                    f"=IFERROR((B{len(formulas)+2}-B{len(formulas)+3})/B{len(formulas)+3},0)",
                    kpi_values.get("MoM Volume Growth %"),
                    "Percent",
                    "Growth percentage versus previous month",
                ),
            ]
        )

    financial = result.financial_columns
    revenue_col = financial.get("revenue")
    cost_col = financial.get("cost")
    profit_col = financial.get("profit")
    if revenue_col:
        formulas.append(("Revenue", f"=SUM({sheet_col_range(df, revenue_col)})", None, "Currency", f"Detected from {revenue_col}"))
    else:
        formulas.append(("Revenue", None, None, "Currency", "Not available in current source data"))
    if cost_col:
        formulas.append(("Cost", f"=SUM({sheet_col_range(df, cost_col)})", None, "Currency", f"Detected from {cost_col}"))
    else:
        formulas.append(("Cost", None, None, "Currency", "Not available in current source data"))
    if profit_col:
        formulas.append(("Profit", f"=SUM({sheet_col_range(df, profit_col)})", None, "Currency", f"Detected from {profit_col}"))
    elif revenue_col and cost_col:
        formulas.append(("Profit", "=B21-B22", None, "Currency", "Computed as revenue minus cost"))
    else:
        formulas.append(("Profit", None, None, "Currency", "Not available in current source data"))
    if revenue_col and (cost_col or profit_col):
        formulas.append(("Profit Margin", "=IFERROR(B23/B21,0)", None, "Percent", "Profit divided by revenue"))

    for row, (kpi, formula, cached_value, fmt_name, definition) in enumerate(formulas, start=1):
        worksheet.write(row, 0, kpi, formats["label"])
        value_format = excel_format_for(fmt_name, formats)
        if formula:
            worksheet.write_formula(row, 1, formula, value_format, cached_formula_value(cached_value, fmt_name))
            worksheet.write(row, 2, formula, formats["formula"])
        else:
            worksheet.write(row, 1, "Not available", formats["text"])
            worksheet.write(row, 2, "Source column not found", formats["formula"])
        worksheet.write(row, 3, fmt_name, formats["text"])
        worksheet.write(row, 4, definition, formats["text"])

    lookup_start = len(formulas) + 3
    write_lookup_examples(worksheet, df, columns, lookup_start, formats)
    apply_kpi_conditional_formats(worksheet, len(formulas), formats)


def cached_formula_value(value: Any, fmt_name: str) -> Any:
    if value is None or pd.isna(value):
        return 0 if fmt_name in {"Count", "Number", "Currency", "Percent"} else ""
    return value


def excel_format_for(fmt_name: str, formats: dict[str, Any]) -> Any:
    if fmt_name == "Percent":
        return formats["percent"]
    if fmt_name == "Currency":
        return formats["currency"]
    if fmt_name == "Count":
        return formats["integer"]
    if fmt_name == "Number":
        return formats["number"]
    return formats["text"]


def sheet_col_range(df: pd.DataFrame, column: str | None, sheet_name: str = SHEET_CLEAN) -> str:
    if not column or column not in df.columns:
        column = df.columns[0]
    col_idx = df.columns.get_loc(column)
    letter = xl_col_to_name(col_idx)
    last_row = max(len(df) + 1, 2)
    return f"'{sheet_name}'!${letter}$2:${letter}${last_row}"


def sheet_cell(df: pd.DataFrame, column: str, row_number: int, sheet_name: str = SHEET_CLEAN) -> str:
    col_idx = df.columns.get_loc(column)
    letter = xl_col_to_name(col_idx)
    return f"'{sheet_name}'!${letter}${row_number}"


def write_lookup_examples(
    worksheet: Any,
    df: pd.DataFrame,
    columns: dict[str, str | None],
    start_row: int,
    formats: dict[str, Any],
) -> None:
    worksheet.write(start_row, 0, "Lookup Formula Examples", formats["section"])
    worksheet.write(start_row, 1, "Value", formats["section"])
    worksheet.write(start_row, 2, "Formula", formats["section"])

    project_code = columns.get("project_code") or df.columns[0]
    project_name = columns.get("project_name") or df.columns[min(1, len(df.columns) - 1)]
    status = "Status Group"
    engineer = columns.get("assigned_engineer")

    first_code = df[project_code].dropna().iloc[0] if df[project_code].notna().any() else ""
    code_range = sheet_col_range(df, project_code)
    name_range = sheet_col_range(df, project_name)
    status_range = sheet_col_range(df, status)

    worksheet.write(start_row + 1, 0, "Project Code Input", formats["label"])
    worksheet.write(start_row + 1, 1, first_code, formats["text"])
    worksheet.write(start_row + 1, 2, "Input cell used by formulas below", formats["formula"])

    xlookup = f'=XLOOKUP(B{start_row + 2},{code_range},{name_range},"Not found")'
    worksheet.write(start_row + 2, 0, "Project Name via XLOOKUP", formats["label"])
    worksheet.write_formula(start_row + 2, 1, xlookup, formats["text"])
    worksheet.write(start_row + 2, 2, xlookup, formats["formula"])

    index_match = f'=INDEX({status_range},MATCH(B{start_row + 2},{code_range},0))'
    worksheet.write(start_row + 3, 0, "Status via INDEX/MATCH", formats["label"])
    worksheet.write_formula(start_row + 3, 1, index_match, formats["text"])
    worksheet.write(start_row + 3, 2, index_match, formats["formula"])

    if engineer and engineer in df.columns:
        code_idx = df.columns.get_loc(project_code)
        engineer_idx = df.columns.get_loc(engineer)
        left_idx = min(code_idx, engineer_idx)
        right_idx = max(code_idx, engineer_idx)
        left_letter = xl_col_to_name(left_idx)
        right_letter = xl_col_to_name(right_idx)
        table_range = f"'{SHEET_CLEAN}'!${left_letter}$2:${right_letter}${len(df) + 1}"
        vlookup_col = abs(engineer_idx - code_idx) + 1
        if code_idx <= engineer_idx:
            vlookup = f'=VLOOKUP(B{start_row + 2},{table_range},{vlookup_col},FALSE)'
        else:
            vlookup = "VLOOKUP requires lookup key to be the left-most table column"
        worksheet.write(start_row + 4, 0, "Engineer via VLOOKUP", formats["label"])
        if vlookup.startswith("="):
            worksheet.write_formula(start_row + 4, 1, vlookup, formats["text"])
        else:
            worksheet.write(start_row + 4, 1, vlookup, formats["text"])
        worksheet.write(start_row + 4, 2, vlookup, formats["formula"])


def apply_kpi_conditional_formats(worksheet: Any, formula_count: int, formats: dict[str, Any]) -> None:
    worksheet.conditional_format(1, 1, formula_count, 1, {"type": "cell", "criteria": ">", "value": 0, "format": formats["good"]})
    worksheet.conditional_format(1, 1, formula_count, 1, {"type": "cell", "criteria": "<", "value": 0, "format": formats["bad"]})
    for row in range(1, formula_count + 1):
        kpi_name = row + 1
        if kpi_name in {6, 7, 8, 9}:
            worksheet.conditional_format(row, 1, row, 1, {"type": "cell", "criteria": ">=", "value": 0.8, "format": formats["good"]})
            worksheet.conditional_format(row, 1, row, 1, {"type": "cell", "criteria": "<", "value": 0.5, "format": formats["bad"]})


def write_trend_analysis(writer: pd.ExcelWriter, result: AnalysisResult, formats: dict[str, Any]) -> None:
    worksheet = writer.sheets.get(SHEET_TREND) or writer.book.add_worksheet(SHEET_TREND)
    writer.sheets[SHEET_TREND] = worksheet
    monthly = result.monthly_trend
    weekly = result.weekly_trend
    outliers = result.outliers if not result.outliers.empty else pd.DataFrame(columns=["Row Number", "Column", "Value", "Reason"])
    quality = result.data_quality

    setup_table_block(worksheet, monthly, 0, 0, formats, "Monthly Trend")
    if not monthly.empty and "MoM Growth %" in monthly.columns:
        mom_col = monthly.columns.get_loc("MoM Growth %")
        total_col = monthly.columns.get_loc("Total Projects")
        total_letter = xl_col_to_name(total_col)
        for excel_row in range(2, len(monthly) + 2):
            if excel_row == 2:
                worksheet.write_number(excel_row - 1, mom_col, 0, formats["percent"])
            else:
                formula = f"=IFERROR(({total_letter}{excel_row}-{total_letter}{excel_row - 1})/{total_letter}{excel_row - 1},0)"
                worksheet.write_formula(excel_row - 1, mom_col, formula, formats["percent"])
        worksheet.conditional_format(1, mom_col, len(monthly), mom_col, {"type": "cell", "criteria": ">=", "value": 0, "format": formats["good"]})
        worksheet.conditional_format(1, mom_col, len(monthly), mom_col, {"type": "cell", "criteria": "<", "value": 0, "format": formats["bad"]})

    weekly_start = len(monthly) + 4
    setup_table_block(worksheet, weekly, weekly_start, 0, formats, "Weekly Trend")
    quality_start = weekly_start + len(weekly) + 4
    setup_table_block(worksheet, quality, quality_start, 0, formats, "Data Quality Issues")
    outlier_start = quality_start + len(quality) + 4
    setup_table_block(worksheet, outliers, outlier_start, 0, formats, "Outlier Detection")
    worksheet.freeze_panes(1, 0)
    worksheet.set_column("A:J", 18)


def setup_table_block(
    worksheet: Any,
    df: pd.DataFrame,
    start_row: int,
    start_col: int,
    formats: dict[str, Any],
    title: str,
) -> None:
    worksheet.write(start_row, start_col, title, formats["section"])
    if df.empty:
        worksheet.write(start_row + 1, start_col, "No records", formats["text"])
        return
    for col_idx, column in enumerate(df.columns):
        worksheet.write(start_row + 1, start_col + col_idx, column, formats["header"])
    for row_idx, row in enumerate(df.itertuples(index=False), start=start_row + 2):
        for col_idx, value in enumerate(row):
            column_name = str(df.columns[col_idx]).lower()
            cell_format = formats["text"]
            if isinstance(value, (int, float)) and not pd.isna(value):
                cell_format = formats["percent"] if "rate" in column_name or "%" in column_name else formats["number"]
            worksheet.write(row_idx, start_col + col_idx, clean_excel_value(value), cell_format)


def clean_excel_value(value: Any) -> Any:
    if pd.isna(value):
        return ""
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def write_charts_sheet(writer: pd.ExcelWriter, result: AnalysisResult, formats: dict[str, Any]) -> None:
    workbook = writer.book
    worksheet = writer.sheets.get(SHEET_CHARTS) or workbook.add_worksheet(SHEET_CHARTS)
    writer.sheets[SHEET_CHARTS] = worksheet
    worksheet.set_column("A:H", 20)
    worksheet.freeze_panes(1, 0)

    row = 0
    for label, table in result.top_performers.items():
        if table.empty:
            continue
        setup_table_block(worksheet, table, row, 0, formats, f"Top {label}")
        row += len(table) + 4

    row = 0
    start_col = 10
    for label, table in result.bottom_performers.items():
        if table.empty:
            continue
        setup_table_block(worksheet, table, row, start_col, formats, f"Bottom {label}")
        row += len(table) + 4

    worksheet.conditional_format(0, 5, 1000, 5, {"type": "3_color_scale"})
    worksheet.conditional_format(0, 15, 1000, 15, {"type": "3_color_scale"})


def write_dashboard(
    writer: pd.ExcelWriter,
    result: AnalysisResult,
    formats: dict[str, Any],
    config: dict[str, Any],
) -> None:
    workbook = writer.book
    worksheet = writer.sheets.get(SHEET_DASHBOARD) or workbook.add_worksheet(SHEET_DASHBOARD)
    writer.sheets[SHEET_DASHBOARD] = worksheet
    worksheet.set_column("A:A", 2)
    worksheet.set_column("B:M", 13)
    worksheet.set_row(0, 30)
    worksheet.merge_range("B1:M1", f"Monthly Business Dashboard - {result.report_month} {result.report_year}", formats["title"])
    worksheet.write("B2", f"{config.get('company_name', '')} | Source: {result.source.file_path.name}", formats["subtitle"])

    card_specs = [
        ("B4:C4", "B5:C6", "Total Projects", "='KPI_Summary'!B2", formats["card_value"]),
        ("D4:E4", "D5:E6", "Completed", "='KPI_Summary'!B3", formats["card_value"]),
        ("F4:G4", "F5:G6", "Completion Rate", "='KPI_Summary'!B7", formats["card_percent"]),
        ("H4:I4", "H5:I6", "SLA Met Rate", "='KPI_Summary'!B10", formats["card_percent"]),
        ("J4:K4", "J5:K6", "Ageing Risk", "='KPI_Summary'!B11", formats["card_value"]),
        ("L4:M4", "L5:M6", "MoM Growth", find_kpi_formula_cell(result, "MoM Volume Growth %", fallback="0"), formats["card_percent"]),
    ]
    for title_range, value_range, title, formula, value_format in card_specs:
        worksheet.merge_range(title_range, title, formats["card_title"])
        worksheet.merge_range(value_range, "", value_format)
        first_cell = value_range.split(":")[0]
        worksheet.write_formula(first_cell, formula, value_format)

    insert_trend_chart(workbook, worksheet, result, "B8", formats)
    insert_status_chart(workbook, worksheet, "H8", formats)
    insert_top_chart(workbook, worksheet, result, "B24", formats)

    worksheet.write("H24", "Management Focus", formats["section"])
    focus_items = result.insights["recommendations"][:5]
    for offset, item in enumerate(focus_items, start=25):
        worksheet.write(f"H{offset}", item, formats["text"])


def find_kpi_formula_cell(result: AnalysisResult, kpi_name: str, fallback: str = "0") -> str:
    kpis = list(result.kpi_summary["KPI"])
    if kpi_name in kpis:
        return f"='KPI_Summary'!B{kpis.index(kpi_name) + 2}"
    return fallback


def insert_trend_chart(workbook: Any, worksheet: Any, result: AnalysisResult, cell: str, formats: dict[str, Any]) -> None:
    if result.monthly_trend.empty:
        return
    chart = workbook.add_chart({"type": "line"})
    theme = formats["theme"]
    rows = len(result.monthly_trend)
    chart.add_series(
        {
            "name": "Total Projects",
            "categories": f"='{SHEET_TREND}'!$A$2:$A${rows + 1}",
            "values": f"='{SHEET_TREND}'!$B$2:$B${rows + 1}",
            "line": {"color": theme["primary"], "width": 2.25},
        }
    )
    chart.add_series(
        {
            "name": "Completed",
            "categories": f"='{SHEET_TREND}'!$A$2:$A${rows + 1}",
            "values": f"='{SHEET_TREND}'!$C$2:$C${rows + 1}",
            "line": {"color": theme["status_colors"].get("Completed", "#D9EAF7"), "width": 2.25},
        }
    )
    chart.set_title({"name": "Monthly Trend"})
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 560, "height": 280})
    worksheet.insert_chart(cell, chart)


def insert_status_chart(workbook: Any, worksheet: Any, cell: str, formats: dict[str, Any]) -> None:
    chart = workbook.add_chart({"type": "doughnut"})
    status_colors = formats["theme"]["status_colors"]
    chart.add_series(
        {
            "name": "Status Mix",
            "categories": f"='{SHEET_KPI}'!$A$3:$A$6",
            "values": f"='{SHEET_KPI}'!$B$3:$B$6",
            "points": [
                {"fill": {"color": status_colors.get("Completed", "#D9EAF7")}},
                {"fill": {"color": status_colors.get("Active", "#70AD47")}},
                {"fill": {"color": status_colors.get("Lost", "#C00000")}},
                {"fill": {"color": status_colors.get("Held Up", "#ED7D31")}},
            ],
        }
    )
    chart.set_title({"name": "Status Mix"})
    chart.set_legend({"position": "bottom"})
    chart.set_size({"width": 360, "height": 280})
    worksheet.insert_chart(cell, chart)


def insert_top_chart(workbook: Any, worksheet: Any, result: AnalysisResult, cell: str, formats: dict[str, Any]) -> None:
    if "Assigned Engineer" not in result.top_performers or result.top_performers["Assigned Engineer"].empty:
        return
    table = result.top_performers["Assigned Engineer"]
    rows = min(len(table), 10)
    chart = workbook.add_chart({"type": "bar"})
    chart.add_series(
        {
            "name": "Completed",
            "categories": f"='{SHEET_CHARTS}'!$A$3:$A${rows + 2}",
            "values": f"='{SHEET_CHARTS}'!$C$3:$C${rows + 2}",
            "fill": {"color": formats["theme"]["status_colors"].get("Active", "#70AD47")},
        }
    )
    chart.set_title({"name": "Top Engineers by Completions"})
    chart.set_legend({"none": True})
    chart.set_size({"width": 560, "height": 300})
    worksheet.insert_chart(cell, chart)


def write_executive_summary(
    writer: pd.ExcelWriter,
    result: AnalysisResult,
    formats: dict[str, Any],
    config: dict[str, Any],
) -> None:
    workbook = writer.book
    worksheet = writer.sheets.get(SHEET_EXEC) or workbook.add_worksheet(SHEET_EXEC)
    writer.sheets[SHEET_EXEC] = worksheet
    worksheet.set_column("A:A", 24)
    worksheet.set_column("B:B", 95)
    worksheet.merge_range("A1:B1", f"Executive Summary - {result.report_month} {result.report_year}", formats["title"])
    worksheet.write("A2", "Company/Team", formats["label"])
    worksheet.write("B2", f"{config.get('company_name', '')} / {config.get('team_name', '')}", formats["text"])
    worksheet.write("A3", "Source File", formats["label"])
    worksheet.write("B3", str(result.source.file_path), formats["text"])

    insights_df = flatten_insights(result.insights)
    setup_table_block(worksheet, insights_df, 5, 0, formats, "Strategic Insights")
    overview_start = len(insights_df) + 9
    setup_table_block(worksheet, result.overview, overview_start, 0, formats, "Dataset Overview")


def apply_positive_negative_conditioning(worksheet: Any, cell_range: str, formats: dict[str, Any]) -> None:
    worksheet.conditional_format(cell_range, {"type": "cell", "criteria": ">=", "value": 0, "format": formats["good"]})
    worksheet.conditional_format(cell_range, {"type": "cell", "criteria": "<", "value": 0, "format": formats["bad"]})
