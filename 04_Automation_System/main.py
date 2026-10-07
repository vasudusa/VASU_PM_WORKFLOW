from __future__ import annotations

import argparse
import logging
import sys
import traceback
from datetime import datetime
from pathlib import Path

from excel_generator import create_excel_report
from ppt_generator import create_powerpoint
from reporting_engine import (
    analyze_source,
    detect_latest_input_file,
    flatten_insights,
    load_config,
    load_source_file,
    setup_logging,
)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Automated monthly business reporting workflow.")
    parser.add_argument("--config", default="config.json", help="Path to config.json.")
    parser.add_argument("--input-file", default=None, help="Optional explicit CSV/XLSX input file override.")
    parser.add_argument("--report-month", default=None, help="Optional report month in YYYY-MM format, for example 2026-06.")
    return parser.parse_args()


def replace_existing_output(path: Path) -> None:
    if not path.exists():
        return
    try:
        path.unlink()
    except PermissionError as exc:
        raise PermissionError(f"Please close the existing output file before rerunning: {path}") from exc


def main() -> int:
    args = parse_args()
    base_dir = Path(__file__).resolve().parent
    config_path = Path(args.config)
    if not config_path.is_absolute():
        config_path = base_dir / config_path

    config = load_config(config_path)
    if args.input_file:
        config["explicit_input_files"] = [str(Path(args.input_file).resolve())]
    if args.report_month:
        config["report_month_mode"] = "manual"
        config["manual_report_month"] = args.report_month

    log_path = setup_logging(Path(config["logs_directory"]))
    logging.info("Monthly reporting automation started")
    logging.info("Config path: %s", config_path)

    try:
        input_file = Path(args.input_file).resolve() if args.input_file else detect_latest_input_file(config)
        source = load_source_file(input_file)
        result = analyze_source(source, config)

        output_root = Path(config["output_directory"]) / str(result.report_year) / result.report_month
        output_root.mkdir(parents=True, exist_ok=True)

        excel_path = output_root / f"Monthly_Business_Report_{result.report_label}.xlsx"
        ppt_path = output_root / f"Monthly_Executive_Review_{result.report_label}.pptx"
        quality_path = output_root / f"Data_Quality_Report_{result.report_label}.csv"
        insights_path = output_root / f"Strategic_Insights_{result.report_label}.csv"

        for output_path in [excel_path, ppt_path, quality_path, insights_path]:
            replace_existing_output(output_path)

        create_excel_report(result, excel_path, config)
        create_powerpoint(result, ppt_path, config)
        result.data_quality.to_csv(quality_path, index=False)
        flatten_insights(result.insights).to_csv(insights_path, index=False)

        logging.info("Excel report created: %s", excel_path)
        logging.info("PowerPoint created: %s", ppt_path)
        logging.info("Data quality report created: %s", quality_path)
        logging.info("Strategic insights created: %s", insights_path)
        logging.info("Monthly reporting automation completed")

        print("REPORT_GENERATION_COMPLETE")
        print(f"INPUT_FILE={input_file}")
        print(f"REPORT_PERIOD={result.report_month} {result.report_year}")
        print(f"EXCEL_REPORT={excel_path}")
        print(f"POWERPOINT_REPORT={ppt_path}")
        print(f"DATA_QUALITY_REPORT={quality_path}")
        print(f"STRATEGIC_INSIGHTS={insights_path}")
        print(f"LOG_FILE={log_path}")
        return 0
    except Exception as exc:
        error_path = Path(config["logs_directory"]) / f"error_report_{datetime.now():%Y%m%d_%H%M%S}.txt"
        logging.exception("Monthly reporting automation failed")
        error_path.write_text(
            "Monthly reporting automation failed\n\n"
            f"Error: {exc}\n\n"
            f"Traceback:\n{traceback.format_exc()}",
            encoding="utf-8",
        )
        print("REPORT_GENERATION_FAILED", file=sys.stderr)
        print(f"ERROR_REPORT={error_path}", file=sys.stderr)
        print(str(exc), file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
