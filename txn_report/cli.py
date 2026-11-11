"""Command-line entry point: ``python -m txn_report <input> [options]``."""

from __future__ import annotations

import argparse
import sys

from .summary import ReportConfig, build_report, load_transactions
from .writer import render_console, write_csvs, write_excel


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="txn_report",
        description="Build a monthly transaction volume summary report with pandas.",
    )
    p.add_argument("input", help="Transactions file (.csv, .xlsx, .xls, .parquet)")
    p.add_argument("-o", "--output", default="monthly_txn_summary.xlsx",
                   help="Excel output path (default: %(default)s)")
    p.add_argument("--csv-dir", help="Also write one CSV per table into this directory")
    p.add_argument("--sheet", default=0, help="Sheet name/index when input is Excel")
    p.add_argument("--date-col", default="txn_date", help="Date column (default: %(default)s)")
    p.add_argument("--amount-col", default="amount", help="Amount column (default: %(default)s)")
    p.add_argument("--date-format", help="Explicit date format, e.g. %%d/%%m/%%Y")
    p.add_argument("--group-by", nargs="*", default=[], metavar="COL",
                   help="Columns to break down by (e.g. channel txn_type)")
    p.add_argument("--status-col", help="Status column; only rows matching --success-values are kept")
    p.add_argument("--success-values", nargs="+", default=["SUCCESS"], metavar="VAL",
                   help="Status values counted as successful (default: SUCCESS)")
    p.add_argument("--start", help="Include transactions on/after this date (YYYY-MM-DD)")
    p.add_argument("--end", help="Include transactions on/before this date (YYYY-MM-DD)")
    p.add_argument("-q", "--quiet", action="store_true", help="Don't print the summary")
    return p.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    sheet = int(args.sheet) if str(args.sheet).isdigit() else args.sheet
    cfg = ReportConfig(
        date_col=args.date_col,
        amount_col=args.amount_col,
        group_by=args.group_by,
        status_col=args.status_col,
        success_values=args.success_values,
        start=args.start,
        end=args.end,
        date_format=args.date_format,
    )
    try:
        df = load_transactions(args.input, sheet_name=sheet)
        report = build_report(df, cfg)
    except (FileNotFoundError, KeyError, ValueError) as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    xlsx = write_excel(report, args.output)
    if not args.quiet:
        print(render_console(report))
    print(f"Excel report written to {xlsx}")
    if args.csv_dir:
        files = write_csvs(report, args.csv_dir)
        print(f"{len(files)} CSV file(s) written to {args.csv_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
