"""Command line: ``python -m txn_report "<path to 01_Results>" [options]``."""

from __future__ import annotations

import argparse
import sys

import pandas as pd

from .report import DEFAULT_MAX_AVG_AMOUNT, ReportError, build


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="txn_report",
        description="Build TXN_count_RPP_and_Duitnow_<MON>_<YYYY>_CLIENT.xlsx "
                    "from the monthly 01_Results folder.",
    )
    p.add_argument("results_dir", help="Path to the 01_Results folder")
    p.add_argument("-o", "--output-dir",
                   help="Folder to save the report in (default: current folder)")
    p.add_argument("-m", "--month",
                   help="Data month, e.g. OCT_2026 (default: read from the file names)")
    p.add_argument("--allow-missing", action="store_true",
                   help="Build the report even if some input files are missing")
    p.add_argument("--max-avg-amount", type=float, default=DEFAULT_MAX_AVG_AMOUNT,
                   help="Warn when amount / volume exceeds this (default: %(default)s)")
    p.add_argument("-q", "--quiet", action="store_true", help="Only print warnings and the output path")
    return p.parse_args(argv)


def _count(v) -> str:
    return "" if v is None or pd.isna(v) else f"{v:,.0f}"


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        result = build(args.results_dir, args.output_dir, args.month,
                       args.allow_missing, args.max_avg_amount)
    except ReportError as exc:
        print(f"error: {exc}", file=sys.stderr)
        return 1

    if not args.quiet:
        with pd.option_context("display.width", 200, "display.max_columns", None,
                               "display.float_format", "{:,.2f}".format):
            print(f"=== Inputs for {result.period} ===")
            print(result.values.drop(columns=["volume_cell"])
                  .to_string(index=False, formatters={"volume": _count}))
            print("\n=== Totals (as Excel will calculate them) ===")
            print(result.totals.to_string(
                index=False, na_rep="",
                formatters={"out_volume": _count, "in_volume": _count}))
            print()
    for w in result.warnings:
        print(f"WARNING {w}", file=sys.stderr)
    print(f"Report written to {result.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
