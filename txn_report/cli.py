"""Command line: ``python -m txn_report "<path to 01_Results>" [options]``."""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path

import pandas as pd

from .logs import close_logging, setup_logging
from .report import DEFAULT_MAX_AVG_AMOUNT, ReportError, build

log = logging.getLogger("txn_report")


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        prog="txn_report",
        description="Build the monthly RPP / DuitNow transaction volume report "
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
    p.add_argument("--no-hints", action="store_true",
                   help="Leave out the DATA SOURCES box under the table")
    p.add_argument("--log-dir",
                   help="Folder for log files (default: 'logs' inside the output folder)")
    p.add_argument("--log-level", default="DEBUG", choices=["DEBUG", "INFO", "WARNING"],
                   help="Detail level of the log file (default: %(default)s)")
    p.add_argument("-q", "--quiet", action="store_true",
                   help="Console shows only warnings, errors and the output path")
    return p.parse_args(argv)


def _count(v) -> str:
    return "" if v is None or v == "" or pd.isna(v) else f"{v:,.0f}"


def _money(v) -> str:
    return "" if v is None or v == "" or pd.isna(v) else f"{v:,.2f}"


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    out_dir = Path(args.output_dir) if args.output_dir else Path.cwd()
    log_file = setup_logging(args.log_dir or out_dir / "logs", args.log_level,
                             "WARNING" if args.quiet else "INFO")
    log.info("Run started: %s", " ".join(sys.argv))
    log.info("Log file: %s", log_file)
    try:
        result = build(args.results_dir, out_dir, args.month,
                       args.allow_missing, args.max_avg_amount,
                       hints=not args.no_hints)
    except ReportError as exc:
        log.error("Report NOT created: %s", exc)
        print(f"See log: {log_file}", file=sys.stderr)
        close_logging()
        return 1
    except Exception:
        log.exception("Unexpected error; report NOT created")
        print(f"See log: {log_file}", file=sys.stderr)
        close_logging()
        return 2

    inputs = result.values.drop(columns=["volume_cell"]).to_string(
        index=False, formatters={"volume": _count, "amount": _money})
    totals = result.totals.to_string(
        index=False, na_rep="",
        formatters={"out_volume": _count, "in_volume": _count,
                    "out_amount": _money, "in_amount": _money})
    log.info("Inputs for %s:\n%s", result.period, inputs)
    log.info("Totals calculated by the tool:\n%s", totals)
    log.info("Finished with %d warning(s)", len(result.warnings))
    print(f"Report written to {result.output}")
    print(f"Log written to {log_file}")
    close_logging()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
