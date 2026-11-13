"""Logging setup: a timestamped log file per run plus console output."""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path

LOGGER_NAME = "txn_report"
FILE_FORMAT = "%(asctime)s | %(levelname)-7s | %(message)s"


def setup_logging(log_dir: str | Path, file_level: str = "DEBUG",
                  console_level: str = "INFO") -> Path:
    """Send everything to ``<log_dir>/txn_report_<YYYYMMDD_HHMMSS>.log``.

    The file gets every detail (DEBUG by default: raw cell values and each
    calculation); the console gets INFO and above. Returns the log file path.
    """
    log_dir = Path(log_dir)
    log_dir.mkdir(parents=True, exist_ok=True)
    log_file = log_dir / f"txn_report_{datetime.now():%Y%m%d_%H%M%S}.log"

    logger = logging.getLogger(LOGGER_NAME)
    for h in list(logger.handlers):  # avoid duplicate handlers on repeat calls
        logger.removeHandler(h)
        h.close()
    logger.setLevel(logging.DEBUG)
    logger.propagate = False

    fh = logging.FileHandler(log_file, encoding="utf-8")
    fh.setLevel(file_level.upper())
    fh.setFormatter(logging.Formatter(FILE_FORMAT, "%Y-%m-%d %H:%M:%S"))
    logger.addHandler(fh)

    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(console_level.upper())
    ch.setFormatter(logging.Formatter("%(levelname)s %(message)s"))
    logger.addHandler(ch)
    return log_file


def close_logging() -> None:
    logger = logging.getLogger(LOGGER_NAME)
    for h in list(logger.handlers):
        logger.removeHandler(h)
        h.close()
