"""Build the monthly RPP / DuitNow volume report (OUTPUT_NAME) from the 01_Results folder."""

from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from decimal import Decimal
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from .mapping import (AMOUNT_CELL, CALCULATIONS, SOURCES, TOTAL_LABELS,
                      VOLUME_CELL, VOLUME_COLUMNS, Source)
from .hints import apply_hints, remove_hints
from .reader import cell_value, read_sheet, to_number

log = logging.getLogger("txn_report")

TEMPLATE = Path(__file__).parent / "templates" / "TXN_count_RPP_and_Duitnow_TEMPLATE.xlsx"
OUTPUT_NAME = "amb_txn_count_rpp_{period}.xlsx"
EXTENSIONS = (".xls", ".xlsx")

MONTHS = {
    "JAN": "JANUARY", "FEB": "FEBRUARY", "MAR": "MARCH", "APR": "APRIL",
    "MAY": "MAY", "JUN": "JUNE", "JUL": "JULY", "AUG": "AUGUST",
    "SEP": "SEPTEMBER", "OCT": "OCTOBER", "NOV": "NOVEMBER", "DEC": "DECEMBER",
}
_PERIOD_RE = re.compile(r"_(%s)_(\d{4})_(?:ARCH|LIV)$" % "|".join(MONTHS), re.I)

# An average transaction above this amount is flagged as suspicious.
DEFAULT_MAX_AVG_AMOUNT = 1_000_000


class ReportError(Exception):
    """Raised when the report cannot be built (bad folder, missing files...)."""


@dataclass
class ReportResult:
    period: str
    values: pd.DataFrame          # one row per mapped input file (Decimal values)
    totals: pd.DataFrame          # rows 17-22, calculated by the tool
    cells: dict[str, Decimal] = field(default_factory=dict)  # every number written
    warnings: list[str] = field(default_factory=list)
    output: Path | None = None


# --------------------------------------------------------------------------- #
# Period handling
# --------------------------------------------------------------------------- #
def normalise_period(text: str) -> str:
    """Accept 'OCT_2026', 'oct-2026', 'OCT 2026', 'October 2026', '2026-10'."""
    t = text.strip().upper().replace("-", "_").replace(" ", "_")
    m = re.fullmatch(r"(\d{4})_(\d{1,2})", t)
    if m:
        year, month = m.group(1), int(m.group(2))
        if not 1 <= month <= 12:
            raise ReportError(f"Invalid month in {text!r}")
        return f"{list(MONTHS)[month - 1]}_{year}"
    m = re.fullmatch(r"([A-Z]+)_(\d{4})", t)
    if m and m.group(1)[:3] in MONTHS:
        return f"{m.group(1)[:3]}_{m.group(2)}"
    raise ReportError(f"Cannot understand month {text!r}; use e.g. OCT_2026")


def detect_period(root: Path) -> str:
    """Find the data month from the file names under the mapped folders."""
    found: set[str] = set()
    for folder in {(s.db_folder, s.product) for s in SOURCES}:
        directory = root.joinpath(*folder)
        if not directory.is_dir():
            continue
        for f in directory.iterdir():
            if f.suffix.lower() in EXTENSIONS:
                m = _PERIOD_RE.search(f.stem)
                if m:
                    found.add(f"{m.group(1).upper()}_{m.group(2)}")
    if not found:
        raise ReportError(
            f"No files named like '*_<MON>_<YYYY>_ARCH.xls' found under {root}. "
            "Is this the 01_Results folder?"
        )
    if len(found) > 1:
        raise ReportError(
            f"Files for more than one month found: {sorted(found)}. "
            "Choose one with --month."
        )
    period = found.pop()
    log.info("Data month detected from file names: %s", period)
    return period


def sheet_title(period: str) -> str:
    mon, year = period.split("_")
    return f"{MONTHS[mon]}_{year}"


# --------------------------------------------------------------------------- #
# Reading inputs
# --------------------------------------------------------------------------- #
def find_file(root: Path, source: Source, period: str) -> Path | None:
    """Locate a mapped file; the name match ignores case and accepts .xls/.xlsx."""
    directory = root / source.db_folder / source.product
    if not directory.is_dir():
        return None
    wanted = source.name_for(period).lower()
    for f in directory.iterdir():
        if f.is_file() and f.suffix.lower() in EXTENSIONS and f.stem.lower() == wanted:
            return f
    return None


def collect_values(root: str | Path, period: str) -> tuple[pd.DataFrame, list[str]]:
    """Read volume (A2) and amount (B2) from every mapped file, as exact Decimals."""
    root = Path(root)
    rows, warnings = [], []

    def problem(msg: str) -> None:
        warnings.append(msg)
        log.error(msg)

    for s in SOURCES:
        path = find_file(root, s, period)
        volume = amount = None
        status = "ok"
        if path is None:
            status = "missing"
            problem(f"MISSING: {s.db_folder}/{s.product}/{s.name_for(period)}.xls "
                    f"(cells {s.volume_cell}/{s.amount_cell} left blank)")
        else:
            log.debug("Reading %s", path)
            try:
                grid = read_sheet(path)
                raw_volume = cell_value(grid, VOLUME_CELL)
                raw_amount = cell_value(grid, AMOUNT_CELL)
                log.debug("  raw %s=%r %s=%r", VOLUME_CELL, raw_volume, AMOUNT_CELL, raw_amount)
                volume = to_number(raw_volume)
                amount = to_number(raw_amount)
            except Exception as exc:  # report the file, keep going
                status = "error"
                log.debug("Read failure details", exc_info=True)
                problem(f"ERROR reading {path.name}: {exc}")
            else:
                if volume is None:
                    status = "error"
                    problem(f"ERROR: {path.name} has no volume in {VOLUME_CELL}")
                elif amount is None and volume == 0:
                    # No transactions: SQL SUM() returns NULL, so B2 is blank.
                    amount = Decimal(0)
                    log.info("%s: no transactions, blank amount taken as 0", path.name)
                elif amount is None:
                    status = "error"
                    problem(f"ERROR: {path.name} has volume {volume:,} but no amount in {AMOUNT_CELL}")
            if status == "ok":
                log.info("Read %-45s %s=%s  %s=%s", path.name,
                         s.volume_cell, f"{volume:,}", s.amount_cell, f"{amount:,}")
        rows.append({
            "direction": s.direction,
            "section": s.section,
            "line": s.line,
            "file": s.name_for(period),
            "volume_cell": s.volume_cell,
            "amount_cell": s.amount_cell,
            "volume": volume,
            "amount": amount,
            "status": status,
        })
    return pd.DataFrame(rows), warnings


def sanity_checks(values: pd.DataFrame, max_avg_amount: float) -> list[str]:
    """Flag values that look wrong (negative, fractional volumes, huge averages)."""
    warnings = []
    limit = Decimal(str(max_avg_amount))
    for _, r in values.iterrows():
        vol, amt = r["volume"], r["amount"]
        if pd.isna(vol) or pd.isna(amt):
            continue
        label = f"{r['direction']} {r['section']} {r['line']} ({r['amount_cell']})"
        if vol < 0 or amt < 0:
            warnings.append(f"CHECK: negative value for {label}")
        if vol != vol.to_integral_value():
            warnings.append(f"CHECK: volume {vol} is not a whole number for {label}")
        if vol > 0 and amt / vol > limit:
            warnings.append(
                f"CHECK: average amount per transaction is {amt / vol:,.2f} for {label} "
                f"(amount {amt:,.2f} / volume {vol:,})"
            )
        if vol == 0 and amt != 0:
            warnings.append(f"CHECK: amount without volume for {label}")
    for w in warnings:
        log.warning(w)
    return warnings


def calculate(values: pd.DataFrame) -> dict[str, Decimal]:
    """Every report cell (inputs and totals) as an exact Decimal.

    Missing inputs (only possible with --allow-missing) count as 0 in the totals.
    """
    cells: dict[str, Decimal] = {}
    for _, r in values.iterrows():
        cells[r["volume_cell"]] = Decimal(0) if pd.isna(r["volume"]) else r["volume"]
        cells[r["amount_cell"]] = Decimal(0) if pd.isna(r["amount"]) else r["amount"]
    for target, sources in CALCULATIONS:
        cells[target] = sum((cells[c] for c in sources), Decimal(0))
        log.debug("%s = %s = %s", target, " + ".join(sources), f"{cells[target]:,}")
    return cells


def totals_table(cells: dict[str, Decimal]) -> pd.DataFrame:
    """Rows 17-22 as a table for the console and log."""
    rows = []
    for row in range(17, 22):
        rows.append({"line": TOTAL_LABELS[row],
                     "out_volume": cells[f"C{row}"], "out_amount": cells[f"D{row}"],
                     "in_volume": cells[f"F{row}"], "in_amount": cells[f"G{row}"]})
    rows.append({"line": TOTAL_LABELS[22], "out_volume": "", "out_amount": "",
                 "in_volume": cells["F22"], "in_amount": cells["G22"]})
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Writing the report
# --------------------------------------------------------------------------- #
EXCEL_DIGITS = 15  # Excel keeps at most 15 significant digits of any number


def _excel_number(cell: str, value: Decimal):
    """Volumes become whole numbers; amounts become numbers Excel can store."""
    if cell[0] in VOLUME_COLUMNS and value == value.to_integral_value():
        return int(value)
    return float(value)


def write_report(values: pd.DataFrame, cells: dict[str, Decimal], period: str,
                 output: Path, template: Path = TEMPLATE, hints: bool = True) -> Path:
    """Fill the template with the 16 inputs and the calculated totals (no formulas)."""
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    title = sheet_title(period)
    ws.title = title
    ws["B2"] = f"TRANSACTION VOLUME - {title}"
    ws["E18"] = "Total CT SETTLED"  # the client sample had just "T" here

    missing = set(values.loc[values["volume"].isna(), "volume_cell"]) | \
        set(values.loc[values["amount"].isna(), "amount_cell"])
    # The report must not contain formulas: clear any the template still has.
    # Every one of them is a cell the tool calculates itself (CALCULATIONS).
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and c.value.startswith("="):
                if c.coordinate not in cells:
                    log.warning("Template formula in %s (%s) is not a cell the tool "
                                "calculates; it was removed", c.coordinate, c.value)
                c.value = None

    for ref, value in cells.items():
        # A missing input stays blank; its totals still count it as 0.
        ws[ref] = None if ref in missing else _excel_number(ref, value)
        digits = len(value.normalize().as_tuple().digits)
        if ref not in missing and digits > EXCEL_DIGITS:
            log.warning("%s has %d significant digits; Excel can only show %d, so it will "
                        "display %s. Exact value: %s", ref, digits, EXCEL_DIGITS,
                        f"{float(f'{value:.{EXCEL_DIGITS}g}'):,.2f}", f"{value:,}")

    if hints:
        apply_hints(ws, period)
        log.info("Added source notes to the value cells")
    else:
        remove_hints(ws)

    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    log.info("Report written to %s", output)
    return output


def build(root: str | Path, output_dir: str | Path | None = None,
          period: str | None = None, allow_missing: bool = False,
          max_avg_amount: float = DEFAULT_MAX_AVG_AMOUNT,
          output_name: str = OUTPUT_NAME, hints: bool = True) -> ReportResult:
    """Read the 01_Results folder and write the client report."""
    root = Path(root)
    log.info("Input folder: %s", root)
    if not root.is_dir():
        raise ReportError(f"Input folder not found: {root}")
    period = normalise_period(period) if period else detect_period(root)
    log.info("Building report for %s", period)

    values, warnings = collect_values(root, period)
    bad = values[values["status"].isin(["missing", "error"])]
    if len(bad) and not allow_missing:
        raise ReportError(
            "Cannot build the report:\n  " + "\n  ".join(warnings) +
            "\nFix the files, or pass --allow-missing to leave those cells blank."
        )
    if len(bad):
        log.warning("%d input(s) missing; continuing because --allow-missing was given", len(bad))
    warnings += sanity_checks(values, max_avg_amount)

    cells = calculate(values)
    result = ReportResult(period, values, totals_table(cells), cells, warnings)
    out_dir = Path(output_dir) if output_dir else Path.cwd()
    result.output = write_report(values, cells, period,
                                 out_dir / output_name.format(period=period),
                                 hints=hints)
    return result
