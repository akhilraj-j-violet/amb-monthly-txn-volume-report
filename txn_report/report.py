"""Build TXN_count_RPP_and_Duitnow_<MON>_<YYYY>_CLIENT.xlsx from the 01_Results folder."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

from .mapping import AMOUNT_CELL, SOURCES, TOTAL_ROWS, VOLUME_CELL, Source
from .reader import cell_value, read_sheet, to_number

TEMPLATE = Path(__file__).parent / "templates" / "TXN_count_RPP_and_Duitnow_TEMPLATE.xlsx"
OUTPUT_NAME = "TXN_count_RPP_and_Duitnow_{period}_CLIENT.xlsx"
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
    values: pd.DataFrame          # one row per mapped input file
    totals: pd.DataFrame          # the formula rows, recomputed in pandas
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
    return found.pop()


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
    """Read volume (A2) and amount (B2) from every mapped file."""
    root = Path(root)
    rows, warnings = [], []
    for s in SOURCES:
        path = find_file(root, s, period)
        volume = amount = None
        status = "ok"
        if path is None:
            status = "missing"
            warnings.append(
                f"MISSING: {s.db_folder}/{s.product}/{s.name_for(period)}.xls "
                f"(cells {s.volume_cell}/{s.amount_cell} left blank)"
            )
        else:
            try:
                grid = read_sheet(path)
                volume = to_number(cell_value(grid, VOLUME_CELL))
                amount = to_number(cell_value(grid, AMOUNT_CELL))
            except Exception as exc:  # report the file, keep going
                status = "error"
                warnings.append(f"ERROR reading {path.name}: {exc}")
            else:
                if volume is None:
                    status = "error"
                    warnings.append(f"ERROR: {path.name} has no volume in {VOLUME_CELL}")
                elif amount is None and volume == 0:
                    # No transactions: SQL SUM() returns NULL, so B2 is blank.
                    amount = 0.0
                elif amount is None:
                    status = "error"
                    warnings.append(
                        f"ERROR: {path.name} has volume {volume:,.0f} but no amount in {AMOUNT_CELL}"
                    )
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
    for _, r in values.dropna(subset=["volume", "amount"]).iterrows():
        label = f"{r['direction']} {r['section']} {r['line']} ({r['amount_cell']})"
        if r["volume"] < 0 or r["amount"] < 0:
            warnings.append(f"CHECK: negative value for {label}")
        if r["volume"] != int(r["volume"]):
            warnings.append(f"CHECK: volume {r['volume']} is not a whole number for {label}")
        if r["volume"] > 0 and r["amount"] / r["volume"] > max_avg_amount:
            warnings.append(
                f"CHECK: average amount per transaction is "
                f"{r['amount'] / r['volume']:,.2f} for {label} "
                f"(amount {r['amount']:,.2f} / volume {int(r['volume']):,})"
            )
        if r["volume"] == 0 and r["amount"] != 0:
            warnings.append(f"CHECK: amount without volume for {label}")
    return warnings


def compute_totals(values: pd.DataFrame) -> pd.DataFrame:
    """Recompute the template's formula rows (17-22) with pandas for display."""
    by_cell = {}
    for _, r in values.iterrows():
        by_cell[r["volume_cell"]] = r["volume"] or 0.0
        by_cell[r["amount_cell"]] = r["amount"] or 0.0

    labels = {17: "Total AE SETTLED", 18: "Total CT SETTLED",
              19: "Total AE NOT SETTLED", 20: "Total CT NOT SETTLED"}
    rows = []
    for row, (duit, qr) in TOTAL_ROWS.items():
        rows.append({
            "line": labels[row],
            "out_volume": by_cell[f"C{duit}"] + by_cell[f"C{qr}"],
            "out_amount": by_cell[f"D{duit}"] + by_cell[f"D{qr}"],
            "in_volume": by_cell[f"F{duit}"] + by_cell[f"F{qr}"],
            "in_amount": by_cell[f"G{duit}"] + by_cell[f"G{qr}"],
        })
    totals = pd.DataFrame(rows)
    sums = totals.drop(columns="line").sum()
    totals.loc[len(totals)] = {"line": "Total Outgoing / Incoming", **sums}
    totals.loc[len(totals)] = {
        "line": "Grand Total (Outgoing + Incoming)",
        "out_volume": float("nan"), "out_amount": float("nan"),
        "in_volume": sums["out_volume"] + sums["in_volume"],
        "in_amount": sums["out_amount"] + sums["in_amount"],
    }
    return totals


# --------------------------------------------------------------------------- #
# Writing the report
# --------------------------------------------------------------------------- #
def write_report(values: pd.DataFrame, period: str, output: Path,
                 template: Path = TEMPLATE) -> Path:
    """Fill the template's 16 input cells; rows 17-22 keep their formulas."""
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    title = sheet_title(period)
    ws.title = title
    ws["B2"] = f"TRANSACTION VOLUME - {title}"
    ws["E18"] = "Total CT SETTLED"  # the client sample had just "T" here

    for _, r in values.iterrows():
        if r["volume"] is not None and not pd.isna(r["volume"]):
            v = r["volume"]
            ws[r["volume_cell"]] = int(v) if v == int(v) else v
        if r["amount"] is not None and not pd.isna(r["amount"]):
            ws[r["amount_cell"]] = r["amount"]

    # Make Excel recalculate the totals when the file is opened.
    wb.calculation.fullCalcOnLoad = True
    output.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output)
    return output


def build(root: str | Path, output_dir: str | Path | None = None,
          period: str | None = None, allow_missing: bool = False,
          max_avg_amount: float = DEFAULT_MAX_AVG_AMOUNT) -> ReportResult:
    """Read the 01_Results folder and write the client report."""
    root = Path(root)
    if not root.is_dir():
        raise ReportError(f"Input folder not found: {root}")
    period = normalise_period(period) if period else detect_period(root)

    values, warnings = collect_values(root, period)
    bad = values[values["status"].isin(["missing", "error"])]
    if len(bad) and not allow_missing:
        raise ReportError(
            "Cannot build the report:\n  " + "\n  ".join(warnings) +
            "\nFix the files, or pass --allow-missing to leave those cells blank."
        )
    warnings += sanity_checks(values, max_avg_amount)

    result = ReportResult(period, values, compute_totals(values), warnings)
    out_dir = Path(output_dir) if output_dir else Path.cwd()
    result.output = write_report(values, period, out_dir / OUTPUT_NAME.format(period=period))
    return result
