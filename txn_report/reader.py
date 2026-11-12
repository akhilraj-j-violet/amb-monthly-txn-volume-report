"""Read single cells (A2 / B2) from the extract files with pandas.

The extracts are named ``.xls`` but database export tools often write other
formats under that extension, so the real format is detected from the file's
first bytes: real Excel 97-2003 (xlrd), Excel 2007+ (openpyxl), an HTML table,
or delimited text.
"""

from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

_OLE2 = b"\xd0\xcf\x11\xe0"   # Excel 97-2003 .xls
_ZIP = b"PK\x03\x04"          # .xlsx


def read_sheet(path: str | Path) -> pd.DataFrame:
    """Return the first sheet as a raw grid (no header row; row 0 = Excel row 1)."""
    path = Path(path)
    head = path.read_bytes()[:512]
    if head.startswith(_OLE2):
        return pd.read_excel(path, header=None, sheet_name=0, engine="xlrd")
    if head.startswith(_ZIP):
        return pd.read_excel(path, header=None, sheet_name=0, engine="openpyxl")
    text = path.read_bytes().decode("utf-8-sig", errors="replace")
    if re.search(r"<\s*table", text, re.I):
        return pd.read_html(io.StringIO(text), header=None)[0]
    return pd.read_csv(io.StringIO(text), header=None, sep=None, engine="python")


def cell_value(grid: pd.DataFrame, ref: str):
    """Value at an Excel reference such as ``"A2"`` (None if outside the data)."""
    col_letters, row = coordinate_from_string(ref)
    r, c = row - 1, column_index_from_string(col_letters) - 1
    if r >= grid.shape[0] or c >= grid.shape[1]:
        return None
    value = grid.iat[r, c]
    return None if pd.isna(value) else value


def to_number(value) -> float | None:
    """Convert a cell value to a number; strips commas, spaces and currency text."""
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = re.sub(r"[^0-9.\-]", "", str(value))
    if cleaned in {"", "-", ".", "-."}:
        return None
    return float(cleaned)
