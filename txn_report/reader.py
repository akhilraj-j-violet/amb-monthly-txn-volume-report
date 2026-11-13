"""Read single cells (A2 / B2) from the extract files with pandas.

The extracts are named ``.xls`` but database export tools often write other
formats under that extension, so the real format is detected from the file's
first bytes: real Excel 97-2003 (xlrd), Excel 2007+ (openpyxl), an HTML table,
or delimited text.
"""

from __future__ import annotations

import io
import re
from decimal import Decimal, InvalidOperation
from html.parser import HTMLParser
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
        return _html_grid(text)
    return pd.read_csv(io.StringIO(text), header=None, sep=None, engine="python")


class _TableParser(HTMLParser):
    """Collect the first <table> as rows of cell text; <th> and <td> both count
    as cells, so the header row is Excel row 1, as when Excel opens the file."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.rows: list[list[str]] = []
        self._depth = 0
        self._done = False
        self._cell: list[str] | None = None

    def handle_starttag(self, tag, attrs):
        if self._done:
            return
        if tag == "table":
            self._depth += 1
        elif self._depth == 1 and tag == "tr":
            self.rows.append([])
        elif self._depth == 1 and tag in ("td", "th") and self.rows:
            self._cell = []

    def handle_endtag(self, tag):
        if self._done:
            return
        if tag == "table":
            self._depth -= 1
            self._done = self._depth == 0
        elif tag in ("td", "th") and self._cell is not None:
            self.rows[-1].append(" ".join("".join(self._cell).split()))
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def _html_grid(text: str) -> pd.DataFrame:
    """Grid from an HTML table such as a SQL*Plus ``SET MARKUP HTML`` export."""
    parser = _TableParser()
    parser.feed(text)
    rows = [r for r in parser.rows if r]
    width = max((len(r) for r in rows), default=0)
    grid = pd.DataFrame([r + [None] * (width - len(r)) for r in rows])
    return grid.replace("", None)


def cell_value(grid: pd.DataFrame, ref: str):
    """Value at an Excel reference such as ``"A2"`` (None if outside the data)."""
    col_letters, row = coordinate_from_string(ref)
    r, c = row - 1, column_index_from_string(col_letters) - 1
    if r >= grid.shape[0] or c >= grid.shape[1]:
        return None
    value = grid.iat[r, c]
    return None if pd.isna(value) else value


def to_number(value) -> Decimal | None:
    """Convert a cell value to an exact Decimal.

    Text such as ``"45881698693822.41"`` is converted digit for digit (no float
    rounding). Commas, spaces and currency text are stripped first.
    """
    if value is None:
        return None
    if isinstance(value, bool):
        raise ValueError(f"unexpected boolean cell value {value!r}")
    if isinstance(value, int):
        return Decimal(value)
    if isinstance(value, float):
        # repr() gives the shortest text that round-trips, e.g. 99.1 -> "99.1".
        return Decimal(repr(value))
    cleaned = re.sub(r"[^0-9.\-]", "", str(value))
    if cleaned in {"", "-", ".", "-."}:
        return None
    try:
        return Decimal(cleaned)
    except InvalidOperation:
        raise ValueError(f"not a number: {value!r}") from None
