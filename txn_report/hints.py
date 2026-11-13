"""Plain-language notes explaining where every number in the report comes from.

Each value cell gets an Excel note (hover over the cell to read it; cells
with a note show a small red triangle in the corner). A visible line under
the table tells readers this. The same notes are put in the template, using
``<MON>_<YYYY>`` in place of the month.
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.styles import Font

from .mapping import AMOUNT_CELL, CALCULATIONS, SOURCES, VOLUME_CELL

AUTHOR = "Report tool"
HINT_CELLS = ("B24", "B25")   # visible explanation lines under the table
PLACEHOLDER = "<MON>_<YYYY>"

_SIDE = {"C": "Outgoing", "D": "Outgoing", "F": "Incoming", "G": "Incoming"}
_WHAT = {"C": "number of transactions", "D": "total amount",
         "F": "number of transactions", "G": "total amount"}
_SECTION_ROWS = {7: "DuitNow", 8: "DuitNow", 9: "DuitNow", 10: "DuitNow",
                 13: "RPPQR (DuitNow QR)", 14: "RPPQR (DuitNow QR)",
                 15: "RPPQR (DuitNow QR)", 16: "RPPQR (DuitNow QR)"}
_LINE = {7: "AE settled", 8: "CT settled", 9: "AE not settled", 10: "CT not settled",
         13: "AE settled", 14: "CT settled", 15: "AE not settled", 16: "CT not settled",
         17: "Total AE settled", 18: "Total CT settled",
         19: "Total AE not settled", 20: "Total CT not settled"}


def _row(ref: str) -> int:
    return int(ref[1:])


def _short_name(ref: str) -> str:
    """How a cell is called inside an explanation, e.g. 'DuitNow AE settled (C7)'."""
    col, row = ref[0], _row(ref)
    if row in _SECTION_ROWS:
        section = "DuitNow" if row <= 10 else "RPPQR"
        return f"{section} {_LINE[row]} ({ref})"
    if row in (17, 18, 19, 20):
        return f"{_LINE[row]} ({ref})"
    if row == 21:
        return f"Total {_SIDE[col].lower()} ({ref})"
    return ref


def _title(ref: str) -> str:
    col, row = ref[0], _row(ref)
    what = _WHAT[col]
    if row in _SECTION_ROWS:
        return f"{_SIDE[col]} {_SECTION_ROWS[row]}, {_LINE[row]}: {what}."
    if row in (17, 18, 19, 20):
        return f"{_SIDE[col]} {_LINE[row]} (DuitNow + RPPQR): {what}."
    if row == 21:
        return f"Total {_SIDE[col].lower()} (all lines added up): {what}."
    return f"Grand total, outgoing + incoming: {what}."


def hint_texts(period: str | None = None) -> dict[str, str]:
    """Note text for every value cell. ``period`` like 'OCT_2026', or None for the template."""
    month = period or PLACEHOLDER
    notes: dict[str, str] = {}
    for s in SOURCES:
        for ref, src_cell, column in ((s.volume_cell, VOLUME_CELL, "Total Transaction"),
                                      (s.amount_cell, AMOUNT_CELL, "Total Amount")):
            notes[ref] = (
                f"{_title(ref)}\n\n"
                f"Where it comes from:\n"
                f"Folder: 01_Results > {s.db_folder} > {s.product}\n"
                f"File: {s.file_name.format(period=month)}.xls\n"
                f"Value: cell {src_cell} (column '{column}').\n\n"
                f"Copied into this report exactly as it is in the file."
            )
    for target, sources in CALCULATIONS:
        parts = "\n + ".join(_short_name(r) for r in sources)
        notes[target] = (
            f"{_title(target)}\n\n"
            f"How it is worked out:\n   {parts}\n\n"
            f"Added up by the report tool (not an Excel formula)."
        )
    return notes


def apply_hints(ws, period: str | None = None) -> None:
    """Put the notes on the value cells and the explanation lines under the table."""
    for ref, text in hint_texts(period).items():
        lines = text.count("\n") + 1
        longest = max(len(line) for line in text.split("\n"))
        note = Comment(text, AUTHOR)
        note.width = max(260, min(560, longest * 6.5))
        note.height = 20 + lines * 15
        ws[ref].comment = note

    source = (f"the {period} extract files" if period
              else "that month's extract files")
    ws[HINT_CELLS[0]] = ("Tip: hover over any number (cells with a small red triangle) "
                         "to see where it comes from.")
    ws[HINT_CELLS[1]] = (f"Rows 7-16 are copied from {source} in "
                         "01_Results > ARCH DB > RTP1. Rows 17-22 are totals "
                         "added up by the report tool.")
    for ref in HINT_CELLS:
        ws[ref].font = Font(name="Calibri Light", italic=True, size=8, color="595959")


def remove_hints(ws) -> None:
    """Strip notes and explanation lines (used for --no-hints)."""
    for ref in hint_texts():
        ws[ref].comment = None
    for ref in HINT_CELLS:
        ws[ref] = None
