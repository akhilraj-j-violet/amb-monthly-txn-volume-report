"""The "DATA SOURCES" box under the report table.

A bordered box below the table lists where every number comes from, in the
same style the user wrote the mapping, e.g.::

    AE SETTLED VOLUME - ARCH/RTP1/Outgoing_Settled_DuitAE_ARCH/volume

File names are shown without the month and year, so the box is identical in
the template and in every month's report. The text is generated from
``SOURCES`` and ``CALCULATIONS`` in mapping.py, so it always matches what the
tool actually does.

Layout (outgoing on the left under B:D, incoming on the right under E:G,
like the table above it)::

    24        DATA SOURCES - where each number in the report comes from
    25        OUTGOING - DUITNOW          | INCOMING - DUITNOW
    26-33     8 lines                     | 8 lines
    34        OUTGOING - RPPQR            | INCOMING - RPPQR
    35-42     8 lines                     | 8 lines
    43        OUTGOING - TOTALS           | INCOMING - TOTALS
    44-53     10 lines                    | 10 lines
    54        GRAND TOTAL
    55-56     2 lines
"""

from __future__ import annotations

from openpyxl.styles import Alignment, Border, Font, Side

from .mapping import CALCULATIONS, SOURCES, VOLUME_COLUMNS

FIRST_ROW = 24
LEFT, RIGHT = ("B", "D"), ("E", "G")
FONT_NAME = "Calibri Light"

_LINE = {7: "AE SETTLED", 8: "CT SETTLED", 9: "AE NOT SETTLED", 10: "CT NOT SETTLED",
         13: "AE SETTLED", 14: "CT SETTLED", 15: "AE NOT SETTLED", 16: "CT NOT SETTLED",
         17: "AE SETTLED", 18: "CT SETTLED", 19: "AE NOT SETTLED", 20: "CT NOT SETTLED"}
_SIDE = {"C": "OUTGOING", "D": "OUTGOING", "F": "INCOMING", "G": "INCOMING"}


def _kind(ref: str) -> str:
    return "VOLUME" if ref[0] in VOLUME_COLUMNS else "AMOUNT"


def _name(ref: str) -> str:
    """What a report cell is called in the box, e.g. 'DUITNOW AE SETTLED VOLUME'."""
    col, row = ref[0], int(ref[1:])
    if row <= 10:
        return f"DUITNOW {_LINE[row]} {_kind(ref)}"
    if row <= 16:
        return f"RPPQR {_LINE[row]} {_kind(ref)}"
    if row <= 20:
        return f"TOTAL {_LINE[row]} {_kind(ref)}"
    if row == 21:
        return f"TOTAL {_SIDE[col]} {_kind(ref)}"
    return f"GRAND TOTAL {_kind(ref)}"


def _source_line(ref: str, source, kind: str) -> str:
    db = source.db_folder.split()[0]                       # "ARCH DB" -> "ARCH"
    file_name = source.file_name.replace("_{period}", "")  # no month / year
    suffix = "volume" if kind == "VOLUME" else "Amount"   # as the user writes it
    return f"{_LINE[int(ref[1:])]} {kind} - {db}/{source.product}/{file_name}/{suffix}"


def _calc_line(target: str, sources: tuple[str, ...]) -> str:
    row = int(target[1:])
    if row == 21:
        return f"{_name(target)} - SUM OF THE 4 TOTAL {_kind(target)}S ABOVE"
    if row <= 20:
        # Short form so it fits the box: the two sources are always the same
        # line in the DUITNOW and RPPQR sections.
        assert [_name(s).split(" ", 1)[1] for s in sources] == \
            [f"{_LINE[row]} {_kind(target)}"] * 2, sources
        return f"{_name(target)} - DUITNOW + RPPQR {_LINE[row]} {_kind(target)}"
    return f"{_name(target)} - " + " + ".join(_name(s) for s in sources)


def source_box() -> list[tuple[str | None, str | None, str]]:
    """Rows of the box as (left text, right text, style).

    style is "title", "heading" or "line". A row with right text None spans
    the full width (B:G).
    """
    rows: list[tuple[str | None, str | None, str]] = [
        ("DATA SOURCES - where each number in the report comes from", None, "title")]

    def side_lines(section_rows, side_cols):
        out = []
        for s in SOURCES:
            if int(s.volume_cell[1:]) in section_rows and s.volume_cell[0] in side_cols:
                out.append(_source_line(s.volume_cell, s, "VOLUME"))
                out.append(_source_line(s.amount_cell, s, "AMOUNT"))
        return out

    for title, section_rows in (("DUITNOW", range(7, 11)), ("RPPQR", range(13, 17))):
        rows.append((f"OUTGOING - {title}", f"INCOMING - {title}", "heading"))
        left, right = side_lines(section_rows, "CD"), side_lines(section_rows, "FG")
        rows += [(a, b, "line") for a, b in zip(left, right)]

    rows.append(("OUTGOING - TOTALS (added up by the report tool)",
                 "INCOMING - TOTALS (added up by the report tool)", "heading"))
    left = [_calc_line(t, s) for t, s in CALCULATIONS if t[0] in "CD"]
    right = [_calc_line(t, s) for t, s in CALCULATIONS if t[0] in "FG" and int(t[1:]) <= 21]
    rows += [(a, b, "line") for a, b in zip(left, right)]

    rows.append(("GRAND TOTAL (added up by the report tool)", None, "heading"))
    rows += [(_calc_line(t, s), None, "line") for t, s in CALCULATIONS if int(t[1:]) == 22]
    return rows


def clear_below_table(ws) -> None:
    """Remove anything from row FIRST_ROW down in B:G (old box, old tips, merges)."""
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= FIRST_ROW:
            ws.unmerge_cells(str(rng))
    for row in ws.iter_rows(min_row=FIRST_ROW, max_row=max(ws.max_row, FIRST_ROW),
                            min_col=2, max_col=7):
        for c in row:
            c.value = None
            c.border = Border()
            c.comment = None


def remove_cell_notes(ws) -> None:
    """Older versions put hover notes on the value cells; they are not wanted."""
    for row in ws.iter_rows():
        for c in row:
            if c.comment is not None:   # merged cells never carry one
                c.comment = None


def write_source_box(ws) -> int:
    """Draw the box under the table. Returns the number of rows used."""
    clear_below_table(ws)
    thin, medium = Side(style="thin"), Side(style="medium")
    rows = source_box()
    last = FIRST_ROW + len(rows) - 1
    for i, (left, right, style) in enumerate(rows):
        r = FIRST_ROW + i
        spans = ([(LEFT[0], RIGHT[1], left)] if right is None
                 else [(LEFT[0], LEFT[1], left), (RIGHT[0], RIGHT[1], right)])
        for first, end, text in spans:
            ws.merge_cells(f"{first}{r}:{end}{r}")
            cell = ws[f"{first}{r}"]
            cell.value = text
            cell.font = Font(name=FONT_NAME, size=8 if style == "line" else 9,
                             bold=style != "line")
            cell.alignment = Alignment(horizontal="center" if style == "title" else "left",
                                       vertical="center", indent=0 if style == "title" else 1)
        # Borders: box outline, a line under title/headings, divider between sides.
        for col in "BCDEFG":
            c = ws[f"{col}{r}"]
            c.border = Border(
                left=medium if col == "B" else (thin if col == "E" and right is not None else None),
                right=medium if col == "G" else (thin if col == "D" and right is not None else None),
                top=medium if r == FIRST_ROW else (thin if style != "line" else None),
                bottom=medium if r == last else (thin if style != "line" else None),
            )
    return len(rows)
