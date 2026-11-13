"""Results mapping: which input file feeds which cell of the report.

Each input file holds the volume in cell A2 and the amount in cell B2.
``{period}`` in a file name is replaced by the data month, e.g. ``OCT_2026``.
Rows 17-22 of the template are Excel formulas and are not listed here.
"""

from __future__ import annotations

from dataclasses import dataclass

# Cells inside each input file (Excel notation).
VOLUME_CELL = "A2"
AMOUNT_CELL = "B2"


@dataclass(frozen=True)
class Source:
    db_folder: str      # folder directly under 01_Results, e.g. "ARCH DB"
    product: str        # folder under the DB, e.g. "RTP1"
    file_name: str      # with "{period}" placeholder, without extension
    volume_cell: str    # target cell in the report for the volume (from A2)
    amount_cell: str    # target cell in the report for the amount (from B2)
    section: str        # for console/log output
    direction: str
    line: str

    def name_for(self, period: str) -> str:
        return self.file_name.format(period=period)


def _rtp1(direction: str, status: str, kind: str, vol: str, amt: str,
          section: str, line: str) -> Source:
    return Source(
        db_folder="ARCH DB",
        product="RTP1",
        file_name=f"{direction}_{status}_{kind}_{{period}}_ARCH",
        volume_cell=vol,
        amount_cell=amt,
        section=section,
        direction=direction.upper(),
        line=line,
    )


SOURCES: list[Source] = [
    # OUTGOING - DUITNOW
    _rtp1("Outgoing", "Settled",    "DuitAE", "C7",  "D7",  "DUITNOW", "AE SETTLED"),
    _rtp1("Outgoing", "Settled",    "DuitCT", "C8",  "D8",  "DUITNOW", "CT SETTLED"),
    _rtp1("Outgoing", "NonSettled", "DuitAE", "C9",  "D9",  "DUITNOW", "AE NOT SETTLED"),
    _rtp1("Outgoing", "NonSettled", "DuitCT", "C10", "D10", "DUITNOW", "CT NOT SETTLED"),
    # OUTGOING - RPPQR
    _rtp1("Outgoing", "Settled",    "QR_AE",  "C13", "D13", "RPPQR",   "AE SETTLED"),
    _rtp1("Outgoing", "Settled",    "QR_CT",  "C14", "D14", "RPPQR",   "CT SETTLED"),
    _rtp1("Outgoing", "NonSettled", "QR_AE",  "C15", "D15", "RPPQR",   "AE NOT SETTLED"),
    _rtp1("Outgoing", "NonSettled", "QR_CT",  "C16", "D16", "RPPQR",   "CT NOT SETTLED"),
    # INCOMING - DUITNOW
    _rtp1("Incoming", "Settled",    "DuitAE", "F7",  "G7",  "DUITNOW", "AE SETTLED"),
    _rtp1("Incoming", "Settled",    "DuitCT", "F8",  "G8",  "DUITNOW", "CT SETTLED"),
    _rtp1("Incoming", "NonSettled", "DuitAE", "F9",  "G9",  "DUITNOW", "AE NOT SETTLED"),
    _rtp1("Incoming", "NonSettled", "DuitCT", "F10", "G10", "DUITNOW", "CT NOT SETTLED"),
    # INCOMING - RPPQR
    _rtp1("Incoming", "Settled",    "QR_AE",  "F13", "G13", "RPPQR",   "AE SETTLED"),
    _rtp1("Incoming", "Settled",    "QR_CT",  "F14", "G14", "RPPQR",   "CT SETTLED"),
    _rtp1("Incoming", "NonSettled", "QR_AE",  "F15", "G15", "RPPQR",   "AE NOT SETTLED"),
    _rtp1("Incoming", "NonSettled", "QR_CT",  "F16", "G16", "RPPQR",   "CT NOT SETTLED"),
]

# --------------------------------------------------------------------------- #
# Calculated cells (rows 17-22). The tool works these out itself with exact
# decimal arithmetic and writes plain numbers; the report holds no formulas.
# Order matters: later cells use earlier ones.
# --------------------------------------------------------------------------- #
TOTAL_LABELS = {
    17: "Total AE SETTLED", 18: "Total CT SETTLED",
    19: "Total AE NOT SETTLED", 20: "Total CT NOT SETTLED",
    21: "Total Outgoing / Total Incoming",
    22: "Grand Total : Total Outgoing + Total Incoming",
}

CALCULATIONS: list[tuple[str, tuple[str, ...]]] = []
# Rows 17-20: DUITNOW line + RPPQR line (e.g. C17 = C7 + C13).
for _row, (_duit, _qr) in {17: (7, 13), 18: (8, 14), 19: (9, 15), 20: (10, 16)}.items():
    for _col in "CDFG":
        CALCULATIONS.append((f"{_col}{_row}", (f"{_col}{_duit}", f"{_col}{_qr}")))
# Row 21: Total Outgoing (C, D) / Total Incoming (F, G) = rows 17-20.
for _col in "CDFG":
    CALCULATIONS.append((f"{_col}21", tuple(f"{_col}{r}" for r in range(17, 21))))
# Row 22: Grand Total = Total Outgoing + Total Incoming.
CALCULATIONS.append(("F22", ("C21", "F21")))   # volume
CALCULATIONS.append(("G22", ("D21", "G21")))   # amount

VOLUME_COLUMNS = {"C", "F"}
