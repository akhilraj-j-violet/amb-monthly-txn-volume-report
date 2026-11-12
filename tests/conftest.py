import itertools
from pathlib import Path

import pytest
import xlwt

# Values from the OCT 2026 client sample, keyed by report cell of the volume.
SAMPLE = {
    "C7": (5677630, 39848955737.718), "C8": (4255962, 26245866367),
    "C9": (208092, 45881698693822.4), "C10": (6061, 118699235.58),
    "C13": (8437622, 119679040.97), "C14": (7981315, 765318323.33),
    "C15": (140930, 106210811.23), "C16": (13550, 2973519.12),
    "F7": (6448938, 26512169929.46), "F8": (5630252, 28747321380.67),
    "F9": (156343, 209903189.28), "F10": (13285, 85805440.55),
    "F13": (10295200, 366921147.69), "F14": (8683214, 543699645.83),
    "F15": (233990, 5165815.34), "F16": (739, 153452.53),
}


def write_xls(path: Path, volume, amount):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")
    ws.write(0, 0, "COUNT")
    ws.write(0, 1, "AMOUNT")
    ws.write(1, 0, volume)
    ws.write(1, 1, amount)
    wb.save(str(path))


@pytest.fixture
def results_dir(tmp_path):
    """A full 01_Results tree for OCT_2026 with real .xls files."""
    from txn_report.mapping import SOURCES

    root = tmp_path / "01_Results"
    for s in SOURCES:
        folder = root / s.db_folder / s.product
        folder.mkdir(parents=True, exist_ok=True)
        write_xls(folder / f"{s.name_for('OCT_2026')}.xls", *SAMPLE[s.volume_cell])

    # Unmapped files that live in the same tree must be ignored.
    rtp1 = root / "ARCH DB" / "RTP1"
    for d, st, k in itertools.product(["Incoming", "Outgoing"], ["Settled", "NonSettled"],
                                      ["DuitAE_PACS", "DuitAE_PRXY", "DuitCT_PACS", "DuitCT_PRXY"]):
        write_xls(rtp1 / f"{d}_{st}_{k}_OCT_2026_ARCH.xls", 1, 1)
    fpx = root / "LIV DB" / "FPX"
    fpx.mkdir(parents=True)
    write_xls(fpx / "Settled_OCT_2026_LIV.xls", 1, 1)
    return root
