from decimal import Decimal

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from conftest import SAMPLE, write_sqlplus, write_xls
from txn_report import ReportError, build, detect_period
from txn_report.cli import main
from txn_report.reader import read_sheet, cell_value, to_number
from txn_report.report import normalise_period


def test_detect_period(results_dir):
    assert detect_period(results_dir) == "OCT_2026"


def test_full_report(results_dir, tmp_path):
    result = build(results_dir, tmp_path / "out")
    assert result.output.name == "amb_txn_count_rpp_OCT_2026.xlsx"

    ws = load_workbook(result.output).active
    assert ws.title == "OCTOBER_2026"
    assert ws["B2"].value == "TRANSACTION VOLUME - OCTOBER_2026"
    assert ws["E18"].value == "Total CT SETTLED"
    for vol_cell, (vol, amt) in SAMPLE.items():
        amt_cell = chr(ord(vol_cell[0]) + 1) + vol_cell[1:]
        assert ws[vol_cell].value == vol
        assert ws[amt_cell].value == pytest.approx(amt)
    # No formulas anywhere: every cell holds a plain value.
    for row in ws.iter_rows():
        for c in row:
            assert not (isinstance(c.value, str) and c.value.startswith("=")), c.coordinate

    # Totals are exact decimal sums of the source text.
    D = lambda cell: Decimal(str(SAMPLE[cell][1]))
    cells = result.cells
    assert cells["C17"] == 5677630 + 8437622
    assert cells["D17"] == Decimal("39848955737.718") + Decimal("119679040.97")
    assert cells["D17"] == Decimal("39968634778.688")
    out_amt = sum(D(c) for c in ["C7", "C8", "C9", "C10", "C13", "C14", "C15", "C16"])
    in_amt = sum(D(c) for c in ["F7", "F8", "F9", "F10", "F13", "F14", "F15", "F16"])
    assert cells["D21"] == out_amt and cells["G21"] == in_amt
    assert cells["F22"] == sum(v for v, _ in SAMPLE.values())
    assert cells["G22"] == out_amt + in_amt

    # The workbook holds the calculated numbers.
    assert ws["C17"].value == 5677630 + 8437622 and isinstance(ws["C17"].value, int)
    assert ws["F22"].value == sum(v for v, _ in SAMPLE.values())
    assert ws["G22"].value == pytest.approx(float(out_amt + in_amt))

    totals = result.totals.set_index("line")
    assert totals.loc["Total AE SETTLED", "out_volume"] == 5677630 + 8437622


def test_flags_suspicious_amount(results_dir, tmp_path):
    result = build(results_dir, tmp_path)
    assert any("(D9)" in w and "average amount" in w for w in result.warnings)


def test_missing_file_stops_unless_allowed(results_dir, tmp_path):
    (results_dir / "ARCH DB" / "RTP1" / "Outgoing_Settled_QR_CT_OCT_2026_ARCH.xls").unlink()
    with pytest.raises(ReportError, match="Outgoing_Settled_QR_CT_OCT_2026_ARCH"):
        build(results_dir, tmp_path)
    result = build(results_dir, tmp_path, allow_missing=True)
    ws = load_workbook(result.output).active
    assert ws["C14"].value is None and ws["D14"].value is None
    # Totals treat the missing input as 0.
    assert ws["C18"].value == SAMPLE["C8"][0]


def test_two_months_need_choice(results_dir, tmp_path):

    write_sqlplus(results_dir / "ARCH DB" / "RTP1" / "Outgoing_Settled_DuitAE_NOV_2026_ARCH.xls", 1, 1)
    with pytest.raises(ReportError, match="more than one month"):
        build(results_dir, tmp_path)
    assert build(results_dir, tmp_path, period="oct-2026").period == "OCT_2026"


@pytest.mark.parametrize("text,expected", [
    ("OCT_2026", "OCT_2026"), ("october 2026", "OCT_2026"), ("2026-10", "OCT_2026"),
])
def test_normalise_period(text, expected):
    assert normalise_period(text) == expected


def test_reader_handles_disguised_formats(tmp_path):
    # .xlsx content saved with .xls extension
    wb = Workbook(); ws = wb.active
    ws.append(["COUNT", "AMOUNT"]); ws.append([12, "1,234.50"])
    p = tmp_path / "a.xls"; wb.save(p)
    g = read_sheet(p)
    assert to_number(cell_value(g, "A2")) == 12
    assert to_number(cell_value(g, "B2")) == Decimal("1234.50")
    # Real Excel 97-2003
    p = tmp_path / "b.xls"; write_xls(p, 7, 99.1)
    g = read_sheet(p)
    assert to_number(cell_value(g, "A2")) == 7
    assert to_number(cell_value(g, "B2")) == Decimal("99.1")
    # SQL*Plus HTML: the <th> header row must count as row 1
    p = tmp_path / "d.xls"; write_sqlplus(p, 208092, "45881698693822.41")
    g = read_sheet(p)
    assert cell_value(g, "A1") == "Total Transaction"
    assert to_number(cell_value(g, "A2")) == 208092
    assert to_number(cell_value(g, "B2")) == Decimal("45881698693822.41")  # exact
    # Tab separated text
    p = tmp_path / "c.xls"; p.write_text("COUNT\tAMOUNT\n5\t10.5\n")
    g = read_sheet(p)
    assert to_number(cell_value(g, "A2")) == 5
    assert to_number(cell_value(g, "B2")) == Decimal("10.5")


def test_cli_writes_log(results_dir, tmp_path, capsys):
    assert main([str(results_dir), "-o", str(tmp_path)]) == 0
    out = capsys.readouterr().out
    assert "amb_txn_count_rpp_OCT_2026.xlsx" in out
    logs = list((tmp_path / "logs").glob("txn_report_*.log"))
    assert len(logs) == 1
    text = logs[0].read_text(encoding="utf-8")
    assert "Data month detected from file names: OCT_2026" in text
    assert "Outgoing_Settled_DuitAE_OCT_2026_ARCH.xls" in text
    assert "C17 = C7 + C13 = 14,115,252" in text          # each calculation
    assert "WARNING" in text and "(D9)" in text
    assert "Report written to" in text


def test_cli_logs_failure(results_dir, tmp_path, capsys):
    (results_dir / "ARCH DB" / "RTP1" / "Incoming_Settled_QR_AE_OCT_2026_ARCH.xls").unlink()
    log_dir = tmp_path / "mylogs"
    assert main([str(results_dir), "-o", str(tmp_path), "--log-dir", str(log_dir)]) == 1
    text = next(log_dir.glob("*.log")).read_text(encoding="utf-8")
    assert "MISSING" in text and "Report NOT created" in text
    assert not list(tmp_path.glob("*.xlsx"))
    assert main([str(tmp_path / "nope"), "-o", str(tmp_path)]) == 1


RTP1 = ("ARCH DB", "RTP1")


def test_zero_transactions_blank_amount_is_zero(results_dir, tmp_path):
    # SUM() over no rows is NULL, which SQL*Plus prints as &nbsp;
    write_sqlplus(results_dir.joinpath(*RTP1, "Incoming_NonSettled_QR_CT_OCT_2026_ARCH.xls"),
                  0, "&nbsp;")
    result = build(results_dir, tmp_path)
    ws = load_workbook(result.output).active
    assert ws["F16"].value == 0 and ws["G16"].value == 0


@pytest.mark.parametrize("volume,amount", [("&nbsp;", "5.00"), (12, "&nbsp;")])
def test_blank_values_stop_the_run(results_dir, tmp_path, volume, amount):
    write_sqlplus(results_dir.joinpath(*RTP1, "Outgoing_Settled_DuitCT_OCT_2026_ARCH.xls"),
                  volume, amount)
    with pytest.raises(ReportError, match="Outgoing_Settled_DuitCT"):
        build(results_dir, tmp_path)


def test_source_notes(results_dir, tmp_path):
    result = build(results_dir, tmp_path)
    ws = load_workbook(result.output).active
    note = ws["C7"].comment.text
    assert "ARCH DB > RTP1" in note
    assert "Outgoing_Settled_DuitAE_OCT_2026_ARCH.xls" in note and "cell A2" in note
    assert "cell B2" in ws["G16"].comment.text
    assert "DuitNow AE settled (D7)" in ws["D17"].comment.text
    assert "RPPQR AE settled (D13)" in ws["D17"].comment.text
    assert "Total outgoing (D21)" in ws["G22"].comment.text
    # Every value cell (32 copied from 16 files + 22 totals) has a note.
    assert sum(1 for r in ws.iter_rows() for c in r if c.comment) == 54
    assert "hover over any number" in ws["B24"].value
    assert "OCT_2026" in ws["B25"].value


def test_no_hints_option(results_dir, tmp_path):
    assert main([str(results_dir), "-o", str(tmp_path), "--no-hints", "-q"]) == 0
    ws = load_workbook(next(tmp_path.glob("*.xlsx"))).active
    assert not any(c.comment for r in ws.iter_rows() for c in r)
    assert ws["B24"].value is None and ws["B25"].value is None
