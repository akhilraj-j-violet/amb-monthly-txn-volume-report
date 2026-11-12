import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from conftest import SAMPLE
from txn_report import ReportError, build, detect_period
from txn_report.cli import main
from txn_report.reader import read_sheet, cell_value, to_number
from txn_report.report import normalise_period


def test_detect_period(results_dir):
    assert detect_period(results_dir) == "OCT_2026"


def test_full_report(results_dir, tmp_path):
    result = build(results_dir, tmp_path / "out")
    assert result.output.name == "TXN_count_RPP_and_Duitnow_OCT_2026_CLIENT.xlsx"

    ws = load_workbook(result.output).active
    assert ws.title == "OCTOBER_2026"
    assert ws["B2"].value == "TRANSACTION VOLUME - OCTOBER_2026"
    assert ws["E18"].value == "Total CT SETTLED"
    for vol_cell, (vol, amt) in SAMPLE.items():
        amt_cell = chr(ord(vol_cell[0]) + 1) + vol_cell[1:]
        assert ws[vol_cell].value == vol
        assert ws[amt_cell].value == pytest.approx(amt)
    # Formula rows untouched.
    assert ws["C17"].value == "=C7+C13"
    assert ws["G22"].value == "=SUM(D21,G21)"

    totals = result.totals.set_index("line")
    assert totals.loc["Total AE SETTLED", "out_volume"] == 5677630 + 8437622
    grand = totals.loc["Grand Total (Outgoing + Incoming)"]
    assert grand["in_volume"] == sum(v for v, _ in SAMPLE.values())
    assert grand["in_amount"] == pytest.approx(sum(a for _, a in SAMPLE.values()))


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


def test_two_months_need_choice(results_dir, tmp_path):
    from conftest import write_xls
    write_xls(results_dir / "ARCH DB" / "RTP1" / "Outgoing_Settled_DuitAE_NOV_2026_ARCH.xls", 1, 1)
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
    assert to_number(cell_value(g, "A2")) == 12 and to_number(cell_value(g, "B2")) == 1234.5
    # HTML table export
    p = tmp_path / "b.xls"
    p.write_text("<html><table><tr><td>COUNT</td><td>AMOUNT</td></tr>"
                 "<tr><td>7</td><td>99.10</td></tr></table></html>")
    g = read_sheet(p)
    assert to_number(cell_value(g, "A2")) == 7 and to_number(cell_value(g, "B2")) == 99.1
    # Tab separated text
    p = tmp_path / "c.xls"; p.write_text("COUNT\tAMOUNT\n5\t10.5\n")
    g = read_sheet(p)
    assert to_number(cell_value(g, "A2")) == 5 and to_number(cell_value(g, "B2")) == 10.5


def test_cli(results_dir, tmp_path, capsys):
    assert main([str(results_dir), "-o", str(tmp_path)]) == 0
    out = capsys.readouterr()
    assert "TXN_count_RPP_and_Duitnow_OCT_2026_CLIENT.xlsx" in out.out
    assert main([str(tmp_path / "nope")]) == 1
