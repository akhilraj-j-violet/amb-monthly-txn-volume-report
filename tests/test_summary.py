import pandas as pd
import pytest

from txn_report import ReportConfig, build_report, prepare, write_excel
from txn_report.cli import main


@pytest.fixture
def df():
    return pd.DataFrame({
        "txn_date": ["2026-01-05", "2026-01-20", "2026-03-02", "2026-03-15", "bad", "2026-03-20"],
        "amount": ["1,000.00", "500", "200", "300", "50", "999"],
        "channel": ["ATM", "POS", "ATM", "ATM", "POS", "POS"],
        "status": ["SUCCESS", "success", "SUCCESS", "SUCCESS", "SUCCESS", "FAILED"],
    })


def test_prepare_cleans_and_filters(df):
    out = prepare(df, ReportConfig(status_col="status"))
    assert len(out) == 4  # bad date and FAILED dropped
    assert out["amount"].tolist() == [1000.0, 500.0, 200.0, 300.0]


def test_monthly_summary_fills_gaps_and_growth(df):
    monthly = build_report(df, ReportConfig(status_col="status"))["Monthly Summary"]
    assert monthly["month"].astype(str).tolist() == ["2026-01", "2026-02", "2026-03"]
    assert monthly["txn_count"].tolist() == [2, 0, 2]
    assert monthly["total_amount"].tolist() == [1500.0, 0.0, 500.0]
    assert monthly.loc[1, "amount_mom_pct"] == -100.0
    assert pd.isna(monthly.loc[2, "amount_mom_pct"])  # previous month was zero


def test_group_by_and_pivots(df):
    report = build_report(df, ReportConfig(status_col="status", group_by=["channel"]))
    pivot = report["channel Count Pivot"]
    assert pivot.loc["Total", "Total"] == 4
    assert report["channel Amount Pivot"].loc["Total", "ATM"] == 1500.0


def test_date_range_filter(df):
    report = build_report(df, ReportConfig(start="2026-03-01", end="2026-03-15"))
    assert report["Monthly Summary"]["txn_count"].tolist() == [2]


def test_missing_column_raises(df):
    with pytest.raises(KeyError):
        build_report(df, ReportConfig(amount_col="value"))


def test_excel_and_cli(tmp_path, df):
    src = tmp_path / "txns.csv"
    df.to_csv(src, index=False)
    out = tmp_path / "report.xlsx"
    assert main([str(src), "-o", str(out), "--group-by", "channel",
                 "--status-col", "status", "-q"]) == 0
    sheets = pd.read_excel(out, sheet_name=None)
    assert {"Overview", "Monthly Summary", "By channel"} <= set(sheets)
