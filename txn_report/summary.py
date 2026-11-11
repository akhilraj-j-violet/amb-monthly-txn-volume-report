"""Core pandas logic for building the monthly transaction volume summary."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Sequence

import pandas as pd


@dataclass
class ReportConfig:
    """Describes how to read the input data and what to summarise."""

    date_col: str = "txn_date"
    amount_col: str = "amount"
    group_by: Sequence[str] = field(default_factory=list)
    status_col: str | None = None
    success_values: Sequence[str] = field(default_factory=lambda: ["SUCCESS"])
    start: str | None = None  # inclusive, e.g. "2026-01-01"
    end: str | None = None  # inclusive, e.g. "2026-06-30"
    date_format: str | None = None  # explicit strptime format; None = infer


# --------------------------------------------------------------------------- #
# Loading & cleaning
# --------------------------------------------------------------------------- #
def load_transactions(path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """Read a CSV / Excel / Parquet file into a DataFrame."""
    path = Path(path)
    suffix = path.suffix.lower()
    if suffix in {".csv", ".txt"}:
        return pd.read_csv(path)
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path, sheet_name=sheet_name)
    if suffix == ".parquet":
        return pd.read_parquet(path)
    raise ValueError(f"Unsupported input file type: {suffix!r}")


def _require_columns(df: pd.DataFrame, columns: Iterable[str]) -> None:
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise KeyError(
            f"Missing required column(s): {missing}. Available: {list(df.columns)}"
        )


def prepare(df: pd.DataFrame, cfg: ReportConfig) -> pd.DataFrame:
    """Validate, type-cast, filter and add a ``month`` period column."""
    required = [cfg.date_col, cfg.amount_col, *cfg.group_by]
    if cfg.status_col:
        required.append(cfg.status_col)
    _require_columns(df, required)

    out = df.copy()
    out[cfg.date_col] = pd.to_datetime(
        out[cfg.date_col], format=cfg.date_format, errors="coerce"
    )
    if isinstance(out[cfg.amount_col].dtype, pd.StringDtype) or out[
        cfg.amount_col
    ].dtype == object:
        # Strip thousands separators / currency symbols such as "1,234.50" or "$12".
        out[cfg.amount_col] = (
            out[cfg.amount_col].astype(str).str.replace(r"[^0-9.\-]", "", regex=True)
        )
    out[cfg.amount_col] = pd.to_numeric(out[cfg.amount_col], errors="coerce")

    # Drop rows we cannot place in a month or value.
    out = out.dropna(subset=[cfg.date_col, cfg.amount_col])

    if cfg.status_col:
        wanted = {str(v).strip().upper() for v in cfg.success_values}
        status = out[cfg.status_col].astype(str).str.strip().str.upper()
        out = out[status.isin(wanted)]

    if cfg.start:
        out = out[out[cfg.date_col] >= pd.Timestamp(cfg.start)]
    if cfg.end:
        # Inclusive of the whole end day.
        out = out[out[cfg.date_col] < pd.Timestamp(cfg.end) + pd.Timedelta(days=1)]

    out["month"] = out[cfg.date_col].dt.to_period("M")
    return out


# --------------------------------------------------------------------------- #
# Aggregations
# --------------------------------------------------------------------------- #
def _aggregate(df: pd.DataFrame, keys: list[str], amount_col: str) -> pd.DataFrame:
    return (
        df.groupby(keys, observed=True)[amount_col]
        .agg(
            txn_count="count",
            total_amount="sum",
            avg_amount="mean",
            min_amount="min",
            max_amount="max",
        )
        .reset_index()
    )


def monthly_summary(df: pd.DataFrame, cfg: ReportConfig) -> pd.DataFrame:
    """One row per month with volume/value stats and month-over-month growth.

    Months with no transactions inside the observed range are included with
    zero counts so gaps are visible in the report.
    """
    summary = _aggregate(df, ["month"], cfg.amount_col)
    if summary.empty:
        return summary.assign(count_mom_pct=pd.Series(dtype=float),
                              amount_mom_pct=pd.Series(dtype=float),
                              amount_share_pct=pd.Series(dtype=float))

    full_range = pd.period_range(summary["month"].min(), summary["month"].max(), freq="M")
    summary = (
        summary.set_index("month")
        .reindex(full_range)
        .rename_axis("month")
        .reset_index()
    )
    summary[["txn_count", "total_amount"]] = summary[["txn_count", "total_amount"]].fillna(0)
    summary["txn_count"] = summary["txn_count"].astype(int)

    summary["count_mom_pct"] = _pct_change(summary["txn_count"])
    summary["amount_mom_pct"] = _pct_change(summary["total_amount"])
    total = summary["total_amount"].sum()
    summary["amount_share_pct"] = (
        (summary["total_amount"] / total * 100).round(2) if total else 0.0
    )
    return _round(summary)


def group_summary(df: pd.DataFrame, cfg: ReportConfig, group_col: str) -> pd.DataFrame:
    """Long-format month x group stats."""
    summary = _aggregate(df, ["month", group_col], cfg.amount_col)
    return _round(summary.sort_values(["month", group_col]).reset_index(drop=True))


def group_pivot(
    df: pd.DataFrame, cfg: ReportConfig, group_col: str, value: str = "txn_count"
) -> pd.DataFrame:
    """Wide month x group table (rows = month, columns = group values) with totals."""
    aggfunc = "count" if value == "txn_count" else "sum"
    pivot = pd.pivot_table(
        df,
        index="month",
        columns=group_col,
        values=cfg.amount_col,
        aggfunc=aggfunc,
        fill_value=0,
        margins=True,
        margins_name="Total",
        observed=True,
    )
    pivot.columns.name = None
    return pivot.round(2)


def overall_kpis(df: pd.DataFrame, cfg: ReportConfig, monthly: pd.DataFrame) -> pd.DataFrame:
    """Headline numbers for the whole reporting period."""
    if df.empty:
        rows = {"Total transactions": 0, "Total amount": 0.0}
    else:
        busiest = monthly.loc[monthly["txn_count"].idxmax()]
        top_value = monthly.loc[monthly["total_amount"].idxmax()]
        rows = {
            "Period start": df[cfg.date_col].min().date().isoformat(),
            "Period end": df[cfg.date_col].max().date().isoformat(),
            "Months covered": len(monthly),
            "Total transactions": int(df[cfg.amount_col].count()),
            "Total amount": round(float(df[cfg.amount_col].sum()), 2),
            "Average transaction amount": round(float(df[cfg.amount_col].mean()), 2),
            "Average monthly transactions": round(float(monthly["txn_count"].mean()), 2),
            "Average monthly amount": round(float(monthly["total_amount"].mean()), 2),
            "Busiest month (by count)": f"{busiest['month']} ({int(busiest['txn_count'])})",
            "Top month (by amount)": f"{top_value['month']} ({top_value['total_amount']:,.2f})",
        }
    return pd.DataFrame({"metric": list(rows), "value": [str(v) for v in rows.values()]})


def build_report(df: pd.DataFrame, cfg: ReportConfig) -> dict[str, pd.DataFrame]:
    """Run all aggregations and return ``{sheet_name: DataFrame}``."""
    clean = prepare(df, cfg)
    monthly = monthly_summary(clean, cfg)
    report: dict[str, pd.DataFrame] = {
        "Overview": overall_kpis(clean, cfg, monthly),
        "Monthly Summary": monthly,
    }
    for col in cfg.group_by:
        report[_sheet(f"By {col}")] = group_summary(clean, cfg, col)
        report[_sheet(f"{col} Count Pivot")] = group_pivot(clean, cfg, col, "txn_count")
        report[_sheet(f"{col} Amount Pivot")] = group_pivot(clean, cfg, col, "total_amount")
    return report


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def _pct_change(s: pd.Series) -> pd.Series:
    prev = s.shift(1)
    pct = (s - prev) / prev.where(prev != 0) * 100
    return pct


def _round(df: pd.DataFrame) -> pd.DataFrame:
    num = df.select_dtypes("number").columns
    df[num] = df[num].round(2)
    return df


def _sheet(name: str) -> str:
    """Excel sheet names: max 31 chars, no []:*?/\\ characters."""
    for ch in "[]:*?/\\":
        name = name.replace(ch, "_")
    return name[:31]
