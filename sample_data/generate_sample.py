"""Generate a synthetic transactions CSV for trying out the report."""

import argparse

import numpy as np
import pandas as pd


def generate(rows: int = 5000, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    start = pd.Timestamp("2026-01-01")
    dates = start + pd.to_timedelta(rng.integers(0, 304, rows), unit="D")
    return pd.DataFrame({
        "txn_id": [f"T{i:07d}" for i in range(1, rows + 1)],
        "txn_date": dates.strftime("%Y-%m-%d"),
        "amount": rng.gamma(2.0, 150.0, rows).round(2),
        "channel": rng.choice(["ATM", "POS", "Online", "Branch"], rows, p=[.3, .35, .25, .1]),
        "txn_type": rng.choice(["Debit", "Credit"], rows, p=[.7, .3]),
        "status": rng.choice(["SUCCESS", "FAILED", "REVERSED"], rows, p=[.92, .06, .02]),
    }).sort_values("txn_date", ignore_index=True)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("-n", "--rows", type=int, default=5000)
    ap.add_argument("-o", "--output", default="sample_data/transactions.csv")
    a = ap.parse_args()
    generate(a.rows).to_csv(a.output, index=False)
    print(f"Wrote {a.rows} rows to {a.output}")
