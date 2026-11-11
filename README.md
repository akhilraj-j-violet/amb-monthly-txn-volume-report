# Monthly Transaction Volume Report

A small Python + pandas tool that turns a raw transactions file into a monthly
summary report (Excel workbook, optional CSVs, and a console printout).

## What the report contains

| Sheet | Contents |
|---|---|
| **Overview** | Period covered, total count/amount, averages, busiest and top-value months |
| **Monthly Summary** | Per month: `txn_count`, `total_amount`, `avg/min/max_amount`, month-over-month % change for count and amount, share of total amount. Empty months inside the range show as 0. |
| **By `<col>`** | Month × group stats (one per `--group-by` column) |
| **`<col>` Count Pivot / Amount Pivot** | Wide month × group tables with row/column totals |

## Setup

```bash
pip install -r requirements.txt
```

## Usage

```bash
python -m txn_report transactions.csv -o reports/monthly_summary.xlsx
```

Common options:

| Option | Default | Description |
|---|---|---|
| `--date-col` | `txn_date` | Transaction date column |
| `--amount-col` | `amount` | Amount column (`"1,234.50"`, `"$12"` are cleaned automatically) |
| `--date-format` | inferred | Explicit format, e.g. `%d/%m/%Y` |
| `--group-by COL ...` | none | Break down by one or more columns (e.g. `channel txn_type`) |
| `--status-col` | none | Keep only rows whose status is in `--success-values` |
| `--success-values` | `SUCCESS` | Case-insensitive list of statuses to keep |
| `--start` / `--end` | none | Inclusive date range filter (`YYYY-MM-DD`) |
| `--csv-dir DIR` | none | Also write each table as a CSV |
| `--sheet` | `0` | Sheet name/index when the input is Excel |
| `-q` | off | Don't print the summary to the console |

Input can be `.csv`, `.xlsx`/`.xls`, or `.parquet`. Rows with unparseable dates
or amounts are dropped.

### Try it with sample data

```bash
python sample_data/generate_sample.py
```

```bash
python -m txn_report sample_data/transactions.csv --group-by channel txn_type --status-col status -o reports/sample.xlsx
```

### Use from Python

```python
from txn_report import ReportConfig, build_report, load_transactions, write_excel

df = load_transactions("transactions.csv")
report = build_report(df, ReportConfig(group_by=["channel"], status_col="status"))
print(report["Monthly Summary"])
write_excel(report, "monthly_summary.xlsx")
```

## Tests

```bash
pip install -r requirements-dev.txt
```

```bash
python -m pytest
```
