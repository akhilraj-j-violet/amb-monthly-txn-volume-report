# Monthly RPP / DuitNow Transaction Volume Report

Builds `TXN_count_RPP_and_Duitnow_<MON>_<YYYY>_CLIENT.xlsx` from the monthly
`01_Results` extract folder, using pandas.

## Setup

```bash
pip install -r requirements.txt
```

## Usage

Point the tool at the month's `01_Results` folder:

```bash
python -m txn_report "C:\...\Monthly Volume Reports\Nov 2026\01_Results" -o "C:\...\Reports"
```

The tool then:
1. Works out the data month from the file names (e.g. `OCT_2026`). Use
   `--month OCT_2026` if files for more than one month are in the folder.
2. Opens the 16 mapped files in `ARCH DB/RTP1/` and reads **A2 = volume** and
   **B2 = amount** from each.
3. Fills cells C/D/F/G of rows 7–10 (DUITNOW) and 13–16 (RPPQR) in the report
   template. The total rows 17–22 stay as Excel formulas.
4. Names the sheet `OCTOBER_2026` and sets the title to
   `TRANSACTION VOLUME - OCTOBER_2026`.
5. Prints every input value and the totals, plus warnings.

| Option | Meaning |
|---|---|
| `-o, --output-dir` | Where to save the report (default: current folder) |
| `-m, --month` | Data month, e.g. `OCT_2026`, `2026-10`, `October 2026` |
| `--allow-missing` | Build even if mapped files are missing (their cells stay blank) |
| `--max-avg-amount` | Warn when amount ÷ volume is above this (default 1,000,000) |
| `-q` | Print only warnings and the output path |

Without `--allow-missing`, a missing or unreadable file stops the run and lists
what is wrong.

### Checks
- A mapped file is missing or can't be read: the run stops.
- A2 (volume) is empty, or B2 (amount) is empty while the volume isn't 0: the run stops.
- Volume 0 with an empty amount (SQL `SUM` of no rows): the amount is taken as 0.
- A value is negative, a volume isn't a whole number, an amount has no volume,
  or the average amount per transaction is unusually high: a warning is shown.

### Input file formats
The extracts are **SQL\*Plus HTML reports saved with a `.xls` name**. Excel
opens them as a sheet where row 1 is the heading row (`Total Transaction`,
`Total Amount`) and row 2 holds the values. The tool reads them the same way.
Real Excel 97-2003, Excel 2007+ and tab/comma-separated text files are also
accepted; the format is detected from the file's contents, not its extension.

## Changing the mapping
The file → cell mapping is in `txn_report/mapping.py`. The report layout is
`txn_report/templates/TXN_count_RPP_and_Duitnow_TEMPLATE.xlsx`.

## Tests

```bash
pip install -r requirements-dev.txt
```

```bash
python -m pytest
```
