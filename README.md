# Monthly RPP / DuitNow Transaction Volume Report

Builds `amb_txn_count_rpp_<MON>_<YYYY>.xlsx` (e.g. `amb_txn_count_rpp_OCT_2026.xlsx`) from the monthly
`01_Results` extract folder, using pandas.

## Running on Windows

### Which command does your PC have?
Open PowerShell and run both of these:
```
py --version
python --version
```
- **`py` shows a version** (e.g. `Python 3.14.5`): follow **Option A**.
- **Only `python` shows a version**: follow **Option B**.
- **Both show a version**: either option works.
- **Neither works**, or `python` says *"Python was not found ... Microsoft
  Store"*: install Python from https://www.python.org/downloads/ (tick
  **Add python.exe to PATH** on the first screen), open a new PowerShell
  window and check again.

`py` is the Python Launcher for Windows: it finds Python from the Windows
registry, so it works even when Python is not on PATH. `python` only works when
Python's folder is on PATH; otherwise Windows runs a Microsoft Store placeholder
that just prints a message.

### Easiest: run_report.bat (any machine)
`run_report.bat` checks which of `py` / `python` works, installs the libraries
the first time, and runs the report.
- **Double-click it**, paste the path of the `01_Results` folder and press
  Enter. The report and a `logs` folder are saved on your Desktop.
- Or run it from PowerShell, in the tool's folder, with any options:
  ```
  .\run_report.bat "C:\...\Monthly Volume Reports\Nov 2026\01_Results" -o "C:\Reports"
  ```

### Option A: machines with `py`
In PowerShell, in the tool's folder (the one containing `requirements.txt`):

1. Install the libraries (first time only):
   ```
   py -m pip install -r requirements.txt
   ```
2. Run the report:
   ```
   py -m txn_report "C:\...\Monthly Volume Reports\Nov 2026\01_Results" -o "$HOME\Desktop"
   ```

### Option B: machines with `python`
In PowerShell, in the tool's folder:

1. Install the libraries (first time only):
   ```
   python -m pip install -r requirements.txt
   ```
2. Run the report:
   ```
   python -m txn_report "C:\...\Monthly Volume Reports\Nov 2026\01_Results" -o "$HOME\Desktop"
   ```

### Tips
- Get the folder path by opening `01_Results` in File Explorer, clicking the
  address bar and copying it. Keep the quotes around paths, because they
  contain spaces.
- In Command Prompt (cmd) instead of PowerShell, write
  `"%USERPROFILE%\Desktop"` instead of `"$HOME\Desktop"`.
- Use `python -m pip` / `py -m pip`, not plain `pip`: plain `pip` is usually
  not on PATH and gives *"The term 'pip' is not recognized"*.
- A yellow warning that scripts are *"installed in ... which is not on PATH"*
  during installation is harmless.
- If OneDrive files show only a cloud icon and can't be read, right-click
  `01_Results` and choose **Always keep on this device**.

## Usage (all options)

Point the tool at the month's `01_Results` folder (use `py` or `python` as
above):

```bash
python -m txn_report "<path to 01_Results>" -o "<output folder>"
```

The tool then:
1. Works out the data month from the file names (e.g. `OCT_2026`). Use
   `--month OCT_2026` if files for more than one month are in the folder.
2. Opens the 16 mapped files in `ARCH DB/RTP1/` and reads **A2 = volume** and
   **B2 = amount** from each.
3. Fills cells C/D/F/G of rows 7–10 (DUITNOW) and 13–16 (RPPQR) in the report
   template, then **calculates the totals in rows 17–22 itself** and writes
   them as plain numbers. The report contains no formulas.
4. Names the sheet `OCTOBER_2026` and sets the title to
   `TRANSACTION VOLUME - OCTOBER_2026`.
5. Writes a log file for every run (see **Logs** below).

### DATA SOURCES box
Under the table (from row 24) the report and the template have a bordered
**DATA SOURCES** box saying where every number comes from. Outgoing is on the
left and incoming on the right, like the table:

```
OUTGOING - DUITNOW
AE SETTLED VOLUME - ARCH/RTP1/Outgoing_Settled_DuitAE_ARCH/volume
AE SETTLED AMOUNT - ARCH/RTP1/Outgoing_Settled_DuitAE_ARCH/Amount
CT SETTLED VOLUME - ARCH/RTP1/Outgoing_Settled_DuitCT_ARCH/volume
...
OUTGOING - TOTALS (added up by the report tool)
TOTAL AE SETTLED VOLUME - DUITNOW + RPPQR AE SETTLED VOLUME
...
GRAND TOTAL VOLUME - TOTAL OUTGOING VOLUME + TOTAL INCOMING VOLUME
```

File names are shown without the month and year, so the box is the same every
month. `volume` is cell A2 and `Amount` is cell B2 of that file. The box is
generated from the mapping in `txn_report/hints.py`, so it always matches what
the tool does; edits made to the box in the template are replaced when a report
is written. Use `--no-hints` to leave the box out.

### Calculations
The values are read from the files as text and added with Python's `Decimal`
type, so the sums are exact (no floating-point rounding):

| Cells | Calculation |
|---|---|
| C17:G20 | DUITNOW + RPPQR for each line, e.g. `C17 = C7 + C13` |
| C21, D21, F21, G21 | Total Outgoing / Incoming = rows 17 + 18 + 19 + 20 |
| F22 | Grand total volume = C21 + F21 |
| G22 | Grand total amount = D21 + G21 |

Excel can only store 15 significant digits. If a value has more (e.g. October's
D9 = 45,881,698,693,822.41), Excel shows the last digit as 0, so the tool logs
a warning with the exact value.

### Logs
Each run creates `logs/txn_report_<YYYYMMDD_HHMMSS>.log` inside the output
folder (change it with `--log-dir`). The log contains:
- the command, input folder and detected month,
- each file read, its raw A2/B2 text and the values used,
- every calculation (e.g. `C17 = C7 + C13 = 14,115,252`),
- all warnings and errors (with full details if something crashes),
- the input and totals tables and the output file path.

The console shows the same information without the DEBUG details. Use `-q`
to see only warnings and errors there; the log file still gets everything.

| Option | Meaning |
|---|---|
| `-o, --output-dir` | Where to save the report (default: current folder) |
| `-m, --month` | Data month, e.g. `OCT_2026`, `2026-10`, `October 2026` |
| `--allow-missing` | Build even if mapped files are missing (their cells stay blank) |
| `--max-avg-amount` | Warn when amount ÷ volume is above this (default 1,000,000) |
| `--no-hints` | Leave out the DATA SOURCES box under the table |
| `--log-dir` | Folder for log files (default: `logs` inside the output folder) |
| `--log-level` | Detail in the log file: `DEBUG` (default), `INFO` or `WARNING` |
| `-q` | Console shows only warnings, errors and the output path |

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

## Changing the output file name
Edit `OUTPUT_NAME` near the top of `txn_report/report.py`. Keep `{period}` in
it; it becomes the data month, e.g. `OCT_2026`.

## Changing the mapping
The file → cell mapping is in `txn_report/mapping.py`. The report layout is
`txn_report/templates/TXN_count_RPP_and_Duitnow_TEMPLATE.xlsx`. You can edit
the template's look in Excel; any formulas in it are removed when the report is
written, because the tool calculates rows 17–22 itself.

## Tests

```bash
pip install -r requirements-dev.txt
```

```bash
python -m pytest
```
