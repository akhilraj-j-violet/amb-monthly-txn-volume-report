# Project memory: AMB monthly transaction volume report

## Input rules (confirmed by the user)

- **`01_Results` is the root input directory.** The tool should take this folder
  as its input and scan every `<DB>/<Product>/` subfolder under it for Excel files.
- **File names are fixed every month. Only the month/year part changes**, e.g.
  `Settled_OCT_2026_ARCH.xls` next month becomes `Settled_NOV_2026_ARCH.xls`.
  Match files by the fixed name with the `<MON>_<YYYY>` part treated as a
  variable, and read the data month from the file name.
- The file lists below use `OCT_2026` as the example month.

## Source file layout (OneDrive)

Base path:
`Solution Enablement / AMB - RPP / 01_Volume Extraction / Monthly Volume Reports / <Run Month> / 01_Results / <DB> / <Product> /`

- `<Run Month>` is the folder for the month the report is run, e.g. `Nov 2026`.
  **It holds the previous month's data**: files in `Nov 2026` are tagged `OCT_2026`.
- `<DB>` is `ARCH DB` or `LIV DB`. File names end in `_ARCH` or `_LIV` to match.
- Files are named `.xls` but are really **SQL\*Plus HTML reports**
  (`SET MARKUP HTML ON`, generator "SQL*Plus 19.32.0"), about 1.3 KB each.
  Each has one `<table>`: a `<th>` heading row, then one `<td>` data row.
  Excel shows the heading as row 1, so A2/B2 are the first data row. The reader
  must count the `<th>` row as row 1 (`pd.read_html` makes it the column header
  instead, which shifts every cell; that bug was found and fixed with real files).
- RTP1, RTP2 and FPX files have 2 columns: `Total Transaction`, `Total Amount`.
  NAD and Mandate files have 1 column only: `Total Transaction` (no amount).
- A NULL in SQL*Plus output is `&nbsp;`, e.g. `SUM()` over zero rows.
- File name pattern: `<Name>_<MON>_<YYYY>_<DB>.xls`, e.g. `Settled_OCT_2026_ARCH.xls`.

### ARCH DB

**FPX** (2 files)
- `Not_Settled_OCT_2026_ARCH.xls`
- `Settled_OCT_2026_ARCH.xls`

**NAD** (3 files)
- `NAD_Accepted_OCT_2026_ARCH.xls`
- `NAD_Rejected_OCT_2026_ARCH.xls`
- `PRXY_Trans_Notify_OCT_2026_ARCH.xls`

**RTP1** (32 files): `<Direction>_<Status>_<Type>_OCT_2026_ARCH.xls`
- Direction: `Incoming`, `Outgoing`
- Status: `Settled`, `NonSettled`
- Type (8): `DuitAE`, `DuitAE_PACS`, `DuitAE_PRXY`, `DuitCT`, `DuitCT_PACS`,
  `DuitCT_PRXY`, `QR_AE`, `QR_CT`
- Every combination exists: 2 × 2 × 8 = 32 files.

**RTP2** (6 files)
- `Incoming_NonSettled_OCT_2026_ARCH.xls`
- `Incoming_NonSettled_PAYNET_OCT_2026_ARCH.xls`
- `Incoming_Settled_OCT_2026_ARCH.xls`
- `Incoming_Settled_PAYNET_OCT_2026_ARCH.xls`
- `Outgoing_NonSettled_OCT_2026_ARCH.xls`
- `Outgoing_Settled_OCT_2026_ARCH.xls`
- PAYNET variants exist for Incoming only.

### LIV DB

**FPX** (2 files)
- `Not_Settled_OCT_2026_LIV.xls`
- `Settled_OCT_2026_LIV.xls`

**Mandate** (4 files; this folder exists only in LIV DB, not in ARCH DB)
- `Incoming_NonSettled_OCT_2026_LIV.xls`
- `Incoming_Settled_OCT_2026_LIV.xls`
- `Outgoing_NonSettled_OCT_2026_LIV.xls`
- `Outgoing_Settled_OCT_2026_LIV.xls`

**RTP2** (6 files, same set as ARCH DB RTP2)
- `Incoming_NonSettled_OCT_2026_LIV.xls`
- `Incoming_NonSettled_PAYNET_OCT_2026_LIV.xls`
- `Incoming_Settled_OCT_2026_LIV.xls`
- `Incoming_Settled_PAYNET_OCT_2026_LIV.xls`
- `Outgoing_NonSettled_OCT_2026_LIV.xls`
- `Outgoing_Settled_OCT_2026_LIV.xls`

LIV DB has no NAD or RTP1 folder (confirmed by the OCT 2026 zip the user sent).

### Summary of folders

| DB | Products |
|---|---|
| ARCH DB | FPX (2), NAD (3), RTP1 (32), RTP2 (6) |
| LIV DB | FPX (2), Mandate (4), RTP2 (6) |

## Expected output: `amb_txn_count_rpp_<MON>_<YYYY>.xlsx` (was `TXN_count_RPP_and_Duitnow_<MON>_<YYYY>_CLIENT.xlsx`)

The user supplied `TXN_count_RPP_and_Duitnow_OCT_2026_CLIENT.xlsx` as the
target format. A blank copy, with all formatting kept and the input numbers
cleared, is in `txn_report/templates/TXN_count_RPP_and_Duitnow_TEMPLATE.xlsx`.
The client figures themselves are not committed.

- One sheet, named `<MONTHNAME>_<YYYY>` (e.g. `OCTOBER_2026`). The full month
  name is used here, while the file name uses the short form (`OCT`).
- `B2` (merged B2:G2): `TRANSACTION VOLUME - OCTOBER_2026`
- `B3` (merged B3:G3): `RPP PHASE -I`
- Left half is **OUTGOING** (B4:D4), right half is **INCOMING** (E4:G4).
  Each half has label | VOLUME | AMOUNT columns: outgoing = B | C | D,
  incoming = E | F | G.
- Volume format `#,##0`; amount format is accounting, 2 decimals.

| Rows | Section | Row labels (B and E) |
|---|---|---|
| 5–6 | `DUITNOW` header, then VOLUME/AMOUNT headings | |
| 7–10 | DUITNOW inputs | AE SETTLED, CT SETTLED, AE NOT SETTLED, CT NOT SETTLED |
| 11–12 | `RPPQR` header, then VOLUME/AMOUNT headings | |
| 13–16 | RPPQR inputs | AE SETTLED, CT SETTLED, AE NOT SETTLED, CT NOT SETTLED |
| 17–20 | Formulas: DUITNOW + RPPQR per line (e.g. `=C7+C13`) | Total AE SETTLED … Total CT NOT SETTLED |
| 21 | `=SUM(rows 17:20)` | Total Outgoing / Total Incoming |
| 22 | `F22 = C21+F21` (volume), `G22 = D21+G21` (amount) | Grand Total : Total Outgoing + Total Incoming |

**The report must contain no formulas (user request).** The tool fills the 16
input cells and calculates rows 17–22 itself (`CALCULATIONS` in
`txn_report/mapping.py`, done with exact `Decimal` arithmetic), then writes plain
numbers. The user maintains the template themselves (they uploaded a new one
with E18 fixed; it still has formulas), so `write_report` clears every formula
in the template at run time instead of relying on a formula-free template.
Output file name is `OUTPUT_NAME` in `txn_report/report.py`; the user changed
it to `amb_txn_count_rpp_{period}.xlsx`. Keep their edits. Volumes are written as ints and
amounts as floats; Excel shows only 15 significant digits, so the tool logs the
exact value for any cell with more digits.

Quirks in the sample:
- `E18` reads just `T`, which looks like a typo for `Total CT SETTLED`.
- `D9` (Outgoing DUITNOW AE NOT SETTLED amount) is 45,881,698,693,822.40,
  far larger than every other figure. The source file really contains
  `45881698693822.41` (Excel shows `.40` because it keeps only 15 significant
  digits), so the outlier comes from the extract query or data, not the tool.

### Results mapping (confirmed by the user, implemented in `txn_report/mapping.py`)

Every input file has **volume in A2 and amount in B2** (row 1 is a header).
All 16 inputs come from `ARCH DB/RTP1`. The `_PACS` and `_PRXY` files are **not** used.

| Report cells (vol/amt) | Outgoing file | Report cells (vol/amt) | Incoming file |
|---|---|---|---|
| C7/D7 | Outgoing_Settled_DuitAE | F7/G7 | Incoming_Settled_DuitAE |
| C8/D8 | Outgoing_Settled_DuitCT | F8/G8 | Incoming_Settled_DuitCT |
| C9/D9 | Outgoing_NonSettled_DuitAE | F9/G9 | Incoming_NonSettled_DuitAE |
| C10/D10 | Outgoing_NonSettled_DuitCT | F10/G10 | Incoming_NonSettled_DuitCT |
| C13/D13 | Outgoing_Settled_QR_AE | F13/G13 | Incoming_Settled_QR_AE |
| C14/D14 | Outgoing_Settled_QR_CT | F14/G14 | Incoming_Settled_QR_CT |
| C15/D15 | Outgoing_NonSettled_QR_AE | F15/G15 | Incoming_NonSettled_QR_AE |
| C16/D16 | Outgoing_NonSettled_QR_CT | F16/G16 | Incoming_NonSettled_QR_CT |

(File names all end in `_<MON>_<YYYY>_ARCH`.)

Calculations: rows 17–20 = DUITNOW + RPPQR (e.g. C17 = C7+C13); row 21 = sum
of rows 17–20; F22 = C21+F21; G22 = D21+G21. The user's written mapping said
`G22 = D21+G22`, which would be circular; D21+G21 is used.

## Source notes (user request)
Every value cell (32 copied cells + 22 totals) gets a plain-language Excel note
saying where its value comes from (folder, file, cell A2/B2) or how it is
added up. Texts are generated from `SOURCES`/`CALCULATIONS` in
`txn_report/hints.py`, so they stay correct if the mapping changes. B24/B25
hold a visible tip line. The template carries the same notes with
`<MON>_<YYYY>`; the report gets the real month. `--no-hints` removes them.
The template no longer contains formulas (cleared after the user's upload).

The template footer `# Classified as Confidential` comes from the user's
template (a sensitivity label); keep it.

## Logging (user request)
Every run writes `<output dir>/logs/txn_report_<YYYYMMDD_HHMMSS>.log`
(`--log-dir`, `--log-level`). See `txn_report/logs.py`. Use the `txn_report`
logger for any new messages; the file gets DEBUG detail (raw cell values, each
calculation), the console gets INFO.

## Open questions
- Tested against the real OCT 2026 `01_Results` zip: the generated report
  matches the client file in every cell except the E18 label fix and D9's
  last digit (see above). Client input files are not committed.
- Whether other products (FPX, NAD, RTP2, Mandate) get their own report or sheet.
