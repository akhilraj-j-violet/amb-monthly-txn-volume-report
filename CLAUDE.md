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
- All files are `.xls` (Excel 97-2003) and about 2 KB each, so they hold small
  extract results, not raw transaction dumps. Reading `.xls` with pandas needs `xlrd`.
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

No NAD or RTP1 folder has been shown for LIV DB. Confirm with the user whether
they exist before assuming they do.

### Summary of folders

| DB | Products |
|---|---|
| ARCH DB | FPX (2), NAD (3), RTP1 (32), RTP2 (6) |
| LIV DB | FPX (2), Mandate (4), RTP2 (6); NAD and RTP1 not confirmed |

## Expected output: `TXN_count_RPP_and_Duitnow_<MON>_<YYYY>_CLIENT.xlsx`

The user supplied `TXN_count_RPP_and_Duitnow_OCT_2026_CLIENT.xlsx` as the
target format. A blank copy, with all formatting kept and the input numbers
cleared, is in `templates/TXN_count_RPP_and_Duitnow_TEMPLATE.xlsx`.
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

Only the 16 input cells (C/D/F/G × rows 7–10 and 13–16) need filling. Rows
17–22 are formulas and must be left as they are.

Quirks in the sample:
- `E18` reads just `T`, which looks like a typo for `Total CT SETTLED`.
- `D9` (Outgoing DUITNOW AE NOT SETTLED amount) is 45,881,698,693,822.40,
  far larger than every other figure. This may be a data issue in the source.

The mapping from input files to the 16 cells is still to come from the user.
Do not guess it. It probably uses RTP1 (`DuitAE`/`DuitCT` → DUITNOW,
`QR_AE`/`QR_CT` → RPPQR), but it is unconfirmed how the `_PACS`/`_PRXY` files fit.

## Open questions
- The column layout inside the input `.xls` files is still unknown. Ask for a
  sample before building the parser.
- Results mapping: input file → output cell (the user will provide it).
- Whether other products (FPX, NAD, RTP2, Mandate) get their own report or sheet.
