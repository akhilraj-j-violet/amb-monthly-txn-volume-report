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

## Open questions
- The column layout inside the `.xls` files is still unknown. Ask for a sample
  before building the parser.
- The expected format of the final summary report is still unknown.
