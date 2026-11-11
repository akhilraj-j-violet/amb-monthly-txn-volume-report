"""Write the report to Excel / CSV and render it for the console."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def _for_export(df: pd.DataFrame) -> pd.DataFrame:
    """Convert Period values (index or columns) to plain strings."""
    out = df.copy()
    if isinstance(out.index, pd.PeriodIndex) or out.index.dtype == object:
        out.index = out.index.astype(str)
    for col in out.columns:
        if isinstance(out[col].dtype, pd.PeriodDtype):
            out[col] = out[col].astype(str)
    return out


def write_excel(report: dict[str, pd.DataFrame], path: str | Path) -> Path:
    """Write every report table to its own sheet, with readable column widths."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        for name, df in report.items():
            data = _for_export(df)
            keep_index = "Pivot" in name
            data.to_excel(xw, sheet_name=name, index=keep_index)
            ws = xw.sheets[name]
            ws.freeze_panes = "B2" if keep_index else "A2"
            for column_cells in ws.columns:
                width = max(len(str(c.value)) if c.value is not None else 0 for c in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = min(width + 2, 50)
                for cell in column_cells[1:]:
                    if isinstance(cell.value, float):
                        cell.number_format = "#,##0.00"
                    elif isinstance(cell.value, int):
                        cell.number_format = "#,##0"
    return path


def write_csvs(report: dict[str, pd.DataFrame], directory: str | Path) -> list[Path]:
    """Write one CSV per report table into ``directory``."""
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    written = []
    for name, df in report.items():
        file = directory / f"{name.lower().replace(' ', '_')}.csv"
        _for_export(df).to_csv(file, index="Pivot" in name)
        written.append(file)
    return written


def render_console(report: dict[str, pd.DataFrame]) -> str:
    """Pretty text rendering of the overview and monthly summary."""
    parts = []
    with pd.option_context(
        "display.max_columns", None,
        "display.width", 200,
        "display.float_format", "{:,.2f}".format,
    ):
        for name in ("Overview", "Monthly Summary"):
            if name in report:
                parts.append(f"=== {name} ===")
                parts.append(_for_export(report[name]).to_string(index=False))
                parts.append("")
    return "\n".join(parts)
