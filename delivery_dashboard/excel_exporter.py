"""Render the processed dashboard into a formatted .xlsx workbook.

Sheet order and formatting follow the spec: dark-blue headers, frozen header
row, Excel auto-filters, MM/DD/YYYY dates, Canadian-currency totals, whole-
percent picking, wrapped long text, status/priority highlighting and totals
rows on the customer detail reports.
"""
from __future__ import annotations

from datetime import date, datetime
from io import BytesIO

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from . import risk_engine as R

# --- styles ------------------------------------------------------------------
HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
HEADER_FONT = Font(color="FFFFFF", bold=True)
TOTAL_FONT = Font(bold=True)
NOTE_FONT = Font(italic=True, color="9C0006")

FILL_CRITICAL = PatternFill("solid", fgColor="FFC7CE")   # red
FILL_AT_RISK = PatternFill("solid", fgColor="FFD9A0")     # orange
FILL_NOT_STARTED = PatternFill("solid", fgColor="FFF2CC")  # yellow
FILL_READY = PatternFill("solid", fgColor="D9EAD3")       # green-ish

CURRENCY_FMT = '"$"#,##0.00'
DATE_FMT = "MM/DD/YYYY"
PCT_FMT = '0"%"'

_CURRENCY_COLS = {"Total Price", "Sales Order Total"}
_DATE_COLS = {
    "Warehouse Task Creat", "Route Depart. Date", "Planned Dlv. Date",
    "Customer Dock Appointment Date", "Due Date",
}
_PCT_COLS = {"Picking in %", "picking_pct"}
_WRAP_COLS = {"Issue", "Consequence / Risk", "Latest Update"}

# Sheets that receive a Sales Order Total totals row.
_TOTALS_SHEETS = {"AMZ", "PFG WHSE", "Calgary CO-OP", "Sysco", "GFS", "Pratts"}
# Customer detail sheets that get late/not-started row highlighting.
_HIGHLIGHT_DETAIL_SHEETS = _TOTALS_SHEETS

AMZ_NOTE = ("HIGHLIGHTED ORDERS ARE CURRENTLY LATE AND MAY BE SUBJECT TO THE "
            "CONFIGURED ON-TIME DELIVERY CHARGEBACK")

# Canonical download sheet order.
SHEET_ORDER = [
    "Issue Tracker", "Not Started Orders", "SAPUI5 Export", "DSD Priorities",
    "AMZ", "Andrew Tab", "Andrew Tab 2", "PFG WHSE", "Calgary CO-OP",
    "Sysco", "GFS", "Pratts",
]


def _sanitize(value, is_numeric_col: bool):
    """Convert a pandas cell into something openpyxl can write cleanly."""
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return 0 if is_numeric_col else None
    if isinstance(value, pd.Timestamp):
        return None if pd.isna(value) else value.to_pydatetime()
    try:
        if pd.isna(value):
            return 0 if is_numeric_col else None
    except (TypeError, ValueError):
        pass
    return value


def _detail_highlight(row: dict, as_of_ts: pd.Timestamp):
    """Row fill for a customer detail sheet, from raw canonical columns."""
    gi = str(row.get("Goods Issue", "") or "").strip().upper()
    planned = pd.to_datetime(row.get("Planned Dlv. Date"), errors="coerce")
    route = pd.to_datetime(row.get("Route Depart. Date"), errors="coerce")
    pct = pd.to_numeric(pd.Series([row.get("Picking in %")]), errors="coerce").fillna(0).iloc[0]
    late = pd.notna(planned) and planned < as_of_ts and gi != "COMPLETED"
    route_missed = pd.notna(route) and route < as_of_ts and pct < 100
    if late or route_missed:
        return FILL_CRITICAL
    if pct == 0 and gi != "COMPLETED":
        return FILL_NOT_STARTED
    return None


def _write_sheet(
    wb: Workbook,
    name: str,
    df: pd.DataFrame,
    *,
    as_of_ts: pd.Timestamp,
    add_totals: bool = False,
    note: str | None = None,
    status_col: str | None = None,
    priority_col: str | None = None,
    highlight_detail: bool = False,
) -> None:
    ws = wb.create_sheet(title=name[:31])
    if df is None or df.empty:
        ws.append([f"No data for: {name}"])
        return

    columns = list(df.columns)
    numeric_cols = {c for c in columns if c in _CURRENCY_COLS or c in _PCT_COLS
                    or c in {"Cases", "Line Items", "Pallet Quantity", "Num. Lifts Pallets",
                             "Gross Weight", "Order Count", "Days to Route Departure",
                             "Days to Planned Delivery"}}

    ws.append(columns)
    records = df.to_dict("records")
    for rec in records:
        ws.append([_sanitize(rec.get(c), c in numeric_cols) for c in columns])

    # Header styling.
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(records) + 1}"

    # Column formats + widths.
    for idx, col in enumerate(columns, start=1):
        letter = get_column_letter(idx)
        fmt = None
        if col in _CURRENCY_COLS:
            fmt = CURRENCY_FMT
        elif col in _DATE_COLS:
            fmt = DATE_FMT
        elif col in _PCT_COLS:
            fmt = PCT_FMT
        if fmt:
            for r in range(2, len(records) + 2):
                ws.cell(row=r, column=idx).number_format = fmt
        if col in _WRAP_COLS:
            for r in range(2, len(records) + 2):
                ws.cell(row=r, column=idx).alignment = Alignment(wrap_text=True, vertical="top")
            ws.column_dimensions[letter].width = 46
        else:
            sample = [str(rec.get(col, "")) for rec in records[:200]]
            width = max([len(str(col))] + [len(s) for s in sample]) + 2
            ws.column_dimensions[letter].width = min(max(width, 12), 50)

    # Row highlighting.
    for r_i, rec in enumerate(records, start=2):
        fill = None
        if status_col and status_col in rec:
            s = rec[status_col]
            fill = {R.CRITICAL: FILL_CRITICAL, R.AT_RISK: FILL_AT_RISK,
                    R.NOT_STARTED: FILL_NOT_STARTED, R.READY: FILL_READY}.get(s)
        elif priority_col and priority_col in rec:
            p = rec[priority_col]
            fill = {"Critical": FILL_CRITICAL, "High": FILL_AT_RISK,
                    "Medium": FILL_NOT_STARTED}.get(p)
        elif highlight_detail:
            fill = _detail_highlight(rec, as_of_ts)
        if fill:
            for c_i in range(1, len(columns) + 1):
                ws.cell(row=r_i, column=c_i).fill = fill

    # Totals row for customer reports.
    if add_totals and "Sales Order Total" in columns:
        total = float(pd.to_numeric(df["Sales Order Total"], errors="coerce").fillna(0).sum())
        total_row = len(records) + 2
        label_col = 1
        ws.cell(row=total_row, column=label_col, value="TOTAL").font = TOTAL_FONT
        tcol = columns.index("Sales Order Total") + 1
        cell = ws.cell(row=total_row, column=tcol, value=total)
        cell.number_format = CURRENCY_FMT
        cell.font = TOTAL_FONT

    if note:
        note_row = ws.max_row + 2
        c = ws.cell(row=note_row, column=1, value=note)
        c.font = NOTE_FONT


def build_workbook(result) -> bytes:
    """Build the full workbook from a ``ProcessResult``. Returns .xlsx bytes."""
    as_of_ts = pd.Timestamp(result.as_of)
    wb = Workbook()
    wb.remove(wb.active)  # drop the default empty sheet

    _write_sheet(wb, "Issue Tracker", result.issue_tracker, as_of_ts=as_of_ts,
                 status_col="Status")
    _write_sheet(wb, "Not Started Orders", result.not_started, as_of_ts=as_of_ts,
                 priority_col="Priority")

    reports = result.reports
    for name in ["SAPUI5 Export", "DSD Priorities", "AMZ", "Andrew Tab", "Andrew Tab 2",
                 "PFG WHSE", "Calgary CO-OP", "Sysco", "GFS", "Pratts"]:
        if name in reports:
            _write_sheet(
                wb, name, reports[name], as_of_ts=as_of_ts,
                add_totals=name in _TOTALS_SHEETS,
                note=AMZ_NOTE if name == "AMZ" else None,
                highlight_detail=name in _HIGHLIGHT_DETAIL_SHEETS,
            )

    # Any extra customer reports not in the canonical order (future-proofing).
    for name, df in reports.items():
        if name not in SHEET_ORDER and name not in wb.sheetnames:
            _write_sheet(wb, name, df, as_of_ts=as_of_ts)

    # Validation errors last, only if there are any.
    if result.validation_errors is not None and not result.validation_errors.empty:
        _write_sheet(wb, "Validation Errors", result.validation_errors, as_of_ts=as_of_ts)

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
