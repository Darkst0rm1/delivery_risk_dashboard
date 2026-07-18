"""Read and validate the SAPUI5 delivery export.

Matches columns by *normalized* header name (not position) so harmless header
differences — extra spaces, embedded newlines, capitalization, trailing
punctuation — all resolve to the same canonical column. Real SAP exports
truncate a few long headers, so an alias table maps the known truncations to
their full canonical names.
"""
from __future__ import annotations

import re
from typing import Any

import pandas as pd


class DeliveryLoadError(Exception):
    """Raised when the file is unreadable, empty, or missing required columns."""


# --- Canonical columns -------------------------------------------------------
# The full set of business columns the export is expected to carry. Order here
# is also the order used for the "SAPUI5 Export" sheet in the download.
CANONICAL_COLUMNS: list[str] = [
    "ys",
    "Facility Code",
    "Warehouse Task Creat",
    "Picking in %",
    "Route Depart. Date",
    "Planned Dlv. Date",
    "LE Delivery",
    "Cust. Reference",
    "STO PO Number",
    "Carrier",
    "Carrier Description",
    "Shipment Type",
    "Shpg Cond. Desc",
    "Ship-to Name",
    "Ship-to City",
    "Ship-to Province",
    "Cases",
    "Line Items",
    "Goods Issue",
    "Pallet Quantity",
    "Num. Lifts Pallets",
    "Packing Status",
    "Picking",
    "Picking (Plan)",
    "COD Paid Flag",
    "Customer Dock Appointment Date",
    "Customer Dock Appointment Time",
    "Customer Dock Reference",
    "Gross Weight",
    "Sales Order",
    "Outb. Del. Ord.",
    "Shipping Cond.",
    "Sales Order Total",
    "Consignee",
    "Purchase Order",
]

# Columns without which the dashboard cannot compute anything meaningful. A
# missing one of these aborts processing with a clear error.
REQUIRED_COLUMNS: list[str] = [
    "Facility Code",
    "Picking in %",
    "Planned Dlv. Date",
    "Route Depart. Date",
    "LE Delivery",
    "Ship-to Name",
    "Goods Issue",
    "Sales Order Total",
]

# Aliases for headers that arrive differently from the canonical name. Keys are
# themselves normalized (see _normalize_header) before comparison, so only the
# *content* differences need listing here — not spacing/case variants.
_HEADER_ALIASES: dict[str, str] = {
    # Known SAP truncations / long-name variants.
    "warehouse task creation": "Warehouse Task Creat",
    "warehouse task creation date": "Warehouse Task Creat",
    "picking in": "Picking in %",
    "picking %": "Picking in %",
    "route departure date": "Route Depart. Date",
    "planned delivery date": "Planned Dlv. Date",
    "planned dlv date": "Planned Dlv. Date",
    "customer dock appointment": "Customer Dock Appointment Date",
    "customer dock appointment date": "Customer Dock Appointment Date",
    "customer dock appointment time": "Customer Dock Appointment Time",
    "customer dock reference": "Customer Dock Reference",
    "num lifts pallets": "Num. Lifts Pallets",
    "outb del ord": "Outb. Del. Ord.",
    "sto po number": "STO PO Number",
    "cust reference": "Cust. Reference",
    "shpg cond desc": "Shpg Cond. Desc",
    "shipping cond": "Shipping Cond.",
    # Some exports label the site column "Warehouse" or "Site" instead of "ys".
    "warehouse": "ys",
    "site": "ys",
}

_WS_RE = re.compile(r"\s+")


def _normalize_header(name: Any) -> str:
    """Collapse whitespace/newlines, strip trailing punctuation, lowercase.

    Used only for *matching* — the canonical display name is what we keep.
    """
    s = "" if name is None else str(name)
    s = s.replace("\n", " ").replace("\r", " ")
    s = _WS_RE.sub(" ", s).strip()
    s = s.strip(".:;,-_ ").strip()
    return s.lower()


# Pre-normalized canonical lookup: normalized header -> canonical name.
_CANON_BY_NORM: dict[str, str] = {_normalize_header(c): c for c in CANONICAL_COLUMNS}


def resolve_column(header: Any) -> str | None:
    """Return the canonical name for a raw header, or None if unrecognized."""
    norm = _normalize_header(header)
    if norm in _CANON_BY_NORM:
        return _CANON_BY_NORM[norm]
    if norm in _HEADER_ALIASES:
        return _HEADER_ALIASES[norm]
    return None


def load_sap_export(file: Any) -> pd.DataFrame:
    """Read the export, map headers to canonical names, and validate.

    Returns a dataframe whose columns are a subset of CANONICAL_COLUMNS (every
    recognized column present in the file). Unrecognized columns are dropped.
    Raises :class:`DeliveryLoadError` on unreadable / empty / invalid files.
    """
    try:
        sheets = pd.read_excel(file, dtype=str, sheet_name=None)
    except ValueError as exc:
        raise DeliveryLoadError(
            "Could not read the workbook. Make sure it is a valid .xlsx export "
            f"with the delivery data. ({exc})"
        ) from exc
    except Exception as exc:  # noqa: BLE001 - surfaced verbatim to the UI
        raise DeliveryLoadError(f"Could not read the file: {exc}") from exc

    if not sheets:
        raise DeliveryLoadError("The workbook contains no sheets.")

    # A daily raw export has the data on its only/first sheet; a prototype-style
    # workbook has it on a "SAPUI5 Export" sheet alongside report sheets. Pick
    # the sheet whose headers resolve the most REQUIRED columns.
    def _score(frame: pd.DataFrame) -> int:
        if frame is None or frame.shape[1] == 0:
            return 0
        resolved = {resolve_column(c) for c in frame.columns}
        return sum(1 for req in REQUIRED_COLUMNS if req in resolved)

    raw = max(sheets.values(), key=_score)
    if _score(raw) == 0:
        raw = next(iter(sheets.values()))  # fall back to the first sheet

    if raw is None or raw.shape[1] == 0:
        raise DeliveryLoadError("The workbook appears to be empty.")

    # Map each raw column to its canonical name; keep the first occurrence of
    # each canonical (defends against duplicate columns after aliasing).
    rename: dict[Any, str] = {}
    seen: set[str] = set()
    for col in raw.columns:
        canon = resolve_column(col)
        if canon and canon not in seen:
            rename[col] = canon
            seen.add(canon)

    df = raw.rename(columns=rename)
    df = df[[c for c in df.columns if c in seen]].copy()

    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise DeliveryLoadError(
            "This does not look like a SAPUI5 delivery export — required "
            "column(s) missing: " + ", ".join(missing) + ".\n\n"
            "Recognized columns in the uploaded file: "
            + (", ".join(sorted(seen)) if seen else "none") + "."
        )

    # Add any canonical columns the export omitted so downstream code sees a
    # consistent shape. They stay blank.
    for col in CANONICAL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    # Reorder to canonical order.
    df = df[CANONICAL_COLUMNS].copy()

    if df.dropna(how="all").empty:
        raise DeliveryLoadError("The export contains headers but no data rows.")

    return df
