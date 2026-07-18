"""Clean and de-duplicate the loaded SAPUI5 export.

Responsibilities (all defensive — never fabricate a value):
  * trim text, preserve identifiers as strings, strip Excel ``.0`` endings
  * parse both Excel serial dates and ordinary datetime strings
  * coerce ``Picking in %`` to 0..100 and other measures to numeric (bad -> 0)
  * drop fully-blank rows
  * quarantine rows that cannot be used at order level (no LE Delivery) into a
    validation-errors frame rather than dropping them silently
  * de-duplicate by LE Delivery, keeping the most complete record so
    Sales Order Total is never double-counted
"""
from __future__ import annotations

import re

import numpy as np
import pandas as pd

# Identifier-like columns: keep as text, strip trailing ".0" from Excel floats,
# never coerce to a number (leading zeros / long ids must survive).
_ID_COLUMNS = [
    "LE Delivery",
    "Sales Order",
    "Outb. Del. Ord.",
    "STO PO Number",
    "Purchase Order",
    "Cust. Reference",
    "Consignee",
    "Customer Dock Reference",
    "Carrier",
]

# Free-text columns: trim only.
_TEXT_COLUMNS = [
    "ys",
    "Facility Code",
    "Carrier Description",
    "Shipment Type",
    "Shpg Cond. Desc",
    "Ship-to Name",
    "Ship-to City",
    "Ship-to Province",
    "Goods Issue",
    "Packing Status",
    "Picking",
    "Picking (Plan)",
    "COD Paid Flag",
    "Shipping Cond.",
]

# Measures coerced to numeric; unparseable -> 0.
_NUMERIC_COLUMNS = [
    "Cases",
    "Line Items",
    "Pallet Quantity",
    "Num. Lifts Pallets",
    "Gross Weight",
    "Sales Order Total",
]

# Date columns parsed leniently (serial or datetime string).
_DATE_COLUMNS = [
    "Warehouse Task Creat",
    "Route Depart. Date",
    "Planned Dlv. Date",
    "Customer Dock Appointment Date",
]

_TRAIL_ZERO_RE = re.compile(r"\.0+$")
# Excel's day-zero (1900 date system, corrected for the 1900 leap-year bug).
_EXCEL_EPOCH = pd.Timestamp("1899-12-30")


def _clean_identifier(value) -> str:
    """Normalize one identifier cell to clean text. Blank/NaN -> ''."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return ""
    # Excel float coercion leaves e.g. "80117458.0" -> "80117458".
    if _TRAIL_ZERO_RE.search(s):
        try:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
        except ValueError:
            pass
    return s


def _parse_dates(series: pd.Series) -> pd.Series:
    """Parse a column that may hold datetime strings OR Excel serial numbers."""
    s = series.astype("string").str.strip()
    # First pass: ordinary datetime strings ("2026-05-18 23:00:00", "05/18/2026").
    parsed = pd.to_datetime(s, errors="coerce", format="mixed")

    # Second pass: cells that are purely numeric are Excel serial day counts.
    still_na = parsed.isna() & s.notna() & (s.str.fullmatch(r"\d+(\.\d+)?"))
    if still_na.any():
        serials = pd.to_numeric(s[still_na], errors="coerce")
        parsed.loc[still_na] = _EXCEL_EPOCH + pd.to_timedelta(serials, unit="D")
    return parsed


def _to_percent(series: pd.Series) -> pd.Series:
    """Coerce Picking in % to a 0..100 float. Bad -> 0. Detects 0..1 fractions."""
    num = pd.to_numeric(series, errors="coerce")
    valid = num.dropna()
    # If the whole column is expressed as a 0..1 fraction, scale to 0..100.
    if not valid.empty and valid.max() <= 1.0 and (valid > 0).any():
        num = num * 100.0
    return num.fillna(0.0).clip(lower=0.0, upper=100.0)


def clean(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Return ``(clean_orders, validation_errors)``.

    ``clean_orders`` is one row per unique LE Delivery, ready for enrichment.
    ``validation_errors`` holds records that could not be used at order level
    (no LE Delivery) plus soft-flagged rows (blank customer / non-numeric
    Sales Order Total), each annotated with a ``Validation Issue`` column.
    """
    df = df.copy()

    # 1) Drop rows that are entirely blank.
    df = df.replace(r"^\s*$", np.nan, regex=True)
    df = df.dropna(how="all").copy()

    # 2) Text + identifier cleaning.
    for col in _TEXT_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v))
                else str(v).strip()
            )
    for col in _ID_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(_clean_identifier)

    # 3) Track whether Sales Order Total was genuinely non-numeric (before we
    #    zero-fill it) so we can flag those records.
    raw_total = df["Sales Order Total"] if "Sales Order Total" in df.columns else pd.Series(dtype=object)
    total_numeric = pd.to_numeric(raw_total, errors="coerce")
    bad_total_mask = raw_total.notna() & (raw_total.astype(str).str.strip() != "") & total_numeric.isna()

    # 4) Numeric coercion (bad -> 0).
    for col in _NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    # 5) Picking percentage.
    if "Picking in %" in df.columns:
        df["Picking in %"] = _to_percent(df["Picking in %"])

    # 6) Dates.
    for col in _DATE_COLUMNS:
        if col in df.columns:
            df[col] = _parse_dates(df[col])

    df = df.reset_index(drop=True)

    # 7) Build the validation-errors frame.
    errors: list[pd.DataFrame] = []

    le = df["LE Delivery"].fillna("").astype(str).str.strip()
    no_delivery_mask = le == ""
    if no_delivery_mask.any():
        miss = df.loc[no_delivery_mask].copy()
        miss.insert(0, "Validation Issue", "Missing LE Delivery — excluded from order calculations")
        errors.append(miss)

    valid = df.loc[~no_delivery_mask].copy()

    blank_cust = valid["Ship-to Name"].fillna("").astype(str).str.strip() == ""
    if blank_cust.any():
        bc = valid.loc[blank_cust].copy()
        bc.insert(0, "Validation Issue", "Blank Ship-to Name — classified as Other")
        errors.append(bc)

    if bad_total_mask.any():
        # Re-align mask onto the surviving rows.
        bt = valid.loc[valid.index.intersection(df.index[bad_total_mask])].copy()
        if not bt.empty:
            bt.insert(0, "Validation Issue", "Sales Order Total was not numeric — treated as $0")
            errors.append(bt)

    validation_errors = (
        pd.concat(errors, ignore_index=True) if errors else pd.DataFrame(columns=["Validation Issue"])
    )

    # 8) De-duplicate by LE Delivery, keeping the most complete / latest record.
    before = len(valid)
    if not valid.empty:
        completeness = valid.notna().sum(axis=1)
        picking = valid.get("Picking in %", pd.Series(0.0, index=valid.index))
        wtc = valid.get("Warehouse Task Creat", pd.Series(pd.NaT, index=valid.index))
        valid = (
            valid.assign(_complete=completeness, _pick=picking, _wtc=wtc)
            .sort_values(["_complete", "_pick", "_wtc"], ascending=[False, False, False])
            .drop_duplicates(subset=["LE Delivery"], keep="first")
            .drop(columns=["_complete", "_pick", "_wtc"])
            .sort_index()
            .reset_index(drop=True)
        )
    duplicates_removed = before - len(valid)
    valid.attrs["duplicates_removed"] = duplicates_removed

    return valid, validation_errors
