"""Tests for the Daily Delivery Risk Dashboard engine.

Covers header normalization, Excel date parsing, identifier cleaning, picking
conversion, LE Delivery de-duplication, customer classification, Not Started /
late logic, Total Price, Amazon deadline logic, Andrew Tab filtering, Excel
report creation, and the "new upload replaces old data" guarantee.

The prototype workbook (Calgary_July_17) is used for end-to-end validation when
present; those checks skip gracefully in environments without the file so the
core unit tests always run.
"""
from __future__ import annotations

import io
from datetime import date
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import Workbook, load_workbook

from delivery_dashboard import cleaner, loader, process_export, risk_engine
from delivery_dashboard.customer_rules import load_ruleset
from delivery_dashboard.excel_exporter import SHEET_ORDER, build_workbook
from delivery_dashboard.loader import CANONICAL_COLUMNS, DeliveryLoadError
from delivery_dashboard.report_builder import (
    build_andrew_tab,
    build_andrew_tab_2,
    build_issue_tracker,
    build_not_started,
)

PROTOTYPE = Path(r"C:/Users/melgh/Downloads/Calgary_July_17 (Prototype).xlsx")
AS_OF = date(2026, 7, 17)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _xlsx(headers: list, rows: list[list], sheet="SAPUI5 Export") -> io.BytesIO:
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws.append(headers)
    for r in rows:
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# A minimal but complete synthetic export builder — one row = kwargs override.
_DEFAULTS = {
    "ys": "Calgary Warehouse", "Facility Code": "Warehouse", "Warehouse Task Creat": "Y",
    "Picking in %": "0", "Route Depart. Date": "2026-07-20 23:00:00",
    "Planned Dlv. Date": "2026-07-22 23:00:00", "LE Delivery": "80100000",
    "Ship-to Name": "SOME CUSTOMER", "Carrier Description": "VERSACOLD LOGISTICS",
    "Cases": "10", "Line Items": "2", "Goods Issue": "Not Started",
    "Sales Order": "3000000", "Sales Order Total": "1000",
    "Customer Dock Appointment Date": None,
}


def _rows(*overrides: dict) -> io.BytesIO:
    headers = list(_DEFAULTS.keys())
    rows = []
    for i, ov in enumerate(overrides):
        rec = dict(_DEFAULTS)
        rec["LE Delivery"] = str(80100000 + i)
        rec.update(ov)
        rows.append([rec[h] for h in headers])
    return _xlsx(headers, rows)


def _process(buf, as_of=AS_OF):
    return process_export(buf, as_of)


# ---------------------------------------------------------------------------
# Header normalization
# ---------------------------------------------------------------------------
def test_header_normalization_tolerates_spacing_case_newlines():
    assert loader.resolve_column("  ship-to name ") == "Ship-to Name"
    assert loader.resolve_column("SHIP-TO NAME") == "Ship-to Name"
    assert loader.resolve_column("Planned Dlv. Date\n") == "Planned Dlv. Date"
    assert loader.resolve_column("Picking in %") == "Picking in %"
    assert loader.resolve_column("Route Departure Date") == "Route Depart. Date"
    assert loader.resolve_column("totally unknown column") is None


def test_missing_required_columns_raise():
    buf = _xlsx(["ys", "LE Delivery", "Cases"], [["Calgary", "1", "3"]])
    with pytest.raises(DeliveryLoadError):
        loader.load_sap_export(buf)


def test_empty_workbook_raises():
    buf = _xlsx(list(_DEFAULTS.keys()), [])
    with pytest.raises(DeliveryLoadError):
        loader.load_sap_export(buf)


# ---------------------------------------------------------------------------
# Cleaning
# ---------------------------------------------------------------------------
def test_identifier_cleaning_strips_dot_zero_and_keeps_string():
    df = loader.load_sap_export(_rows({"LE Delivery": "80117458.0", "Sales Order": "3075382.0"}))
    clean, _ = cleaner.clean(df)
    assert clean["LE Delivery"].iloc[0] == "80117458"
    assert clean["Sales Order"].iloc[0] == "3075382"


def test_excel_serial_date_parsing():
    # 46220 == 2026-07-17 in the Excel 1900 date system (epoch 1899-12-30).
    df = loader.load_sap_export(_rows({"Planned Dlv. Date": "46220"}))
    clean, _ = cleaner.clean(df)
    assert pd.Timestamp(clean["Planned Dlv. Date"].iloc[0]).date() == date(2026, 7, 17)


def test_datetime_string_date_parsing():
    df = loader.load_sap_export(_rows({"Planned Dlv. Date": "2026-07-22 23:00:00"}))
    clean, _ = cleaner.clean(df)
    assert pd.Timestamp(clean["Planned Dlv. Date"].iloc[0]).date() == date(2026, 7, 22)


def test_picking_percentage_conversion():
    df = loader.load_sap_export(_rows(
        {"Picking in %": "70"}, {"Picking in %": "not-a-number"}, {"Picking in %": "150"}))
    clean, _ = cleaner.clean(df)
    pct = clean.sort_values("LE Delivery")["Picking in %"].tolist()
    assert pct[0] == 70.0
    assert pct[1] == 0.0        # invalid -> 0
    assert pct[2] == 100.0      # clipped


def test_picking_fraction_column_is_scaled():
    df = loader.load_sap_export(_rows(
        {"Picking in %": "0.7"}, {"Picking in %": "1"}, {"Picking in %": "0"}))
    clean, _ = cleaner.clean(df)
    assert clean["Picking in %"].max() == 100.0  # 0..1 column scaled to 0..100


def test_bad_sales_order_total_becomes_zero_and_is_flagged():
    # "TBD" is non-numeric junk that pandas does NOT auto-treat as NA.
    df = loader.load_sap_export(_rows({"Sales Order Total": "TBD"}))
    clean, errors = cleaner.clean(df)
    assert clean["Sales Order Total"].iloc[0] == 0.0
    assert (errors["Validation Issue"].str.contains("Sales Order Total")).any()


def test_blank_rows_removed_and_missing_delivery_quarantined():
    df = loader.load_sap_export(_rows(
        {"LE Delivery": "80100001"},
        {"LE Delivery": ""},   # no delivery -> quarantined
    ))
    clean, errors = cleaner.clean(df)
    assert len(clean) == 1
    assert (errors["Validation Issue"].str.contains("Missing LE Delivery")).any()


def test_dedup_by_le_delivery_no_double_count():
    df = loader.load_sap_export(_rows(
        {"LE Delivery": "80100001", "Sales Order Total": "1000", "Picking in %": "0"},
        {"LE Delivery": "80100001", "Sales Order Total": "1000", "Picking in %": "50"},
    ))
    clean, _ = cleaner.clean(df)
    assert len(clean) == 1
    assert clean["Sales Order Total"].sum() == 1000.0


# ---------------------------------------------------------------------------
# Customer classification
# ---------------------------------------------------------------------------
@pytest.mark.parametrize("ship,carrier,expected", [
    ("AMAZON CANADA FULFILLMENT", "VERSACOLD", "amazon"),
    ("SOME STORE", "AMAZON PICKUP", "amazon"),
    ("AMZ WAREHOUSE 3", "VERSACOLD", "amazon"),
    ("GFS EDMONTON", "VERSACOLD", "gfs"),
    ("A&W GFS CALGARY", "GFS CAPSTONE PICKUP", "aw_via_gfs"),
    ("SYSCO CALGARY", "VERSACOLD", "sysco"),
    ("CALGARY CO-OP 704", "VERSACOLD", "calgary_coop"),
    ("SAVE ON C/O TCL SUPPLY", "SAVE ON PICKUP", "save_on"),
    ("ASSOCIATED GROCERS 9221 WHS", "VERSACOLD", "associated_grocers"),
    ("FCL 4840 WAREHOUSE", "VERSACOLD", "federated_coop"),
    ("SOBEYS 3012 MILLWOODS", "VERSACOLD", "sobeys_group"),
    ("CANADA SAFEWAY 8830", "VERSACOLD", "sobeys_group"),
    ("PRATT'S WHOLESALE LTD", "VERIGO", "pratts"),
    ("UNKNOWN CORNER STORE", "SCOOPS", "other"),
])
def test_customer_classification(ship, carrier, expected):
    rules = load_ruleset()
    df = loader.load_sap_export(_rows({"Ship-to Name": ship, "Carrier Description": carrier}))
    clean, _ = cleaner.clean(df)
    classified = rules.classify(clean)
    assert classified["customer_group"].iloc[0] == expected


def test_aw_matches_before_gfs():
    # A&W must win over the generic GFS group (listed first in config).
    rules = load_ruleset()
    df = loader.load_sap_export(_rows({"Ship-to Name": "A&W GFS CALGARY"}))
    clean, _ = cleaner.clean(df)
    assert rules.classify(clean)["customer_group"].iloc[0] == "aw_via_gfs"


# ---------------------------------------------------------------------------
# Risk / Not Started / late logic
# ---------------------------------------------------------------------------
def test_not_started_logic():
    df = loader.load_sap_export(_rows(
        {"Picking in %": "0", "Goods Issue": "Not Started"},      # not started
        {"Picking in %": "0", "Goods Issue": "Completed"},        # GI complete -> not
        {"Picking in %": "30", "Goods Issue": "Not Started"},     # in progress -> not
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    assert enriched.sort_values("LE Delivery")["is_not_started"].tolist() == [True, False, False]


def test_late_logic():
    df = loader.load_sap_export(_rows(
        {"Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Not Started"},  # late
        {"Planned Dlv. Date": "2026-07-25 23:00:00", "Goods Issue": "Not Started"},  # future
        {"Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Completed"},    # done -> not late
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    assert enriched.sort_values("LE Delivery")["is_late"].tolist() == [True, False, False]


def test_route_departure_missed_and_status_critical():
    df = loader.load_sap_export(_rows(
        {"Route Depart. Date": "2026-07-10 23:00:00", "Picking in %": "40",
         "Planned Dlv. Date": "2026-07-25 23:00:00"},
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    assert bool(enriched["is_route_departure_missed"].iloc[0]) is True
    assert enriched["risk_status"].iloc[0] == risk_engine.CRITICAL


def test_ready_awaiting_gi_status():
    df = loader.load_sap_export(_rows(
        {"Picking in %": "100", "Goods Issue": "Not Started",
         "Planned Dlv. Date": "2026-07-25 23:00:00", "Route Depart. Date": "2026-07-24 23:00:00"},
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    assert enriched["risk_status"].iloc[0] == risk_engine.READY


# ---------------------------------------------------------------------------
# Amazon deadline logic
# ---------------------------------------------------------------------------
def test_amazon_deadline_passed_is_critical():
    df = loader.load_sap_export(_rows(
        {"Ship-to Name": "AMAZON CANADA", "Carrier Description": "AMAZON PICKUP",
         "Customer Dock Appointment Date": "2026-07-10 00:00:00",
         "Planned Dlv. Date": "2026-07-25 23:00:00", "Route Depart. Date": "2026-07-24 23:00:00",
         "Picking in %": "50"},
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    assert enriched["risk_status"].iloc[0] == risk_engine.CRITICAL
    assert enriched["days_to_customer_deadline"].iloc[0] < 0


def test_amazon_deadline_within_five_days_is_at_risk():
    df = loader.load_sap_export(_rows(
        {"Ship-to Name": "AMAZON CANADA", "Carrier Description": "AMAZON PICKUP",
         "Customer Dock Appointment Date": "2026-07-21 00:00:00",  # 4 days out
         "Planned Dlv. Date": "2026-07-28 23:00:00", "Route Depart. Date": "2026-07-27 23:00:00",
         "Picking in %": "80"},
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    assert enriched["risk_status"].iloc[0] == risk_engine.AT_RISK


# ---------------------------------------------------------------------------
# Andrew Tab filtering
# ---------------------------------------------------------------------------
def test_andrew_tab_filter():
    df = loader.load_sap_export(_rows(
        {"Facility Code": "Warehouse", "Planned Dlv. Date": "2026-07-10 23:00:00", "Picking in %": "40"},   # in
        {"Facility Code": "Warehouse", "Planned Dlv. Date": "2026-07-10 23:00:00", "Picking in %": "100"},  # out (100%)
        {"Facility Code": "DSD", "Planned Dlv. Date": "2026-07-10 23:00:00", "Picking in %": "40"},         # out (facility)
        {"Facility Code": "Warehouse", "Planned Dlv. Date": "2026-07-25 23:00:00", "Picking in %": "40"},   # out (future)
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    tab = build_andrew_tab(enriched, AS_OF)
    assert len(tab) == 1
    assert list(tab.columns) == [
        "Picking in %", "Planned Dlv. Date", "LE Delivery", "Ship-to Name",
        "Ship-to City", "Ship-to Province", "Cases", "Line Items", "Sales Order Total"]


def test_andrew_tab_2_normalizes_spelling_and_filters():
    df = loader.load_sap_export(_rows(
        {"Facility Code": "Independant", "Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Not Started"},  # in
        {"Facility Code": "Independent", "Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Not Started"},  # in
        {"Facility Code": "DSD", "Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Not Started"},          # in
        {"Facility Code": "DSD", "Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Completed"},            # out
        {"Facility Code": "Warehouse", "Planned Dlv. Date": "2026-07-10 23:00:00", "Goods Issue": "Not Started"},    # out
    ))
    rules = load_ruleset()
    clean, _ = cleaner.clean(df)
    enriched = risk_engine.enrich(rules.classify(clean), AS_OF, rules)
    tab = build_andrew_tab_2(enriched, AS_OF)
    assert len(tab) == 3


# ---------------------------------------------------------------------------
# Total Price (issue tracker sums only affected orders, dedup by LE Delivery)
# ---------------------------------------------------------------------------
def test_issue_tracker_total_price_counts_only_escalating():
    buf = _rows(
        {"Ship-to Name": "SYSCO CALGARY", "Planned Dlv. Date": "2026-07-10 23:00:00",
         "Goods Issue": "Not Started", "Sales Order Total": "500"},                      # late -> escalating
        {"Ship-to Name": "SYSCO EDMONTON", "Planned Dlv. Date": "2026-08-30 23:00:00",
         "Route Depart. Date": "2026-08-29 23:00:00", "Goods Issue": "Not Started",
         "Sales Order Total": "999"},                                                    # far future -> not escalating
    )
    result = _process(buf)
    sysco_row = result.issue_tracker[result.issue_tracker["Customer"] == "Sysco"]
    assert len(sysco_row) == 1
    assert sysco_row["Total Price"].iloc[0] == 500.0   # 999 order excluded
    assert sysco_row["Order Count"].iloc[0] == 1


# ---------------------------------------------------------------------------
# Excel report creation
# ---------------------------------------------------------------------------
def test_excel_workbook_has_sheets_in_order():
    df = _rows({"Ship-to Name": "AMAZON CANADA", "Carrier Description": "AMAZON PICKUP",
                "Planned Dlv. Date": "2026-07-10 23:00:00"})
    result = _process(df)
    xlsx = build_workbook(result)
    wb = load_workbook(io.BytesIO(xlsx))
    for name in SHEET_ORDER:
        assert name in wb.sheetnames
    # Order of the canonical sheets is preserved.
    positions = [wb.sheetnames.index(n) for n in SHEET_ORDER]
    assert positions == sorted(positions)


def test_new_upload_replaces_old_data():
    """Processing a second file must not retain anything from the first."""
    first = _process(_rows({"Ship-to Name": "SYSCO CALGARY",
                            "Planned Dlv. Date": "2026-07-10 23:00:00", "Sales Order Total": "500"}))
    assert "Sysco" in first.issue_tracker["Customer"].values

    second = _process(_rows({"Ship-to Name": "PRATT'S WHOLESALE",
                             "Planned Dlv. Date": "2026-07-10 23:00:00", "Sales Order Total": "800"}))
    # No Sysco carried over; only Pratts present.
    assert "Sysco" not in second.issue_tracker["Customer"].values
    assert "Pratts" in second.issue_tracker["Customer"].values
    assert len(second.orders) == 1


# ---------------------------------------------------------------------------
# End-to-end validation against the prototype (skips if file absent)
# ---------------------------------------------------------------------------
@pytest.mark.skipif(not PROTOTYPE.exists(), reason="prototype workbook not available")
@pytest.mark.parametrize("report,rows,total", [
    ("AMZ", 12, 240636.72),
    ("PFG WHSE", 14, 256270.89),
    ("GFS", 10, None),
    ("Sysco", 11, None),
    ("Calgary CO-OP", 7, None),
    ("Pratts", 4, None),
])
def test_prototype_report_counts(report, rows, total):
    result = process_export(str(PROTOTYPE), AS_OF)
    df = result.reports[report]
    assert len(df) == rows, f"{report}: expected {rows} rows, got {len(df)}"
    if total is not None:
        got = float(pd.to_numeric(df["Sales Order Total"], errors="coerce").fillna(0).sum())
        assert abs(got - total) < 0.05, f"{report}: expected ${total}, got ${got}"


@pytest.mark.skipif(not PROTOTYPE.exists(), reason="prototype workbook not available")
def test_prototype_dedup_and_workbook_builds():
    result = process_export(str(PROTOTYPE), AS_OF)
    assert len(result.orders) == result.orders["LE Delivery"].nunique()
    xlsx = build_workbook(result)
    assert len(xlsx) > 10_000
    wb = load_workbook(io.BytesIO(xlsx))
    assert "Issue Tracker" in wb.sheetnames
