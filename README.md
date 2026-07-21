# Daily Delivery Risk Dashboard

A standalone Streamlit application that turns a daily **SAPUI5 delivery/picking
export** into an actionable risk view: an automatically generated Issue Tracker,
a Not Started report, per-customer detail reports, and one formatted Excel
workbook — all regenerated from scratch on every upload (no state is carried
between days).

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

Then in the browser:

1. **Upload Daily SAPUI5 Export** (`.xlsx`) in the sidebar.
2. Pick an **As of Date** (defaults to today, America/Edmonton) — all overdue,
   deadline and priority calculations use it, so historical testing is
   reproducible.
3. Explore the tabs and **Download** the workbook.

## What it does

- **Cleaning & validation** — header matching by normalized name (tolerates
  extra spaces, newlines, capitalization, truncations), Excel-serial and
  datetime date parsing, identifier preservation, picking→0-100, numeric
  coercion, blank-row removal, and **de-duplication by LE Delivery** so
  Sales Order Total is never double-counted. Unusable records are surfaced in a
  Validation Errors sheet rather than dropped silently.
- **Risk engine** — derives `is_not_started`, `is_late`,
  `is_route_departure_missed`, days-to-deadline fields, `risk_status`
  (Critical / At Risk / In Progress / Ready-Awaiting GI / Not Started /
  Resolved) and `priority`, all relative to the As of Date.
- **Issue Tracker** — one escalation-only row per **warehouse × customer group ×
  issue type**, with deterministic issue text, consequence/risk, owner, earliest
  due date, status, value at risk (affected orders only), order count, and a
  live "Latest Update". An Amazon problem in Calgary never merges with an Amazon
  problem in Mississauga, and a late delivery stays separate from a missed route
  departure.
- **Reports** — Not Started Orders, plus AMZ, Andrew Tab, Andrew Tab 2,
  PFG WHSE, Calgary CO-OP, Sysco, GFS, Pratts and DSD Priorities.
- **Excel export** — organized by warehouse (see below), with professional
  formatting: frozen headers, filters, MM/DD/YYYY dates, Canadian-currency
  totals, status highlighting and per-report totals.

## Workbook structure

The download opens with three combined sheets, then one section per warehouse
**detected in the uploaded export**, then the full export:

```
All Plants Summary            one row per warehouse + an All Plants roll-up
All Plants Issue Tracker      every warehouse, Site column retained
All Plants Not Started
<Warehouse> Issue Tracker     \
<Warehouse> Not Started        |  repeated for each warehouse, in the order
<Warehouse> Orders             |  set by site_order in the YAML
<Warehouse> - AMZ, ...        /   (customer reports, non-empty ones only)
SAPUI5 Export
Validation Warnings           only when a record could not be fully parsed
```

Nothing is hardcoded to a fixed set of warehouses: a new site appearing in
tomorrow's export gets its own section with no code change (sites absent from
`site_order` are appended alphabetically). A warehouse's Not Started sheet is
its deliveries at `Picking in % == 0` whose normalized Goods Issue is not
`Completed`; its Orders sheet is every unique delivery for that warehouse. Each
warehouse's Issue Tracker is exactly the All Plants tracker filtered to that
site, and every count and dollar total is computed over unique `LE Delivery`
values so duplicate records can never inflate a figure.

Excel caps worksheet names at 31 characters, so
[`delivery_dashboard/sheet_names.py`](delivery_dashboard/sheet_names.py) strips
illegal characters, shortens long warehouse/report names on a word boundary
(once per warehouse, so the whole block reads consistently) and guarantees
uniqueness.

## Configuration

All business rules live in [`config/delivery_dashboard_rules.yaml`](config/delivery_dashboard_rules.yaml)
— customer matching patterns, owners, consequence text, detail-report filters
(e.g. PFG WHSE = Save On / Associated Grocers restricted to warehouse
facilities), DSD priority patterns, escalation windows, warehouse labels
(`site_display`) and the order warehouse sections appear in (`site_order`). Edit
the YAML to add customers, add a warehouse or retune thresholds; the Python
engine stays generic.

## Project layout

```
app.py                              # Streamlit entry point (streamlit run app.py)
config/delivery_dashboard_rules.yaml
delivery_dashboard/                 # reusable, UI-free processing engine
    loader.py        # read + validate the export
    cleaner.py       # clean + de-duplicate
    customer_rules.py# YAML rules + classification + warehouse ordering
    risk_engine.py   # derived fields, status, issue type, priority
    report_builder.py# issue tracker, reports, per-warehouse sections, summary
    sheet_names.py   # safe/unique Excel worksheet names
    excel_exporter.py# formatted workbook
tests/test_delivery_dashboard.py
```

## Tests

```bash
python -m pytest -q
```

End-to-end tests validate the engine against a reference export (skipped
automatically if the file is not present).
