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
- **Issue Tracker** — one escalation-only row per customer group, with
  deterministic issue text, consequence/risk, owner, earliest due date, status,
  value at risk (affected orders only), order count, and a live "Latest Update".
- **Reports** — Not Started Orders, plus AMZ, Andrew Tab, Andrew Tab 2,
  PFG WHSE, Calgary CO-OP, Sysco, GFS, Pratts and DSD Priorities.
- **Excel export** — professional formatting: frozen headers, filters,
  MM/DD/YYYY dates, Canadian-currency totals, status highlighting, per-report
  totals, in the required sheet order.

## Configuration

All business rules live in [`config/delivery_dashboard_rules.yaml`](config/delivery_dashboard_rules.yaml)
— customer matching patterns, owners, consequence text, detail-report filters
(e.g. PFG WHSE = Save On / Associated Grocers restricted to warehouse
facilities), DSD priority patterns, and escalation windows. Edit the YAML to add
customers or retune thresholds; the Python engine stays generic.

## Project layout

```
app.py                              # Streamlit entry point (streamlit run app.py)
config/delivery_dashboard_rules.yaml
delivery_dashboard/                 # reusable, UI-free processing engine
    loader.py        # read + validate the export
    cleaner.py       # clean + de-duplicate
    customer_rules.py# YAML rules + classification
    risk_engine.py   # derived fields, status, priority
    report_builder.py# issue tracker + all reports
    excel_exporter.py# formatted workbook
tests/test_delivery_dashboard.py
```

## Tests

```bash
python -m pytest -q
```

End-to-end tests validate the engine against a reference export (skipped
automatically if the file is not present).
