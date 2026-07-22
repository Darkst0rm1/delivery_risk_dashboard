"""Build the Issue Tracker, Not Started report, and all customer detail reports.

Every output is regenerated from the enriched order frame — no state is carried
between uploads. Issue text and "Latest Update" strings are deterministic
functions of the current data (plus the as_of / now timestamps), never
hard-coded from the prototype.

The workbook is organized by warehouse: :func:`build_site_sections` slices the
orders per site, and every combined ("All Plants") output is the union of those
slices, so a warehouse sheet and the combined sheet can never disagree.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Callable
from zoneinfo import ZoneInfo

import pandas as pd

from . import risk_engine as R
from .customer_rules import Ruleset
from .loader import CANONICAL_COLUMNS

EDMONTON = ZoneInfo("America/Edmonton")

# Issue Tracker column order (must match the Excel sheet exactly).
ISSUE_TRACKER_COLUMNS = [
    "Site", "Customer", "Issue Type", "Issue", "Consequence / Risk", "Owner",
    "Due Date", "Status", "Total Price", "Order Count", "Latest Update", "Detail Report",
]

# All Plants Summary column order.
SUMMARY_COLUMNS = [
    "Site", "Total Orders", "Open Orders", "Critical", "At Risk", "Not Started",
    "Late", "Route Departure Missed", "Issues",
]

# Label of the roll-up row at the bottom of the All Plants Summary.
ALL_PLANTS = "All Plants"

# "Late orders — picking" block, written below the All Plants Summary table.
# The three picking buckets partition the late orders exactly (>=100 / 0<x<100
# / 0), so the counts and values always add back to Late Orders / Late Value.
LATE_PICKING_COLUMNS = [
    "Site", "Late Orders", "Late Value",
    "Completed", "Completed $", "Partial", "Partial $", "Not Started", "Not Started $",
]

# One-line headline block: the whole book, count and value.
ORDER_TOTALS_COLUMNS = ["Total Orders", "Total Order Value"]

# Not Started report column order.
NOT_STARTED_COLUMNS = [
    "Priority", "Site", "Customer Group", "Facility Code", "Warehouse Task Creat",
    "Route Depart. Date", "Planned Dlv. Date", "LE Delivery", "Cust. Reference",
    "Carrier Description", "Ship-to Name", "Ship-to City", "Ship-to Province",
    "Cases", "Line Items", "Picking in %", "Picking Status", "Goods Issue",
    "Sales Order", "Sales Order Total", "Days to Route Departure", "Days to Planned Delivery",
]

# The nine columns shared by Andrew Tab / Andrew Tab 2.
ANDREW_COLUMNS = [
    "Picking in %", "Planned Dlv. Date", "LE Delivery", "Ship-to Name",
    "Ship-to City", "Ship-to Province", "Cases", "Line Items", "Sales Order Total",
]


def now_edmonton() -> datetime:
    """Current wall-clock time in America/Edmonton."""
    return datetime.now(EDMONTON)


def _fmt_date(ts) -> str:
    if ts is None or pd.isna(ts):
        return "N/A"
    return pd.Timestamp(ts).strftime("%b %d, %Y")


def _fmt_mmdd(ts) -> str:
    if ts is None or pd.isna(ts):
        return ""
    return pd.Timestamp(ts).strftime("%m/%d")


def _unique_orders(df: pd.DataFrame) -> pd.DataFrame:
    """One row per LE Delivery.

    The cleaner already de-duplicates, so this is a guard rather than a fix:
    every count and every dollar total in the warehouse sheets goes through it
    so a duplicate can never inflate a figure.
    """
    if df is None or df.empty or "LE Delivery" not in df.columns:
        return df
    return df.drop_duplicates(subset=["LE Delivery"])


def _due_dates(df: pd.DataFrame, rules: Ruleset) -> pd.Series:
    """Per-order 'earliest relevant deadline' timestamp.

    Uses the Amazon dock appointment (cancellation deadline) when applicable,
    plus the route departure and planned delivery dates — whichever is earliest.
    """
    planned = pd.to_datetime(df.get("Planned Dlv. Date"), errors="coerce")
    route = pd.to_datetime(df.get("Route Depart. Date"), errors="coerce")
    dock = pd.to_datetime(df.get("Customer Dock Appointment Date"), errors="coerce")
    is_amazon = df["customer_group"].map(lambda k: rules.rule(k).amazon_deadline)
    dock = dock.where(is_amazon)  # only Amazon uses the dock date as a deadline
    stacked = pd.concat([planned, route, dock], axis=1)
    return stacked.min(axis=1)


# ---------------------------------------------------------------------------
# Issue text (per issue type, deterministic)
# ---------------------------------------------------------------------------
# Tracker rows are now split by issue type, so the sentence is driven by the
# issue type and names the customer, rather than one per-customer paragraph
# that would report "0 late, 0 not started" on a deadline-only row. The
# customer-specific angle survives in the Consequence / Risk column, which
# still comes straight from the YAML rules.
def _counts(affected: pd.DataFrame) -> dict:
    picking = pd.to_numeric(affected.get("picking_pct"), errors="coerce").fillna(0)
    return {
        "count": int(affected["LE Delivery"].nunique()),
        "not_started": int(affected["is_not_started"].sum()),
        "late": int(affected["is_late"].sum()),
        "in_progress": int(affected["is_picking_in_progress"].sum()),
        "missed_route": int(affected["is_route_departure_missed"].sum()),
        "unpicked": int((picking < 100).sum()),
    }


def _earliest_planned(affected: pd.DataFrame) -> str:
    planned = pd.to_datetime(affected.get("Planned Dlv. Date"), errors="coerce").dropna()
    return _fmt_date(planned.min()) if not planned.empty else "N/A"


def _nearest_deadline(affected: pd.DataFrame) -> str:
    dock = pd.to_datetime(affected.get("Customer Dock Appointment Date"), errors="coerce").dropna()
    return _fmt_date(dock.min()) if not dock.empty else "N/A"


def _lowest_picked(affected: pd.DataFrame) -> float:
    pct = pd.to_numeric(affected.get("picking_pct"), errors="coerce").dropna()
    return float(pct.min()) if not pct.empty else 0.0


def _delivery_ids(affected: pd.DataFrame, limit: int = 5) -> str:
    """Name the affected deliveries when there are few enough to act on."""
    ids = sorted({str(x) for x in affected["LE Delivery"].dropna() if str(x).strip()})
    if not ids or len(ids) > limit:
        return ""
    return " Affected LE Deliveries: " + ", ".join(ids) + "."


def _issue_deadline_passed(a: pd.DataFrame, customer: str) -> str:
    return (f"{_counts(a)['count']} {customer} order(s) are past the customer cancellation "
            f"deadline (earliest {_nearest_deadline(a)}).")


def _issue_late(a: pd.DataFrame, customer: str) -> str:
    return (f"{_counts(a)['count']} {customer} order(s) are past the planned delivery date "
            f"(earliest {_earliest_planned(a)}) with Goods Issue still open.")


def _issue_route_missed(a: pd.DataFrame, customer: str) -> str:
    return (f"{_counts(a)['count']} {customer} order(s) missed the route departure window and "
            f"are not fully picked (lowest {_lowest_picked(a):.0f}% picked).")


def _issue_deadline_soon(a: pd.DataFrame, customer: str) -> str:
    c = _counts(a)
    return (f"{c['count']} {customer} order(s) face a cancellation deadline on or before "
            f"{_nearest_deadline(a)} and {c['unpicked']} are not fully picked.")


def _issue_not_started(a: pd.DataFrame, customer: str) -> str:
    return (f"{_counts(a)['count']} {customer} order(s) are still at 0% picked. "
            f"Earliest planned delivery {_earliest_planned(a)}.")


def _issue_picking_behind(a: pd.DataFrame, customer: str) -> str:
    return (f"{_counts(a)['count']} {customer} order(s) are only partially picked (lowest "
            f"{_lowest_picked(a):.0f}%) with a deadline approaching. Earliest planned "
            f"delivery {_earliest_planned(a)}.")


def _issue_awaiting_gi(a: pd.DataFrame, customer: str) -> str:
    return (f"{_counts(a)['count']} {customer} order(s) are fully picked and awaiting Goods "
            f"Issue. Earliest planned delivery {_earliest_planned(a)}.")


def _issue_general(a: pd.DataFrame, customer: str) -> str:
    c = _counts(a)
    return (f"{c['count']} {customer} order(s) require attention: {c['not_started']} not "
            f"started, {c['late']} late, {c['in_progress']} in progress. Earliest planned "
            f"delivery: {_earliest_planned(a)}.")


_ISSUE_BUILDERS: dict[str, Callable[[pd.DataFrame, str], str]] = {
    R.IT_DEADLINE_PASSED: _issue_deadline_passed,
    R.IT_LATE: _issue_late,
    R.IT_ROUTE_MISSED: _issue_route_missed,
    R.IT_DEADLINE_SOON: _issue_deadline_soon,
    R.IT_NOT_STARTED: _issue_not_started,
    R.IT_PICKING_BEHIND: _issue_picking_behind,
    R.IT_AWAITING_GI: _issue_awaiting_gi,
}


def _issue_text(issue_type: str, affected: pd.DataFrame, customer: str) -> str:
    builder = _ISSUE_BUILDERS.get(issue_type, _issue_general)
    return builder(affected, customer) + _delivery_ids(affected)


def _latest_update(affected: pd.DataFrame, rules: Ruleset, now: datetime) -> str:
    late = int(affected["is_late"].sum())
    not_started = int(affected["is_not_started"].sum())
    value = float(affected["Sales Order Total"].sum())
    due = _due_dates(affected, rules).dropna()
    nearest = _fmt_date(due.min()) if not due.empty else "N/A"
    # %-I (no leading-zero hour) is not portable to Windows; format 12-hour
    # time manually so the string reads e.g. "Jul 17, 2026 at 8:15 AM".
    hour12 = now.hour % 12 or 12
    stamp = f"{now.strftime('%b %d, %Y')} at {hour12}:{now.minute:02d} {now.strftime('%p')}"
    return (f"As of {stamp}: {late} orders are late, {not_started} are not started and the "
            f"current value at risk is ${value:,.0f}. The nearest due date is {nearest}.")


# ---------------------------------------------------------------------------
# Issue Tracker
# ---------------------------------------------------------------------------
def build_issue_tracker(
    orders: pd.DataFrame,
    rules: Ruleset,
    as_of: date,
    now: datetime | None = None,
    include_non_escalated: bool = False,
) -> pd.DataFrame:
    """One summarized row per Site x Customer Group x Issue Type.

    Splitting on all three is what keeps an Amazon problem in Calgary on its
    own row rather than merged with an Amazon problem in Mississauga, and a
    late delivery separate from a missed route departure. It also means a
    warehouse tracker is exactly the subset of this tracker for that site.

    By default only groups with at least one Critical/At Risk order appear
    (escalation-only). ``include_non_escalated=True`` also lists groups whose
    orders are merely In Progress / Ready.
    """
    now = now or now_edmonton()
    cols = ISSUE_TRACKER_COLUMNS
    if orders is None or orders.empty:
        return pd.DataFrame(columns=cols)

    if include_non_escalated:
        wanted = orders[orders["is_open"]].copy()
    else:
        wanted = orders[orders["risk_status"].isin(R.ESCALATION_STATUSES)].copy()
    if wanted.empty:
        return pd.DataFrame(columns=cols)

    rows = []
    for (site, key, issue_type), grp in wanted.groupby(
        ["site", "customer_group", "issue_type"], sort=False,
    ):
        rule = rules.rule(key)
        affected = _unique_orders(grp)
        has_critical = (affected["risk_status"] == R.CRITICAL).any()
        status = R.CRITICAL if has_critical else (
            R.AT_RISK if (affected["risk_status"] == R.AT_RISK).any() else affected["risk_status"].iloc[0]
        )
        due = _due_dates(affected, rules).dropna()
        report_mode = affected["customer_report"].mode()
        rows.append({
            "Site": site,
            "Customer": rule.display,
            "Issue Type": issue_type,
            "Issue": _issue_text(issue_type, affected, rule.display),
            "Consequence / Risk": rule.consequence or rules.default_consequence,
            "Owner": rule.owner_or_unassigned(),
            "Due Date": due.min() if not due.empty else pd.NaT,
            "Status": status,
            "Total Price": float(affected["Sales Order Total"].sum()),
            "Order Count": int(affected["LE Delivery"].nunique()),
            "Latest Update": _latest_update(affected, rules, now),
            "Detail Report": report_mode.iloc[0] if not report_mode.empty else "",
            "_status_rank": 0 if has_critical else 1,
            "_issue_rank": R.issue_type_rank(issue_type),
        })

    out = pd.DataFrame(rows)
    # Stable so that filtering the All Plants tracker down to one site gives
    # the same row order as building that site's tracker on its own.
    out = out.sort_values(
        ["_status_rank", "_issue_rank", "Total Price"], ascending=[True, True, False],
        kind="stable",
    )
    out = out.drop(columns=["_status_rank", "_issue_rank"]).reset_index(drop=True)
    return out[cols]


# ---------------------------------------------------------------------------
# Not Started Orders
# ---------------------------------------------------------------------------
def build_not_started(orders: pd.DataFrame) -> pd.DataFrame:
    """Orders at 0% picking with Goods Issue not completed, ranked by priority."""
    if orders is None or orders.empty:
        return pd.DataFrame(columns=NOT_STARTED_COLUMNS)

    ns = orders[orders["is_not_started"]].copy()
    if ns.empty:
        return pd.DataFrame(columns=NOT_STARTED_COLUMNS)

    ns["Priority"] = ns["priority"]
    ns["Site"] = ns["site"]
    ns["Customer Group"] = ns["customer_display"]
    ns["Picking Status"] = ns.get("Picking", "")
    ns["Days to Route Departure"] = ns["days_to_route_departure"]
    ns["Days to Planned Delivery"] = ns["days_to_planned_delivery"]

    ns = ns.sort_values(
        by=["priority_rank", "Planned Dlv. Date", "Sales Order Total"],
        ascending=[True, True, False],
        na_position="last",
    ).reset_index(drop=True)

    for col in NOT_STARTED_COLUMNS:
        if col not in ns.columns:
            ns[col] = pd.NA
    return ns[NOT_STARTED_COLUMNS]


# ---------------------------------------------------------------------------
# Customer detail reports
# ---------------------------------------------------------------------------
def _canonical_slice(df: pd.DataFrame) -> pd.DataFrame:
    """Return the original business columns only, in canonical order."""
    cols = [c for c in CANONICAL_COLUMNS if c in df.columns]
    return df[cols].copy()


def _amazon_comment(row) -> str:
    dock = row.get("Customer Dock Appointment Date")
    if dock is not None and not pd.isna(dock):
        return f"Cancellation deadline {_fmt_mmdd(dock)}"
    return ""


def _build_customer_report(orders: pd.DataFrame, spec, rules: Ruleset) -> pd.DataFrame:
    sub = orders[orders["customer_group"].isin(spec.customers)].copy()
    if spec.facility_in:
        facs = {f.strip().upper() for f in spec.facility_in}
        sub = sub[sub["Facility Code"].fillna("").astype(str).str.strip().str.upper().isin(facs)]
    body = _canonical_slice(sub)
    if spec.comments_column:
        if spec.amazon_deadline:
            comments = sub["Customer Dock Appointment Date"].map(
                lambda d: f"Cancellation deadline {_fmt_mmdd(d)}" if pd.notna(d) else ""
            )
        else:
            comments = pd.Series("", index=body.index)
        body.insert(0, "Comments", comments.values)
    body = body.reset_index(drop=True)
    # Flags the exporter reads instead of matching on sheet name — the sheets
    # are now called e.g. "Calgary - AMZ", so name matching would miss them.
    body.attrs["is_customer_report"] = True
    body.attrs["amazon_note"] = bool(spec.amazon_deadline)
    return body


def build_andrew_tab(orders: pd.DataFrame, as_of: date) -> pd.DataFrame:
    as_of_ts = pd.Timestamp(as_of)
    fac = orders["Facility Code"].fillna("").astype(str).str.strip()
    planned = pd.to_datetime(orders.get("Planned Dlv. Date"), errors="coerce")
    picking = pd.to_numeric(orders.get("Picking in %"), errors="coerce").fillna(0)
    mask = (fac == "Warehouse") & (planned < as_of_ts) & (picking < 100)
    sub = orders[mask].copy()
    sub = sub.sort_values(
        by=["Planned Dlv. Date", "Picking in %", "Sales Order Total"],
        ascending=[True, True, False], na_position="last",
    )
    return sub[[c for c in ANDREW_COLUMNS if c in sub.columns]].reset_index(drop=True)


def build_andrew_tab_2(orders: pd.DataFrame, as_of: date) -> pd.DataFrame:
    as_of_ts = pd.Timestamp(as_of)
    fac = orders["Facility Code"].fillna("").astype(str).str.strip().str.title()
    # Normalize both spellings Independant / Independent.
    fac_norm = fac.replace({"Independant": "Independent"})
    planned = pd.to_datetime(orders.get("Planned Dlv. Date"), errors="coerce")
    gi = orders["Goods Issue"].fillna("").astype(str).str.strip().str.upper()
    mask = fac_norm.isin(["Dsd", "Independent"]) & (planned < as_of_ts) & (gi != "COMPLETED")
    sub = orders[mask].copy()
    sub = sub.sort_values(
        by=["Planned Dlv. Date", "Picking in %", "Sales Order Total"],
        ascending=[True, True, False], na_position="last",
    )
    return sub[[c for c in ANDREW_COLUMNS if c in sub.columns]].reset_index(drop=True)


def build_dsd_priorities(orders: pd.DataFrame, rules: Ruleset) -> pd.DataFrame:
    """Orders whose Ship-to Name matches a configured DSD priority pattern."""
    prio = orders["Ship-to Name"].map(rules.dsd_priority_for)
    sub = orders[prio.notna()].copy()
    if sub.empty:
        return pd.DataFrame(columns=["Priority"] + [c for c in CANONICAL_COLUMNS if c in orders.columns])
    sub["Priority"] = prio[prio.notna()].values
    sub["_prio_rank"] = sub["Priority"].map(lambda p: 1 if "1" in str(p) else 2)
    sub = sub.sort_values(by=["_prio_rank", "Planned Dlv. Date"], ascending=[True, True], na_position="last")
    body = _canonical_slice(sub)
    body.insert(0, "Priority", sub["Priority"].values)
    return body.reset_index(drop=True)


def build_detail_reports(orders: pd.DataFrame, rules: Ruleset, as_of: date) -> dict[str, pd.DataFrame]:
    """Return every detail report keyed by sheet name, in workbook order."""
    reports: dict[str, pd.DataFrame] = {}

    # Always-present base sheets.
    reports["SAPUI5 Export"] = _canonical_slice(orders)
    reports["DSD Priorities"] = build_dsd_priorities(orders, rules)

    for spec in rules.detail_reports:
        if spec.special == "andrew_tab":
            reports[spec.name] = build_andrew_tab(orders, as_of)
        elif spec.special == "andrew_tab_2":
            reports[spec.name] = build_andrew_tab_2(orders, as_of)
        elif spec.customers:
            reports[spec.name] = _build_customer_report(orders, spec, rules)
    return reports


# ---------------------------------------------------------------------------
# Per-warehouse sections
# ---------------------------------------------------------------------------
@dataclass
class SiteSection:
    """Every sheet one warehouse contributes to the workbook."""

    site: str
    issue_tracker: pd.DataFrame
    not_started: pd.DataFrame
    orders: pd.DataFrame                        # all unique deliveries for the site
    reports: dict[str, pd.DataFrame] = field(default_factory=dict)  # non-empty only


def build_site_sections(
    orders: pd.DataFrame,
    rules: Ruleset,
    as_of: date,
    now: datetime | None = None,
) -> list[SiteSection]:
    """One :class:`SiteSection` per warehouse present in this export.

    Sections are discovered from the data, so a fourth (or first-ever)
    warehouse appearing in tomorrow's export gets its own sheets with no code
    change. A customer report with no matching records is left out entirely
    rather than written as an empty sheet.
    """
    if orders is None or orders.empty:
        return []
    now = now or now_edmonton()

    sections: list[SiteSection] = []
    for site in rules.ordered_sites(orders["site"]):
        sub = _unique_orders(orders[orders["site"] == site])
        if sub.empty:
            continue
        # "SAPUI5 Export" is dropped here — the site's copy of it is the
        # "<Site> Orders" sheet, and the full export gets its own sheet.
        reports = {
            name: df
            for name, df in build_detail_reports(sub, rules, as_of).items()
            if name != "SAPUI5 Export" and df is not None and not df.empty
        }
        sections.append(SiteSection(
            site=site,
            issue_tracker=build_issue_tracker(sub, rules, as_of, now=now),
            not_started=build_not_started(sub),
            orders=_canonical_slice(sub).reset_index(drop=True),
            reports=reports,
        ))
    return sections


# ---------------------------------------------------------------------------
# All Plants Summary
# ---------------------------------------------------------------------------
def _summary_row(label: str, orders: pd.DataFrame, tracker: pd.DataFrame) -> dict:
    # Counts only — every dollar figure on this sheet lives in the blocks
    # written below the table (see build_order_totals /
    # build_late_picking_breakdown).
    df = _unique_orders(orders)
    not_started = df[df["is_not_started"]]
    return {
        "Site": label,
        "Total Orders": int(df["LE Delivery"].nunique()),
        "Open Orders": int(df.loc[df["is_open"], "LE Delivery"].nunique()),
        "Critical": int(df.loc[df["risk_status"] == R.CRITICAL, "LE Delivery"].nunique()),
        "At Risk": int(df.loc[df["risk_status"] == R.AT_RISK, "LE Delivery"].nunique()),
        "Not Started": int(not_started["LE Delivery"].nunique()),
        "Late": int(df.loc[df["is_late"], "LE Delivery"].nunique()),
        "Route Departure Missed": int(
            df.loc[df["is_route_departure_missed"], "LE Delivery"].nunique()),
        "Issues": int(len(tracker)),
    }


def _late_picking_row(label: str, orders: pd.DataFrame) -> dict:
    """One row of the late-orders picking breakdown."""
    df = _unique_orders(orders)
    late = df[df["is_late"]]
    # is_late already implies the order is open, so picking == 0 is exactly the
    # "not started" bucket here and the three buckets cannot overlap.
    buckets = {
        "Completed": late[late["is_picking_complete"]],
        "Partial": late[late["is_picking_in_progress"]],
        "Not Started": late[late["is_not_started"]],
    }
    row = {
        "Site": label,
        "Late Orders": int(late["LE Delivery"].nunique()),
        "Late Value": float(late["Sales Order Total"].sum()),
    }
    for name, sub in buckets.items():
        row[name] = int(sub["LE Delivery"].nunique())
        row[f"{name} $"] = float(sub["Sales Order Total"].sum())
    return row


def build_late_picking_breakdown(orders: pd.DataFrame, rules: Ruleset) -> pd.DataFrame:
    """Late orders split by picking status — one row per warehouse plus a roll-up."""
    if orders is None or orders.empty:
        return pd.DataFrame(columns=LATE_PICKING_COLUMNS)

    rows = []
    for site in rules.ordered_sites(orders["site"]):
        sub = orders[orders["site"] == site]
        if sub.empty:
            continue
        rows.append(_late_picking_row(site, sub))
    rows.append(_late_picking_row(ALL_PLANTS, orders))
    return pd.DataFrame(rows)[LATE_PICKING_COLUMNS]


def build_order_totals(orders: pd.DataFrame) -> pd.DataFrame:
    """Single-row headline block: total order count and total order value."""
    df = _unique_orders(orders)
    if df is None or df.empty:
        return pd.DataFrame([{"Total Orders": 0, "Total Order Value": 0.0}])[ORDER_TOTALS_COLUMNS]
    return pd.DataFrame([{
        "Total Orders": int(df["LE Delivery"].nunique()),
        "Total Order Value": float(df["Sales Order Total"].sum()),
    }])[ORDER_TOTALS_COLUMNS]


def build_all_plants_summary(
    orders: pd.DataFrame,
    rules: Ruleset,
    tracker: pd.DataFrame | None = None,
) -> pd.DataFrame:
    """One row per warehouse plus an ``All Plants`` roll-up row.

    ``tracker`` is the All Plants Issue Tracker; its rows are attributed to a
    warehouse by their Site column, so the "Issues" count on a warehouse row
    always equals the number of rows on that warehouse's tracker sheet.
    """
    if orders is None or orders.empty:
        return pd.DataFrame(columns=SUMMARY_COLUMNS)
    if tracker is None:
        tracker = pd.DataFrame(columns=ISSUE_TRACKER_COLUMNS)

    rows = []
    for site in rules.ordered_sites(orders["site"]):
        sub = orders[orders["site"] == site]
        if sub.empty:
            continue
        site_issues = (tracker[tracker["Site"] == site] if "Site" in tracker.columns
                       else tracker.iloc[0:0])
        rows.append(_summary_row(site, sub, site_issues))
    rows.append(_summary_row(ALL_PLANTS, orders, tracker))
    return pd.DataFrame(rows)[SUMMARY_COLUMNS]
