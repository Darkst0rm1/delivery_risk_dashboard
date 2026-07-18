"""Build the Issue Tracker, Not Started report, and all customer detail reports.

Every output is regenerated from the enriched order frame — no state is carried
between uploads. Issue text and "Latest Update" strings are deterministic
functions of the current data (plus the as_of / now timestamps), never
hard-coded from the prototype.
"""
from __future__ import annotations

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
    "Site", "Customer", "Issue", "Consequence / Risk", "Owner",
    "Due Date", "Status", "Total Price", "Order Count", "Latest Update", "Detail Report",
]

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
# Issue text (per customer, deterministic)
# ---------------------------------------------------------------------------
def _counts(affected: pd.DataFrame) -> dict:
    return {
        "count": int(affected["LE Delivery"].nunique()),
        "not_started": int(affected["is_not_started"].sum()),
        "late": int(affected["is_late"].sum()),
        "in_progress": int(affected["is_picking_in_progress"].sum()),
        "missed_route": int(affected["is_route_departure_missed"].sum()),
    }


def _earliest_planned(affected: pd.DataFrame) -> str:
    planned = pd.to_datetime(affected.get("Planned Dlv. Date"), errors="coerce").dropna()
    return _fmt_date(planned.min()) if not planned.empty else "N/A"


def _nearest_deadline(affected: pd.DataFrame) -> str:
    dock = pd.to_datetime(affected.get("Customer Dock Appointment Date"), errors="coerce").dropna()
    return _fmt_date(dock.min()) if not dock.empty else "N/A"


def _issue_amazon(a: pd.DataFrame) -> str:
    c = _counts(a)
    return (f"{c['count']} Amazon orders are at risk. {c['not_started']} have not started "
            f"and {c['late']} are late. The nearest cancellation deadline is {_nearest_deadline(a)}.")


def _issue_gfs(a: pd.DataFrame) -> str:
    c = _counts(a)
    return (f"{c['count']} GFS pickup orders require attention. {c['missed_route']} missed the "
            f"route departure window and {c['not_started']} remain at 0% picked.")


def _issue_aw(a: pd.DataFrame) -> str:
    pct = pd.to_numeric(a.get("picking_pct"), errors="coerce").fillna(0)
    worst = pct.min() if not pct.empty else 0
    return (f"A&W order within the GFS pickup group is delayed and remains at "
            f"{worst:.0f}% picked.")


def _issue_sysco(a: pd.DataFrame) -> str:
    c = _counts(a)
    return (f"{c['count']} Sysco orders require attention. {c['not_started']} remain at 0% picked "
            f"and {c['missed_route']} have missed the pickup window.")


def _issue_save_on(a: pd.DataFrame) -> str:
    c = _counts(a)
    return (f"{c['count']} Save On Foods warehouse orders are delayed or at risk. "
            f"{c['not_started']} have not started.")


def _issue_associated(a: pd.DataFrame) -> str:
    c = _counts(a)
    base = (f"{c['count']} Associated Grocers orders require attention. "
            f"{c['not_started']} have not started and {c['late']} are late.")
    ids = [str(x) for x in a["LE Delivery"].dropna().unique()]
    if 0 < len(ids) <= 5:
        base += " Affected LE Deliveries: " + ", ".join(ids) + "."
    return base


def _issue_general(a: pd.DataFrame) -> str:
    c = _counts(a)
    return (f"{c['count']} orders require attention: {c['not_started']} not started, "
            f"{c['late']} late, {c['in_progress']} in progress. "
            f"Earliest planned delivery: {_earliest_planned(a)}.")


_ISSUE_BUILDERS: dict[str, Callable[[pd.DataFrame], str]] = {
    "amazon": _issue_amazon,
    "gfs": _issue_gfs,
    "aw_via_gfs": _issue_aw,
    "sysco": _issue_sysco,
    "save_on": _issue_save_on,
    "associated_grocers": _issue_associated,
}


def _issue_text(group_key: str, affected: pd.DataFrame) -> str:
    return _ISSUE_BUILDERS.get(group_key, _issue_general)(affected)


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
            f"current value at risk is ${value:,.2f}. The nearest due date is {nearest}.")


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
    """One summarized row per escalating customer group.

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
    for key, grp in wanted.groupby("customer_group", sort=False):
        rule = rules.rule(key)
        has_critical = (grp["risk_status"] == R.CRITICAL).any()
        status = R.CRITICAL if has_critical else (
            R.AT_RISK if (grp["risk_status"] == R.AT_RISK).any() else grp["risk_status"].iloc[0]
        )
        due = _due_dates(grp, rules).dropna()
        site_mode = grp["site"].mode()
        rows.append({
            "Site": site_mode.iloc[0] if not site_mode.empty else "",
            "Customer": rule.display,
            "Issue": _issue_text(key, grp),
            "Consequence / Risk": rule.consequence or rules.default_consequence,
            "Owner": rule.owner_or_unassigned(),
            "Due Date": due.min() if not due.empty else pd.NaT,
            "Status": status,
            "Total Price": float(grp["Sales Order Total"].sum()),
            "Order Count": int(grp["LE Delivery"].nunique()),
            "Latest Update": _latest_update(grp, rules, now),
            "Detail Report": grp["customer_report"].mode().iloc[0] if not grp["customer_report"].mode().empty else "",
            "_status_rank": 0 if has_critical else 1,
        })

    out = pd.DataFrame(rows)
    out = out.sort_values(["_status_rank", "Total Price"], ascending=[True, False])
    out = out.drop(columns=["_status_rank"]).reset_index(drop=True)
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
    body.attrs["is_customer_report"] = True
    return body.reset_index(drop=True)


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
