"""Derive order-level status, risk and priority fields.

All calculations are relative to the supplied ``as_of`` date (midnight in the
America/Edmonton sense — the caller passes a naive :class:`datetime.date`).
Everything is vectorized (numpy.select) so it is safe on empty frames — no
``DataFrame.apply(axis=1)`` that would misbehave on pandas 3.0.
"""
from __future__ import annotations

from datetime import date

import numpy as np
import pandas as pd

from .customer_rules import Ruleset

# Status vocabulary.
CRITICAL = "Critical"
AT_RISK = "At Risk"
IN_PROGRESS = "In Progress"
READY = "Ready / Awaiting GI"
NOT_STARTED = "Not Started"
RESOLVED = "Resolved"

# Statuses that put a customer on the escalation-only Issue Tracker.
ESCALATION_STATUSES = {CRITICAL, AT_RISK}

# Issue Type vocabulary — *why* an order is on the tracker, as opposed to how
# bad it is. The Issue Tracker groups on it, so one Amazon row can be "Late
# Delivery" while another stays "Cancellation Deadline Approaching" instead of
# collapsing into a single vague row.
IT_DEADLINE_PASSED = "Cancellation Deadline Passed"
IT_LATE = "Late Delivery"
IT_ROUTE_MISSED = "Route Departure Missed"
IT_DEADLINE_SOON = "Cancellation Deadline Approaching"
IT_NOT_STARTED = "Not Started"
IT_PICKING_BEHIND = "Picking Behind Schedule"
IT_AWAITING_GI = "Awaiting Goods Issue"
IT_RESOLVED = "Resolved"
IT_OTHER = "Other"

# Assignment precedence (an order can trip several conditions at once) and the
# order issue types are ranked in on the tracker.
_ISSUE_TYPE_ORDER = {
    IT_DEADLINE_PASSED: 0,
    IT_LATE: 1,
    IT_ROUTE_MISSED: 2,
    IT_DEADLINE_SOON: 3,
    IT_NOT_STARTED: 4,
    IT_PICKING_BEHIND: 5,
    IT_AWAITING_GI: 6,
    IT_RESOLVED: 7,
    IT_OTHER: 8,
}


def issue_type_rank(value: str) -> int:
    """Sort key for an issue type; unknown values sort last."""
    return _ISSUE_TYPE_ORDER.get(value, 9)

# Priority vocabulary (Not Started report + row ranking).
PRIORITY_CRITICAL = "Critical"
PRIORITY_HIGH = "High"
PRIORITY_MEDIUM = "Medium"
PRIORITY_LOW = "Low"
_PRIORITY_ORDER = {PRIORITY_CRITICAL: 0, PRIORITY_HIGH: 1, PRIORITY_MEDIUM: 2, PRIORITY_LOW: 3}


def priority_rank(value: str) -> int:
    """Sort key: Critical(0) < High(1) < Medium(2) < Low(3)."""
    return _PRIORITY_ORDER.get(value, 9)


def _days_from(series: pd.Series, as_of_ts: pd.Timestamp) -> pd.Series:
    """Whole-day difference (date granularity) from as_of to each date. NaN-safe."""
    s = pd.to_datetime(series, errors="coerce")
    return (s.dt.normalize() - as_of_ts).dt.days


def _within(days: pd.Series, n: int) -> pd.Series:
    """True where a day-count is between 0 and n inclusive (upcoming within n)."""
    return days.notna() & (days >= 0) & (days <= n)


def _map_group_to_report(rules: Ruleset) -> dict[str, str]:
    """customer_group key -> the detail report it primarily belongs to."""
    out: dict[str, str] = {}
    for spec in rules.detail_reports:
        if spec.special:
            continue
        for key in spec.customers:
            out.setdefault(key, spec.name)
    return out


def enrich(df: pd.DataFrame, as_of: date, rules: Ruleset) -> pd.DataFrame:
    """Return ``df`` with all derived order fields added.

    Assumes ``df`` has already been cleaned and classified (customer_group etc).
    """
    out = df.copy()
    as_of_ts = pd.Timestamp(as_of)

    esc = rules.escalation
    at_risk_days = int(esc.get("at_risk_days", 2))
    amz_crit_days = int(esc.get("amazon_cancel_critical_days", 2))
    amz_risk_days = int(esc.get("amazon_cancel_at_risk_days", 5))
    high_days = int(esc.get("not_started_high_days", 2))
    medium_days = int(esc.get("not_started_medium_days", 5))

    picking = pd.to_numeric(out.get("Picking in %", 0), errors="coerce").fillna(0.0)
    goods_issue = out.get("Goods Issue", pd.Series("", index=out.index)).fillna("").astype(str).str.strip()

    out["picking_pct"] = picking
    out["is_goods_issue_complete"] = goods_issue.str.upper().eq("COMPLETED")
    out["is_open"] = ~out["is_goods_issue_complete"]
    out["is_not_started"] = (picking == 0) & ~out["is_goods_issue_complete"]
    out["is_picking_in_progress"] = (picking > 0) & (picking < 100)
    out["is_picking_complete"] = picking >= 100

    planned = pd.to_datetime(out.get("Planned Dlv. Date"), errors="coerce")
    route = pd.to_datetime(out.get("Route Depart. Date"), errors="coerce")

    out["days_to_planned_delivery"] = _days_from(planned, as_of_ts)
    out["days_to_route_departure"] = _days_from(route, as_of_ts)

    out["is_late"] = planned.notna() & (planned < as_of_ts) & out["is_open"]
    out["is_route_departure_missed"] = route.notna() & (route < as_of_ts) & (picking < 100)

    # Customer deadline: Amazon cancellation deadline = Customer Dock Appointment Date.
    is_amazon = out["customer_group"].map(lambda k: rules.rule(k).amazon_deadline)
    dock = pd.to_datetime(out.get("Customer Dock Appointment Date"), errors="coerce")
    days_to_dock = _days_from(dock, as_of_ts)
    out["days_to_customer_deadline"] = np.where(is_amazon, days_to_dock, np.nan)
    deadline_passed = is_amazon & dock.notna() & (dock < as_of_ts)
    deadline_within_crit = is_amazon & _within(days_to_dock, amz_crit_days)
    deadline_within_risk = is_amazon & _within(days_to_dock, amz_risk_days)

    # -- Risk status ---------------------------------------------------------
    critical = (
        out["is_late"]
        | out["is_route_departure_missed"]
        | deadline_passed
        | (deadline_within_crit & (picking < 100))
    )
    at_risk = (
        (_within(out["days_to_planned_delivery"], at_risk_days) & (picking == 0))
        | (_within(out["days_to_route_departure"], at_risk_days) & (picking < 100))
        | deadline_within_risk
        | ((out["customer_priority"] == 1) & _within(out["days_to_planned_delivery"], at_risk_days) & out["is_open"])
    )

    conditions = [
        out["is_goods_issue_complete"],
        critical,
        at_risk,
        out["is_picking_complete"],
        out["is_picking_in_progress"],
    ]
    choices = [RESOLVED, CRITICAL, AT_RISK, READY, IN_PROGRESS]
    out["risk_status"] = np.select(conditions, choices, default=NOT_STARTED)

    # -- Issue type ----------------------------------------------------------
    # Same leading condition as risk_status (a completed Goods Issue is done,
    # whatever else is true of it), then most-to-least severe.
    out["issue_type"] = np.select(
        [
            out["is_goods_issue_complete"],
            deadline_passed,
            out["is_late"],
            out["is_route_departure_missed"],
            deadline_within_crit | deadline_within_risk,
            out["is_not_started"],
            out["is_picking_in_progress"],
            out["is_picking_complete"],
        ],
        [
            IT_RESOLVED, IT_DEADLINE_PASSED, IT_LATE, IT_ROUTE_MISSED,
            IT_DEADLINE_SOON, IT_NOT_STARTED, IT_PICKING_BEHIND, IT_AWAITING_GI,
        ],
        default=IT_OTHER,
    )

    # -- Priority (Not Started report + ranking) -----------------------------
    earliest_due = pd.concat(
        [out["days_to_planned_delivery"], out["days_to_route_departure"], pd.Series(days_to_dock, index=out.index)],
        axis=1,
    ).min(axis=1)

    p_critical = (
        out["is_late"]
        | out["is_route_departure_missed"]
        | deadline_passed
        | deadline_within_crit
    )
    p_high = earliest_due.notna() & (earliest_due <= high_days)
    p_medium = earliest_due.notna() & (earliest_due <= medium_days)
    out["priority"] = np.select(
        [p_critical, p_high, p_medium],
        [PRIORITY_CRITICAL, PRIORITY_HIGH, PRIORITY_MEDIUM],
        default=PRIORITY_LOW,
    )
    out["priority_rank"] = out["priority"].map(priority_rank)
    out["_earliest_due"] = earliest_due

    # Detail report each order maps to (for the Issue Tracker "Detail Report").
    group_to_report = _map_group_to_report(rules)
    out["customer_report"] = out["customer_group"].map(lambda k: group_to_report.get(k, ""))

    return out
