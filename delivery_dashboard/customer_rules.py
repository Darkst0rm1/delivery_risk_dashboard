"""Load the YAML customer rules and classify orders.

Classification is vectorized and deterministic: each order is assigned to the
first customer group (in config order) whose Ship-to Name / Carrier Description
matches. Unmatched orders fall through to the ``fallback`` group so urgent
orders from unconfigured customers are never silently lost.
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import pandas as pd
import yaml

# config/delivery_dashboard_rules.yaml lives at the project root (one level up
# from this package).
DEFAULT_RULES_PATH = Path(__file__).resolve().parents[1] / "config" / "delivery_dashboard_rules.yaml"

_WS_RE = re.compile(r"\s+")


def _norm(value: Any) -> str:
    """Upper-case, whitespace-collapsed key for case-insensitive matching."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    return _WS_RE.sub(" ", str(value)).strip().upper()


@dataclass(frozen=True)
class CustomerRule:
    key: str
    display: str
    priority: int
    owner: str
    consequence: str
    ship_to_contains: tuple[str, ...] = ()
    ship_to_startswith: tuple[str, ...] = ()
    carrier_contains: tuple[str, ...] = ()
    amazon_deadline: bool = False

    def owner_or_unassigned(self) -> str:
        return self.owner.strip() if self.owner and self.owner.strip() else "Unassigned"

    def matches(self, ship_to: str, carrier: str) -> bool:
        """True if the (already normalized) ship-to / carrier hit any pattern."""
        for pat in self.ship_to_contains:
            if pat in ship_to:
                return True
        for pat in self.ship_to_startswith:
            if ship_to.startswith(pat):
                return True
        for pat in self.carrier_contains:
            if pat in carrier:
                return True
        return False


@dataclass(frozen=True)
class DetailReportSpec:
    name: str
    customers: tuple[str, ...] = ()
    facility_in: tuple[str, ...] = ()
    comments_column: bool = False
    amazon_deadline: bool = False
    special: str = ""            # "andrew_tab" / "andrew_tab_2" -> engine-handled


@dataclass
class Ruleset:
    customers: list[CustomerRule]
    fallback: CustomerRule
    detail_reports: list[DetailReportSpec]
    dsd_priorities: dict[str, list[str]]
    site_display: dict[str, str]
    escalation: dict[str, int]
    default_consequence: str
    _by_key: dict[str, CustomerRule] = field(default_factory=dict)

    def __post_init__(self) -> None:
        self._by_key = {r.key: r for r in self.customers}
        self._by_key[self.fallback.key] = self.fallback

    # -- lookups -------------------------------------------------------------
    def rule(self, key: str) -> CustomerRule:
        return self._by_key.get(key, self.fallback)

    def site_label(self, ys: Any) -> str:
        norm = _norm(ys)
        if norm in {k.upper(): v for k, v in self.site_display.items()}:
            return {k.upper(): v for k, v in self.site_display.items()}[norm]
        # Fall back to the first word ("Calgary Warehouse" -> "Calgary").
        raw = "" if ys is None else str(ys).strip()
        return raw.split()[0] if raw else "Unknown"

    # -- classification ------------------------------------------------------
    def classify(self, df: pd.DataFrame) -> pd.DataFrame:
        """Add ``customer_group`` (key), ``customer_display``, ``customer_owner``,
        ``customer_priority``, ``customer_consequence`` and ``site`` columns.
        """
        out = df.copy()
        ship = out.get("Ship-to Name", pd.Series("", index=out.index)).map(_norm)
        carrier = out.get("Carrier Description", pd.Series("", index=out.index)).map(_norm)

        group = pd.Series(self.fallback.key, index=out.index, dtype=object)
        assigned = pd.Series(False, index=out.index)

        for rule in self.customers:
            if assigned.all():
                break
            hit = pd.Series(False, index=out.index)
            for pat in rule.ship_to_contains:
                hit |= ship.str.contains(re.escape(pat), na=False)
            for pat in rule.ship_to_startswith:
                hit |= ship.str.startswith(pat)
            for pat in rule.carrier_contains:
                hit |= carrier.str.contains(re.escape(pat), na=False)
            take = hit & ~assigned
            group.loc[take] = rule.key
            assigned.loc[take] = True

        out["customer_group"] = group
        out["customer_display"] = group.map(lambda k: self.rule(k).display)
        out["customer_owner"] = group.map(lambda k: self.rule(k).owner_or_unassigned())
        out["customer_priority"] = group.map(lambda k: self.rule(k).priority)
        out["customer_consequence"] = group.map(lambda k: self.rule(k).consequence)
        out["site"] = out.get("ys", pd.Series("", index=out.index)).map(self.site_label)
        return out

    def dsd_priority_for(self, ship_to: Any) -> str | None:
        """Return "Priority 1"/"Priority 2"/None for a ship-to name."""
        norm = _norm(ship_to)
        if not norm:
            return None
        for label in sorted(self.dsd_priorities.keys()):  # Priority 1 before Priority 2
            for pat in self.dsd_priorities[label]:
                if _norm(pat) in norm:
                    return label
        return None


def _rule_from_dict(d: dict[str, Any]) -> CustomerRule:
    match = d.get("match", {}) or {}
    return CustomerRule(
        key=str(d["key"]),
        display=str(d.get("display", d["key"])),
        priority=int(d.get("priority", 3)),
        owner=str(d.get("owner", "") or ""),
        consequence=str(d.get("consequence", "") or ""),
        ship_to_contains=tuple(_norm(x) for x in match.get("ship_to_contains", []) or []),
        ship_to_startswith=tuple(_norm(x) for x in match.get("ship_to_startswith", []) or []),
        carrier_contains=tuple(_norm(x) for x in match.get("carrier_contains", []) or []),
        amazon_deadline=bool(d.get("amazon_deadline", False)),
    )


def load_ruleset(path: str | Path | None = None) -> Ruleset:
    """Load and parse the YAML rules file into a :class:`Ruleset`."""
    p = Path(path) if path else DEFAULT_RULES_PATH
    with open(p, "r", encoding="utf-8") as fh:
        cfg = yaml.safe_load(fh) or {}

    customers = [_rule_from_dict(c) for c in cfg.get("customers", [])]

    fb = cfg.get("fallback", {}) or {}
    default_consequence = str(cfg.get("default_consequence", "Delivery delay, customer out-of-stock risk and lost sales."))
    fallback = CustomerRule(
        key=str(fb.get("key", "other")),
        display=str(fb.get("display", "Other / Unclassified")),
        priority=int(fb.get("priority", 3)),
        owner=str(fb.get("owner", "") or ""),
        consequence=str(fb.get("consequence", "") or default_consequence),
    )

    detail_reports = []
    for r in cfg.get("detail_reports", []) or []:
        detail_reports.append(
            DetailReportSpec(
                name=str(r["name"]),
                customers=tuple(r.get("customers", []) or []),
                facility_in=tuple(r.get("facility_in", []) or []),
                comments_column=bool(r.get("comments_column", False)),
                amazon_deadline=bool(r.get("amazon_deadline", False)),
                special=str(r.get("special", "") or ""),
            )
        )

    return Ruleset(
        customers=customers,
        fallback=fallback,
        detail_reports=detail_reports,
        dsd_priorities={str(k): list(v or []) for k, v in (cfg.get("dsd_priorities", {}) or {}).items()},
        site_display={str(k): str(v) for k, v in (cfg.get("site_display", {}) or {}).items()},
        escalation={str(k): int(v) for k, v in (cfg.get("escalation", {}) or {}).items()},
        default_consequence=default_consequence,
    )
