"""Daily Delivery Risk Dashboard engine.

A reusable processing pipeline for the SAPUI5 delivery/picking export that
powers the Daily Delivery Risk Dashboard Streamlit page. Everything except the
Streamlit UI lives here so the logic is testable in isolation.

Pipeline:
    loader.load_sap_export(file)      -> raw dataframe + validation
    cleaner.clean(df)                 -> cleaned, de-duplicated dataframe
    customer_rules.Ruleset.classify() -> customer_group / customer_report
    risk_engine.enrich(df, as_of)     -> derived fields, risk_status, priority
    report_builder.build_reports(...) -> issue tracker + all detail reports
    report_builder.build_site_sections() -> the per-warehouse sheets
    excel_exporter.build_workbook(...)-> formatted .xlsx bytes

The single high-level entry point is :func:`process_export`.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Any

import pandas as pd

from . import cleaner, loader, report_builder, risk_engine
from .customer_rules import Ruleset, load_ruleset


class DeliveryDashboardError(Exception):
    """User-facing error for any recoverable problem while processing the export."""


@dataclass
class ProcessResult:
    """Everything the page and the Excel exporter need for one upload."""

    orders: pd.DataFrame                       # cleaned + enriched, one row per LE Delivery
    issue_tracker: pd.DataFrame                # All Plants: Site x Customer x Issue Type
    not_started: pd.DataFrame                  # All Plants
    reports: dict[str, pd.DataFrame]           # detail report name -> dataframe (all plants)
    validation_errors: pd.DataFrame            # rows/columns that could not be processed
    as_of: date
    site_sections: list = field(default_factory=list)          # list[SiteSection]
    all_plants_summary: pd.DataFrame = field(default_factory=pd.DataFrame)
    warnings: list[str] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)

    @property
    def sites(self) -> list[str]:
        """Warehouses detected in this export, in workbook order."""
        return [s.site for s in self.site_sections]


def process_export(
    file: Any,
    as_of: date,
    ruleset: Ruleset | None = None,
) -> ProcessResult:
    """Run the full pipeline on an uploaded SAPUI5 export.

    Parameters
    ----------
    file:
        A file-like object (or path) pointing at the .xlsx export.
    as_of:
        The "As of Date" used for every overdue / deadline / priority
        calculation. Supplied by the page (default: today in America/Edmonton).
    ruleset:
        Pre-loaded customer :class:`Ruleset`. Loaded from the default YAML if
        omitted.
    """
    rules = ruleset or load_ruleset()

    raw = loader.load_sap_export(file)          # raises loader.DeliveryLoadError
    cleaned, validation_errors = cleaner.clean(raw)
    classified = rules.classify(cleaned)
    orders = risk_engine.enrich(classified, as_of, rules)

    # One timestamp for every "Latest Update" string in this run, so a warehouse
    # tracker and the combined tracker can never be stamped a minute apart.
    now = report_builder.now_edmonton()

    issue_tracker = report_builder.build_issue_tracker(orders, rules, as_of, now=now)
    not_started = report_builder.build_not_started(orders)
    reports = report_builder.build_detail_reports(orders, rules, as_of)
    site_sections = report_builder.build_site_sections(orders, rules, as_of, now=now)
    all_plants_summary = report_builder.build_all_plants_summary(orders, rules, issue_tracker)

    warnings: list[str] = []
    # The whole workbook is organized by warehouse, so silently landing every
    # order in one "Unknown" section would be easy to miss. Say so instead.
    if not orders.empty and set(orders["site"].unique()) == {"Unknown"}:
        warnings.append(
            "No warehouse column was recognized in this export, so every order is "
            "grouped under a single **Unknown** warehouse. The site column should be "
            "named ys, Site, Warehouse or Plant — tell me what yours is called and it "
            "can be added to the header aliases."
        )
    if not validation_errors.empty:
        warnings.append(
            f"{len(validation_errors)} record(s) could not be fully parsed — "
            "see the Validation Warnings sheet in the download."
        )

    return ProcessResult(
        orders=orders,
        issue_tracker=issue_tracker,
        not_started=not_started,
        reports=reports,
        validation_errors=validation_errors,
        as_of=as_of,
        site_sections=site_sections,
        all_plants_summary=all_plants_summary,
        warnings=warnings,
        meta={
            "row_count": int(len(orders)),
            "duplicate_deliveries_removed": int(cleaned.attrs.get("duplicates_removed", 0)),
            "sites": [s.site for s in site_sections],
        },
    )
