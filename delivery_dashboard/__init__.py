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
    issue_tracker: pd.DataFrame                # one summarized row per escalating customer
    not_started: pd.DataFrame
    reports: dict[str, pd.DataFrame]           # detail report name -> dataframe
    validation_errors: pd.DataFrame            # rows/columns that could not be processed
    as_of: date
    warnings: list[str] = field(default_factory=list)
    meta: dict[str, Any] = field(default_factory=dict)


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

    issue_tracker = report_builder.build_issue_tracker(orders, rules, as_of)
    not_started = report_builder.build_not_started(orders)
    reports = report_builder.build_detail_reports(orders, rules, as_of)

    warnings: list[str] = []
    if not validation_errors.empty:
        warnings.append(
            f"{len(validation_errors)} record(s) could not be fully parsed — "
            "see the Validation Errors download."
        )

    return ProcessResult(
        orders=orders,
        issue_tracker=issue_tracker,
        not_started=not_started,
        reports=reports,
        validation_errors=validation_errors,
        as_of=as_of,
        warnings=warnings,
        meta={
            "row_count": int(len(orders)),
            "duplicate_deliveries_removed": int(cleaned.attrs.get("duplicates_removed", 0)),
        },
    )
