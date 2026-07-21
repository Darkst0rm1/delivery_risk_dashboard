"""Daily Delivery Risk Dashboard — standalone Streamlit application.

Upload a fresh SAPUI5 delivery/picking export each day; the app regenerates
every status, issue, risk and report from scratch (no state is carried between
uploads) and lets the user download one formatted Excel workbook.

Run with:  streamlit run app.py
"""
from __future__ import annotations

import io
from datetime import datetime
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

from delivery_dashboard import ProcessResult, process_export
from delivery_dashboard import risk_engine as R
from delivery_dashboard.customer_rules import load_ruleset
from delivery_dashboard.excel_exporter import build_workbook
from delivery_dashboard.loader import DeliveryLoadError
from delivery_dashboard.report_builder import (
    ALL_PLANTS,
    build_all_plants_summary,
    build_detail_reports,
    build_issue_tracker,
    build_not_started,
    build_site_sections,
)

try:
    import plotly.express as px
    _HAS_PLOTLY = True
except Exception:  # noqa: BLE001
    _HAS_PLOTLY = False

EDMONTON = ZoneInfo("America/Edmonton")

st.set_page_config(
    page_title="Daily Delivery Risk Dashboard",
    page_icon="🚚",
    layout="wide",
)

st.title("Daily Delivery Risk Dashboard")
st.caption(
    "Upload today's SAPUI5 delivery export to regenerate the Issue Tracker, "
    "Not Started report, and every customer report — then download one workbook."
)

# ---------------------------------------------------------------------------
# Sidebar — upload + As of Date
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Upload & Settings")
    uploaded = st.file_uploader(
        "Upload Daily SAPUI5 Export", type=["xlsx"], key="ddr_upload",
    )
    as_of = st.date_input(
        "As of Date",
        value=datetime.now(EDMONTON).date(),
        help="All overdue, deadline and priority calculations use this date "
             "(default: today, America/Edmonton). Change it for reproducible "
             "historical testing.",
        key="ddr_asof",
    )
    show_resolved = st.toggle(
        "Show non-escalated groups in Issue Tracker", value=False,
        help="Default tracker is escalation-only (Critical / At Risk).",
    )

if uploaded is None:
    st.info("Upload the daily SAPUI5 export (.xlsx) in the sidebar to begin.")
    st.stop()


# Cache keyed on file bytes + As of Date so a new upload invalidates the old
# result and yesterday's issues never persist.
@st.cache_data(show_spinner=False)
def _run(file_bytes: bytes, as_of_iso: str) -> ProcessResult:
    from datetime import date
    return process_export(io.BytesIO(file_bytes), date.fromisoformat(as_of_iso))


with st.spinner("Cleaning and analyzing the export..."):
    try:
        result = _run(uploaded.getvalue(), as_of.isoformat())
    except DeliveryLoadError as exc:
        st.error(f"**Could not process this file.**\n\n{exc}")
        st.stop()
    except Exception as exc:  # noqa: BLE001
        st.error(f"Unexpected error while processing the file: {exc}")
        st.stop()

orders = result.orders
rules = load_ruleset()

for warn in result.warnings:
    st.warning(warn)

# ---------------------------------------------------------------------------
# Filters (affect dashboard views only, not the default download)
# ---------------------------------------------------------------------------
if "ddr_filter_version" not in st.session_state:
    st.session_state["ddr_filter_version"] = 0
_v = st.session_state["ddr_filter_version"]


def _opts(col: str) -> list[str]:
    if col not in orders.columns:
        return []
    return sorted(orders[col].dropna().astype(str).str.strip().replace("", pd.NA).dropna().unique().tolist())


with st.sidebar:
    st.markdown("---")
    st.header("Filters")
    f_site = st.multiselect("Site", _opts("site"), key=f"ddr_site_{_v}")
    f_customer = st.multiselect("Customer", _opts("customer_display"), key=f"ddr_cust_{_v}")
    f_status = st.multiselect("Status", sorted(orders["risk_status"].unique()), key=f"ddr_status_{_v}")
    f_priority = st.multiselect("Priority", sorted(orders["priority"].unique()), key=f"ddr_prio_{_v}")
    f_facility = st.multiselect("Facility Code", _opts("Facility Code"), key=f"ddr_fac_{_v}")
    f_province = st.multiselect("Ship-to Province", _opts("Ship-to Province"), key=f"ddr_prov_{_v}")
    if st.button("Reset Filters", key="ddr_reset"):
        st.session_state["ddr_filter_version"] += 1
        st.rerun()

fdf = orders.copy()
if f_site:
    fdf = fdf[fdf["site"].isin(f_site)]
if f_customer:
    fdf = fdf[fdf["customer_display"].isin(f_customer)]
if f_status:
    fdf = fdf[fdf["risk_status"].isin(f_status)]
if f_priority:
    fdf = fdf[fdf["priority"].isin(f_priority)]
if f_facility:
    fdf = fdf[fdf["Facility Code"].isin(f_facility)]
if f_province:
    fdf = fdf[fdf["Ship-to Province"].isin(f_province)]

_filtered = any([f_site, f_customer, f_status, f_priority, f_facility, f_province])
if _filtered:
    st.info(f"Filters active — **{len(fdf):,}** of **{len(orders):,}** orders shown. "
            "Reset in the sidebar to clear. (The default Excel download always contains "
            "the complete dataset.)")

# ---------------------------------------------------------------------------
# Derived helpers
# ---------------------------------------------------------------------------
def _money(x: float) -> str:
    return f"${x:,.0f}"


tracker_view = build_issue_tracker(fdf, rules, as_of, include_non_escalated=show_resolved)
not_started_view = fdf[fdf["is_not_started"]]
escalating = fdf[fdf["risk_status"].isin(R.ESCALATION_STATUSES)]

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_exec, tab_issue, tab_ns, tab_cust, tab_raw, tab_dl = st.tabs([
    "Executive Dashboard", "Issue Tracker", "Not Started Orders",
    "Customer Reports", "Raw SAP Data", "Download Report",
])

# ── Executive Dashboard ──────────────────────────────────────────────────────
with tab_exec:
    n_critical = int((escalating["risk_status"] == R.CRITICAL).sum())
    n_atrisk = int((escalating["risk_status"] == R.AT_RISK).sum())
    n_attention = int(len(escalating))
    ns_count = int(len(not_started_view))
    price_risk = float(escalating["Sales Order Total"].sum())
    price_ns = float(not_started_view["Sales Order Total"].sum())

    c1, c2, c3 = st.columns(3)
    c1.metric("Critical Issues", f"{n_critical:,}")
    c2.metric("At-Risk Issues", f"{n_atrisk:,}")
    c3.metric("Orders Requiring Attention", f"{n_attention:,}")
    c4, c5, c6 = st.columns(3)
    c4.metric("Not Started Orders", f"{ns_count:,}")
    c5.metric("Total Price at Risk", _money(price_risk))
    c6.metric("Total Price Not Started", _money(price_ns))

    st.markdown("---")

    st.markdown("**All Plants Summary**")
    st.caption("One row per warehouse plus the roll-up — the first sheet in the download.")
    st.dataframe(
        build_all_plants_summary(fdf, rules, tracker_view),
        use_container_width=True, hide_index=True,
        column_config={
            "Total Order Value": st.column_config.NumberColumn(format="$%.0f"),
            "Value at Risk": st.column_config.NumberColumn(format="$%.0f"),
            "Not Started Value": st.column_config.NumberColumn(format="$%.0f"),
        },
    )

    st.markdown("---")

    if escalating.empty:
        st.success("No escalating orders for the current filters. 🎉")
    else:
        col_l, col_r = st.columns(2)
        by_cust = (escalating.groupby("customer_display")["Sales Order Total"].sum()
                   .sort_values(ascending=False).reset_index())
        ns_by_cust = (not_started_view.groupby("customer_display")["LE Delivery"].nunique()
                      .sort_values(ascending=False).reset_index()
                      .rename(columns={"LE Delivery": "Not Started Orders"}))
        status_dist = escalating["risk_status"].value_counts().reset_index()
        status_dist.columns = ["Status", "Orders"]

        if _HAS_PLOTLY:
            fig1 = px.bar(by_cust, x="Sales Order Total", y="customer_display",
                          orientation="h", title="Total Price at Risk by Customer",
                          labels={"customer_display": "", "Sales Order Total": "Price at Risk"},
                          color="Sales Order Total", color_continuous_scale=["#FEF3C7", "#EF4444"])
            fig1.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False,
                               margin=dict(t=40, b=0, l=0, r=0))
            col_l.plotly_chart(fig1, use_container_width=True)

            fig2 = px.bar(ns_by_cust, x="Not Started Orders", y="customer_display",
                          orientation="h", title="Not Started Order Count by Customer",
                          labels={"customer_display": ""}, color="Not Started Orders",
                          color_continuous_scale=["#DBEAFE", "#1F4E78"])
            fig2.update_layout(yaxis=dict(autorange="reversed"), coloraxis_showscale=False,
                               margin=dict(t=40, b=0, l=0, r=0))
            col_r.plotly_chart(fig2, use_container_width=True)

            fig3 = px.pie(status_dist, names="Status", values="Orders",
                          title="Orders by Risk Status",
                          color="Status", color_discrete_map={
                              R.CRITICAL: "#EF4444", R.AT_RISK: "#F59E0B"})
            fig3.update_layout(margin=dict(t=40, b=0, l=0, r=0))
            col_l.plotly_chart(fig3, use_container_width=True)

            dist = escalating.dropna(subset=["Planned Dlv. Date"]).copy()
            if not dist.empty:
                dist["Planned Day"] = pd.to_datetime(dist["Planned Dlv. Date"]).dt.date
                day_dist = (dist.groupby("Planned Day")["Sales Order Total"].sum()
                            .reset_index())
                fig4 = px.bar(day_dist, x="Planned Day", y="Sales Order Total",
                              title="Planned Delivery Date Risk Distribution",
                              labels={"Sales Order Total": "Price at Risk"},
                              color="Sales Order Total",
                              color_continuous_scale=["#FEF3C7", "#EF4444"])
                fig4.update_layout(coloraxis_showscale=False, margin=dict(t=40, b=0, l=0, r=0))
                col_r.plotly_chart(fig4, use_container_width=True)
        else:
            col_l.caption("Total Price at Risk by Customer")
            col_l.bar_chart(by_cust, x="customer_display", y="Sales Order Total")
            col_r.caption("Not Started Order Count by Customer")
            col_r.bar_chart(ns_by_cust, x="customer_display", y="Not Started Orders")
            col_l.caption("Orders by Risk Status")
            col_l.bar_chart(status_dist, x="Status", y="Orders")

# ── Issue Tracker ────────────────────────────────────────────────────────────
with tab_issue:
    st.subheader("Issue Tracker")
    st.caption("One row per warehouse × customer × issue type, so an Amazon problem in "
               "Calgary never merges with one in Mississauga. Total Price counts only the "
               "affected (at-risk) orders. Select a row to see its underlying orders.")
    if tracker_view.empty:
        st.success("No customer groups are currently escalating for these filters.")
    else:
        display = tracker_view.copy()
        display["Due Date"] = pd.to_datetime(display["Due Date"], errors="coerce").dt.strftime("%m/%d/%Y")
        event = st.dataframe(
            display, use_container_width=True, hide_index=True,
            column_config={
                "Total Price": st.column_config.NumberColumn("Total Price", format="$%.0f"),
                "Issue": st.column_config.TextColumn("Issue", width="large"),
                "Consequence / Risk": st.column_config.TextColumn("Consequence / Risk", width="medium"),
                "Latest Update": st.column_config.TextColumn("Latest Update", width="large"),
            },
            on_select="rerun", selection_mode="single-row", key="ddr_tracker",
        )
        sel = event.get("selection", {}).get("rows", []) if hasattr(event, "get") else []
        if sel:
            picked = tracker_view.iloc[sel[0]]
            # Match all three grouping keys, or the drill-down would show orders
            # from other warehouses and other issue types.
            related = escalating[
                (escalating["site"] == picked["Site"])
                & (escalating["customer_display"] == picked["Customer"])
                & (escalating["issue_type"] == picked["Issue Type"])
            ]
            st.markdown(
                f"**Orders behind — {picked['Site']} · {picked['Customer']} · "
                f"{picked['Issue Type']}** ({len(related)} escalating)"
            )
            st.dataframe(
                related[[
                    "site", "LE Delivery", "Ship-to Name", "Facility Code", "Picking in %",
                    "risk_status", "issue_type", "priority", "Planned Dlv. Date",
                    "Route Depart. Date", "Sales Order Total",
                ]],
                use_container_width=True, hide_index=True,
                column_config={
                    "Sales Order Total": st.column_config.NumberColumn(format="$%.0f"),
                    "Picking in %": st.column_config.NumberColumn(format="%d%%"),
                },
            )

# ── Not Started Orders ───────────────────────────────────────────────────────
with tab_ns:
    st.subheader("Not Started Orders")
    st.caption("Orders at 0% picking with Goods Issue not completed. Once picking "
               "starts, an order drops off this list on the next upload.")
    # Recompute the Not Started report against the filtered frame for the view.
    ns_table = build_not_started(fdf)

    overdue = int(((not_started_view["days_to_planned_delivery"] < 0)).sum())
    due_2 = int(((not_started_view["days_to_planned_delivery"] >= 0) &
                 (not_started_view["days_to_planned_delivery"] <= 2)).sum())
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Not Started Orders", f"{len(ns_table):,}")
    k2.metric("Total Price Not Started", _money(float(not_started_view['Sales Order Total'].sum())))
    k3.metric("Overdue Not Started", f"{overdue:,}")
    k4.metric("Due Within 2 Days", f"{due_2:,}")

    if ns_table.empty:
        st.success("No not-started orders for the current filters.")
    else:
        st.dataframe(
            ns_table, use_container_width=True, hide_index=True,
            column_config={
                "Sales Order Total": st.column_config.NumberColumn(format="$%.0f"),
                "Picking in %": st.column_config.NumberColumn(format="%d%%"),
            },
        )

# ── Customer Reports ─────────────────────────────────────────────────────────
with tab_cust:
    st.subheader("Customer Reports")
    st.caption("Each report is regenerated from today's export. Totals reflect the "
               "full report (not the sidebar filters). Picking a warehouse shows exactly "
               "what its sheet in the download contains.")

    scope = st.selectbox(
        "Warehouse", [ALL_PLANTS] + result.sites, key="ddr_report_site",
        help="All Plants covers every warehouse; a warehouse lists only the reports "
             "that actually have records for it.",
    )
    if scope == ALL_PLANTS:
        available = {n: df for n, df in result.reports.items() if n != "SAPUI5 Export"}
    else:
        available = next(s.reports for s in result.site_sections if s.site == scope)

    if not available:
        st.info(f"No customer reports have records for {scope}.")
    else:
        choice = st.selectbox("Report", list(available), key=f"ddr_report_pick_{scope}")
        rdf = available[choice]
        label = choice if scope == ALL_PLANTS else f"{scope} - {choice}"
        if "Sales Order Total" in rdf.columns:
            tot = float(pd.to_numeric(rdf["Sales Order Total"], errors="coerce").fillna(0).sum())
            st.metric(f"{label} — Rows / Total", f"{len(rdf):,}  •  {_money(tot)}")
        else:
            st.metric(f"{label} — Rows", f"{len(rdf):,}")
        st.dataframe(
            rdf, use_container_width=True, hide_index=True,
            column_config={
                "Sales Order Total": st.column_config.NumberColumn(format="$%.0f"),
                "Picking in %": st.column_config.NumberColumn(format="%d%%"),
            },
        )

# ── Raw SAP Data ─────────────────────────────────────────────────────────────
with tab_raw:
    st.subheader("Raw SAP Data (cleaned)")
    st.caption(f"{len(orders):,} unique deliveries after cleaning and de-duplication. "
               "Original business column names are preserved.")
    st.dataframe(result.reports["SAPUI5 Export"], use_container_width=True, hide_index=True)
    if not result.validation_errors.empty:
        st.markdown(f"**Validation warnings** — {len(result.validation_errors)} record(s) flagged")
        st.dataframe(result.validation_errors, use_container_width=True, hide_index=True)

# ── Download ─────────────────────────────────────────────────────────────────
with tab_dl:
    st.subheader("Download Excel Report")
    st.markdown(
        "One formatted workbook, organized by warehouse: **All Plants Summary → "
        "All Plants Issue Tracker → All Plants Not Started**, then for each warehouse "
        "its **Issue Tracker → Not Started → Orders → customer reports**, then the full "
        "**SAPUI5 Export** (plus a Validation Warnings sheet when needed). Customer "
        "reports with no records for a warehouse are skipped rather than written empty."
    )
    if result.sites:
        st.caption("Warehouses in this export: " + " · ".join(result.sites))
    export_filtered = st.checkbox(
        "Export Filtered View instead of the complete dataset", value=False,
        help="Default export contains the complete processed dataset.",
    )

    if st.button("Generate Excel Workbook", type="primary"):
        with st.spinner("Building workbook..."):
            if export_filtered and _filtered:
                f_tracker = build_issue_tracker(fdf, rules, as_of)
                fr = ProcessResult(
                    orders=fdf,
                    issue_tracker=f_tracker,
                    not_started=build_not_started(fdf),
                    reports=build_detail_reports(fdf, rules, as_of),
                    validation_errors=result.validation_errors,
                    as_of=as_of,
                    site_sections=build_site_sections(fdf, rules, as_of),
                    all_plants_summary=build_all_plants_summary(fdf, rules, f_tracker),
                )
                xlsx = build_workbook(fr)
                fname = f"Daily_Delivery_Risk_{as_of.isoformat()}_filtered.xlsx"
            else:
                xlsx = build_workbook(result)
                fname = f"Daily_Delivery_Risk_{as_of.isoformat()}.xlsx"
        st.download_button(
            "⬇️ Download workbook", data=xlsx, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
