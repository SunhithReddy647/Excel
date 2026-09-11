"""
Auto Dashboard Architect — dynamically analyzes any client dataset (CSV, XLS, XLSX)
and architects the optimal executive Excel plan with real dynamic formulas,
summary breakdown tables, native charts, and interactive slicers.
"""

import re
from typing import Optional
import pandas as pd
import numpy as np

from schemas.schemas import (
    ExcelPlan, SheetSpec, KPISpec, AnalysisSpec,
    ChartSpec, FilterSpec, CalculationSpec, DashboardSpec,
    DatasetProfile,
)


def _sanitize_col(name: str) -> str:
    """Sanitize column name for Excel formulas."""
    return re.sub(r"[\[\]']", "", str(name)).strip()


def architect_executive_plan(
    df: pd.DataFrame,
    profile: Optional[DatasetProfile] = None,
    title: Optional[str] = None,
    user_prompt: Optional[str] = None,
) -> ExcelPlan:
    """
    Examine any client dataset and build a certified executive dashboard plan.
    Identifies primary metrics, categorical dimensions, date hierarchies,
    and constructs a comprehensive, functional Excel workbook architecture.
    """
    cols = list(df.columns)
    num_rows = len(df)

    # 1. Classify columns into measures and dimensions
    measures = []
    dimensions = []
    date_cols = []

    for col in cols:
        series = df[col].dropna()
        if series.empty:
            continue

        # Date detection
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            date_cols.append(col)
            continue
        if "date" in col.lower() or "month" in col.lower() or "year" in col.lower() or "day" in col.lower():
            date_cols.append(col)
            dimensions.append(col)
            continue

        # Numeric detection
        if pd.api.types.is_numeric_dtype(df[col]):
            # Check unique count: if very few integers, could be a category or code
            n_unique = df[col].nunique()
            col_lower = col.lower()
            if n_unique <= 5 and not any(k in col_lower for k in ["sales", "rev", "cost", "price", "amount", "profit", "units", "qty"]):
                dimensions.append(col)
            else:
                measures.append(col)
        else:
            dimensions.append(col)

    # 2. Prioritize primary and secondary measures
    def measure_priority(m: str) -> int:
        ml = m.lower()
        if any(k in ml for k in ["revenue", "sales", "total_sales", "total_revenue", "income", "amount"]):
            return 10
        if any(k in ml for k in ["profit", "net_profit", "margin"]):
            return 9
        if any(k in ml for k in ["cost", "expense", "spend"]):
            return 8
        if any(k in ml for k in ["unit", "qty", "quantity", "volume", "count"]):
            return 7
        if any(k in ml for k in ["price", "rate", "fee"]):
            return 6
        if any(k in ml for k in ["target", "budget", "quota"]):
            return 5
        return 1

    measures.sort(key=measure_priority, reverse=True)

    # Fallback if no numeric measures found
    if not measures:
        # Create a dummy count measure
        primary_measure = cols[0]
        primary_agg = "COUNTA"
        secondary_measure = None
    else:
        primary_measure = measures[0]
        primary_agg = "SUM"
        secondary_measure = measures[1] if len(measures) > 1 else None

    # 3. Prioritize dimensions for breakdown and slicers
    def dimension_priority(d: str) -> int:
        dl = d.lower()
        if any(k in dl for k in ["region", "territory", "country", "area", "zone"]):
            return 10
        if any(k in dl for k in ["category", "type", "segment", "class", "group"]):
            return 9
        if any(k in dl for k in ["product", "item", "sku", "service"]):
            return 8
        if any(k in dl for k in ["salesperson", "rep", "agent", "employee", "manager", "name"]):
            return 7
        if any(k in dl for k in ["month", "quarter", "year", "period", "date"]):
            return 6
        if any(k in dl for k in ["status", "channel", "department", "vendor"]):
            return 5
        # Prefer dimensions with reasonable cardinality (2 to 20 unique values)
        n_unique = df[d].nunique() if d in df.columns else 999
        if 2 <= n_unique <= 20:
            return 4
        return 1

    dimensions.sort(key=dimension_priority, reverse=True)
    primary_dim = dimensions[0] if dimensions else "Category"
    secondary_dim = dimensions[1] if len(dimensions) > 1 else (dimensions[0] if dimensions else "Segment")
    tertiary_dim = dimensions[2] if len(dimensions) > 2 else None

    # Determine Title
    clean_title = title.strip() if title else "Executive Performance Dashboard"

    # 4. Formulate Executive KPIs
    kpis = []
    # KPI 1: Primary Measure Total
    p_format = "currency" if any(k in primary_measure.lower() for k in ["sales", "rev", "cost", "price", "amount", "profit", "spend", "income"]) else "number"
    kpis.append(KPISpec(
        name=f"Total {primary_measure}",
        metric=primary_measure,
        aggregation=primary_agg,
        format=p_format,
        label=f"Total {primary_measure}",
    ))

    # KPI 2: Secondary Measure Total (e.g. Units or Profit)
    if secondary_measure:
        s_format = "currency" if any(k in secondary_measure.lower() for k in ["sales", "rev", "cost", "price", "amount", "profit"]) else "number"
        kpis.append(KPISpec(
            name=f"Total {secondary_measure}",
            metric=secondary_measure,
            aggregation="SUM",
            format=s_format,
            label=f"Total {secondary_measure}",
        ))

    # KPI 3: Average of Primary Measure (e.g. Average Order Value)
    kpis.append(KPISpec(
        name=f"Avg {primary_measure}",
        metric=primary_measure,
        aggregation="AVERAGE",
        format=f"{p_format}_decimal" if p_format == "currency" else "number_decimal",
        label=f"Average {primary_measure}",
    ))

    # KPI 4: Transaction Count / Volume
    kpis.append(KPISpec(
        name="Total Records",
        metric=cols[0],
        aggregation="COUNTA",
        format="integer",
        label="Total Transactions",
    ))

    # 5. Formulate Summary Breakdown Analyses
    analyses = []
    charts = []

    # Analysis 1: Primary Dimension Breakdown
    if primary_dim in df.columns:
        a1_name = f"Breakdown by {primary_dim}"
        analyses.append(AnalysisSpec(
            name=a1_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[primary_dim],
            sort="descending",
            description=f"Performance aggregation by {primary_dim}",
        ))
        # Chart 1: Clustered Column Chart
        charts.append(ChartSpec(
            name=f"{primary_measure} by {primary_dim}",
            chart_type="column",
            title=f"{primary_measure} by {primary_dim}",
            data_source=a1_name,
            category_field=primary_dim,
            value_field=primary_measure,
            x_axis_title=primary_dim,
            y_axis_title=primary_measure,
            width=12,
            height=7,
        ))

    # Analysis 2: Secondary Dimension Breakdown
    if secondary_dim in df.columns and secondary_dim != primary_dim:
        a2_name = f"Distribution by {secondary_dim}"
        analyses.append(AnalysisSpec(
            name=a2_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[secondary_dim],
            sort="descending",
            description=f"Revenue distribution across {secondary_dim}",
        ))
        # Chart 2: Donut or Pie Chart
        charts.append(ChartSpec(
            name=f"{secondary_dim} Distribution",
            chart_type="doughnut" if df[secondary_dim].nunique() <= 6 else "bar",
            title=f"{primary_measure} Distribution by {secondary_dim}",
            data_source=a2_name,
            category_field=secondary_dim,
            value_field=primary_measure,
            width=11,
            height=7,
        ))

    # Analysis 3: Temporal Trend or Leaderboard
    time_dim = next((d for d in date_cols if d in df.columns), None)
    if time_dim:
        a3_name = f"Monthly Trend ({time_dim})"
        analyses.append(AnalysisSpec(
            name=a3_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[time_dim],
            sort="default",
            description=f"Timeline trend by {time_dim}",
        ))
        charts.append(ChartSpec(
            name=f"{primary_measure} Trend",
            chart_type="line",
            title=f"{primary_measure} Over Time",
            data_source=a3_name,
            category_field=time_dim,
            value_field=primary_measure,
            width=14,
            height=7,
        ))
    elif tertiary_dim and tertiary_dim in df.columns:
        a3_name = f"Performance by {tertiary_dim}"
        analyses.append(AnalysisSpec(
            name=a3_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[tertiary_dim],
            sort="descending",
            description=f"Leaderboard by {tertiary_dim}",
        ))
        charts.append(ChartSpec(
            name=f"{tertiary_dim} Leaderboard",
            chart_type="bar",
            title=f"Performance by {tertiary_dim}",
            data_source=a3_name,
            category_field=tertiary_dim,
            value_field=primary_measure,
            width=12,
            height=7,
        ))

    # 6. Slicers / Interactive Filters
    filters = []
    # Pick top 2 categorical dimensions suitable for slicers (between 2 and 15 unique values)
    slicer_candidates = [
        d for d in dimensions
        if d in df.columns and 2 <= df[d].nunique() <= 20
    ]
    if not slicer_candidates and dimensions:
        slicer_candidates = dimensions[:2]

    for sc in slicer_candidates[:2]:
        filters.append(FilterSpec(
            field=sc,
            filter_type="slicer",
            description=f"Interactive Excel Slicer for {sc}",
        ))

    # 7. Dashboard Specification
    dashboard_spec = DashboardSpec(
        title=clean_title,
        kpis=kpis,
        chart_refs=[c.name for c in charts],
        filter_refs=[f.field for f in filters],
    )

    # 8. Ordered Sheets (Executive_Dashboard as first active sheet)
    sheets = [
        SheetSpec(name="Executive_Dashboard", sheet_type="dashboard", description="Interactive executive summary with dynamic KPIs, slicers, and charts"),
        SheetSpec(name="Raw_Data", sheet_type="raw_data", description="Original client dataset formatted as official Excel Table"),
        SheetSpec(name="Summary_Analysis", sheet_type="analysis", description="Dynamic formula summary breakdown tables"),
        SheetSpec(name="Audit_Trail", sheet_type="assignment", description="Dataset architecture metadata and formula verification"),
    ]

    return ExcelPlan(
        workbook_title=clean_title,
        calculations=[],
        analyses=analyses,
        charts=charts,
        dashboard=dashboard_spec,
        filters=filters,
        sheets=sheets,
    )
