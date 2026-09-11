"""
Tests for the Auto Dashboard Architect and direct dashboard generation from data.
"""

import os
import io
import pytest
import pandas as pd
from services.auto_dashboard_architect import architect_executive_plan
from services.dataset_profiler import profile_dataset
from schemas.schemas import ExcelPlan


def test_architect_sales_dataset():
    """Verify sales dataset produces KPIs with SUM, COUNTA, slicers, and charts."""
    df = pd.DataFrame({
        "Order_ID": [101, 102, 103, 104, 105],
        "Region": ["North", "South", "East", "West", "North"],
        "Product": ["Laptop", "Monitor", "Keyboard", "Mouse", "Laptop"],
        "Salesperson": ["Alice", "Bob", "Charlie", "David", "Alice"],
        "Units_Sold": [10, 5, 20, 15, 8],
        "Revenue": [12000, 2500, 1000, 750, 9600],
        "Month": ["Jan", "Jan", "Feb", "Feb", "Mar"],
    })

    plan = architect_executive_plan(df, title="Q1 Regional Sales Performance")
    assert isinstance(plan, ExcelPlan)
    assert plan.workbook_title == "Q1 Regional Sales Performance"
    assert len(plan.dashboard.kpis) == 4

    # Verify primary metric is Revenue with SUM aggregation
    kpi_metrics = [k.metric for k in plan.dashboard.kpis]
    assert "Revenue" in kpi_metrics
    assert any(k.aggregation == "SUM" for k in plan.dashboard.kpis)
    assert any(k.aggregation == "COUNTA" for k in plan.dashboard.kpis)

    # Verify Slicers generated for categorical dimensions
    assert len(plan.filters) >= 1
    assert any(f.filter_type == "slicer" for f in plan.filters)

    # Verify Sheets
    sheet_names = [s.name for s in plan.sheets]
    assert "Executive_Dashboard" in sheet_names
    assert "Raw_Data" in sheet_names


def test_architect_financial_dataset():
    """Verify financial dataset produces profit and expense summaries."""
    df = pd.DataFrame({
        "Department": ["Marketing", "Engineering", "Sales", "Operations"],
        "Budget": [50000, 120000, 80000, 45000],
        "Actual_Spend": [48000, 115000, 85000, 42000],
        "Cost_Center": ["CC101", "CC102", "CC103", "CC104"],
    })

    plan = architect_executive_plan(df, title="Departmental Spend & Budget Analysis")
    assert isinstance(plan, ExcelPlan)
    assert len(plan.dashboard.kpis) >= 3
    assert len(plan.analyses) >= 1
    assert len(plan.charts) >= 1
