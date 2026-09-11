"""
Tests for AI Dashboard Synthesizer — verifying AI question & dataset analysis,
zero-hallucination validation, smart NLP fallback, and multi-table workbook generation.
"""

import os
import pytest
import pandas as pd
from schemas.schemas import ExcelPlan
from services.ai_dashboard_synthesizer import (
    synthesize_ai_dashboard_plan,
    _find_closest_column,
    _smart_nlp_plan_fallback,
)
from services.workbook_generator import generate_workbook


@pytest.fixture
def sample_sales_df():
    return pd.DataFrame({
        "Region": ["North", "South", "East", "West", "North", "South"],
        "Product": ["Laptops", "Tablets", "Phones", "Laptops", "Tablets", "Phones"],
        "Month": ["Jan", "Jan", "Feb", "Feb", "Mar", "Mar"],
        "Salesperson": ["Sarah", "Michael", "David", "Sarah", "Michael", "David"],
        "Revenue": [45000, 28000, 32000, 51000, 31000, 34000],
        "Units": [45, 70, 64, 51, 75, 68],
    })


def test_find_closest_column():
    cols = ["Region", "Product Category", "Total Revenue", "Units Sold"]
    assert _find_closest_column("region", cols) == "Region"
    assert _find_closest_column("Revenue", cols) == "Total Revenue"
    assert _find_closest_column("Units", cols) == "Units Sold"
    assert _find_closest_column("category", cols) == "Product Category"


def test_smart_nlp_fallback(sample_sales_df):
    question = "Show total Revenue by Region, compare Units by Product, and track Month trend"
    plan = _smart_nlp_plan_fallback(sample_sales_df, question_text=question)

    assert isinstance(plan, ExcelPlan)
    assert plan.dashboard is not None
    assert len(plan.dashboard.kpis) >= 3

    # Primary metric should be Revenue because it was mentioned
    kpi_metrics = [k.metric for k in plan.dashboard.kpis]
    assert "Revenue" in kpi_metrics

    # Analyses should contain Region and Product
    analysis_dims = [a.group_by[0] for a in plan.analyses if a.group_by]
    assert "Region" in analysis_dims or "Product" in analysis_dims


def test_ai_dashboard_synthesis_end_to_end(sample_sales_df, tmp_path):
    question = "Create an executive sales dashboard showing Revenue by Region, Units by Product, and add slicers for Region"
    plan, meta = synthesize_ai_dashboard_plan(
        sample_sales_df,
        question_text=question,
        title="Executive Regional Sales",
    )

    assert isinstance(plan, ExcelPlan)
    assert plan.dashboard is not None
    assert len(plan.dashboard.kpis) >= 3
    assert len(plan.analyses) >= 2
    assert len(plan.charts) >= 1

    # Zero hallucination check: every metric and group_by column must exist in sample_sales_df
    for a in plan.analyses:
        assert a.metric in sample_sales_df.columns
        for g in a.group_by:
            assert g in sample_sales_df.columns

    for k in plan.dashboard.kpis:
        assert k.metric in sample_sales_df.columns

    # Verify workbook compilation with multiple tables and dynamic formulas
    out_xlsx = str(tmp_path / "executive_ai_dashboard.xlsx")
    summary = generate_workbook(sample_sales_df, plan, question, out_xlsx)

    assert os.path.exists(out_xlsx)
    assert summary["formula_count"] > 0
    assert "01_Executive_Dashboard" in summary["sheets"]
