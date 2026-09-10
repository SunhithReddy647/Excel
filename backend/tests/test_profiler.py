"""
Tests for the dataset profiler service.
"""

import os
import sys
import pytest
import pandas as pd

# Add parent directory to path
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from services.dataset_profiler import profile_dataset, read_dataset


FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")


class TestDatasetProfiler:
    """Test the dataset profiler with the sales fixture."""

    def test_read_csv(self):
        path = os.path.join(FIXTURES_DIR, "sales_data.csv")
        df = read_dataset(path)
        assert len(df) == 16
        assert len(df.columns) == 10

    def test_profile_sales(self):
        path = os.path.join(FIXTURES_DIR, "sales_data.csv")
        df, profile = profile_dataset(path)

        assert profile.row_count == 16
        assert profile.column_count == 10
        assert len(profile.headers) == 10

        # Check dimensions vs measures classification
        assert "Region" in profile.dimensions
        assert "Month" in profile.dimensions
        assert "Salesperson" in profile.dimensions
        assert "Product" in profile.dimensions

        assert "Units Sold" in profile.measures
        assert "Unit Price" in profile.measures
        assert "Unit Cost" in profile.measures
        assert "Target Units" in profile.measures

    def test_potential_calculations(self):
        path = os.path.join(FIXTURES_DIR, "sales_data.csv")
        df, profile = profile_dataset(path)

        # Should detect Revenue and Expenses as potential calculations
        potential = [p.lower() for p in profile.potential_calculations]
        assert "revenue" in potential or "expenses" in potential

    def test_column_types(self):
        path = os.path.join(FIXTURES_DIR, "sales_data.csv")
        df, profile = profile_dataset(path)

        col_map = {c.name: c for c in profile.columns}

        # Numeric columns should be detected
        assert col_map["Units Sold"].data_type in ("integer", "decimal")
        assert col_map["Unit Price"].data_type in ("integer", "decimal")

        # String columns
        assert col_map["Region"].data_type == "string"
        assert col_map["Product"].data_type == "string"

    def test_statistics(self):
        path = os.path.join(FIXTURES_DIR, "sales_data.csv")
        df, profile = profile_dataset(path)

        col_map = {c.name: c for c in profile.columns}

        units_col = col_map["Units Sold"]
        assert units_col.min_value is not None
        assert units_col.max_value is not None
        assert units_col.mean_value is not None
        assert units_col.unique_count > 0

    def test_duplicate_detection(self):
        path = os.path.join(FIXTURES_DIR, "sales_data.csv")
        df, profile = profile_dataset(path)
        assert profile.duplicate_rows == 0


class TestFormulaEngine:
    """Test the formula engine."""

    def test_template_to_structured_ref(self):
        from services.formula_engine import FormulaEngine
        engine = FormulaEngine("DataTable")

        result = engine.template_to_structured_ref("{Units Sold}*{Unit Price}")
        assert result == "=[@[Units Sold]]*[@[Unit Price]]"

    def test_sumifs(self):
        from services.formula_engine import FormulaEngine
        engine = FormulaEngine("DataTable")

        result = engine.sumifs("Revenue", {"Region": "A2"})
        assert "SUMIFS" in result
        assert "DataTable[Revenue]" in result
        assert "DataTable[Region]" in result

    def test_number_formats(self):
        from services.formula_engine import FormulaEngine
        assert FormulaEngine.get_number_format("currency") == "₹#,##0"
        assert FormulaEngine.get_number_format("percentage") == "0.00%"

    def test_business_formulas(self):
        from services.formula_engine import FormulaEngine

        margin = FormulaEngine.margin_pct("Profit", "Revenue")
        assert "[@[Profit]]" in margin
        assert "[@[Revenue]]" in margin

        achievement = FormulaEngine.achievement_pct("Units Sold", "Target Units")
        assert "[@[Units Sold]]" in achievement
        assert "[@[Target Units]]" in achievement


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
