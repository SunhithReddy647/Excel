"""
Analysis Engine — performs aggregations using pandas and generates
Excel SUMIFS-based summary tables for the workbook.
"""

import pandas as pd
import numpy as np
from typing import Optional
from schemas.schemas import AnalysisSpec
from services.formula_engine import FormulaEngine


# Month ordering for chronological sorting
MONTH_ORDER = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


class AnalysisEngine:
    """Generates summary/analysis tables from data based on AnalysisSpec."""

    def __init__(self, df: pd.DataFrame, table_name: str = "DataTable"):
        self.df = df
        self.table_name = table_name
        self.formula_engine = FormulaEngine(table_name)

    def _resolve_metric(self, metric: str) -> str:
        """Resolve a metric name to an actual column, handling calculated columns."""
        # Direct match
        if metric in self.df.columns:
            return metric
        # Case-insensitive match
        for col in self.df.columns:
            if col.lower() == metric.lower():
                return col
        # Not found — return as-is (might be a calculated column added later)
        return metric

    def _sort_values(self, result_df: pd.DataFrame, group_col: str, sort: str) -> pd.DataFrame:
        """Apply sorting to the aggregation result."""
        if sort == "chronological":
            # Try to sort by month order
            lower_vals = result_df[group_col].astype(str).str.lower().str.strip()
            month_keys = lower_vals.map(MONTH_ORDER)
            if month_keys.notna().all():
                result_df = result_df.assign(_sort_key=month_keys.values).sort_values("_sort_key").drop(columns=["_sort_key"])
            else:
                # Try date parsing
                try:
                    result_df = result_df.sort_values(group_col)
                except Exception:
                    pass
        elif sort == "ascending":
            # Sort by value ascending
            value_cols = [c for c in result_df.columns if c != group_col and c not in self.df.columns]
            if value_cols:
                result_df = result_df.sort_values(value_cols[0], ascending=True)
            else:
                result_df = result_df.sort_values(group_col, ascending=True)
        elif sort == "descending":
            value_cols = [c for c in result_df.columns if c != group_col and c not in self.df.columns]
            if value_cols:
                result_df = result_df.sort_values(value_cols[0], ascending=False)
            else:
                result_df = result_df.sort_values(group_col, ascending=False)

        return result_df.reset_index(drop=True)

    def compute_analysis(self, spec: AnalysisSpec) -> pd.DataFrame:
        """
        Compute the aggregation using pandas.
        Returns a DataFrame with group columns + metric value.
        """
        metric = self._resolve_metric(spec.metric)

        if metric not in self.df.columns:
            raise ValueError(f"Metric column '{metric}' not found in dataset. Available: {list(self.df.columns)}")

        agg_func = spec.aggregation.upper()
        agg_map = {
            "SUM": "sum",
            "AVERAGE": "mean",
            "COUNT": "count",
            "COUNTA": "count",
            "MIN": "min",
            "MAX": "max",
        }
        pd_func = agg_map.get(agg_func, "sum")

        if not spec.group_by:
            # Single aggregate — no grouping
            val = getattr(self.df[metric], pd_func)()
            return pd.DataFrame([{spec.metric: val}])

        group_cols = []
        for g in spec.group_by:
            resolved = self._resolve_metric(g)
            if resolved in self.df.columns:
                group_cols.append(resolved)
            else:
                raise ValueError(f"Group-by column '{g}' not found in dataset.")

        if len(group_cols) == 1:
            result = self.df.groupby(group_cols[0], sort=False)[metric].agg(pd_func).reset_index()
            result.columns = [group_cols[0], spec.metric]
            result = self._sort_values(result, group_cols[0], spec.sort)
        elif len(group_cols) == 2:
            # Cross-tab / pivot for two dimensions
            result = self.df.pivot_table(
                values=metric,
                index=group_cols[0],
                columns=group_cols[1],
                aggfunc=pd_func,
                fill_value=0,
            ).reset_index()
            # Sort if needed
            if spec.sort == "chronological":
                # Try sorting columns by month order
                non_index_cols = [c for c in result.columns if c != group_cols[0]]
                sorted_cols = sorted(non_index_cols, key=lambda c: MONTH_ORDER.get(str(c).lower().strip(), 999))
                result = result[[group_cols[0]] + sorted_cols]
        else:
            result = self.df.groupby(group_cols, sort=False)[metric].agg(pd_func).reset_index()
            result.columns = group_cols + [spec.metric]

        # Convert numpy types for safety
        for col in result.select_dtypes(include=[np.integer]).columns:
            result[col] = result[col].astype(int)
        for col in result.select_dtypes(include=[np.floating]).columns:
            result[col] = result[col].astype(float)

        return result

    def generate_sumifs_formulas(
        self,
        spec: AnalysisSpec,
        group_values: list,
        group_col_letter: str,
        start_row: int,
    ) -> list[str]:
        """
        Generate SUMIFS/COUNTIFS/AVERAGEIFS formulas for a summary table.
        Returns a list of formula strings, one per group value.
        """
        metric = self._resolve_metric(spec.metric)
        formulas = []

        for i, _ in enumerate(group_values):
            row = start_row + i
            criteria = {spec.group_by[0]: f"{group_col_letter}{row}"}

            if spec.aggregation.upper() in ("SUM",):
                formula = self.formula_engine.sumifs(metric, criteria)
            elif spec.aggregation.upper() in ("COUNT", "COUNTA"):
                formula = self.formula_engine.countifs(criteria)
            elif spec.aggregation.upper() == "AVERAGE":
                formula = self.formula_engine.averageifs(metric, criteria)
            else:
                formula = self.formula_engine.sumifs(metric, criteria)

            formulas.append(formula)

        return formulas
