"""
Formula Engine — converts formula templates to Excel structured references,
generates aggregation/conditional/business formulas deterministically.
"""

import re
from typing import Optional


class FormulaEngine:
    """Deterministic formula generator for Excel workbooks."""

    def __init__(self, table_name: str = "DataTable"):
        self.table_name = table_name

    # ─── Template Conversion ────────────────────────────────────────

    def template_to_structured_ref(self, template: str) -> str:
        """
        Convert a formula template like '{Units Sold}*{Unit Price}'
        into an Excel structured reference like '=[@[Units Sold]]*[@[Unit Price]]'.
        """
        def replace_ref(match):
            col_name = match.group(1)
            return f"[@[{col_name}]]"

        formula_body = re.sub(r"\{([^}]+)\}", replace_ref, template)
        return f"={formula_body}"

    def template_to_cell_ref(self, template: str, col_mapping: dict[str, str], row: int) -> str:
        """
        Convert a formula template to cell references.
        col_mapping: {column_name: column_letter}
        row: the Excel row number
        """
        def replace_ref(match):
            col_name = match.group(1)
            col_letter = col_mapping.get(col_name, "A")
            return f"{col_letter}{row}"

        formula_body = re.sub(r"\{([^}]+)\}", replace_ref, template)
        return f"={formula_body}"

    # ─── Aggregation Formulas ───────────────────────────────────────

    def sumifs(self, sum_col: str, criteria: dict[str, str]) -> str:
        """
        Generate a SUMIFS formula.
        criteria: {criteria_column: criteria_cell_ref}
        Example: SUMIFS(DataTable[Revenue],DataTable[Region],A2)
        """
        sum_range = f"{self.table_name}[{sum_col}]"
        criteria_parts = []
        for col, cell_ref in criteria.items():
            criteria_parts.append(f"{self.table_name}[{col}],{cell_ref}")
        return f"=SUMIFS({sum_range},{','.join(criteria_parts)})"

    def countifs(self, criteria: dict[str, str]) -> str:
        """Generate a COUNTIFS formula."""
        criteria_parts = []
        for col, cell_ref in criteria.items():
            criteria_parts.append(f"{self.table_name}[{col}],{cell_ref}")
        return f"=COUNTIFS({','.join(criteria_parts)})"

    def averageifs(self, avg_col: str, criteria: dict[str, str]) -> str:
        """Generate an AVERAGEIFS formula."""
        avg_range = f"{self.table_name}[{avg_col}]"
        criteria_parts = []
        for col, cell_ref in criteria.items():
            criteria_parts.append(f"{self.table_name}[{col}],{cell_ref}")
        return f"=AVERAGEIFS({avg_range},{','.join(criteria_parts)})"

    # ─── Simple Aggregation ─────────────────────────────────────────

    def aggregate(self, func: str, col: str, cell_range: Optional[str] = None) -> str:
        """
        Generate a simple aggregation formula.
        func: SUM, AVERAGE, COUNT, COUNTA, MIN, MAX
        col: column name (uses table reference) or cell_range if provided
        """
        ref = cell_range if cell_range else f"{self.table_name}[{col}]"
        return f"={func}({ref})"

    # ─── Conditional Formulas ───────────────────────────────────────

    def if_formula(self, condition: str, true_val: str, false_val: str) -> str:
        """Generate an IF formula."""
        return f"=IF({condition},{true_val},{false_val})"

    def conditional_structured(
        self,
        column: str,
        operator: str,
        threshold: str,
        true_val: str,
        false_val: str,
    ) -> str:
        """
        Generate an IF formula using structured references.
        Example: =IF([@[Target Achievement]]>1,"Yes","No")
        """
        condition = f"[@[{column}]]{operator}{threshold}"
        return f'=IF({condition},{true_val},{false_val})'

    # ─── Business Formulas ──────────────────────────────────────────

    @staticmethod
    def growth_pct(current_col: str, previous_col: str) -> str:
        """Growth % = (Current - Previous) / Previous"""
        return f"=([@[{current_col}]]-[@[{previous_col}]])/[@[{previous_col}]]"

    @staticmethod
    def margin_pct(profit_col: str, revenue_col: str) -> str:
        """Margin % = Profit / Revenue"""
        return f"=[@[{profit_col}]]/[@[{revenue_col}]]"

    @staticmethod
    def achievement_pct(actual_col: str, target_col: str) -> str:
        """Achievement % = Actual / Target"""
        return f"=[@[{actual_col}]]/[@[{target_col}]]"

    @staticmethod
    def variance(actual_col: str, target_col: str) -> str:
        """Variance = Actual - Target"""
        return f"=[@[{actual_col}]]-[@[{target_col}]]"

    @staticmethod
    def variance_pct(actual_col: str, target_col: str) -> str:
        """Variance % = (Actual - Target) / Target"""
        return f"=([@[{actual_col}]]-[@[{target_col}]])/[@[{target_col}]]"

    @staticmethod
    def net_revenue(revenue_col: str, discount_col: str) -> str:
        """Net Revenue = Revenue * (1 - Discount)"""
        return f"=[@[{revenue_col}]]*(1-[@[{discount_col}]])"

    # ─── Format Codes ───────────────────────────────────────────────

    @staticmethod
    def get_number_format(format_type: str) -> str:
        """Return Excel number format string for a given type."""
        formats = {
            "currency": "₹#,##0",
            "currency_decimal": "₹#,##0.00",
            "percentage": "0.00%",
            "number": "#,##0",
            "number_decimal": "#,##0.00",
            "integer": "#,##0",
            "date": "dd-mmm-yyyy",
            "text": "@",
        }
        return formats.get(format_type, "#,##0")


# Singleton for convenience
formula_engine = FormulaEngine()
