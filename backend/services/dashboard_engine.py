"""
Dashboard Engine — dynamically lays out KPIs, charts, filters, and breakdown tables
on a dashboard sheet using a professional grid-based system with conditional formatting,
data bars, color-coded KPI cards, and polished executive styling.
"""

from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.formatting.rule import DataBarRule, CellIsRule
from schemas.schemas import DashboardSpec, KPISpec
import pandas as pd
import numpy as np


# ─── Professional Color Palette ────────────────────────────────────

COLORS = {
    "primary":       "1B2A4A",
    "secondary":     "2E86AB",
    "accent1":       "A23B72",
    "accent2":       "F18F01",
    "accent3":       "2CA58D",
    "accent4":       "5C4D7D",
    "kpi_bg":        "F0F4F8",
    "kpi_border":    "D1D9E6",
    "white":         "FFFFFF",
    "dark_text":     "1B2A4A",
    "light_text":    "6B7B8D",
    "title_bg":      "1B2A4A",
    "title_text":    "FFFFFF",
    "stripe_even":   "F8FAFC",
    "stripe_odd":    "FFFFFF",
    "green_pos":     "E6F4EA",
    "green_txt":     "137333",
    "red_neg":       "FCE8E6",
    "red_txt":       "C5221F",
    "separator":     "E2E8F0",
    "header_bg":     "334155",
    "header_text":   "FFFFFF",
}

KPI_COLORS = [
    ("2E86AB", "EBF5FA"),  # Teal
    ("A23B72", "FAEEF4"),  # Berry
    ("F18F01", "FFF4E5"),  # Orange
    ("2CA58D", "EBF9F5"),  # Green
    ("5C4D7D", "F0EDF5"),  # Purple
    ("E84855", "FDEBEC"),  # Red
]


class DashboardEngine:
    """Generates a dynamic, polished dashboard layout on an Excel worksheet."""

    def __init__(self, ws):
        self.ws = ws
        self.current_row = 1
        self.max_col = 14  # Columns A through N

    def _set_column_widths(self):
        """Set consistent column widths for the dashboard grid."""
        for i in range(1, self.max_col + 1):
            self.ws.column_dimensions[get_column_letter(i)].width = 12

    def _thin_border(self, color=None):
        c = color or COLORS["kpi_border"]
        return Border(
            left=Side(style="thin", color=c),
            right=Side(style="thin", color=c),
            top=Side(style="thin", color=c),
            bottom=Side(style="thin", color=c),
        )

    def _merge_and_style(
        self,
        start_row: int,
        start_col: int,
        end_row: int,
        end_col: int,
        value: str,
        font: Font = None,
        alignment: Alignment = None,
        fill: PatternFill = None,
        border: Border = None,
    ):
        """Merge cells and apply styling."""
        start_cell = f"{get_column_letter(start_col)}{start_row}"
        end_cell = f"{get_column_letter(end_col)}{end_row}"
        self.ws.merge_cells(f"{start_cell}:{end_cell}")

        cell = self.ws[start_cell]
        cell.value = value
        if font:
            cell.font = font
        if alignment:
            cell.alignment = alignment
        if fill:
            cell.fill = fill
        if border:
            cell.border = border

    def write_title(self, title: str):
        """Write the dashboard title bar with a professional header."""
        # Title bar
        self._merge_and_style(
            self.current_row, 1, self.current_row + 1, self.max_col,
            value=title.upper(),
            font=Font(name="Calibri", size=18, bold=True, color=COLORS["title_text"]),
            alignment=Alignment(horizontal="center", vertical="center"),
            fill=PatternFill(start_color=COLORS["title_bg"], end_color=COLORS["title_bg"], fill_type="solid"),
        )
        self.ws.row_dimensions[self.current_row].height = 25
        self.ws.row_dimensions[self.current_row + 1].height = 25

        # Apply border to all cells in the title bar
        for col in range(1, self.max_col + 1):
            for row in [self.current_row, self.current_row + 1]:
                self.ws.cell(row=row, column=col).border = self._thin_border(COLORS["primary"])

        # Subtitle / timestamp line
        self.current_row += 2
        from datetime import datetime
        timestamp = datetime.now().strftime("%B %d, %Y")
        self._merge_and_style(
            self.current_row, 1, self.current_row, self.max_col,
            value=f"Generated: {timestamp}  |  Powered by ExcelFlow AI",
            font=Font(name="Calibri", size=9, italic=True, color=COLORS["light_text"]),
            alignment=Alignment(horizontal="center", vertical="center"),
            fill=PatternFill(start_color=COLORS["stripe_even"], end_color=COLORS["stripe_even"], fill_type="solid"),
        )
        self.ws.row_dimensions[self.current_row].height = 20
        self.current_row += 2  # Title + subtitle + spacing

    def write_kpis(self, kpis: list[dict]):
        """
        Write KPI cards across the dashboard with color-coded accent bars.
        kpis: list of {"label": str, "value": any, "formula": str|None, "format": str}
        """
        if not kpis:
            return

        num_kpis = min(len(kpis), 6)
        cols_per_kpi = max(2, self.max_col // num_kpis)
        border = self._thin_border()

        for i, kpi in enumerate(kpis[:6]):
            start_col = 1 + i * cols_per_kpi
            end_col = min(start_col + cols_per_kpi - 1, self.max_col)
            fg_color, bg_color = KPI_COLORS[i % len(KPI_COLORS)]

            # Colored accent bar (top)
            self._merge_and_style(
                self.current_row, start_col, self.current_row, end_col,
                value="",
                fill=PatternFill(start_color=fg_color, end_color=fg_color, fill_type="solid"),
            )
            self.ws.row_dimensions[self.current_row].height = 4

            # KPI Label
            self._merge_and_style(
                self.current_row + 1, start_col, self.current_row + 1, end_col,
                value=kpi["label"].upper(),
                font=Font(name="Calibri", size=9, bold=True, color=COLORS["light_text"]),
                alignment=Alignment(horizontal="center", vertical="center"),
                fill=PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid"),
                border=border,
            )

            # KPI Value & Dynamic Formula
            formula = kpi.get("formula")
            value = kpi.get("value")
            fmt = kpi.get("format", "number")

            num_formats = {
                "currency": "$#,##0",
                "currency_decimal": "$#,##0.00",
                "percentage": "0.0%",
                "integer": "#,##0",
                "number": "#,##0",
                "number_decimal": "#,##0.00",
            }
            excel_num_format = num_formats.get(fmt, "#,##0")

            # If we have a formula, write it as a live Excel formula
            if formula:
                cell_val = formula
            else:
                cell_val = value if isinstance(value, (int, float)) else str(value)

            self._merge_and_style(
                self.current_row + 2, start_col, self.current_row + 3, end_col,
                value=cell_val,
                font=Font(name="Calibri", size=22, bold=True, color=fg_color),
                alignment=Alignment(horizontal="center", vertical="center"),
                fill=PatternFill(start_color=COLORS["white"], end_color=COLORS["white"], fill_type="solid"),
                border=border,
            )

            # Apply number format to the value cell
            start_cell_ref = f"{get_column_letter(start_col)}{self.current_row + 2}"
            self.ws[start_cell_ref].number_format = excel_num_format

        self.ws.row_dimensions[self.current_row + 1].height = 22
        self.ws.row_dimensions[self.current_row + 2].height = 30
        self.ws.row_dimensions[self.current_row + 3].height = 15
        self.current_row += 5  # accent + label + value(2) + spacing

    def write_breakdown_table(
        self,
        title: str,
        result_df,
        metric_col: str,
        dim_col: str,
        aggregation: str,
        table_name: str,
        is_currency: bool = False,
        start_col: int = 1,
        col_span: int = 6,
    ) -> int:
        """
        Write a compact breakdown table with header, data rows, grand total,
        and data bars on the value column.  Returns the row after the table.
        """
        ws = self.ws
        end_col = start_col + col_span - 1

        # Section header
        self._merge_and_style(
            self.current_row, start_col, self.current_row, end_col,
            value=title,
            font=Font(name="Calibri", size=12, bold=True, color=COLORS["primary"]),
            alignment=Alignment(horizontal="left", vertical="center"),
        )
        # Separator line under the header
        for c in range(start_col, end_col + 1):
            ws.cell(row=self.current_row, column=c).border = Border(
                bottom=Side(style="medium", color=COLORS["secondary"])
            )
        self.current_row += 1

        if result_df is None or result_df.empty:
            ws.cell(row=self.current_row, column=start_col, value="No data available")
            ws.cell(row=self.current_row, column=start_col).font = Font(
                name="Calibri", size=10, italic=True, color=COLORS["light_text"]
            )
            self.current_row += 2
            return self.current_row

        # Table header
        header_fill = PatternFill(start_color=COLORS["header_bg"], end_color=COLORS["header_bg"], fill_type="solid")
        header_font = Font(name="Calibri", size=10, bold=True, color=COLORS["header_text"])
        header_align = Alignment(horizontal="center", vertical="center")

        for col_idx, col_name in enumerate(result_df.columns):
            c = start_col + col_idx
            cell = ws.cell(row=self.current_row, column=c, value=col_name)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = header_align
            cell.border = self._thin_border(COLORS["header_bg"])
        self.current_row += 1

        # Data rows with alternating stripes and SUMIFS formulas
        data_start = self.current_row
        num_fmt = "$#,##0" if is_currency else "#,##0"

        for row_idx, (_, row) in enumerate(result_df.iterrows()):
            fill_color = COLORS["stripe_even"] if row_idx % 2 == 0 else COLORS["stripe_odd"]
            row_fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type="solid")

            cat_val = row[result_df.columns[0]]
            ws.cell(row=self.current_row, column=start_col, value=cat_val)
            ws.cell(row=self.current_row, column=start_col).font = Font(name="Calibri", size=10)
            ws.cell(row=self.current_row, column=start_col).fill = row_fill

            # Dynamic Excel formula instead of static value
            agg_upper = aggregation.upper()
            if agg_upper == "AVERAGE":
                formula = f'=AVERAGEIFS({table_name}[{metric_col}], {table_name}[{dim_col}], "{cat_val}")'
            elif agg_upper in ("COUNT", "COUNTA"):
                formula = f'=COUNTIFS({table_name}[{dim_col}], "{cat_val}")'
            else:
                formula = f'=SUMIFS({table_name}[{metric_col}], {table_name}[{dim_col}], "{cat_val}")'

            val_cell = ws.cell(row=self.current_row, column=start_col + 1, value=formula)
            val_cell.number_format = num_fmt
            val_cell.font = Font(name="Calibri", size=10, bold=True)
            val_cell.fill = row_fill

            self.current_row += 1

        data_end = self.current_row - 1

        # Grand Total row
        tot_fill = PatternFill(start_color=COLORS["kpi_bg"], end_color=COLORS["kpi_bg"], fill_type="solid")
        ws.cell(row=self.current_row, column=start_col, value="GRAND TOTAL")
        ws.cell(row=self.current_row, column=start_col).font = Font(name="Calibri", size=10, bold=True, color=COLORS["primary"])
        ws.cell(row=self.current_row, column=start_col).fill = tot_fill
        ws.cell(row=self.current_row, column=start_col).border = Border(top=Side(style="double", color=COLORS["primary"]))

        val_letter = get_column_letter(start_col + 1)
        if aggregation.upper() == "AVERAGE":
            total_formula = f"=AVERAGE({val_letter}{data_start}:{val_letter}{data_end})"
        else:
            total_formula = f"=SUM({val_letter}{data_start}:{val_letter}{data_end})"

        tot_cell = ws.cell(row=self.current_row, column=start_col + 1, value=total_formula)
        tot_cell.font = Font(name="Calibri", size=10, bold=True, color=COLORS["primary"])
        tot_cell.number_format = num_fmt
        tot_cell.fill = tot_fill
        tot_cell.border = Border(top=Side(style="double", color=COLORS["primary"]))

        self.current_row += 1

        # Add data bars (conditional formatting) on the value column
        if data_end >= data_start:
            bar_range = f"{val_letter}{data_start}:{val_letter}{data_end}"
            try:
                rule = DataBarRule(
                    start_type="min", end_type="max",
                    color="2E86AB",
                    showValue=True,
                    minLength=None,
                    maxLength=None,
                )
                ws.conditional_formatting.add(bar_range, rule)
            except Exception:
                pass

        self.current_row += 1  # spacing after table
        return self.current_row

    def get_chart_position(self, chart_index: int, charts_per_row: int = 2) -> str:
        """Calculate chart position based on index in the chart grid."""
        row_in_grid = chart_index // charts_per_row
        col_in_grid = chart_index % charts_per_row

        chart_height_rows = 15
        chart_row = self.current_row + row_in_grid * (chart_height_rows + 1)

        if charts_per_row == 1:
            chart_col = 1
        else:
            cols_per_chart = self.max_col // charts_per_row
            chart_col = 1 + col_in_grid * cols_per_chart

        return f"{get_column_letter(chart_col)}{chart_row}"

    def advance_past_charts(self, num_charts: int, charts_per_row: int = 2):
        """Advance current_row past the chart grid area."""
        chart_height_rows = 15
        num_rows_of_charts = (num_charts + charts_per_row - 1) // charts_per_row
        self.current_row += num_rows_of_charts * (chart_height_rows + 1) + 1

    def write_section_header(self, title: str):
        """Write a section header with separator."""
        self._merge_and_style(
            self.current_row, 1, self.current_row, self.max_col,
            value=title,
            font=Font(name="Calibri", size=12, bold=True, color=COLORS["primary"]),
            alignment=Alignment(horizontal="left", vertical="center"),
        )
        for c in range(1, self.max_col + 1):
            self.ws.cell(row=self.current_row, column=c).border = Border(
                bottom=Side(style="medium", color=COLORS["secondary"])
            )
        self.current_row += 1

    def write_filters_section(self, filters: list):
        """Write a professional filters/slicers reference section."""
        if not filters:
            return

        self.write_section_header("Interactive Filters & Slicers")

        for f in filters:
            field = f.field if hasattr(f, 'field') else f.get("field", "")
            desc = f.description if hasattr(f, 'description') else f.get("description", "")
            filter_type = f.filter_type if hasattr(f, 'filter_type') else f.get("filter_type", "")

            row = self.current_row
            self.ws.cell(row=row, column=1, value=f"✓ {field}")
            self.ws.cell(row=row, column=1).font = Font(name="Calibri", size=10, bold=True, color=COLORS["accent3"])
            self.ws.cell(row=row, column=4, value=f"Type: {filter_type}")
            self.ws.cell(row=row, column=4).font = Font(name="Calibri", size=10, color=COLORS["light_text"])
            self.ws.cell(row=row, column=7, value=desc)
            self.ws.cell(row=row, column=7).font = Font(name="Calibri", size=10, italic=True, color=COLORS["light_text"])
            self.current_row += 1

        self.current_row += 1  # spacing


dashboard_engine_cls = DashboardEngine
