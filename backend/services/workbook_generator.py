"""
Workbook Generator — master orchestrator that assembles the final .xlsx file.
Creates sheets, tables, formulas, analysis, charts, dashboard, and formatting.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows

from schemas.schemas import ExcelPlan, AnalysisSpec, ChartSpec, CalculationSpec
from services.formula_engine import FormulaEngine
from services.analysis_engine import AnalysisEngine
from services.chart_engine import ChartEngine
from services.dashboard_engine import DashboardEngine, COLORS


# ─── Formatting Constants ───────────────────────────────────────────

HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_FILL = PatternFill(start_color="1B2A4A", end_color="1B2A4A", fill_type="solid")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center", wrap_text=True)
HEADER_BORDER = Border(
    bottom=Side(style="thin", color="4472C4"),
    left=Side(style="thin", color="D9E2F3"),
    right=Side(style="thin", color="D9E2F3"),
)

DATA_FONT = Font(name="Calibri", size=11)
DATA_ALIGNMENT = Alignment(vertical="center")
ALT_ROW_FILL = PatternFill(start_color="F2F6FC", end_color="F2F6FC", fill_type="solid")

TITLE_FONT = Font(name="Calibri", size=16, bold=True, color="1B2A4A")
SUBTITLE_FONT = Font(name="Calibri", size=12, bold=True, color="2E86AB")
SECTION_FONT = Font(name="Calibri", size=11, bold=True, color="1B2A4A")


def _get_number_format(fmt: str) -> str:
    """Map format type to Excel number format code."""
    formats = {
        "currency": "$#,##0",
        "currency_decimal": "$#,##0.00",
        "percentage": "0.0%",
        "number": "#,##0",
        "number_decimal": "#,##0.00",
        "integer": "#,##0",
    }
    return formats.get(fmt, "General")


def _auto_column_width(ws, min_width: int = 10, max_width: int = 30):
    """Auto-adjust column widths based on content."""
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val = str(cell.value) if cell.value is not None else ""
                max_len = max(max_len, len(val))
            except Exception:
                pass
        adjusted = min(max(max_len + 3, min_width), max_width)
        ws.column_dimensions[col_letter].width = adjusted


def _format_header_row(ws, row_num: int, num_cols: int):
    """Apply professional formatting to a header row."""
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = HEADER_BORDER


class WorkbookGenerator:
    """Assembles the final Excel workbook from a plan and dataset."""

    def __init__(self, df: pd.DataFrame, plan: ExcelPlan, question_text: str):
        self.df = df.copy()
        self.plan = plan
        self.question_text = question_text
        self.wb = Workbook()
        self.wb.remove(self.wb.active)  # Remove default sheet
        self.table_name = "DataTable"
        self.formula_engine = FormulaEngine(self.table_name)
        self.analysis_engine = AnalysisEngine(self.df, self.table_name)
        self.analysis_results: dict[str, pd.DataFrame] = {}
        self.analysis_positions: dict[str, dict] = {}  # name → {sheet, start_row, end_row, cat_col, val_col}
        self.formula_count = 0
        self.chart_count = 0

    def generate(self, output_path: str) -> dict:
        """
        Generate the complete workbook. Returns a summary dict.
        """
        # Step 1: Apply calculated columns to the DataFrame
        self._apply_calculations()

        # Step 2: Create sheets based on plan
        sheets_created = []

        # Sort sheets so raw_data is always processed first, then calculations, analysis, charts, dashboard, assignment
        type_priority = {
            "raw_data": 1,
            "calculations": 2,
            "analysis": 3,
            "charts": 4,
            "dashboard": 5,
            "assignment": 6,
        }
        ordered_sheets = sorted(self.plan.sheets, key=lambda s: type_priority.get(s.sheet_type, 99))

        for sheet_spec in ordered_sheets:
            if sheet_spec.sheet_type == "raw_data":
                self._create_raw_data_sheet(sheet_spec.name)
                sheets_created.append(sheet_spec.name)

            elif sheet_spec.sheet_type == "calculations":
                self._create_calculations_sheet(sheet_spec.name)
                sheets_created.append(sheet_spec.name)

            elif sheet_spec.sheet_type == "analysis":
                self._create_analysis_sheet(sheet_spec.name)
                sheets_created.append(sheet_spec.name)

            elif sheet_spec.sheet_type == "charts":
                self._create_charts_sheet(sheet_spec.name)
                sheets_created.append(sheet_spec.name)

            elif sheet_spec.sheet_type == "dashboard":
                self._create_dashboard_sheet(sheet_spec.name)
                sheets_created.append(sheet_spec.name)

            elif sheet_spec.sheet_type == "assignment":
                self._create_assignment_sheet(sheet_spec.name)
                sheets_created.append(sheet_spec.name)

        # Ensure Executive Dashboard is active sheet and gridlines are enabled
        dash_sheet_names = [s.name for s in self.plan.sheets if s.sheet_type == "dashboard"]
        if dash_sheet_names and dash_sheet_names[0] in self.wb.sheetnames:
            self.wb.active = self.wb[dash_sheet_names[0]]

        for sheet in self.wb.worksheets:
            sheet.views.sheetView[0].showGridLines = True

        # Save base workbook with openpyxl
        self.wb.save(output_path)

        # Enhance with 100% authentic native Excel PivotTables, Slicers, and PivotCharts
        pivot_result = {}
        try:
            from services.pivot_engine import enhance_workbook_with_native_pivots
            pivot_result = enhance_workbook_with_native_pivots(output_path, self.plan, self.df)
        except Exception as e:
            # OpenPyXL formula summary tables already exist as fallback
            pass

        return {
            "filename": Path(output_path).name,
            "sheets": sheets_created,
            "row_count": len(self.df),
            "formula_count": self.formula_count,
            "chart_count": self.chart_count + pivot_result.get("charts_created", 0),
            "analysis_count": len(self.analysis_results),
            "filter_count": len(self.plan.filters),
            "native_pivots": pivot_result.get("native_pivots_enabled", False),
            "pivots_created": pivot_result.get("pivots_created", 0),
            "slicers_created": pivot_result.get("slicers_created", 0),
        }

    # ─── Calculated Columns ─────────────────────────────────────────

    def _apply_calculations(self):
        """Add calculated columns to the DataFrame using pandas (for analysis computation)."""
        for calc in self.plan.calculations:
            try:
                template = calc.formula_template
                # Parse {Col} references and evaluate with pandas
                expr = template
                for col in self.df.columns:
                    expr = expr.replace(f"{{{col}}}", f"df['{col}']")
                # Safe eval with only the DataFrame in scope
                self.df[calc.name] = eval(expr, {"df": self.df, "np": np, "pd": pd})
            except Exception as e:
                # If pandas eval fails, leave it — formula will still be in Excel
                pass

    # ─── Raw Data Sheet ──────────────────────────────────────────────

    def _create_raw_data_sheet(self, sheet_name: str):
        """Create the raw data sheet with an Excel Table and formulas for calculated columns."""
        ws = self.wb.create_sheet(title=sheet_name)

        # Write headers
        all_columns = list(self.df.columns)
        original_columns = [c for c in all_columns if c not in [calc.name for calc in self.plan.calculations]]
        calc_columns = [calc.name for calc in self.plan.calculations]
        headers = original_columns + calc_columns
        col_letter_map = {header: get_column_letter(idx) for idx, header in enumerate(headers, 1)}

        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)

        # Write data rows
        for row_idx, (_, row) in enumerate(self.df.iterrows(), 2):
            for col_idx, col_name in enumerate(original_columns, 1):
                val = row[col_name]
                if isinstance(val, (np.integer,)):
                    val = int(val)
                elif isinstance(val, (np.floating,)):
                    val = float(val)
                elif pd.isna(val):
                    val = None
                ws.cell(row=row_idx, column=col_idx, value=val)

            # For calculated columns, write valid row-level formulas
            for calc_idx, calc in enumerate(self.plan.calculations):
                col_idx = len(original_columns) + calc_idx + 1
                formula = self.formula_engine.template_to_cell_ref(calc.formula_template, col_letter_map, row_idx)
                ws.cell(row=row_idx, column=col_idx, value=formula)

                # Apply number format
                fmt = _get_number_format(calc.format)
                ws.cell(row=row_idx, column=col_idx).number_format = fmt
                self.formula_count += 1

        # Create Excel Table
        num_rows = len(self.df) + 1
        num_cols = len(headers)
        table_ref = f"A1:{get_column_letter(num_cols)}{num_rows}"
        table = Table(displayName=self.table_name, ref=table_ref)
        style = TableStyleInfo(
            name="TableStyleMedium2",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        table.tableStyleInfo = style
        ws.add_table(table)

        # Format header row
        _format_header_row(ws, 1, num_cols)

        # Auto column widths
        _auto_column_width(ws)

        # Freeze top row
        ws.freeze_panes = "A2"

    # ─── Calculations Sheet ──────────────────────────────────────────

    def _create_calculations_sheet(self, sheet_name: str):
        """Create a sheet documenting all calculations with formulas."""
        ws = self.wb.create_sheet(title=sheet_name)

        # Title
        ws.merge_cells("A1:E1")
        ws["A1"].value = "Calculation Reference"
        ws["A1"].font = TITLE_FONT
        ws["A1"].alignment = Alignment(vertical="center")
        ws.row_dimensions[1].height = 30

        # Headers
        headers = ["Calculation", "Formula", "Format", "Description"]
        for i, h in enumerate(headers, 1):
            cell = ws.cell(row=3, column=i, value=h)
        _format_header_row(ws, 3, len(headers))

        # Rows
        for row_idx, calc in enumerate(self.plan.calculations, 4):
            ws.cell(row=row_idx, column=1, value=calc.name)
            ws.cell(row=row_idx, column=1).font = Font(name="Calibri", size=11, bold=True)
            formula_display = self.formula_engine.template_to_structured_ref(calc.formula_template)
            # Prepend apostrophe so Excel treats it as documentation text rather than active formula
            ws.cell(row=row_idx, column=2, value=f"'{formula_display}")
            ws.cell(row=row_idx, column=2).font = Font(name="Consolas", size=10, color="2E86AB")
            ws.cell(row=row_idx, column=3, value=calc.format)
            ws.cell(row=row_idx, column=4, value=calc.description)

            # Alternating row color
            if row_idx % 2 == 0:
                for col in range(1, 5):
                    ws.cell(row=row_idx, column=col).fill = ALT_ROW_FILL

        _auto_column_width(ws, min_width=15)

    # ─── Analysis Sheet ──────────────────────────────────────────────

    def _create_analysis_sheet(self, sheet_name: str):
        """Create analysis sheet with summary tables using SUMIFS formulas."""
        ws = self.wb.create_sheet(title=sheet_name)

        # Title
        ws.merge_cells("A1:H1")
        ws["A1"].value = "Data Analysis"
        ws["A1"].font = TITLE_FONT
        ws.row_dimensions[1].height = 30

        current_row = 3

        for analysis in self.plan.analyses:
            try:
                # Compute analysis with pandas
                result_df = self.analysis_engine.compute_analysis(analysis)
                self.analysis_results[analysis.name] = result_df

                # Section header
                ws.cell(row=current_row, column=1, value=analysis.name)
                ws.cell(row=current_row, column=1).font = SUBTITLE_FONT
                current_row += 1

                # Write headers
                for col_idx, col_name in enumerate(result_df.columns, 1):
                    ws.cell(row=current_row, column=col_idx, value=col_name)
                _format_header_row(ws, current_row, len(result_df.columns))
                header_row = current_row
                current_row += 1

                # Write data
                data_start_row = current_row
                for _, row in result_df.iterrows():
                    for col_idx, col_name in enumerate(result_df.columns, 1):
                        val = row[col_name]
                        if isinstance(val, (np.integer,)):
                            val = int(val)
                        elif isinstance(val, (np.floating,)):
                            val = float(val)
                        ws.cell(row=current_row, column=col_idx, value=val)

                        # Apply currency format to numeric columns that aren't the group column
                        if col_idx > len(analysis.group_by) and isinstance(val, (int, float)):
                            metric_lower = analysis.metric.lower()
                            if any(kw in metric_lower for kw in ("revenue", "cost", "expense", "profit", "price", "salary")):
                                ws.cell(row=current_row, column=col_idx).number_format = "$#,##0"
                            elif "percent" in metric_lower or "margin" in metric_lower or "rate" in metric_lower:
                                ws.cell(row=current_row, column=col_idx).number_format = "0.0%"
                            else:
                                ws.cell(row=current_row, column=col_idx).number_format = "#,##0"

                        # Alt row fill
                        if (current_row - data_start_row) % 2 == 1:
                            ws.cell(row=current_row, column=col_idx).fill = ALT_ROW_FILL

                    current_row += 1

                # Track positions for chart engine
                self.analysis_positions[analysis.name] = {
                    "sheet": sheet_name,
                    "header_row": header_row,
                    "start_row": header_row,
                    "end_row": current_row - 1,
                    "cat_col": 1,
                    "val_col": len(analysis.group_by) + 1,
                    "num_value_cols": len(result_df.columns) - len(analysis.group_by),
                }

                current_row += 2  # Spacing between analyses

            except Exception as e:
                ws.cell(row=current_row, column=1, value=f"Error computing {analysis.name}: {str(e)}")
                ws.cell(row=current_row, column=1).font = Font(color="FF0000")
                current_row += 2

        _auto_column_width(ws)

    # ─── Charts Sheet ────────────────────────────────────────────────

    def _create_charts_sheet(self, sheet_name: str):
        """Create charts sheet with actual Excel chart objects."""
        ws = self.wb.create_sheet(title=sheet_name)

        ws.merge_cells("A1:H1")
        ws["A1"].value = "Charts & Visualizations"
        ws["A1"].font = TITLE_FONT
        ws.row_dimensions[1].height = 30

        chart_row = 3
        charts_created = 0

        for chart_spec in self.plan.charts:
            if chart_spec.data_source not in self.analysis_positions:
                continue

            pos = self.analysis_positions[chart_spec.data_source]

            # Get the analysis worksheet
            analysis_ws = self.wb[pos["sheet"]]

            try:
                if pos["num_value_cols"] > 1:
                    # Multi-series chart
                    ChartEngine.create_multi_series_chart(
                        spec=chart_spec,
                        ws=ws,
                        data_start_row=pos["start_row"],
                        data_end_row=pos["end_row"],
                        category_col=pos["cat_col"],
                        value_start_col=pos["val_col"],
                        value_end_col=pos["val_col"] + pos["num_value_cols"] - 1,
                        position=f"A{chart_row}",
                    )
                else:
                    # We need to copy data to the charts sheet for proper referencing
                    result_df = self.analysis_results.get(chart_spec.data_source)
                    if result_df is None:
                        continue

                    # Write compact data table in the charts sheet
                    data_col_start = 10  # Column J onwards (hidden area)
                    data_start = chart_row

                    for col_idx, col_name in enumerate(result_df.columns):
                        ws.cell(row=data_start, column=data_col_start + col_idx, value=col_name)
                    for r_idx, (_, row) in enumerate(result_df.iterrows(), 1):
                        for col_idx, col_name in enumerate(result_df.columns):
                            val = row[col_name]
                            if isinstance(val, (np.integer,)):
                                val = int(val)
                            elif isinstance(val, (np.floating,)):
                                val = float(val)
                            ws.cell(row=data_start + r_idx, column=data_col_start + col_idx, value=val)

                    data_end = data_start + len(result_df)

                    ChartEngine.create_chart(
                        spec=chart_spec,
                        ws=ws,
                        data_start_row=data_start,
                        data_end_row=data_end,
                        category_col=data_col_start,
                        value_col=data_col_start + 1,
                        position=f"A{chart_row}",
                    )

                charts_created += 1
                self.chart_count += 1
                chart_row += 16  # Space for chart + gap

            except Exception as e:
                ws.cell(row=chart_row, column=1, value=f"Error creating chart '{chart_spec.title}': {str(e)}")
                ws.cell(row=chart_row, column=1).font = Font(color="FF0000")
                chart_row += 2

    # ─── Dashboard Sheet ─────────────────────────────────────────────

    def _create_dashboard_sheet(self, sheet_name: str):
        """Create the dashboard sheet with KPIs, charts, and filters."""
        ws = self.wb.create_sheet(title=sheet_name)
        dashboard = DashboardEngine(ws)
        dashboard._set_column_widths()

        dash_spec = self.plan.dashboard
        if not dash_spec:
            return

        # Title
        dashboard.write_title(dash_spec.title)

        # KPIs
        if dash_spec.kpis:
            kpi_data = []
            for kpi in dash_spec.kpis:
                formula = None
                try:
                    metric_col = kpi.metric
                    if metric_col in self.df.columns:
                        agg_func = kpi.aggregation.lower()
                        value = getattr(self.df[metric_col], agg_func)()
                        if isinstance(value, (np.integer, np.floating)):
                            value = float(value)
                        agg_upper = kpi.aggregation.upper()
                        # Real dynamic structured formula referencing Excel table
                        formula = f"={agg_upper}({self.table_name}[{metric_col}])"
                        self.formula_count += 1
                    else:
                        value = "N/A"
                except Exception:
                    value = "N/A"

                kpi_data.append({
                    "label": kpi.label or kpi.name,
                    "value": value,
                    "formula": formula,
                    "format": kpi.format,
                })
            dashboard.write_kpis(kpi_data)

        # Summary Breakdown Tables on Dashboard
        if self.plan.analyses and self.analysis_results:
            for analysis in self.plan.analyses[:3]:
                result_df = self.analysis_results.get(analysis.name)
                if result_df is None or result_df.empty:
                    continue
                dashboard.write_section_header(f"Performance Breakdown: {analysis.name}")
                tbl_header_row = dashboard.current_row
                for col_idx, col_name in enumerate(result_df.columns, 1):
                    ws.cell(row=tbl_header_row, column=col_idx, value=col_name)
                _format_header_row(ws, tbl_header_row, len(result_df.columns))
                dashboard.current_row += 1

                tbl_start = dashboard.current_row
                is_curr = any(kw in analysis.metric.lower() for kw in ("revenue", "cost", "sales", "amount", "profit", "price"))
                for _, r in result_df.iterrows():
                    r_idx = dashboard.current_row
                    cat_val = r[result_df.columns[0]]
                    ws.cell(row=r_idx, column=1, value=cat_val)

                    # Dynamic formula referencing DataTable
                    dim_col = analysis.group_by[0] if analysis.group_by else self.df.columns[0]
                    metric_col = analysis.metric
                    agg = analysis.aggregation.upper()
                    if agg == "AVERAGE":
                        formula = f'=AVERAGEIFS({self.table_name}[{metric_col}], {self.table_name}[{dim_col}], "{cat_val}")'
                    elif agg in ("COUNT", "COUNTA"):
                        formula = f'=COUNTIFS({self.table_name}[{dim_col}], "{cat_val}")'
                    else:
                        formula = f'=SUMIFS({self.table_name}[{metric_col}], {self.table_name}[{dim_col}], "{cat_val}")'
                    metric_cell = ws.cell(row=r_idx, column=2, value=formula)
                    metric_cell.number_format = "$#,##0" if is_curr else "#,##0"
                    self.formula_count += 1
                    dashboard.current_row += 1

                # Grand Total row
                tot_row = dashboard.current_row
                ws.cell(row=tot_row, column=1, value="Grand Total").font = Font(name="Calibri", size=11, bold=True)
                tot_formula = f"=SUM(B{tbl_start}:B{tot_row-1})" if agg != "AVERAGE" else f"=AVERAGE(B{tbl_start}:B{tot_row-1})"
                tot_cell = ws.cell(row=tot_row, column=2, value=tot_formula)
                tot_cell.font = Font(name="Calibri", size=11, bold=True)
                tot_cell.number_format = "$#,##0" if is_curr else "#,##0"
                self.formula_count += 1
                dashboard.current_row += 2

        # Charts on dashboard
        if dash_spec.chart_refs:
            chart_specs = [c for c in self.plan.charts if c.name in dash_spec.chart_refs]
            if not chart_specs:
                chart_specs = self.plan.charts[:4]  # Fallback: first 4 charts

            charts_per_row = 2 if len(chart_specs) > 1 else 1

            for i, chart_spec in enumerate(chart_specs):
                result_df = self.analysis_results.get(chart_spec.data_source)
                if result_df is None:
                    continue

                position = dashboard.get_chart_position(i, charts_per_row)

                # Write data for this chart in a hidden area
                data_col = 16 + i * 3  # Far right columns
                data_row = 3

                for col_idx, col_name in enumerate(result_df.columns):
                    ws.cell(row=data_row, column=data_col + col_idx, value=col_name)
                for r_idx, (_, row) in enumerate(result_df.iterrows(), 1):
                    for col_idx, col_name in enumerate(result_df.columns):
                        val = row[col_name]
                        if isinstance(val, (np.integer,)):
                            val = int(val)
                        elif isinstance(val, (np.floating,)):
                            val = float(val)
                        ws.cell(row=data_row + r_idx, column=data_col + col_idx, value=val)

                data_end = data_row + len(result_df)

                try:
                    # Adjust chart size for dashboard
                    dash_chart_spec = ChartSpec(
                        name=chart_spec.name,
                        chart_type=chart_spec.chart_type,
                        title=chart_spec.title,
                        data_source=chart_spec.data_source,
                        category_field=chart_spec.category_field,
                        value_field=chart_spec.value_field,
                        x_axis_title=chart_spec.x_axis_title,
                        y_axis_title=chart_spec.y_axis_title,
                        width=14 if charts_per_row == 1 else 10,
                        height=8,
                    )
                    ChartEngine.create_chart(
                        spec=dash_chart_spec,
                        ws=ws,
                        data_start_row=data_row,
                        data_end_row=data_end,
                        category_col=data_col,
                        value_col=data_col + 1,
                        position=position,
                    )
                    self.chart_count += 1
                except Exception:
                    pass

            dashboard.advance_past_charts(len(chart_specs), charts_per_row)

        # Filters note
        if self.plan.filters:
            dashboard.write_section_header("Applied Filters")
            for f in self.plan.filters:
                row = dashboard.current_row
                ws.cell(row=row, column=1, value=f"✓ {f.field}")
                ws.cell(row=row, column=1).font = Font(name="Calibri", size=10, color="2CA58D")
                ws.cell(row=row, column=3, value=f"Type: {f.filter_type} (AutoFilter on Raw Data sheet)")
                ws.cell(row=row, column=3).font = Font(name="Calibri", size=10, color="6B7B8D")
                dashboard.current_row += 1

    # ─── Assignment Sheet ────────────────────────────────────────────

    def _create_assignment_sheet(self, sheet_name: str):
        """Create a sheet with the original assignment question."""
        ws = self.wb.create_sheet(title=sheet_name)

        ws.merge_cells("A1:H1")
        ws["A1"].value = "ASSIGNMENT"
        ws["A1"].font = TITLE_FONT
        ws.row_dimensions[1].height = 30

        ws.merge_cells("A2:H2")
        ws["A2"].value = ""
        ws.row_dimensions[2].height = 10

        # Write question text, splitting by lines
        lines = self.question_text.split("\n")
        for i, line in enumerate(lines):
            row = i + 3
            ws.merge_cells(f"A{row}:H{row}")
            ws.cell(row=row, column=1, value=line)
            ws.cell(row=row, column=1).font = Font(name="Calibri", size=11)
            ws.cell(row=row, column=1).alignment = Alignment(wrap_text=True)

        ws.column_dimensions["A"].width = 80


def generate_workbook(
    df: pd.DataFrame,
    plan: ExcelPlan,
    question_text: str,
    output_path: str,
) -> dict:
    """
    Top-level function to generate a workbook.
    Returns summary dict.
    """
    generator = WorkbookGenerator(df, plan, question_text)
    return generator.generate(output_path)
