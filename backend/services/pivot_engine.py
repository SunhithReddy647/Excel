"""
Native Pivot Table Engine — creates 100% authentic, interactive Microsoft Excel
PivotTables, Slicers, and PivotCharts via Excel COM automation (pywin32) on Windows.

If Excel COM is unavailable, gracefully logs and preserves the OpenPyXL
formula summary tables as a clean fallback.
"""

import os
import sys
import logging
from pathlib import Path
from typing import Optional
import pandas as pd

from schemas.schemas import ExcelPlan, AnalysisSpec, ChartSpec, FilterSpec

logger = logging.getLogger("excelflow.pivot_engine")

# Excel COM Constants
xlDatabase = 1
xlRowField = 1
xlColumnField = 2
xlPageField = 3
xlDataField = 4

xlSum = -4157
xlAverage = -4106
xlCount = -4112
xlMin = -4139
xlMax = -4136

xlTabularRow = 1
xlRepeatLabels = 2

# Chart type constants
CHART_TYPE_MAP = {
    "column": 51,    # xlColumnClustered
    "bar": 57,       # xlBarClustered
    "line": 4,       # xlLine
    "pie": 5,        # xlPie
    "doughnut": -4120, # xlDoughnut
    "area": 1,       # xlArea
}


def is_excel_com_available() -> bool:
    """Check if Microsoft Excel COM automation is available on this system."""
    if sys.platform != "win32":
        return False
    try:
        import pythoncom
        pythoncom.CoInitialize()
        import win32com.client
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.Quit()
        return True
    except Exception as e:
        logger.warning(f"Excel COM automation not available: {e}")
        return False
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


class NativePivotEngine:
    """
    Builds true native interactive Microsoft Excel PivotTables, Slicers,
    and PivotCharts directly inside the generated .xlsx workbook.
    """

    def __init__(self, workbook_path: str, plan: ExcelPlan, df: pd.DataFrame):
        self.workbook_path = os.path.abspath(workbook_path)
        self.plan = plan
        self.df = df
        self.pivots_created = 0
        self.slicers_created = 0
        self.charts_created = 0

    def enhance(self) -> dict:
        """
        Enhance the workbook with native PivotTables, Slicers, and PivotCharts.
        Returns a summary of objects created.
        """
        if not is_excel_com_available():
            logger.info("Excel COM not available; keeping OpenPyXL formula summary tables.")
            return {
                "native_pivots_enabled": False,
                "reason": "Excel COM automation not available in current environment",
                "pivots_created": 0,
                "slicers_created": 0,
                "charts_created": 0,
            }

        import win32com.client
        import pythoncom

        pythoncom.CoInitialize()
        excel = None
        wb = None

        try:
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False

            wb = excel.Workbooks.Open(self.workbook_path)

            # 1. Locate Raw Data sheet and DataTable
            raw_sheet = self._find_raw_data_sheet(wb)
            if not raw_sheet:
                logger.warning("Raw data worksheet not found for native PivotTable creation.")
                return {"native_pivots_enabled": False, "reason": "Raw data sheet not found"}

            raw_table = None
            if raw_sheet.ListObjects.Count > 0:
                raw_table = raw_sheet.ListObjects(1)

            data_source_range = raw_table.Range if raw_table else raw_sheet.UsedRange

            # 2. Locate or create the Analysis/PivotTables sheet
            pivot_sheet = self._get_or_create_pivot_sheet(wb)

            # 3. Clear any existing static tables on the pivot sheet so real pivot tables take center stage
            pivot_sheet.Cells.Clear()

            # 4. Format Executive Header for the PivotTables sheet
            self._write_sheet_title(pivot_sheet, "Executive Pivot Table & Multi-Dimensional Analysis")

            # 5. Build native Pivot Tables for each analysis in the plan
            current_row = 4
            slicer_left = 650  # Place interactive slicers to the right of the pivot tables

            for idx, analysis in enumerate(self.plan.analyses):
                try:
                    pt_info = self._create_single_pivot_table(
                        wb=wb,
                        ws=pivot_sheet,
                        source_range=data_source_range,
                        analysis=analysis,
                        start_row=current_row,
                        index=idx + 1,
                    )
                    if pt_info:
                        pt = pt_info["pivot_table"]
                        pt_end_row = pt_info["end_row"]
                        self.pivots_created += 1

                        # Add Slicer for the primary group column if applicable
                        if analysis.group_by and self.slicers_created < 3:
                            slicer_col = analysis.group_by[0]
                            slicer_top = current_row * 20
                            if self._create_slicer(wb, pivot_sheet, pt, slicer_col, slicer_top, slicer_left):
                                self.slicers_created += 1
                                slicer_left += 165

                        # Add linked PivotChart
                        matching_charts = [
                            c for c in self.plan.charts
                            if c.data_source.lower() == analysis.name.lower() or c.name.lower() == analysis.name.lower()
                        ]
                        if matching_charts:
                            chart_spec = matching_charts[0]
                            chart_top = (pt_end_row + 2) * 20
                            self._create_pivot_chart(pivot_sheet, pt, chart_spec, chart_top=chart_top, chart_left=30)
                            current_row = pt_end_row + 18
                        else:
                            current_row = pt_end_row + 4

                except Exception as ex:
                    logger.error(f"Error creating native PivotTable for {analysis.name}: {ex}", exc_info=True)
                    current_row += 6

            # Auto-fit columns for professional appearance
            try:
                pivot_sheet.Columns("A:Z").AutoFit()
            except Exception:
                pass

            # Refresh all pivot caches
            wb.RefreshAll()

            # Save enhanced workbook
            wb.Save()

            logger.info(
                f"Successfully enhanced workbook with {self.pivots_created} native PivotTables, "
                f"{self.slicers_created} Slicers, and {self.charts_created} PivotCharts."
            )

            return {
                "native_pivots_enabled": True,
                "pivots_created": self.pivots_created,
                "slicers_created": self.slicers_created,
                "charts_created": self.charts_created,
            }

        except Exception as e:
            logger.error(f"Failed to enhance workbook with native PivotTables: {e}", exc_info=True)
            return {
                "native_pivots_enabled": False,
                "error": str(e),
                "pivots_created": self.pivots_created,
            }
        finally:
            if wb:
                try:
                    wb.Close(False)
                except Exception:
                    pass
            if excel:
                try:
                    excel.Quit()
                except Exception:
                    pass
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _find_raw_data_sheet(self, wb):
        """Find the raw data sheet by name or convention."""
        for sheet in wb.Worksheets:
            name_lower = sheet.Name.lower()
            if "raw" in name_lower or "data" in name_lower:
                return sheet
        return wb.Worksheets(1) if wb.Worksheets.Count > 0 else None

    def _get_or_create_pivot_sheet(self, wb):
        """Find existing analysis sheet or create a dedicated PivotTables sheet."""
        for sheet in wb.Worksheets:
            name_lower = sheet.Name.lower()
            if "analysis" in name_lower or "pivot" in name_lower:
                return sheet
        # If not found, add one after raw data
        return wb.Worksheets.Add(After=wb.Worksheets(1))

    def _write_sheet_title(self, ws, title: str):
        """Format an executive title bar at the top of the worksheet."""
        try:
            ws.Range("B1:J1").Merge()
            title_cell = ws.Range("B1")
            title_cell.Value = title
            title_cell.Font.Name = "Calibri"
            title_cell.Font.Size = 16
            title_cell.Font.Bold = True
            title_cell.Font.Color = 0x4A2A1B  # Dark Navy
            ws.Rows(1).RowHeight = 35
        except Exception:
            pass

    def _create_single_pivot_table(
        self,
        wb,
        ws,
        source_range,
        analysis: AnalysisSpec,
        start_row: int,
        index: int,
    ) -> Optional[dict]:
        """Create and format an authentic Microsoft Excel PivotTable."""
        # 1. Create Pivot Cache
        pivot_cache = wb.PivotCaches().Create(SourceType=xlDatabase, SourceData=source_range)

        # 2. Table Destination
        dest_cell = ws.Range(f"B{start_row}")
        table_name = f"PivotTable_{index}_{analysis.metric[:10]}".replace(" ", "_")
        pivot_table = pivot_cache.CreatePivotTable(TableDestination=dest_cell, TableName=table_name)

        # Section Header above Pivot Table
        header_cell = ws.Range(f"B{start_row - 1}")
        header_cell.Value = f"{analysis.name} — Interactive Pivot Table"
        header_cell.Font.Name = "Calibri"
        header_cell.Font.Size = 12
        header_cell.Font.Bold = True
        header_cell.Font.Color = 0xAB862E  # Teal Accent

        # 3. Configure Row Fields
        for pos, group_field in enumerate(analysis.group_by[:2], 1):
            try:
                rf = pivot_table.PivotFields(group_field)
                if pos == 1:
                    rf.Orientation = xlRowField
                    rf.Position = 1
                elif pos == 2:
                    # If 2nd dimension is suitable for column pivot (e.g. Month, Region, Product), make it column field
                    rf.Orientation = xlColumnField
                    rf.Position = 1
            except Exception as e:
                logger.warning(f"Could not add field '{group_field}' to pivot table: {e}")

        # 4. Configure Data Field (Values)
        metric = analysis.metric
        agg_upper = analysis.aggregation.upper()
        agg_code = xlSum
        agg_label = "Sum"

        if agg_upper == "AVERAGE":
            agg_code = xlAverage
            agg_label = "Average"
        elif agg_upper in ("COUNT", "COUNTA"):
            agg_code = xlCount
            agg_label = "Count"
        elif agg_upper == "MIN":
            agg_code = xlMin
            agg_label = "Min"
        elif agg_upper == "MAX":
            agg_code = xlMax
            agg_label = "Max"

        try:
            data_field = pivot_table.AddDataField(
                pivot_table.PivotFields(metric),
                f"{agg_label} of {metric}",
                agg_code,
            )

            # Apply professional number format to data field
            metric_lower = metric.lower()
            if any(k in metric_lower for k in ("revenue", "profit", "sales", "price", "cost", "salary", "expense", "amount")):
                data_field.NumberFormat = "₹#,##0"
            elif any(k in metric_lower for k in ("margin", "pct", "percent", "rate", "growth", "ratio")):
                data_field.NumberFormat = "0.00%"
            else:
                data_field.NumberFormat = "#,##0"
        except Exception as e:
            logger.warning(f"Could not add data field '{metric}' to pivot table: {e}")

        # 5. Apply Human-Crafted Styling & Layout
        try:
            pivot_table.TableStyle2 = "PivotStyleMedium9"
            pivot_table.ShowTableStyleRowStripes = True
            pivot_table.ShowTableStyleColumnStripes = False
            pivot_table.RowAxisLayout(xlTabularRow)  # Clean tabular layout
            pivot_table.RowGrand = True
            pivot_table.ColumnGrand = True
        except Exception:
            pass

        try:
            pivot_table.RepeatAllLabels(xlRepeatLabels)
        except Exception:
            pass

        # Estimate end row based on data range
        end_row = start_row + 10
        try:
            tbl_range = pivot_table.TableRange2
            end_row = tbl_range.Row + tbl_range.Rows.Count
        except Exception:
            pass

        return {
            "pivot_table": pivot_table,
            "start_row": start_row,
            "end_row": end_row,
        }

    def _create_slicer(self, wb, ws, pivot_table, field_name: str, top: int, left: int) -> bool:
        """Create an authentic Excel interactive Slicer connected to the Pivot Table."""
        try:
            slicer_cache = wb.SlicerCaches.Add2(pivot_table, field_name)
            slicer = slicer_cache.Slicers.Add(
                SlicerDestination=ws,
                Name=f"Slicer_{field_name}_{self.slicers_created + 1}",
                Caption=f"Filter: {field_name}",
                Top=top,
                Left=left,
                Width=150,
                Height=170,
            )
            slicer.Style = "SlicerStyleLight2"
            return True
        except Exception as e:
            logger.warning(f"Could not create slicer for '{field_name}': {e}")
            return False

    def _create_pivot_chart(self, ws, pivot_table, chart_spec: ChartSpec, chart_top: int, chart_left: int):
        """Create a native PivotChart linked directly to the PivotTable."""
        try:
            chart_type_code = CHART_TYPE_MAP.get(chart_spec.chart_type.lower(), 51)
            # Shapes.AddChart2(Style, Type, Left, Top, Width, Height)
            chart_shape = ws.Shapes.AddChart2(201, chart_type_code)
            chart = chart_shape.Chart
            chart.SetSourceData(pivot_table.TableRange2)

            chart_shape.Top = chart_top
            chart_shape.Left = chart_left
            chart_shape.Width = 480
            chart_shape.Height = 280

            chart.HasTitle = True
            chart.ChartTitle.Text = chart_spec.title
            chart.ChartStyle = 245  # Sleek modern Excel chart style
            self.charts_created += 1
        except Exception as e:
            logger.warning(f"Could not create pivot chart for '{chart_spec.title}': {e}")


def enhance_workbook_with_native_pivots(workbook_path: str, plan: ExcelPlan, df: pd.DataFrame) -> dict:
    """
    Top-level helper to enhance a workbook with real native Excel PivotTables,
    Slicers, and PivotCharts.
    """
    engine = NativePivotEngine(workbook_path, plan, df)
    return engine.enhance()
