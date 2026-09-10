"""
Pydantic schemas for API request/response models and the Excel generation plan.
"""

from __future__ import annotations
from typing import Optional
from pydantic import BaseModel, Field


# ─── Dataset Schemas ────────────────────────────────────────────────

class ColumnProfile(BaseModel):
    name: str
    data_type: str  # string, integer, decimal, date, percentage, currency, boolean
    semantic_type: str = ""  # unit_price, quantity, region, etc.
    classification: str = ""  # dimension or measure
    unique_count: int = 0
    missing_count: int = 0
    sample_values: list = Field(default_factory=list)
    min_value: Optional[str] = None
    max_value: Optional[str] = None
    mean_value: Optional[float] = None


class DatasetProfile(BaseModel):
    dataset_id: str = ""
    filename: str = ""
    row_count: int = 0
    column_count: int = 0
    headers: list[str] = Field(default_factory=list)
    columns: list[ColumnProfile] = Field(default_factory=list)
    dimensions: list[str] = Field(default_factory=list)
    measures: list[str] = Field(default_factory=list)
    potential_calculations: list[str] = Field(default_factory=list)
    duplicate_rows: int = 0


class DatasetUploadResponse(BaseModel):
    dataset_id: str
    filename: str
    row_count: int
    column_count: int
    headers: list[str]
    profile: DatasetProfile


# ─── Question Schemas ───────────────────────────────────────────────

class QuestionUploadResponse(BaseModel):
    assignment_id: str
    question_text: str
    source_type: str


class QuestionPasteRequest(BaseModel):
    dataset_id: str
    question_text: str


# ─── Plan Schemas (LLM output & generation input) ──────────────────

class CalculationSpec(BaseModel):
    """A calculated column to add to the data."""
    name: str
    formula_template: str  # e.g. "{Units Sold}*{Unit Price}"
    format: str = "number"  # number, currency, percentage, integer
    description: str = ""


class AnalysisSpec(BaseModel):
    """An aggregation/summary table to generate."""
    name: str
    metric: str  # Column name or calculated field
    aggregation: str = "SUM"  # SUM, AVERAGE, COUNT, COUNTA, MIN, MAX
    group_by: list[str] = Field(default_factory=list)
    sort: str = "default"  # default, ascending, descending, chronological
    description: str = ""


class ChartSpec(BaseModel):
    """A chart to create."""
    name: str
    chart_type: str = "column"  # column, bar, line, pie, doughnut, combo, area
    title: str
    data_source: str  # Analysis name this chart is based on
    category_field: str = ""  # X-axis field
    value_field: str = ""  # Y-axis field
    x_axis_title: str = ""
    y_axis_title: str = ""
    width: int = 12
    height: int = 7


class KPISpec(BaseModel):
    """A KPI card for the dashboard."""
    name: str
    metric: str  # Column name or calculated field
    aggregation: str = "SUM"
    format: str = "number"
    label: str = ""


class FilterSpec(BaseModel):
    """A filter/slicer request."""
    field: str
    filter_type: str = "autofilter"  # autofilter, table_filter, slicer (future)
    description: str = ""


class DashboardSpec(BaseModel):
    """Dashboard layout configuration."""
    title: str = "Dashboard"
    kpis: list[KPISpec] = Field(default_factory=list)
    chart_refs: list[str] = Field(default_factory=list)  # References to ChartSpec names
    filter_refs: list[str] = Field(default_factory=list)  # References to FilterSpec fields


class SheetSpec(BaseModel):
    """A sheet to include in the workbook."""
    name: str
    sheet_type: str  # raw_data, calculations, analysis, charts, dashboard, assignment
    description: str = ""


class ExcelPlan(BaseModel):
    """The master plan for workbook generation — produced by the LLM, validated by schema."""
    workbook_title: str = "Generated Workbook"
    calculations: list[CalculationSpec] = Field(default_factory=list)
    analyses: list[AnalysisSpec] = Field(default_factory=list)
    charts: list[ChartSpec] = Field(default_factory=list)
    dashboard: Optional[DashboardSpec] = None
    filters: list[FilterSpec] = Field(default_factory=list)
    sheets: list[SheetSpec] = Field(default_factory=list)


# ─── Generation Schemas ─────────────────────────────────────────────

class AnalyzeRequest(BaseModel):
    dataset_id: str
    assignment_id: str


class AnalyzeResponse(BaseModel):
    generation_id: str
    plan: ExcelPlan
    dataset_profile: DatasetProfile


class GenerateRequest(BaseModel):
    generation_id: str
    plan: ExcelPlan  # Possibly user-modified


class GenerationProgress(BaseModel):
    step: str
    status: str  # pending, in_progress, completed, failed
    message: str = ""


class GenerationStatusResponse(BaseModel):
    generation_id: str
    status: str
    progress: list[GenerationProgress] = Field(default_factory=list)
    error: Optional[str] = None


# ─── Validation Schemas ─────────────────────────────────────────────

class ValidationCheck(BaseModel):
    check_name: str
    passed: bool
    message: str = ""
    details: list[str] = Field(default_factory=list)


class ValidationResult(BaseModel):
    workbook_id: str
    overall_status: str  # passed, failed
    checks: list[ValidationCheck] = Field(default_factory=list)
    formula_count: int = 0
    chart_count: int = 0
    sheet_count: int = 0
    error_count: int = 0


# ─── Workbook Summary ──────────────────────────────────────────────

class WorkbookSummary(BaseModel):
    workbook_id: str
    filename: str
    sheets: list[str] = Field(default_factory=list)
    row_count: int = 0
    formula_count: int = 0
    chart_count: int = 0
    analysis_count: int = 0
    filter_count: int = 0
    validation: Optional[ValidationResult] = None
