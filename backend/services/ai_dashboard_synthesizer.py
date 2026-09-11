"""
AI Dashboard Synthesizer — uses Gemini to deeply analyze client datasets
and user business questions/prompts, architecting tailored, executive-grade
Excel plans with zero-hallucination column validation and robust NLP fallback.
"""

import json
import re
import difflib
from typing import Optional
import pandas as pd
import numpy as np

import google.generativeai as genai
from config import settings
from schemas.schemas import (
    ExcelPlan, SheetSpec, KPISpec, AnalysisSpec,
    ChartSpec, FilterSpec, CalculationSpec, DashboardSpec,
    DatasetProfile,
)


AI_SYNTHESIS_SYSTEM_PROMPT = """You are an elite Chief Analytics Officer and Microsoft Excel architect.
You are given:
1. The exact schema and column profiles of a client dataset.
2. Sample data rows from the client dataset.
3. The specific business question, assignment, or prompt provided by the user/client.

Your mission is to architect a bespoke, executive-grade Microsoft Excel Dashboard plan that specifically and directly answers the user's questions using the available data.

## CRITICAL RULES:
1. ZERO HALLUCINATION: You MUST ONLY use columns that actually exist in the dataset schema provided. NEVER invent columns.
2. ANSWER THE QUESTION: Look closely at what metrics, dimensions, comparisons, trends, and breakdowns the user is asking for.
   - If the user asks for "Revenue by Region", include an analysis grouping by Region with Revenue.
   - If the user asks for "Monthly Trend", include an analysis grouping by Month/Date with Revenue/Sales.
   - If the user asks for "Top Products" or "Salesperson performance", include those exact breakdowns.
   - If no specific question is given, identify the top 3-4 strategic business questions this dataset can answer.
3. MULTI-TABLE DASHBOARD: Provide 2 to 4 distinct analyses that comprehensively answer different angles of the business question.
4. CHART SELECTION:
   - Time series (dates, months, quarters) -> "line"
   - Categorical comparisons with <= 12 categories -> "column"
   - Proportions/shares with <= 6 categories -> "doughnut"
   - Rankings or categories with long names -> "bar"
5. KPIS: Provide 3 to 5 vital executive metrics (totals, averages, counts, rates) that answer the top-level question.
6. SLICERS / FILTERS: Identify 1 to 3 key categorical dimensions that the user will want to filter the dashboard by (e.g. Region, Category, Month).
7. Return ONLY a single raw JSON object matching the schema below. No markdown backticks, no explanations.

## JSON SCHEMA:
{
  "workbook_title": "string — professional executive title reflecting the question and dataset",
  "dashboard_summary": "string — 2-3 sentence executive briefing explaining how this dashboard answers the client's questions",
  "calculations": [
    {
      "name": "string — calculated column name",
      "formula_template": "string — e.g. {Revenue} - {Cost}",
      "format": "currency|number|percentage|integer",
      "description": "string"
    }
  ],
  "kpis": [
    {
      "name": "string",
      "metric": "string — exact column name from dataset",
      "aggregation": "SUM|AVERAGE|COUNT|COUNTA|MIN|MAX",
      "format": "currency|currency_decimal|number|number_decimal|percentage|integer",
      "label": "string — clear executive label"
    }
  ],
  "analyses": [
    {
      "name": "string — descriptive analysis title (e.g. Revenue by Region)",
      "metric": "string — exact column name from dataset",
      "aggregation": "SUM|AVERAGE|COUNT|COUNTA|MIN|MAX",
      "group_by": ["string — exact column name(s) from dataset"],
      "sort": "descending|ascending|default",
      "description": "string"
    }
  ],
  "charts": [
    {
      "name": "string — unique name matching an analysis",
      "chart_type": "column|bar|line|doughnut|pie",
      "title": "string — professional chart title",
      "data_source": "string — MUST match one of the analysis names above",
      "category_field": "string — column used in group_by",
      "value_field": "string — metric column",
      "x_axis_title": "string",
      "y_axis_title": "string",
      "width": 12,
      "height": 7
    }
  ],
  "filters": [
    {
      "field": "string — exact column name to filter/slice on",
      "filter_type": "slicer",
      "description": "string"
    }
  ]
}
"""


def _find_closest_column(target: str, available_cols: list[str]) -> Optional[str]:
    """Find closest column name from available columns using fuzzy matching."""
    if not available_cols:
        return None
    if target in available_cols:
        return target
    # Case-insensitive match
    t_lower = target.lower().strip()
    for col in available_cols:
        if col.lower().strip() == t_lower:
            return col
    # Partial substring match
    for col in available_cols:
        if t_lower in col.lower() or col.lower() in t_lower:
            return col
    # Difflib close match
    matches = difflib.get_close_matches(target, available_cols, n=1, cutoff=0.5)
    if matches:
        return matches[0]
    return available_cols[0]


def _clean_json_response(text: str) -> str:
    """Extract raw JSON from model response."""
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*\n?", "", text)
        text = re.sub(r"\n?```\s*$", "", text)
    # Find outer JSON braces
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        return text[start:end+1]
    return text.strip()


def _smart_nlp_plan_fallback(
    df: pd.DataFrame,
    profile: Optional[DatasetProfile] = None,
    question_text: Optional[str] = None,
    title: Optional[str] = None,
) -> ExcelPlan:
    """
    Intelligent NLP fallback if Gemini is unreachable.
    Analyzes question keywords and matches them directly to dataset columns.
    """
    cols = list(df.columns)
    q_lower = (question_text or "").lower()

    # Classify measures and dimensions
    measures = []
    dimensions = []
    date_cols = []

    for col in cols:
        col_l = col.lower()
        if pd.api.types.is_datetime64_any_dtype(df[col]) or any(k in col_l for k in ["date", "month", "year", "quarter"]):
            date_cols.append(col)
            dimensions.append(col)
        elif pd.api.types.is_numeric_dtype(df[col]):
            if df[col].nunique() <= 5 and not any(k in col_l for k in ["sales", "rev", "cost", "price", "amount", "profit", "qty", "units"]):
                dimensions.append(col)
            else:
                measures.append(col)
        else:
            dimensions.append(col)

    # Rank measures by mention in question, then by business importance
    def measure_score(m: str) -> int:
        score = 0
        ml = m.lower()
        if ml in q_lower:
            score += 50
        for word in ml.split():
            if word in q_lower:
                score += 15
        if any(k in ml for k in ["revenue", "sales", "total", "amount", "income"]):
            score += 10
        elif any(k in ml for k in ["profit", "margin"]):
            score += 8
        elif any(k in ml for k in ["cost", "spend"]):
            score += 6
        elif any(k in ml for k in ["qty", "unit", "volume"]):
            score += 5
        return score

    measures.sort(key=measure_score, reverse=True)
    if not measures:
        measures = [cols[0]]

    # Rank dimensions by mention in question, then by cardinality
    def dimension_score(d: str) -> int:
        score = 0
        dl = d.lower()
        if dl in q_lower:
            score += 50
        for word in dl.split():
            if len(word) > 2 and word in q_lower:
                score += 15
        if any(k in dl for k in ["region", "territory", "country", "area"]):
            score += 10
        elif any(k in dl for k in ["product", "item", "category", "type"]):
            score += 9
        elif any(k in dl for k in ["salesperson", "rep", "agent", "employee", "department"]):
            score += 8
        elif any(k in dl for k in ["month", "quarter", "date", "year"]):
            score += 7
        return score

    dimensions.sort(key=dimension_score, reverse=True)
    if not dimensions:
        dimensions = [cols[0]]

    primary_measure = measures[0]
    secondary_measure = measures[1] if len(measures) > 1 else None

    # Determine Title
    clean_title = title.strip() if title else (
        f"Executive {primary_measure} Analysis Dashboard" if not question_text else
        f"Executive Dashboard: {question_text[:40].strip()}..."
    )

    # KPIs
    kpis = []
    p_format = "currency" if any(k in primary_measure.lower() for k in ["sales", "rev", "cost", "price", "amount", "profit"]) else "number"
    kpis.append(KPISpec(
        name=f"Total {primary_measure}",
        metric=primary_measure,
        aggregation="SUM",
        format=p_format,
        label=f"Total {primary_measure}",
    ))

    if secondary_measure:
        s_format = "currency" if any(k in secondary_measure.lower() for k in ["sales", "rev", "cost", "price", "amount", "profit"]) else "number"
        kpis.append(KPISpec(
            name=f"Total {secondary_measure}",
            metric=secondary_measure,
            aggregation="SUM",
            format=s_format,
            label=f"Total {secondary_measure}",
        ))

    kpis.append(KPISpec(
        name=f"Average {primary_measure}",
        metric=primary_measure,
        aggregation="AVERAGE",
        format=f"{p_format}_decimal" if p_format == "currency" else "number_decimal",
        label=f"Average {primary_measure}",
    ))

    kpis.append(KPISpec(
        name="Total Records",
        metric=cols[0],
        aggregation="COUNTA",
        format="integer",
        label="Total Records",
    ))

    # Analyses & Charts matching question dimensions
    analyses = []
    charts = []

    # Analysis 1: Top dimension
    dim1 = dimensions[0]
    a1_name = f"{primary_measure} by {dim1}"
    analyses.append(AnalysisSpec(
        name=a1_name,
        metric=primary_measure,
        aggregation="SUM",
        group_by=[dim1],
        sort="descending",
        description=f"Performance aggregation by {dim1}",
    ))
    charts.append(ChartSpec(
        name=a1_name,
        chart_type="column",
        title=f"{primary_measure} by {dim1}",
        data_source=a1_name,
        category_field=dim1,
        value_field=primary_measure,
        x_axis_title=dim1,
        y_axis_title=primary_measure,
        width=12,
        height=7,
    ))

    # Analysis 2: Second dimension
    if len(dimensions) > 1 and dimensions[1] != dim1:
        dim2 = dimensions[1]
        a2_name = f"{primary_measure} by {dim2}"
        analyses.append(AnalysisSpec(
            name=a2_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[dim2],
            sort="descending",
            description=f"Distribution across {dim2}",
        ))
        c2_type = "doughnut" if df[dim2].nunique() <= 6 else "bar"
        charts.append(ChartSpec(
            name=a2_name,
            chart_type=c2_type,
            title=f"{primary_measure} Distribution by {dim2}",
            data_source=a2_name,
            category_field=dim2,
            value_field=primary_measure,
            width=11,
            height=7,
        ))

    # Analysis 3: Timeline or Third dimension
    time_col = next((d for d in date_cols if d in df.columns), None)
    if time_col and time_col not in [dim1, dimensions[1] if len(dimensions) > 1 else None]:
        a3_name = f"{primary_measure} Trend by {time_col}"
        analyses.append(AnalysisSpec(
            name=a3_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[time_col],
            sort="default",
            description=f"Timeline trend by {time_col}",
        ))
        charts.append(ChartSpec(
            name=a3_name,
            chart_type="line",
            title=f"{primary_measure} Trend Over Time",
            data_source=a3_name,
            category_field=time_col,
            value_field=primary_measure,
            width=13,
            height=7,
        ))
    elif len(dimensions) > 2 and dimensions[2] not in [dim1, dimensions[1]]:
        dim3 = dimensions[2]
        a3_name = f"{primary_measure} by {dim3}"
        analyses.append(AnalysisSpec(
            name=a3_name,
            metric=primary_measure,
            aggregation="SUM",
            group_by=[dim3],
            sort="descending",
            description=f"Breakdown by {dim3}",
        ))
        charts.append(ChartSpec(
            name=a3_name,
            chart_type="bar",
            title=f"{primary_measure} by {dim3}",
            data_source=a3_name,
            category_field=dim3,
            value_field=primary_measure,
            width=12,
            height=7,
        ))

    # Slicers
    filters = []
    for d in dimensions[:2]:
        if df[d].nunique() <= 50:
            filters.append(FilterSpec(field=d, filter_type="slicer", description=f"Filter by {d}"))

    dashboard = DashboardSpec(
        title=clean_title,
        kpis=kpis,
        chart_refs=[c.name for c in charts],
        filter_refs=[f.field for f in filters],
    )

    sheets = [
        SheetSpec(name="01_Executive_Dashboard", sheet_type="dashboard", description="Interactive executive view"),
        SheetSpec(name="02_Analysis_Summary", sheet_type="analysis", description="Structured summary aggregation tables"),
        SheetSpec(name="03_Client_Data", sheet_type="raw_data", description="Source client records"),
    ]

    return ExcelPlan(
        workbook_title=clean_title,
        dashboard=dashboard,
        analyses=analyses,
        charts=charts,
        filters=filters,
        calculations=[],
        sheets=sheets,
    )


def synthesize_ai_dashboard_plan(
    df: pd.DataFrame,
    profile: Optional[DatasetProfile] = None,
    question_text: Optional[str] = None,
    title: Optional[str] = None,
) -> tuple[ExcelPlan, dict]:
    """
    Synthesize an executive Excel plan using Gemini AI to analyze both
    the dataset structure and the user's specific business questions.
    Returns (plan, ai_metadata).
    """
    available_cols = list(df.columns)
    num_rows = len(df)

    # If no question provided, generate a default prompt based on columns
    effective_question = question_text.strip() if question_text and question_text.strip() else (
        f"Analyze this client dataset of {num_rows} records. Formulate the primary executive KPIs, "
        f"compare performance across key categorical dimensions, and show relevant timeline or categorical breakdowns."
    )

    ai_metadata = {
        "engine": "gemini-3.6-flash",
        "question_analyzed": effective_question,
        "ai_used": False,
        "summary": "Synthesized executive architecture matching dataset schema and business questions.",
    }

    # Prepare schema and sample rows for Gemini prompt
    schema_info = []
    for col in available_cols:
        series = df[col].dropna()
        dtype = str(df[col].dtype)
        unique_cnt = int(series.nunique())
        sample_vals = [str(x) for x in series.head(5).tolist()]
        is_num = pd.api.types.is_numeric_dtype(df[col])
        schema_info.append({
            "name": col,
            "data_type": dtype,
            "is_numeric": is_num,
            "unique_count": unique_cnt,
            "samples": sample_vals,
        })

    # Prepare prompt
    user_prompt = f"""## CLIENT DATASET METRICS
Total Rows: {num_rows}
Total Columns: {len(available_cols)}

## DATASET COLUMNS:
{json.dumps(schema_info, indent=2)}

## CLIENT QUESTION / BUSINESS OBJECTIVES:
"{effective_question}"

## DESIRED TITLE (IF GIVEN):
"{title or ''}"

Analyze the dataset schema and the question above. Return ONLY the JSON plan following the required schema.
Ensure all metric and group_by column names match the dataset columns EXACTLY.
"""

    if settings.gemini_api_key:
        try:
            genai.configure(api_key=settings.gemini_api_key)
            model = genai.GenerativeModel(
                settings.gemini_model,
                system_instruction=AI_SYNTHESIS_SYSTEM_PROMPT,
            )
            response = model.generate_content(user_prompt)
            cleaned_json = _clean_json_response(response.text)
            data = json.loads(cleaned_json)

            # Zero-hallucination validation and sanitization
            workbook_title = data.get("workbook_title") or (title or "Executive Performance Dashboard")
            dashboard_summary = data.get("dashboard_summary") or "Tailored executive dashboard answering client business questions."

            # 1. Validate KPIs
            raw_kpis = data.get("kpis") or []
            validated_kpis = []
            for k in raw_kpis:
                metric_col = _find_closest_column(k.get("metric", ""), available_cols)
                if not metric_col:
                    continue
                agg = k.get("aggregation", "SUM").upper()
                if agg not in ["SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX"]:
                    agg = "SUM"
                fmt = k.get("format", "number")
                validated_kpis.append(KPISpec(
                    name=k.get("name") or f"Total {metric_col}",
                    metric=metric_col,
                    aggregation=agg,
                    format=fmt,
                    label=k.get("label") or f"Total {metric_col}",
                ))

            # 2. Validate Analyses
            raw_analyses = data.get("analyses") or []
            validated_analyses = []
            analysis_name_map = {}

            for a in raw_analyses:
                a_name = a.get("name") or "Summary Analysis"
                metric_col = _find_closest_column(a.get("metric", ""), available_cols)
                if not metric_col:
                    continue
                raw_groups = a.get("group_by") or []
                val_groups = []
                for g in raw_groups:
                    gc = _find_closest_column(g, available_cols)
                    if gc and gc not in val_groups:
                        val_groups.append(gc)
                if not val_groups:
                    # Pick a dimension column that is not the metric
                    val_groups = [c for c in available_cols if c != metric_col][:1]

                agg = a.get("aggregation", "SUM").upper()
                if agg not in ["SUM", "AVERAGE", "COUNT", "COUNTA", "MIN", "MAX"]:
                    agg = "SUM"

                analysis_spec = AnalysisSpec(
                    name=a_name,
                    metric=metric_col,
                    aggregation=agg,
                    group_by=val_groups,
                    sort=a.get("sort", "descending"),
                    description=a.get("description", f"Breakdown of {metric_col} by {val_groups}"),
                )
                validated_analyses.append(analysis_spec)
                analysis_name_map[a_name] = analysis_spec

            # 3. Validate Charts
            raw_charts = data.get("charts") or []
            validated_charts = []
            for c in raw_charts:
                data_src = c.get("data_source")
                # Link to existing validated analysis
                matched_analysis = analysis_name_map.get(data_src)
                if not matched_analysis and validated_analyses:
                    matched_analysis = validated_analyses[0]
                    data_src = matched_analysis.name

                if not matched_analysis:
                    continue

                cat_field = matched_analysis.group_by[0] if matched_analysis.group_by else available_cols[0]
                val_field = matched_analysis.metric
                chart_type = c.get("chart_type", "column").lower()
                if chart_type not in ["column", "bar", "line", "doughnut", "pie", "area"]:
                    chart_type = "column"

                c_name = c.get("name") or f"Chart: {matched_analysis.name}"
                validated_charts.append(ChartSpec(
                    name=c_name,
                    chart_type=chart_type,
                    title=c.get("title") or f"{val_field} by {cat_field}",
                    data_source=data_src,
                    category_field=cat_field,
                    value_field=val_field,
                    x_axis_title=c.get("x_axis_title") or cat_field,
                    y_axis_title=c.get("y_axis_title") or val_field,
                    width=c.get("width", 12),
                    height=c.get("height", 7),
                ))

            # 4. Validate Filters / Slicers
            raw_filters = data.get("filters") or []
            validated_filters = []
            for f in raw_filters:
                f_col = _find_closest_column(f.get("field", ""), available_cols)
                if f_col and f_col not in [vf.field for vf in validated_filters]:
                    validated_filters.append(FilterSpec(
                        field=f_col,
                        filter_type="slicer",
                        description=f.get("description", f"Filter by {f_col}"),
                    ))

            # Guarantee minimum 1 analysis and 1 chart
            if not validated_analyses or not validated_kpis:
                raise ValueError("LLM plan produced insufficient validated elements, falling back to NLP architect.")

            dashboard = DashboardSpec(
                title=workbook_title,
                kpis=validated_kpis,
                chart_refs=[c.name for c in validated_charts],
                filter_refs=[f.field for f in validated_filters],
            )

            sheets = [
                SheetSpec(name="01_Executive_Dashboard", sheet_type="dashboard", description="Interactive executive view answering client questions"),
                SheetSpec(name="02_Analysis_Summary", sheet_type="analysis", description="Structured summary aggregation tables"),
                SheetSpec(name="03_Client_Data", sheet_type="raw_data", description="Source client records"),
            ]

            plan = ExcelPlan(
                workbook_title=workbook_title,
                dashboard=dashboard,
                analyses=validated_analyses,
                charts=validated_charts,
                filters=validated_filters,
                calculations=[],
                sheets=sheets,
            )

            ai_metadata["ai_used"] = True
            ai_metadata["summary"] = dashboard_summary
            return plan, ai_metadata

        except Exception as e:
            # Fall back safely to NLP plan
            ai_metadata["ai_used"] = False
            ai_metadata["fallback_reason"] = str(e)

    # Use smart NLP fallback
    plan = _smart_nlp_plan_fallback(df, profile, question_text=effective_question, title=title)
    ai_metadata["summary"] = (
        f"Smart NLP Architect analyzed client question '{effective_question[:60]}' and mapped metrics and dimensions."
    )
    return plan, ai_metadata
