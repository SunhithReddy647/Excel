"""
LLM prompts for ExcelFlow question parsing and plan generation.
"""

SYSTEM_PROMPT = """You are an expert Excel assignment analyzer. Given a dataset schema and a user's assignment/question, you produce a structured JSON plan for generating an Excel workbook.

## Rules
1. Return ONLY valid JSON matching the schema below. No markdown, no explanations, no code fences.
2. Every calculation must reference actual columns from the dataset.
3. Formula templates use curly braces for column references: {Column Name}
4. Choose appropriate aggregation functions (SUM, AVERAGE, COUNT, MIN, MAX).
5. Select chart types based on the data pattern:
   - Time series → line chart
   - Category comparison → column or bar chart
   - Part-of-whole → pie or doughnut chart
   - Actual vs target → combo chart
6. Only create sheets, calculations, analyses, and charts that the assignment actually requests.
7. If the assignment asks for a dashboard, include KPIs, chart references, and filter references.
8. Do NOT invent columns that don't exist in the dataset.
9. For calculated columns, write the formula template using actual column names from the dataset.

## Output JSON Schema
{
  "workbook_title": "string — title for the workbook",
  "calculations": [
    {
      "name": "string — name of the calculated column",
      "formula_template": "string — e.g. {Units Sold}*{Unit Price}",
      "format": "number|currency|percentage|integer",
      "description": "string — what this calculates"
    }
  ],
  "analyses": [
    {
      "name": "string — display name",
      "metric": "string — column or calculated field to aggregate",
      "aggregation": "SUM|AVERAGE|COUNT|COUNTA|MIN|MAX",
      "group_by": ["string — column(s) to group by"],
      "sort": "default|ascending|descending|chronological",
      "description": "string"
    }
  ],
  "charts": [
    {
      "name": "string — unique chart name",
      "chart_type": "column|bar|line|pie|doughnut|combo|area",
      "title": "string — chart title",
      "data_source": "string — must match an analysis name",
      "category_field": "string — x-axis grouping field",
      "value_field": "string — y-axis metric",
      "x_axis_title": "string",
      "y_axis_title": "string",
      "width": 12,
      "height": 7
    }
  ],
  "dashboard": {
    "title": "string — dashboard title",
    "kpis": [
      {
        "name": "string",
        "metric": "string — column or calculated field",
        "aggregation": "SUM|AVERAGE|COUNT|MIN|MAX",
        "format": "number|currency|percentage|integer",
        "label": "string — display label"
      }
    ],
    "chart_refs": ["string — references to chart names"],
    "filter_refs": ["string — references to filter fields"]
  },
  "filters": [
    {
      "field": "string — column name to filter on",
      "filter_type": "autofilter|table_filter",
      "description": "string"
    }
  ],
  "sheets": [
    {
      "name": "string — sheet display name",
      "sheet_type": "raw_data|calculations|analysis|charts|dashboard|assignment",
      "description": "string"
    }
  ]
}
"""


def build_analysis_prompt(
    dataset_schema: list[dict],
    sample_rows: list[dict],
    statistics: dict,
    question_text: str,
    potential_calculations: list[str],
) -> str:
    """Build the user prompt for the LLM with dataset context and the assignment."""

    schema_text = "## Dataset Columns\n"
    for col in dataset_schema:
        schema_text += f"- **{col['name']}** (type: {col['data_type']}, classification: {col['classification']}, unique values: {col.get('unique_count', '?')}, semantic: {col.get('semantic_type', 'unknown')})\n"
        if col.get('sample_values'):
            samples = ", ".join(str(v) for v in col['sample_values'][:5])
            schema_text += f"  Sample values: {samples}\n"

    stats_text = f"""## Dataset Statistics
- Rows: {statistics.get('row_count', 0)}
- Columns: {statistics.get('column_count', 0)}
- Dimensions: {', '.join(statistics.get('dimensions', []))}
- Measures: {', '.join(statistics.get('measures', []))}
- Duplicate rows: {statistics.get('duplicate_rows', 0)}
"""

    if potential_calculations:
        stats_text += f"- Potential calculated measures: {', '.join(potential_calculations)}\n"

    sample_text = "## Sample Data (first 3 rows)\n```\n"
    if sample_rows:
        headers = list(sample_rows[0].keys())
        sample_text += " | ".join(headers) + "\n"
        sample_text += " | ".join(["---"] * len(headers)) + "\n"
        for row in sample_rows[:3]:
            sample_text += " | ".join(str(row.get(h, "")) for h in headers) + "\n"
    sample_text += "```\n"

    prompt = f"""{schema_text}

{stats_text}

{sample_text}

## User Assignment/Question
{question_text}

## Instructions
Analyze the assignment above and produce a complete JSON plan for generating an Excel workbook.
- Include ONLY what the assignment asks for. Do not add extra analyses or charts.
- If the assignment mentions a dashboard, include appropriate KPIs based on the data.
- If the assignment mentions slicers or filters, map them to the "filters" section.
- Every analysis must have a corresponding chart unless the assignment says otherwise.
- Ensure all formula templates reference exact column names from the dataset.
- Always include a raw_data sheet and an assignment sheet.
- If calculations are needed (e.g., Revenue = Units Sold * Unit Price), include them in the calculations list.

Return ONLY the JSON object. No other text.
"""
    return prompt
