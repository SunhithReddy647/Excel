"""
Dataset Profiler — reads uploaded Excel/CSV, detects types, classifies columns,
computes statistics, and identifies potential calculated measures.
"""

import pandas as pd
import numpy as np
from pathlib import Path
from schemas.schemas import DatasetProfile, ColumnProfile


# Heuristic keywords for semantic classification
DIMENSION_KEYWORDS = {
    "region", "country", "state", "city", "area", "zone", "territory",
    "department", "division", "team", "group", "category", "type", "class",
    "name", "employee", "salesperson", "customer", "vendor", "supplier",
    "product", "item", "sku", "brand", "model",
    "month", "year", "quarter", "date", "week", "day", "period",
    "status", "gender", "designation", "level", "grade", "rating",
    "channel", "segment", "source", "medium",
}

MEASURE_KEYWORDS = {
    "revenue", "sales", "income", "amount", "total", "value",
    "cost", "expense", "price", "fee", "charge",
    "profit", "margin", "earnings", "net",
    "units", "quantity", "count", "volume", "stock", "inventory",
    "target", "budget", "forecast", "actual", "plan",
    "salary", "wage", "bonus", "commission",
    "score", "rating", "percentage", "rate", "ratio",
    "discount", "tax", "interest",
    "roi", "growth", "achievement",
}

# Common calculated-measure patterns: (result_name, required_cols_substrings)
CALC_PATTERNS = [
    ("Revenue", [("unit", "price"), ("unit", "sold")]),
    ("Revenue", [("price",), ("quantity",)]),
    ("Revenue", [("price",), ("units",)]),
    ("Expenses", [("unit", "cost"), ("unit", "sold")]),
    ("Expenses", [("cost",), ("quantity",)]),
    ("Profit", [("revenue",), ("expense",)]),
    ("Profit", [("revenue",), ("cost",)]),
    ("Profit Margin", [("profit",), ("revenue",)]),
    ("Target Achievement", [("actual",), ("target",)]),
    ("Target Achievement", [("unit", "sold"), ("target",)]),
]


def read_dataset(file_path: str) -> pd.DataFrame:
    """Read Excel or CSV into a DataFrame."""
    p = Path(file_path)
    ext = p.suffix.lower()
    if ext == ".csv":
        return pd.read_csv(p)
    elif ext in (".xlsx", ".xls"):
        return pd.read_excel(p, engine="openpyxl")
    else:
        raise ValueError(f"Unsupported file format: {ext}")


def _detect_data_type(series: pd.Series) -> str:
    """Detect the semantic data type of a pandas Series."""
    if series.dropna().empty:
        return "string"

    dtype = series.dtype

    # Check for boolean
    unique_vals = set(series.dropna().unique())
    if unique_vals <= {True, False, 0, 1, "Yes", "No", "yes", "no", "TRUE", "FALSE"}:
        return "boolean"

    # Check for date
    if pd.api.types.is_datetime64_any_dtype(series):
        return "date"

    # Numeric checks
    if pd.api.types.is_integer_dtype(dtype):
        return "integer"
    if pd.api.types.is_float_dtype(dtype):
        # Check if it's percentage-like (values between 0 and 1 with few unique)
        non_null = series.dropna()
        if len(non_null) > 0 and non_null.min() >= 0 and non_null.max() <= 1:
            return "percentage"
        return "decimal"

    # String column — check for currency symbols
    sample = series.dropna().astype(str)
    if sample.str.contains(r"[₹$€£¥]", regex=True).any():
        return "currency"

    # Check if string column holds dates
    try:
        pd.to_datetime(sample.head(10), infer_datetime_format=True)
        return "date"
    except (ValueError, TypeError):
        pass

    return "string"


def _classify_column(col_name: str, data_type: str) -> str:
    """Classify a column as 'dimension' or 'measure'."""
    lower = col_name.lower().replace("_", " ").replace("-", " ")
    tokens = set(lower.split())

    if data_type in ("integer", "decimal", "currency", "percentage"):
        # Numeric columns are usually measures unless they're IDs
        if any(kw in lower for kw in ("id", "code", "number", "no.", "no ")):
            return "dimension"
        return "measure"

    if data_type == "date":
        return "dimension"

    # String — check keywords
    if tokens & MEASURE_KEYWORDS:
        return "measure"
    return "dimension"


def _detect_semantic_type(col_name: str) -> str:
    """Map a column name to a canonical semantic type."""
    lower = col_name.lower().replace("_", " ").replace("-", " ")

    mapping = {
        "unit price": "unit_price",
        "price": "unit_price",
        "unit cost": "unit_cost",
        "cost": "unit_cost",
        "units sold": "quantity",
        "units": "quantity",
        "quantity": "quantity",
        "volume": "quantity",
        "revenue": "revenue",
        "sales": "revenue",
        "total sales": "revenue",
        "sales amount": "revenue",
        "profit": "profit",
        "margin": "margin",
        "target": "target",
        "target units": "target",
        "budget": "target",
        "region": "region",
        "area": "region",
        "territory": "region",
        "zone": "region",
        "month": "month",
        "year": "year",
        "quarter": "quarter",
        "date": "date",
        "product": "product",
        "item": "product",
        "category": "category",
        "department": "department",
        "employee": "employee",
        "salesperson": "employee",
        "customer": "customer",
        "customer type": "customer_type",
        "salary": "salary",
        "performance": "performance",
        "stock": "stock",
        "inventory": "stock",
        "warehouse": "warehouse",
        "discount": "discount",
    }

    for key, semantic in mapping.items():
        if key in lower:
            return semantic
    return lower.replace(" ", "_")


def _find_potential_calculations(columns: list[ColumnProfile], headers: list[str]) -> list[str]:
    """Identify calculated measures that could be derived from existing columns."""
    lower_headers = [h.lower().replace("_", " ") for h in headers]
    found = []

    for calc_name, required_groups in CALC_PATTERNS:
        all_groups_found = True
        for keyword_group in required_groups:
            group_found = False
            for header in lower_headers:
                if all(kw in header for kw in keyword_group):
                    group_found = True
                    break
            if not group_found:
                all_groups_found = False
                break
        if all_groups_found and calc_name not in found:
            # Only suggest if the result doesn't already exist
            if calc_name.lower() not in lower_headers:
                found.append(calc_name)

    return found


def profile_dataset(file_path: str) -> tuple[pd.DataFrame, DatasetProfile]:
    """
    Profile a dataset file. Returns the DataFrame and a DatasetProfile.
    """
    df = read_dataset(file_path)

    headers = list(df.columns)
    columns: list[ColumnProfile] = []

    for col in headers:
        series = df[col]
        data_type = _detect_data_type(series)
        classification = _classify_column(col, data_type)
        semantic_type = _detect_semantic_type(col)

        non_null = series.dropna()
        sample_values = non_null.head(5).tolist()
        # Convert numpy types for JSON serialization
        sample_values = [
            v.item() if isinstance(v, (np.integer, np.floating)) else v
            for v in sample_values
        ]

        min_val = None
        max_val = None
        mean_val = None

        if data_type in ("integer", "decimal", "currency", "percentage"):
            numeric = pd.to_numeric(series, errors="coerce").dropna()
            if len(numeric) > 0:
                min_val = str(numeric.min())
                max_val = str(numeric.max())
                mean_val = round(float(numeric.mean()), 2)

        columns.append(ColumnProfile(
            name=col,
            data_type=data_type,
            semantic_type=semantic_type,
            classification=classification,
            unique_count=int(series.nunique()),
            missing_count=int(series.isna().sum()),
            sample_values=sample_values,
            min_value=min_val,
            max_value=max_val,
            mean_value=mean_val,
        ))

    dimensions = [c.name for c in columns if c.classification == "dimension"]
    measures = [c.name for c in columns if c.classification == "measure"]
    potential_calcs = _find_potential_calculations(columns, headers)

    profile = DatasetProfile(
        filename=Path(file_path).name,
        row_count=len(df),
        column_count=len(headers),
        headers=headers,
        columns=columns,
        dimensions=dimensions,
        measures=measures,
        potential_calculations=potential_calcs,
        duplicate_rows=int(df.duplicated().sum()),
    )

    return df, profile
