"""
Question Parser — extracts text from uploaded files and calls the LLM
to produce a validated ExcelPlan JSON.
"""

import json
import re
from pathlib import Path
from typing import Optional

import google.generativeai as genai
from config import settings
from schemas.schemas import ExcelPlan, DatasetProfile
from prompts.prompts import SYSTEM_PROMPT, build_analysis_prompt


def _extract_text_from_pdf(file_path: str) -> str:
    from PyPDF2 import PdfReader
    reader = PdfReader(file_path)
    text_parts = []
    for page in reader.pages:
        text_parts.append(page.extract_text() or "")
    return "\n".join(text_parts).strip()


def _extract_text_from_docx(file_path: str) -> str:
    from docx import Document
    doc = Document(file_path)
    return "\n".join(para.text for para in doc.paragraphs).strip()


def _extract_text_from_txt(file_path: str) -> str:
    return Path(file_path).read_text(encoding="utf-8").strip()


def _extract_text_from_image(file_path: str) -> str:
    """Use Gemini Vision to extract text from an image."""
    import PIL.Image
    genai.configure(api_key=settings.gemini_api_key)
    model = genai.GenerativeModel(settings.gemini_model)
    img = PIL.Image.open(file_path)
    response = model.generate_content([
        "Extract all text from this image. Return only the text content, no commentary.",
        img,
    ])
    return response.text.strip()


def extract_question_text(file_path: str, source_type: str) -> str:
    """Extract question text from an uploaded file."""
    extractors = {
        "pdf": _extract_text_from_pdf,
        "docx": _extract_text_from_docx,
        "txt": _extract_text_from_txt,
        "image": _extract_text_from_image,
    }
    extractor = extractors.get(source_type)
    if not extractor:
        raise ValueError(f"Unsupported source type: {source_type}")
    return extractor(file_path)


def _clean_json_response(text: str) -> str:
    """Strip markdown code fences and extra text from LLM response."""
    # Remove ```json ... ``` wrappers
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*\n?", "", text)
        text = re.sub(r"\n?```\s*$", "", text)
    return text.strip()


def generate_plan(
    profile: DatasetProfile,
    question_text: str,
    max_retries: int = 3,
) -> ExcelPlan:
    """
    Call the LLM with dataset context + question, get back a validated ExcelPlan.
    Retries on validation failure.
    """
    genai.configure(api_key=settings.gemini_api_key)
    model = genai.GenerativeModel(
        settings.gemini_model,
        system_instruction=SYSTEM_PROMPT,
    )

    # Build the structured dataset context for the prompt
    schema_dicts = [col.model_dump() for col in profile.columns]
    stats = {
        "row_count": profile.row_count,
        "column_count": profile.column_count,
        "dimensions": profile.dimensions,
        "measures": profile.measures,
        "duplicate_rows": profile.duplicate_rows,
    }

    # Use sample values from columns as pseudo-rows
    sample_rows = []
    if profile.columns:
        num_samples = min(3, max(len(c.sample_values) for c in profile.columns) if profile.columns else 0)
        for i in range(num_samples):
            row = {}
            for col in profile.columns:
                row[col.name] = col.sample_values[i] if i < len(col.sample_values) else ""
            sample_rows.append(row)

    prompt = build_analysis_prompt(
        dataset_schema=schema_dicts,
        sample_rows=sample_rows,
        statistics=stats,
        question_text=question_text,
        potential_calculations=profile.potential_calculations,
    )

    last_error = None
    for attempt in range(max_retries):
        try:
            response = model.generate_content(prompt)
            raw_text = response.text
            cleaned = _clean_json_response(raw_text)
            plan_dict = json.loads(cleaned)
            plan = ExcelPlan.model_validate(plan_dict)

            # Post-validation: ensure sheets list is populated
            if not plan.sheets:
                plan.sheets = _generate_default_sheets(plan)

            return plan

        except (json.JSONDecodeError, Exception) as e:
            last_error = e
            # On retry, append error feedback to the prompt
            prompt += f"\n\n## PREVIOUS ATTEMPT FAILED\nError: {str(e)}\nPlease fix the JSON and try again. Return ONLY valid JSON."

    raise RuntimeError(f"Failed to generate valid plan after {max_retries} attempts: {last_error}")


def _generate_default_sheets(plan: ExcelPlan) -> list:
    """Generate a default sheet list based on what the plan contains."""
    from schemas.schemas import SheetSpec
    sheets = [SheetSpec(name="01_Raw_Data", sheet_type="raw_data", description="Original dataset")]

    if plan.calculations:
        sheets.append(SheetSpec(name="02_Calculations", sheet_type="calculations", description="Calculated columns"))

    if plan.analyses:
        sheets.append(SheetSpec(name="03_Analysis", sheet_type="analysis", description="Summary tables and analysis"))

    if plan.charts:
        sheets.append(SheetSpec(name="04_Charts", sheet_type="charts", description="Data visualizations"))

    if plan.dashboard:
        sheets.append(SheetSpec(name="05_Dashboard", sheet_type="dashboard", description="Interactive dashboard"))

    sheets.append(SheetSpec(name="06_Assignment", sheet_type="assignment", description="Original assignment question"))

    return sheets
