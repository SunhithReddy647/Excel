"""
API routes for ExcelFlow — handles all HTTP endpoints with SQLite database persistence.
"""

import os
import uuid
import json
import asyncio
from pathlib import Path
from datetime import datetime
from typing import Optional

from fastapi import APIRouter, UploadFile, File, Form, HTTPException, BackgroundTasks, Depends
from fastapi.responses import FileResponse, StreamingResponse
from pydantic import ValidationError
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from config import settings
from database import get_session, async_session
from models.models import Dataset as DBDataset, Assignment as DBAssignment, Generation as DBGeneration, Workbook as DBWorkbook
from schemas.schemas import (
    DatasetUploadResponse, QuestionUploadResponse, QuestionPasteRequest,
    AnalyzeRequest, AnalyzeResponse, GenerateRequest, ExcelPlan,
    GenerationStatusResponse, GenerationProgress, WorkbookSummary,
    ValidationResult, DatasetProfile,
)
from services.dataset_profiler import profile_dataset, read_dataset
from services.question_parser import extract_question_text, generate_plan
from services.workbook_generator import generate_workbook
from services.validator import validate_workbook
from services.auto_dashboard_architect import architect_executive_plan
from services.ai_dashboard_synthesizer import synthesize_ai_dashboard_plan

router = APIRouter()

# In-memory stores for active progress tracking during background generation
_active_progress: dict[str, list] = {}


# ─── Dataset Upload ──────────────────────────────────────────────────

ALLOWED_DATA_EXTENSIONS = {".xlsx", ".xls", ".csv"}
ALLOWED_QUESTION_EXTENSIONS = {".pdf", ".docx", ".txt", ".png", ".jpg", ".jpeg", ".webp"}


@router.post("/api/datasets/upload", response_model=DatasetUploadResponse)
async def upload_dataset(
    file: UploadFile = File(...),
    session: AsyncSession = Depends(get_session),
):
    """Upload a dataset file (.xlsx, .xls, .csv) and persist metadata to SQLite."""
    # Validate extension
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_DATA_EXTENSIONS:
        raise HTTPException(400, f"Unsupported file type: {ext}. Allowed: {', '.join(ALLOWED_DATA_EXTENSIONS)}")

    # Validate size
    content = await file.read()
    if len(content) > settings.max_file_size_bytes:
        raise HTTPException(400, f"File too large. Maximum: {settings.max_file_size_mb}MB")

    # Save file to storage
    dataset_id = str(uuid.uuid4())
    safe_filename = f"{dataset_id}{ext}"
    file_path = str(settings.upload_dir / safe_filename)

    with open(file_path, "wb") as f:
        f.write(content)

    # Profile dataset
    try:
        df, profile = profile_dataset(file_path)
        profile.dataset_id = dataset_id
        profile.filename = file.filename
    except Exception as e:
        if os.path.exists(file_path):
            os.remove(file_path)
        raise HTTPException(400, f"Failed to read dataset: {str(e)}")

    # Persist in SQLite
    db_dataset = DBDataset(
        id=dataset_id,
        filename=safe_filename,
        original_filename=file.filename,
        file_path=file_path,
        row_count=profile.row_count,
        column_count=profile.column_count,
        headers=profile.headers,
        profile_json=profile.model_dump(),
    )
    session.add(db_dataset)
    await session.commit()

    return DatasetUploadResponse(
        dataset_id=dataset_id,
        filename=file.filename,
        row_count=profile.row_count,
        column_count=profile.column_count,
        headers=profile.headers,
        profile=profile,
    )


# ─── Question Upload / Paste ────────────────────────────────────────

@router.post("/api/questions/upload")
async def upload_question(
    dataset_id: str = Form(...),
    file: UploadFile = File(...),
    session: AsyncSession = Depends(get_session),
):
    """Upload a question file (PDF, DOCX, TXT, image) and persist to SQLite."""
    stmt = select(DBDataset).where(DBDataset.id == dataset_id)
    res = await session.execute(stmt)
    dataset = res.scalar_one_or_none()
    if not dataset:
        raise HTTPException(404, "Dataset not found")

    ext = Path(file.filename).suffix.lower()
    source_type_map = {
        ".pdf": "pdf",
        ".docx": "docx",
        ".txt": "txt",
        ".png": "image",
        ".jpg": "image",
        ".jpeg": "image",
        ".webp": "image",
    }
    source_type = source_type_map.get(ext)
    if not source_type:
        raise HTTPException(400, f"Unsupported question file type: {ext}")

    # Save file
    content = await file.read()
    assignment_id = str(uuid.uuid4())
    safe_filename = f"{assignment_id}{ext}"
    file_path = str(settings.upload_dir / safe_filename)

    with open(file_path, "wb") as f:
        f.write(content)

    # Extract text
    try:
        question_text = extract_question_text(file_path, source_type)
    except Exception as e:
        if os.path.exists(file_path):
            os.remove(file_path)
        raise HTTPException(400, f"Failed to extract text: {str(e)}")

    db_assignment = DBAssignment(
        id=assignment_id,
        dataset_id=dataset_id,
        question_text=question_text,
        source_type=source_type,
        source_filename=file.filename,
    )
    session.add(db_assignment)
    await session.commit()

    return QuestionUploadResponse(
        assignment_id=assignment_id,
        question_text=question_text,
        source_type=source_type,
    )


@router.post("/api/questions/paste")
async def paste_question(
    request: QuestionPasteRequest,
    session: AsyncSession = Depends(get_session),
):
    """Submit a question as pasted text and persist to SQLite."""
    stmt = select(DBDataset).where(DBDataset.id == request.dataset_id)
    res = await session.execute(stmt)
    dataset = res.scalar_one_or_none()
    if not dataset:
        raise HTTPException(404, "Dataset not found")

    assignment_id = str(uuid.uuid4())
    db_assignment = DBAssignment(
        id=assignment_id,
        dataset_id=request.dataset_id,
        question_text=request.question_text,
        source_type="text",
        source_filename=None,
    )
    session.add(db_assignment)
    await session.commit()

    return QuestionUploadResponse(
        assignment_id=assignment_id,
        question_text=request.question_text,
        source_type="text",
    )


# ─── Analyze ─────────────────────────────────────────────────────────

@router.post("/api/analyze")
async def analyze(
    request: AnalyzeRequest,
    session: AsyncSession = Depends(get_session),
):
    """Analyze dataset + question and produce an ExcelPlan, saved to SQLite."""
    # Query dataset
    stmt_d = select(DBDataset).where(DBDataset.id == request.dataset_id)
    res_d = await session.execute(stmt_d)
    dataset = res_d.scalar_one_or_none()
    if not dataset:
        raise HTTPException(404, "Dataset not found")

    # Query assignment
    stmt_a = select(DBAssignment).where(DBAssignment.id == request.assignment_id)
    res_a = await session.execute(stmt_a)
    assignment = res_a.scalar_one_or_none()
    if not assignment:
        raise HTTPException(404, "Assignment not found")

    profile = DatasetProfile.model_validate(dataset.profile_json)
    question_text = assignment.question_text

    # Call LLM to generate plan
    try:
        plan = generate_plan(profile, question_text)
    except Exception as e:
        raise HTTPException(500, f"Failed to generate plan: {str(e)}")

    generation_id = str(uuid.uuid4())
    db_gen = DBGeneration(
        id=generation_id,
        assignment_id=request.assignment_id,
        plan_json=plan.model_dump(),
        status="planned",
        progress_step="Plan generated",
        progress_pct=15,
    )
    session.add(db_gen)
    await session.commit()

    return AnalyzeResponse(
        generation_id=generation_id,
        plan=plan,
        dataset_profile=profile,
    )


# ─── Generate Workbook ───────────────────────────────────────────────

@router.post("/api/workbooks/generate")
async def generate(
    request: GenerateRequest,
    background_tasks: BackgroundTasks,
    session: AsyncSession = Depends(get_session),
):
    """Generate an Excel workbook from a plan, tracked in SQLite."""
    stmt = select(DBGeneration).where(DBGeneration.id == request.generation_id)
    res = await session.execute(stmt)
    gen = res.scalar_one_or_none()
    if not gen:
        raise HTTPException(404, "Generation not found")

    # Update plan in SQLite
    gen.plan_json = request.plan.model_dump()
    gen.status = "generating"
    gen.progress_step = "Queued generation"
    gen.progress_pct = 20
    await session.commit()

    _active_progress[request.generation_id] = [
        {"step": "Queued generation", "status": "in_progress", "message": "Starting generation pipeline"}
    ]

    # Run generation in background
    background_tasks.add_task(
        _run_generation,
        request.generation_id,
    )

    return {"generation_id": request.generation_id, "status": "generating"}


async def _run_generation(generation_id: str):
    """Background task that runs the generation pipeline and records to SQLite."""
    def log_progress(step: str, status: str, message: str = ""):
        if generation_id not in _active_progress:
            _active_progress[generation_id] = []
        _active_progress[generation_id].append({"step": step, "status": status, "message": message})

    async with async_session() as session:
        try:
            stmt = select(DBGeneration).where(DBGeneration.id == generation_id)
            res = await session.execute(stmt)
            gen = res.scalar_one_or_none()
            if not gen:
                return

            log_progress("Analyzing dataset", "completed")
            gen.progress_step = "Loading dataset"
            gen.progress_pct = 30
            await session.commit()

            # Load assignment and dataset
            stmt_a = select(DBAssignment).where(DBAssignment.id == gen.assignment_id)
            res_a = await session.execute(stmt_a)
            assignment = res_a.scalar_one_or_none()

            stmt_d = select(DBDataset).where(DBDataset.id == assignment.dataset_id)
            res_d = await session.execute(stmt_d)
            dataset = res_d.scalar_one_or_none()

            df = read_dataset(dataset.file_path)
            log_progress("Loading data", "completed")

            plan = ExcelPlan.model_validate(gen.plan_json)
            question_text = assignment.question_text

            log_progress("Building calculation plan", "completed")
            gen.progress_step = "Generating workbook & PivotTables"
            gen.progress_pct = 50
            await session.commit()

            # Generate workbook
            log_progress("Generating workbook", "in_progress")
            workbook_id = str(uuid.uuid4())
            title = plan.workbook_title.replace(" ", "_").replace("/", "_")
            filename = f"{title}_Generated.xlsx"
            output_path = str(settings.generated_dir / f"{workbook_id}.xlsx")

            summary = generate_workbook(df, plan, question_text, output_path)
            log_progress(
                "Generating workbook",
                "completed",
                f"Generated with {summary.get('pivots_created', 0)} native PivotTables"
            )

            # Validate
            log_progress("Validating workbook", "in_progress")
            gen.progress_step = "Validating workbook"
            gen.progress_pct = 85
            await session.commit()

            validation = validate_workbook(output_path, plan)
            validation.workbook_id = workbook_id
            log_progress("Validating workbook", "completed", f"Status: {validation.overall_status}")

            # Persist Workbook record to SQLite
            db_wb = DBWorkbook(
                id=workbook_id,
                generation_id=generation_id,
                filename=filename,
                file_path=output_path,
                validation_json=validation.model_dump(),
                validation_status=validation.overall_status,
                summary_json=summary,
            )
            session.add(db_wb)

            gen.status = "completed"
            gen.progress_step = "Completed"
            gen.progress_pct = 100
            gen.completed_at = datetime.utcnow()
            await session.commit()

        except Exception as e:
            if gen:
                gen.status = "failed"
                gen.error = str(e)
                gen.progress_step = "Failed"
                await session.commit()
            log_progress("Generation failed", "failed", str(e))


# ─── Status & Download ───────────────────────────────────────────────

@router.get("/api/workbooks/{generation_id}/status")
async def get_status(
    generation_id: str,
    session: AsyncSession = Depends(get_session),
):
    """Get the current generation status from SQLite."""
    stmt = select(DBGeneration).where(DBGeneration.id == generation_id)
    res = await session.execute(stmt)
    gen = res.scalar_one_or_none()
    if not gen:
        raise HTTPException(404, "Generation not found")

    # Fetch associated workbook if completed
    stmt_wb = select(DBWorkbook).where(DBWorkbook.generation_id == generation_id)
    res_wb = await session.execute(stmt_wb)
    wb = res_wb.scalar_one_or_none()

    progress = _active_progress.get(generation_id, [])

    return {
        "generation_id": generation_id,
        "status": gen.status,
        "progress": progress,
        "error": gen.error,
        "progress_step": gen.progress_step,
        "progress_pct": gen.progress_pct,
        "workbook_id": wb.id if wb else None,
    }


@router.get("/api/workbooks/{workbook_id}/summary")
async def get_summary(
    workbook_id: str,
    session: AsyncSession = Depends(get_session),
):
    """Get workbook summary and validation results from SQLite."""
    stmt = select(DBWorkbook).where(DBWorkbook.id == workbook_id)
    res = await session.execute(stmt)
    wb = res.scalar_one_or_none()
    if not wb:
        raise HTTPException(404, "Workbook not found")

    validation = wb.validation_json or {}
    summary = wb.summary_json or {}

    return {
        "workbook_id": wb.id,
        "filename": wb.filename,
        "sheets": summary.get("sheets", []),
        "row_count": summary.get("row_count", 0),
        "formula_count": summary.get("formula_count", 0),
        "chart_count": summary.get("chart_count", 0),
        "analysis_count": summary.get("analysis_count", 0),
        "filter_count": summary.get("filter_count", 0),
        "native_pivots": summary.get("native_pivots", False),
        "pivots_created": summary.get("pivots_created", 0),
        "slicers_created": summary.get("slicers_created", 0),
        "validation": validation,
    }


@router.get("/api/workbooks/{workbook_id}/download")
async def download_workbook(
    workbook_id: str,
    session: AsyncSession = Depends(get_session),
):
    """Download the generated workbook from SQLite path."""
    stmt = select(DBWorkbook).where(DBWorkbook.id == workbook_id)
    res = await session.execute(stmt)
    wb = res.scalar_one_or_none()
    if not wb:
        raise HTTPException(404, "Workbook not found")

    file_path = wb.file_path
    if not os.path.exists(file_path):
        raise HTTPException(404, "Workbook file not found on disk")

    return FileResponse(
        path=file_path,
        filename=wb.filename,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ─── Example Questions ───────────────────────────────────────────────

EXAMPLE_QUESTIONS = {
    "sales_dashboard": """Create a Sales Dashboard using the given dataset.

The dashboard should contain:
1. Revenue by Region
2. Revenue by Product
3. Revenue by Month
4. Units Sold by Salesperson
5. A filter for Month
6. A filter for Region""",

    "employee_analysis": """Perform a comprehensive Employee Analysis.

Include:
1. Average Salary by Department
2. Employee Count by Department
3. Performance Rating Distribution
4. Salary Distribution by Gender
5. Top performers by Department""",

    "financial_analysis": """Create a Financial Analysis Report.

Include:
1. Monthly Revenue Trend
2. Profit Margin by Product
3. Cost vs Revenue comparison
4. Quarterly Performance Summary
5. Year-over-Year Growth""",

    "inventory_analysis": """Create an Inventory Analysis Dashboard.

Include:
1. Stock Levels by Product
2. Sales vs Stock comparison
3. Warehouse Utilization
4. Low Stock Alerts
5. Top Selling Products""",

    "marketing_analysis": """Create a Marketing Campaign Analysis.

Include:
1. Campaign Performance by Channel
2. ROI by Campaign
3. Conversion Rate Trends
4. Budget vs Actual Spend
5. Customer Acquisition Cost""",

    "customer_analysis": """Create a Customer Analysis Dashboard.

Include:
1. Customer Segmentation
2. Revenue by Customer Type
3. Customer Retention Rates
4. Average Order Value by Segment
5. Geographic Distribution""",
}


@router.get("/api/examples")
async def get_examples():
    """Get available example questions."""
    return {
        name: {
            "title": name.replace("_", " ").title(),
            "question": question,
        }
        for name, question in EXAMPLE_QUESTIONS.items()
    }


# ─── Direct Client Dashboard Generation ─────────────────────────────

@router.post("/api/dashboard/generate-from-data")
async def generate_dashboard_from_data(
    file: UploadFile = File(...),
    title: Optional[str] = Form(None),
    question: Optional[str] = Form(None),
    prompt: Optional[str] = Form(None),
    question_file: Optional[UploadFile] = File(None),
    session: AsyncSession = Depends(get_session),
):
    """
    Direct client endpoint: Upload any dataset (CSV, XLSX, XLS) and optional
    question/assignment text or file. Uses Gemini AI to deeply analyze both
    the data and question asked to compile a tailored, executive Excel dashboard.
    """
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_DATA_EXTENSIONS:
        raise HTTPException(400, f"Unsupported file type: {ext}. Allowed: {', '.join(ALLOWED_DATA_EXTENSIONS)}")

    content = await file.read()
    if len(content) > settings.max_file_size_bytes:
        raise HTTPException(400, f"File too large. Maximum: {settings.max_file_size_mb}MB")

    # Save uploaded file
    dataset_id = str(uuid.uuid4())
    safe_filename = f"{dataset_id}{ext}"
    file_path = str(settings.upload_dir / safe_filename)

    with open(file_path, "wb") as f:
        f.write(content)

    # Handle question file if provided
    extracted_q_text = ""
    if question_file and question_file.filename:
        q_ext = Path(question_file.filename).suffix.lower()
        source_type_map = {
            ".pdf": "pdf",
            ".docx": "docx",
            ".txt": "txt",
            ".png": "image",
            ".jpg": "image",
            ".jpeg": "image",
            ".webp": "image",
        }
        source_type = source_type_map.get(q_ext)
        if source_type:
            q_id = str(uuid.uuid4())
            q_file_path = str(settings.upload_dir / f"{q_id}{q_ext}")
            q_bytes = await question_file.read()
            with open(q_file_path, "wb") as f:
                f.write(q_bytes)
            try:
                extracted_q_text = extract_question_text(q_file_path, source_type)
            except Exception:
                pass

    raw_question = (question or prompt or "").strip()
    full_question = f"{raw_question}\n{extracted_q_text}".strip() if extracted_q_text else raw_question

    try:
        # 1. Profile dataset
        df, profile = profile_dataset(file_path)
        profile.dataset_id = dataset_id
        profile.filename = file.filename

        # 2. Architect tailored executive plan with AI synthesis
        dashboard_title = title.strip() if title else f"Executive {Path(file.filename).stem.replace('_', ' ').title()} Dashboard"
        plan, ai_meta = synthesize_ai_dashboard_plan(
            df=df,
            profile=profile,
            question_text=full_question,
            title=dashboard_title,
        )

        # 3. Generate workbook
        workbook_id = str(uuid.uuid4())
        safe_title = re.sub(r'[\\/*?:"<>| ]', '_', plan.workbook_title)
        gen_filename = f"{safe_title}.xlsx"
        output_path = str(settings.generated_dir / f"{workbook_id}.xlsx")

        summary = generate_workbook(df, plan, full_question or "Executive Dashboard Generation", output_path)

        # 4. Validate workbook
        validation = validate_workbook(output_path, plan)
        validation.workbook_id = workbook_id

        # 5. Persist to SQLite
        db_dataset = DBDataset(
            id=dataset_id,
            filename=safe_filename,
            original_filename=file.filename,
            file_path=file_path,
            row_count=profile.row_count,
            column_count=profile.column_count,
            headers=profile.headers,
            profile_json=profile.model_dump(),
        )
        session.add(db_dataset)

        db_wb = DBWorkbook(
            id=workbook_id,
            generation_id=dataset_id,
            filename=gen_filename,
            file_path=output_path,
            validation_json=validation.model_dump(),
            validation_status=validation.overall_status,
            summary_json=summary,
        )
        session.add(db_wb)
        await session.commit()

        return {
            "success": True,
            "workbook_id": workbook_id,
            "filename": gen_filename,
            "download_url": f"/api/workbooks/{workbook_id}/download",
            "row_count": len(df),
            "column_count": len(df.columns),
            "sheets": summary.get("sheets", []),
            "formula_count": summary.get("formula_count", 0),
            "chart_count": summary.get("chart_count", 0),
            "native_pivots": summary.get("native_pivots", False),
            "slicers_created": summary.get("slicers_created", 0),
            "validation_status": validation.overall_status,
            "ai_used": ai_meta.get("ai_used", False),
            "ai_summary": ai_meta.get("summary", ""),
            "question_analyzed": ai_meta.get("question_analyzed", full_question),
            "kpis_created": len(plan.dashboard.kpis) if plan.dashboard else 0,
            "analyses_created": len(plan.analyses),
        }
    except Exception as e:
        if os.path.exists(file_path):
            try:
                os.remove(file_path)
            except Exception:
                pass
        raise HTTPException(500, f"Dashboard generation failed: {str(e)}")
