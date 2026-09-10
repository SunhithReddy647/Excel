"""
Tests for SQLite persistence and API endpoints in ExcelFlow.
"""

import os
import pytest
import pytest_asyncio
from httpx import AsyncClient, ASGITransport
from sqlalchemy import select

from main import app
from database import init_db, async_session
from models.models import Dataset, Assignment, Generation, Workbook


@pytest_asyncio.fixture(autouse=True)
async def setup_database():
    """Ensure database tables exist before running tests."""
    await init_db()


@pytest.mark.asyncio
async def test_dataset_upload_and_sqlite_persistence():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # Upload fixture dataset
        fixture_path = os.path.join(os.path.dirname(__file__), "fixtures", "sales_data.csv")
        with open(fixture_path, "rb") as f:
            response = await client.post(
                "/api/datasets/upload",
                files={"file": ("sales_data.csv", f, "text/csv")},
            )

        assert response.status_code == 200
        data = response.json()
        dataset_id = data["dataset_id"]
        assert dataset_id is not None
        assert data["row_count"] == 16
        assert "Region" in data["headers"]

        # Verify record exists in SQLite
        async with async_session() as session:
            stmt = select(Dataset).where(Dataset.id == dataset_id)
            result = await session.execute(stmt)
            db_dataset = result.scalar_one_or_none()
            assert db_dataset is not None
            assert db_dataset.row_count == 16
            assert db_dataset.original_filename == "sales_data.csv"


@pytest.mark.asyncio
async def test_question_paste_and_sqlite_persistence():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # First upload dataset
        fixture_path = os.path.join(os.path.dirname(__file__), "fixtures", "sales_data.csv")
        with open(fixture_path, "rb") as f:
            upload_res = await client.post(
                "/api/datasets/upload",
                files={"file": ("sales_data.csv", f, "text/csv")},
            )
        dataset_id = upload_res.json()["dataset_id"]

        # Paste question
        paste_res = await client.post(
            "/api/questions/paste",
            json={
                "dataset_id": dataset_id,
                "question_text": "Calculate revenue and profit by region with interactive pivot tables.",
            },
        )
        assert paste_res.status_code == 200
        paste_data = paste_res.json()
        assignment_id = paste_data["assignment_id"]
        assert assignment_id is not None

        # Verify in SQLite
        async with async_session() as session:
            stmt = select(Assignment).where(Assignment.id == assignment_id)
            result = await session.execute(stmt)
            db_assignment = result.scalar_one_or_none()
            assert db_assignment is not None
            assert db_assignment.dataset_id == dataset_id
            assert "revenue" in db_assignment.question_text.lower()


@pytest.mark.asyncio
async def test_full_workflow_with_sqlite_and_pivots():
    transport = ASGITransport(app=app)
    async with AsyncClient(transport=transport, base_url="http://test") as client:
        # 1. Upload dataset
        fixture_path = os.path.join(os.path.dirname(__file__), "fixtures", "sales_data.csv")
        with open(fixture_path, "rb") as f:
            upload_res = await client.post(
                "/api/datasets/upload",
                files={"file": ("sales_data.csv", f, "text/csv")},
            )
        assert upload_res.status_code == 200
        dataset_id = upload_res.json()["dataset_id"]

        # 2. Paste question
        paste_res = await client.post(
            "/api/questions/paste",
            json={
                "dataset_id": dataset_id,
                "question_text": "Create a Sales report analyzing Revenue by Region and Product.",
            },
        )
        assert paste_res.status_code == 200
        assignment_id = paste_res.json()["assignment_id"]

        # 3. Analyze with Gemini
        analyze_res = await client.post(
            "/api/analyze",
            json={"dataset_id": dataset_id, "assignment_id": assignment_id},
        )
        assert analyze_res.status_code == 200
        analyze_data = analyze_res.json()
        generation_id = analyze_data["generation_id"]
        plan = analyze_data["plan"]
        assert generation_id is not None
        assert plan["workbook_title"] is not None

        # 4. Generate workbook
        gen_res = await client.post(
            "/api/workbooks/generate",
            json={"generation_id": generation_id, "plan": plan},
        )
        assert gen_res.status_code == 200

        # Wait for background generation to complete
        import asyncio
        for _ in range(30):
            await asyncio.sleep(1)
            status_res = await client.get(f"/api/workbooks/{generation_id}/status")
            status_data = status_res.json()
            if status_data["status"] in ("completed", "failed"):
                break

        assert status_data["status"] == "completed"
        workbook_id = status_data["workbook_id"]
        assert workbook_id is not None

        # 5. Summary from SQLite
        sum_res = await client.get(f"/api/workbooks/{workbook_id}/summary")
        assert sum_res.status_code == 200
        sum_data = sum_res.json()
        assert sum_data["workbook_id"] == workbook_id
        assert sum_data["row_count"] == 16
        assert sum_data["pivots_created"] >= 1
        assert sum_data["native_pivots"] is True

        # 6. Download workbook
        down_res = await client.get(f"/api/workbooks/{workbook_id}/download")
        assert down_res.status_code == 200
        assert len(down_res.content) > 10000

