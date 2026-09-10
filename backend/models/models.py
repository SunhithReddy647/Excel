"""
SQLAlchemy models for ExcelFlow.
"""

import uuid
from datetime import datetime
from sqlalchemy import Column, String, Integer, Text, DateTime, ForeignKey, JSON
from sqlalchemy.orm import relationship
from database import Base


def generate_uuid() -> str:
    return str(uuid.uuid4())


class Dataset(Base):
    __tablename__ = "datasets"

    id = Column(String, primary_key=True, default=generate_uuid)
    filename = Column(String, nullable=False)
    original_filename = Column(String, nullable=False)
    file_path = Column(String, nullable=False)
    row_count = Column(Integer, default=0)
    column_count = Column(Integer, default=0)
    headers = Column(JSON, default=list)
    schema_json = Column(JSON, default=dict)
    profile_json = Column(JSON, default=dict)
    created_at = Column(DateTime, default=datetime.utcnow)

    assignments = relationship("Assignment", back_populates="dataset")


class Assignment(Base):
    __tablename__ = "assignments"

    id = Column(String, primary_key=True, default=generate_uuid)
    dataset_id = Column(String, ForeignKey("datasets.id"), nullable=False)
    question_text = Column(Text, nullable=False)
    source_type = Column(String, default="text")  # text, pdf, docx, image
    source_filename = Column(String, nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    dataset = relationship("Dataset", back_populates="assignments")
    generations = relationship("Generation", back_populates="assignment")


class Generation(Base):
    __tablename__ = "generations"

    id = Column(String, primary_key=True, default=generate_uuid)
    assignment_id = Column(String, ForeignKey("assignments.id"), nullable=False)
    plan_json = Column(JSON, default=dict)
    status = Column(String, default="pending")  # pending, analyzing, generating, validating, completed, failed
    error = Column(Text, nullable=True)
    progress_step = Column(String, nullable=True)
    progress_pct = Column(Integer, default=0)
    created_at = Column(DateTime, default=datetime.utcnow)
    completed_at = Column(DateTime, nullable=True)

    assignment = relationship("Assignment", back_populates="generations")
    workbook = relationship("Workbook", back_populates="generation", uselist=False)


class Workbook(Base):
    __tablename__ = "workbooks"

    id = Column(String, primary_key=True, default=generate_uuid)
    generation_id = Column(String, ForeignKey("generations.id"), nullable=False)
    filename = Column(String, nullable=False)
    file_path = Column(String, nullable=False)
    validation_json = Column(JSON, default=dict)
    validation_status = Column(String, default="pending")  # pending, passed, failed
    summary_json = Column(JSON, default=dict)
    created_at = Column(DateTime, default=datetime.utcnow)

    generation = relationship("Generation", back_populates="workbook")
