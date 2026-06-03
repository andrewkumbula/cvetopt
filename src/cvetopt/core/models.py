from __future__ import annotations

from datetime import date, datetime
from enum import Enum

from pydantic import BaseModel, Field


class JobStatus(str, Enum):
    pending = "pending"
    running = "running"
    completed = "completed"
    failed = "failed"
    cancelled = "cancelled"


class Order(BaseModel):
    portal_id: str
    order_id: str
    name: str | None = None
    created_at: date | None = None
    flight_date: date
    cargo: str | None = None


class JobState(BaseModel):
    id: str
    portal_id: str
    status: JobStatus = JobStatus.pending
    created_at: datetime = Field(default_factory=datetime.utcnow)
    finished_at: datetime | None = None
    logs: list[str] = Field(default_factory=list)
    downloaded_paths: list[str] = Field(default_factory=list)
    error: str | None = None
