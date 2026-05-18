from __future__ import annotations

import asyncio
import uuid
from collections.abc import Coroutine
from datetime import datetime
from typing import Any

from loguru import logger

from cvetopt.core.models import JobState, JobStatus


def _job_log_timestamp() -> str:
    """Локальное время сервера с миллисекундами — по разнице между строками видно длительность шагов."""
    now = datetime.now()
    ms = now.microsecond // 1000
    return now.strftime("%Y-%m-%d %H:%M:%S") + f".{ms:03d}"


class JobManager:
    def __init__(self) -> None:
        self._jobs: dict[str, JobState] = {}
        self._lock = asyncio.Lock()

    def create_job(self, portal_id: str) -> JobState:
        jid = str(uuid.uuid4())
        job = JobState(id=jid, portal_id=portal_id)
        self._jobs[jid] = job
        return job

    def get(self, job_id: str) -> JobState | None:
        return self._jobs.get(job_id)

    def has_active_job(self) -> bool:
        """True, если есть хотя бы один RUNNING/PENDING job."""
        return any(
            j.status in (JobStatus.pending, JobStatus.running) for j in self._jobs.values()
        )

    def list_recent(self, limit: int = 10) -> list[JobState]:
        items = sorted(self._jobs.values(), key=lambda j: j.created_at, reverse=True)
        return items[:limit]

    async def append_log(self, job_id: str, line: str) -> None:
        stamped = f"{_job_log_timestamp()} | {line}"
        logger.info("[job {}] {}", job_id, stamped)
        async with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job.logs.append(stamped)
                if len(job.logs) > 2000:
                    job.logs = job.logs[-1500:]

    async def set_status(
        self,
        job_id: str,
        status: JobStatus,
        error: str | None = None,
    ) -> None:
        async with self._lock:
            job = self._jobs.get(job_id)
            if not job:
                return
            job.status = status
            job.error = error
            if status in (JobStatus.completed, JobStatus.failed):
                from datetime import datetime

                job.finished_at = datetime.utcnow()

    async def add_downloaded(self, job_id: str, path: str) -> None:
        async with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job.downloaded_paths.append(path)


job_manager = JobManager()


async def job_log(job_id: str, msg: str) -> None:
    await job_manager.append_log(job_id, msg)


async def run_coro_logged(
    job_id: str,
    coro: Coroutine[Any, Any, None],
) -> None:
    await job_manager.set_status(job_id, JobStatus.running)
    try:
        await coro
    except Exception as e:
        logger.exception("Job {} failed", job_id)
        await job_manager.append_log(job_id, f"Ошибка: {e}")
        await job_manager.set_status(job_id, JobStatus.failed, error=str(e))
    else:
        await job_manager.set_status(job_id, JobStatus.completed)
