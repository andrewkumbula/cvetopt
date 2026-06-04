from __future__ import annotations

import asyncio
import subprocess
import sys
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


def _is_ui_log_line(line: str) -> bool:
    """
    Фильтр для лога в UI:
    - в файл/консоль пишем ВСЁ,
    - в интерфейсе показываем только ключевые события для человека.
    """
    text = (line or "").strip()
    low = text.lower()
    if not text:
        return False

    noisy_prefixes = (
        "клик по селектору:",
        "селектор ",
        "таблица движений: вариант",
        "строка заголовка таблицы",
        "колонки: операция=",
        "загружена сохранённая сессия",
    )
    if low.startswith(noisy_prefixes):
        return False

    key_markers = (
        "открываю ",
        "выполняю вход",
        "сессия сохранена",
        "переопределён период",
        "ожидаемые даты вылета",
        "всего уникальных заказов",
        "подходящих imp-записей",
        "найдено целевых строк",
        "записано ",
        "сохранено:",
        "файл сохранён:",
        "готово",
        "ошибка",
        "пропуск",
        "не найден",
        "не удалось",
        "не создан",
        "нет данных",
        "перезапуск",
        "эквадор",
        "архив",
        "в архив перенесено",
        "папка скачивания",
        "уже в реестре",
        "останов",
        "дубликат",
        "постобработка",
        "номер ",
        "папку 1",
        "папку 2",
        "папка 1",
        "папка 2",
        "очищены",
        "price",
        "auto1",
        "шаг «",
        "scan",
        "import",
        "calculate",
        "склад",
        "книга:",
        "макрос",
        "excel",
    )
    if any(marker in low for marker in key_markers):
        return True

    # Сохраняем краткие сводки (например "Страница 3: ... 2")
    if low.startswith("страница ") and "заказов в диапазоне" in low:
        return True

    return False


def _kill_excel_processes() -> None:
    if sys.platform != "win32":
        return
    subprocess.run(
        ["taskkill", "/im", "EXCEL.EXE", "/f"],
        capture_output=True,
        check=False,
    )


class JobManager:
    def __init__(self) -> None:
        self._jobs: dict[str, JobState] = {}
        self._tasks: dict[str, asyncio.Task[None]] = {}
        self._cancel_requested: set[str] = set()
        self._lock = asyncio.Lock()

    def create_job(self, portal_id: str) -> JobState:
        jid = str(uuid.uuid4())
        job = JobState(id=jid, portal_id=portal_id)
        self._jobs[jid] = job
        return job

    def get(self, job_id: str) -> JobState | None:
        return self._jobs.get(job_id)

    def cancel_requested(self, job_id: str) -> bool:
        return job_id in self._cancel_requested

    def has_active_job(self) -> bool:
        """True, если есть хотя бы один RUNNING/PENDING job."""
        return any(
            j.status in (JobStatus.pending, JobStatus.running) for j in self._jobs.values()
        )

    def list_recent(self, limit: int = 10) -> list[JobState]:
        items = sorted(self._jobs.values(), key=lambda j: j.created_at, reverse=True)
        return items[:limit]

    def schedule(self, job_id: str, coro: Coroutine[Any, Any, None]) -> asyncio.Task[None]:
        """Запускает прогон в фоне; возвращает asyncio.Task для отмены."""

        async def _run() -> None:
            try:
                await run_coro_logged(job_id, coro)
            finally:
                self._tasks.pop(job_id, None)
                self._cancel_requested.discard(job_id)

        task = asyncio.create_task(_run())
        self._tasks[job_id] = task
        return task

    async def cancel_job(self, job_id: str) -> tuple[bool, str]:
        job = self.get(job_id)
        if job is None:
            return False, "Прогон не найден."
        if job.status not in (JobStatus.pending, JobStatus.running):
            return False, "Прогон уже завершён."

        self._cancel_requested.add(job_id)
        await self.append_log(job_id, "Запрошена остановка прогона…")
        _kill_excel_processes()

        task = self._tasks.get(job_id)
        if task is not None and not task.done():
            task.cancel()

        await self.set_status(
            job_id,
            JobStatus.cancelled,
            error="Остановлено пользователем",
        )
        await self.append_log(
            job_id,
            "Прогон остановлен. Если завис Excel — процесс завершён принудительно.",
        )
        return True, "ok"

    async def append_log(self, job_id: str, line: str) -> None:
        stamped = f"{_job_log_timestamp()} | {line}"
        logger.info("[job {}] {}", job_id, stamped)
        async with self._lock:
            job = self._jobs.get(job_id)
            if job:
                if _is_ui_log_line(line):
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
            if status in (
                JobStatus.completed,
                JobStatus.failed,
                JobStatus.cancelled,
            ):
                job.finished_at = datetime.utcnow()

    async def add_downloaded(self, job_id: str, path: str) -> None:
        async with self._lock:
            job = self._jobs.get(job_id)
            if job:
                job.downloaded_paths.append(path)


job_manager = JobManager()


async def job_log(job_id: str, msg: str) -> None:
    await job_manager.append_log(job_id, msg)


async def raise_if_cancelled(job_id: str) -> None:
    """Прерывает прогон, если пользователь нажал «Остановить»."""
    if job_manager.cancel_requested(job_id):
        raise asyncio.CancelledError()


async def run_coro_logged(
    job_id: str,
    coro: Coroutine[Any, Any, None],
) -> None:
    await job_manager.set_status(job_id, JobStatus.running)
    try:
        await coro
    except asyncio.CancelledError:
        logger.info("Job {} cancelled", job_id)
        job = job_manager.get(job_id)
        if job is not None and job.status == JobStatus.running:
            await job_manager.set_status(
                job_id,
                JobStatus.cancelled,
                error="Остановлено пользователем",
            )
        await job_manager.append_log(job_id, "Прогон прерван.")
        raise
    except Exception as e:
        logger.exception("Job {} failed", job_id)
        await job_manager.append_log(job_id, f"Ошибка: {e}")
        await job_manager.set_status(job_id, JobStatus.failed, error=str(e))
    else:
        if job_manager.cancel_requested(job_id):
            await job_manager.set_status(
                job_id,
                JobStatus.cancelled,
                error="Остановлено пользователем",
            )
        else:
            await job_manager.set_status(job_id, JobStatus.completed)
