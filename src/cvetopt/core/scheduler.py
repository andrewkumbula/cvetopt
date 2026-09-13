from __future__ import annotations

import asyncio
import json
from collections.abc import Awaitable, Callable
from datetime import date, datetime, timedelta
from pathlib import Path

from loguru import logger

from cvetopt.core.job_chains import run_balance_auto_and_delmir
from cvetopt.core.job_manager import job_manager
from cvetopt.core.models import JobStatus
from cvetopt.core.runtime_settings import load_runtime_settings
from cvetopt.core.settings import (
    EnvSettings,
    ScheduleConfig,
    ScheduleTaskConfig,
    SelectionOverride,
)
from cvetopt.scrapers.biflorica import run_biflorica_job
from cvetopt.scrapers.holland_full_cycle import run_holland_full_cycle_job
from cvetopt.scrapers.sklad_template_copy import run_sklad_template_copy_job

_POLL_INTERVAL_SEC = 30

ScheduledTask = tuple[str, ScheduleTaskConfig, Callable[[], Awaitable[None]]]

# Иначе asyncio может собрать задачу сборщиком мусора на середине прогона.
_running: set[asyncio.Task[None]] = set()


async def _run_and_wait(portal_id: str, coro_factory) -> None:
    job = job_manager.create_job(portal_id)
    task = job_manager.schedule(job.id, coro_factory(job.id))
    await task


async def _run_biflorica_auto() -> None:
    env = EnvSettings()
    settings = load_runtime_settings(env)
    selection = SelectionOverride(
        min_age_days=settings.biflorica_min_age_days,
        max_age_days=settings.biflorica_max_age_days,
    )
    await _run_and_wait(
        f"auto:biflorica:{settings.biflorica_min_age_days}-{settings.biflorica_max_age_days}d",
        lambda job_id: run_biflorica_job(job_id, env, selection_override=selection),
    )


async def _run_tuesday_chain() -> None:
    """По очереди: Auto_new.xls (+ del-mir) → Голландия → Шаблон на сегодня."""
    env = EnvSettings()
    settings = load_runtime_settings(env)

    await _run_and_wait(
        f"auto:balance_auto+delmir:{settings.delmir_lookback_days}d",
        lambda job_id: run_balance_auto_and_delmir(job_id, env, settings.delmir_lookback_days),
    )
    await _run_and_wait(
        f"auto:holland_full_cycle:{settings.mail_lookback_days}d",
        lambda job_id: run_holland_full_cycle_job(
            job_id, env, mail_lookback_days_override=settings.mail_lookback_days
        ),
    )
    await _run_and_wait(
        "auto:sklad_template_copy",
        lambda job_id: run_sklad_template_copy_job(job_id, env),
    )


def _state_path(env: EnvSettings) -> Path:
    return env.project_root / "data" / "state" / "scheduler_state.json"


def _load_last_fired(env: EnvSettings) -> dict[str, date]:
    """Дата последнего срабатывания каждой задачи — переживает перезапуск сервера."""
    try:
        raw = json.loads(_state_path(env).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return {}
    except (OSError, json.JSONDecodeError):
        logger.warning("Состояние планировщика нечитаемо — считаем, что запусков не было.")
        return {}
    if not isinstance(raw, dict):
        return {}
    result: dict[str, date] = {}
    for name, value in raw.items():
        try:
            result[name] = date.fromisoformat(value)
        except (TypeError, ValueError):
            continue
    return result


def _save_last_fired(env: EnvSettings, last_fired: dict[str, date]) -> None:
    path = _state_path(env)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp = path.with_suffix(".json.tmp")
        tmp.write_text(
            json.dumps({k: v.isoformat() for k, v in last_fired.items()}, indent=2),
            encoding="utf-8",
        )
        tmp.replace(path)
    except OSError as e:
        # Без записи перезапуск сервера в то же окно запустит задачу второй раз.
        logger.error("Не удалось сохранить состояние планировщика ({}): {}", path, e)


async def _cancel_stuck_jobs(max_hours: int) -> None:
    """Зависший Excel держит прогон в running вечно, а планировщик из-за этого молча
    пропускает все следующие запуски. Снимаем такой прогон, чтобы очередь освободилась."""
    deadline = datetime.utcnow() - timedelta(hours=max_hours)
    for job in job_manager.list_recent(limit=20):
        if job.status not in (JobStatus.pending, JobStatus.running):
            continue
        if job.created_at > deadline:
            continue
        logger.error(
            "Прогон {} ({}) идёт дольше {} ч — снимаю, иначе расписание встанет.",
            job.id, job.portal_id, max_hours,
        )
        await job_manager.cancel_job(job.id)


def _spawn(name: str, runner: Callable[[], Awaitable[None]]) -> None:
    task = asyncio.create_task(runner())
    _running.add(task)

    def _done(finished: asyncio.Task[None]) -> None:
        _running.discard(finished)
        if finished.cancelled():
            return
        error = finished.exception()
        if error is not None:
            logger.error("Автозапуск «{}» прерван ошибкой: {}", name, error)

    task.add_done_callback(_done)


async def _tick(
    env: EnvSettings,
    schedule: ScheduleConfig,
    tasks: list[ScheduledTask],
    last_fired: dict[str, date],
) -> None:
    await _cancel_stuck_jobs(schedule.max_job_hours)

    now = datetime.now()
    for name, cfg, runner in tasks:
        if not cfg.enabled:
            continue
        if now.weekday() != cfg.weekday_index or now.hour != cfg.hour or now.minute < cfg.minute:
            continue
        if last_fired.get(name) == now.date():
            continue
        if job_manager.has_active_job():
            # День не отмечаем: попробуем снова, пока не кончится час запуска.
            logger.warning("Автозапуск «{}» отложен — уже выполняется другой прогон.", name)
            continue
        last_fired[name] = now.date()
        _save_last_fired(env, last_fired)
        logger.info("Автозапуск: {}", name)
        _spawn(name, runner)


async def start_scheduler() -> None:
    """Фоновая петля: раз в 30 с проверяет, не пора ли запустить задачу по расписанию
    из config.yaml (schedule:). Живёт, пока запущен сервер."""
    env = EnvSettings()
    schedule = env.yaml_config().schedule
    if not schedule.enabled:
        logger.info("Автозапуск по расписанию выключен (schedule.enabled: false в config.yaml).")
        return

    tasks: list[ScheduledTask] = [
        ("Скачать отчёты (Biflorica + Эквадор)", schedule.biflorica, _run_biflorica_auto),
        ("Auto_new.xls → Голландия → Шаблон", schedule.weekly_chain, _run_tuesday_chain),
    ]
    for name, cfg, _ in tasks:
        status = "включён" if cfg.enabled else "выключен"
        logger.info(
            "Автозапуск «{}»: {} ({} {:02d}:{:02d})",
            name, status, cfg.weekday, cfg.hour, cfg.minute,
        )

    last_fired = _load_last_fired(env)
    while True:
        try:
            await _tick(env, schedule, tasks, last_fired)
        except Exception:
            # Иначе одна ошибка навсегда убивает расписание в этом процессе.
            logger.exception("Сбой в цикле планировщика — продолжаю со следующей проверки.")
        await asyncio.sleep(_POLL_INTERVAL_SEC)
