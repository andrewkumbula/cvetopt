from __future__ import annotations

import asyncio
from collections.abc import Awaitable, Callable
from datetime import date, datetime

from loguru import logger

from cvetopt.core.job_chains import run_balance_auto_and_delmir
from cvetopt.core.job_manager import job_manager
from cvetopt.core.runtime_settings import load_runtime_settings
from cvetopt.core.settings import EnvSettings, ScheduleTaskConfig, SelectionOverride
from cvetopt.scrapers.biflorica import run_biflorica_job
from cvetopt.scrapers.holland_full_cycle import run_holland_full_cycle_job
from cvetopt.scrapers.sklad_template_copy import run_sklad_template_copy_job

_POLL_INTERVAL_SEC = 30


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


async def start_scheduler() -> None:
    """Фоновая петля: раз в 30 с проверяет, не пора ли запустить задачу по расписанию
    из config.yaml (schedule:). Живёт, пока запущен сервер."""
    env = EnvSettings()
    schedule = env.yaml_config().schedule
    if not schedule.enabled:
        logger.info("Автозапуск по расписанию выключен (schedule.enabled: false в config.yaml).")
        return

    tasks: list[tuple[str, ScheduleTaskConfig, Callable[[], Awaitable[None]]]] = [
        ("Скачать отчёты (Biflorica + Эквадор)", schedule.biflorica, _run_biflorica_auto),
        (
            "Auto_new.xls → Голландия → Шаблон",
            schedule.tuesday_chain,
            _run_tuesday_chain,
        ),
    ]
    for name, cfg, _ in tasks:
        status = "включён" if cfg.enabled else "выключен"
        logger.info(
            "Автозапуск «{}»: {} ({} {:02d}:{:02d})",
            name, status, cfg.weekday, cfg.hour, cfg.minute,
        )

    last_fired: dict[str, date] = {}
    while True:
        now = datetime.now()
        for name, cfg, runner in tasks:
            if not cfg.enabled:
                continue
            if now.weekday() != cfg.weekday_index or now.hour != cfg.hour or now.minute < cfg.minute:
                continue
            if last_fired.get(name) == now.date():
                continue
            last_fired[name] = now.date()
            if job_manager.has_active_job():
                logger.warning(
                    "Автозапуск «{}» отложен — уже выполняется другой прогон.", name
                )
                continue
            logger.info("Автозапуск: {}", name)
            asyncio.create_task(runner())
        await asyncio.sleep(_POLL_INTERVAL_SEC)
