from __future__ import annotations

from cvetopt.core.job_manager import job_log, job_manager, run_coro_logged
from cvetopt.core.models import JobStatus
from cvetopt.core.settings import EnvSettings
from cvetopt.scrapers.balance_auto import run_balance_auto_job
from cvetopt.scrapers.delmir import run_delmir_transport_job


async def run_balance_auto_and_delmir(
    job_id: str,
    env: EnvSettings,
    delmir_lookback_days: int,
) -> None:
    """
    Объединённый шаг 2+3: сначала balance_auto (Biflorica → перелёты),
    затем delmir_transport (Транспорт трак). del-mir стартует ТОЛЬКО при успехе balance_auto.
    """
    await run_coro_logged(job_id, run_balance_auto_job(job_id, env))
    current = job_manager.get(job_id)
    if current is None or current.status != JobStatus.completed:
        await job_log(
            job_id,
            "Шаг 2 (баланс Biflorica) завершился неудачно — Транспорт трак с del-mir пропущен.",
        )
        return
    await job_log(
        job_id,
        f"Шаг 2 завершён успешно. Запускаю Транспорт трак с del-mir.com ({delmir_lookback_days} дн.)…",
    )
    await run_coro_logged(
        job_id,
        run_delmir_transport_job(
            job_id,
            env,
            lookback_days_override=delmir_lookback_days,
        ),
    )
