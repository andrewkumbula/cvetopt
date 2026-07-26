from __future__ import annotations

from cvetopt.core.job_manager import job_step
from cvetopt.core.settings import EnvSettings
from cvetopt.scrapers.auto1_pipeline import run_auto1_pipeline_job
from cvetopt.scrapers.holland_translate import run_holland_translate_job
from cvetopt.scrapers.mail_attachments import run_mail_attachments_job


async def run_holland_full_cycle_job(
    job_id: str,
    env: EnvSettings,
    *,
    mail_lookback_days_override: int | None = None,
) -> None:
    """Почта → auto1 (Scan … for sklad) → перевод Description."""
    await job_step(job_id, "Шаг 1 из 3: забираю вложения из почты…")
    await run_mail_attachments_job(
        job_id,
        env,
        lookback_days_override=mail_lookback_days_override,
    )

    await job_step(job_id, "Шаг 2 из 3: обработка в Excel (auto1) и файл для склада…")
    await run_auto1_pipeline_job(job_id, env)

    await job_step(job_id, "Шаг 3 из 3: перевод описаний…")
    await run_holland_translate_job(job_id, env)

    await job_step(job_id, "Готово: Голландия (почта → обработка → перевод).")
