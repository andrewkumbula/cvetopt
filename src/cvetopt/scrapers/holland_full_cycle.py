from __future__ import annotations

from cvetopt.core.job_manager import job_log
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
    await job_log(job_id, "Голландия: шаг 1/3 — вложения из почты…")
    await run_mail_attachments_job(
        job_id,
        env,
        lookback_days_override=mail_lookback_days_override,
    )

    await job_log(job_id, "Голландия: шаг 2/3 — Auto1 (Scan → Import → Calculate → Sort → для склада)…")
    await run_auto1_pipeline_job(job_id, env)

    await job_log(job_id, "Голландия: шаг 3/3 — перевод Description…")
    await run_holland_translate_job(job_id, env)

    await job_log(job_id, "Готово: Голландия (почта → auto1 → перевод).")
