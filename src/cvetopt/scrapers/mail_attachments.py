from __future__ import annotations

import asyncio

from loguru import logger

from cvetopt.core.job_manager import job_log, job_manager
from cvetopt.core.runtime_settings import load_runtime_settings
from cvetopt.core.settings import EnvSettings
from cvetopt.mail.attachments import collect_mail_attachments


async def run_mail_attachments_job(
    job_id: str,
    env: EnvSettings,
    lookback_days_override: int | None = None,
) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.mail

    if not cfg.enabled:
        await job_log(job_id, "Почта отключена в config.yaml: mail.enabled=false")
        return
    if lookback_days_override is not None:
        cfg = cfg.model_copy(update={"lookback_days": lookback_days_override})
        await job_log(job_id, f"Переопределён период поиска писем: {cfg.lookback_days} дн.")

    try:
        runtime = load_runtime_settings(env)
        paths, log_lines = await asyncio.to_thread(
            collect_mail_attachments, cfg, env, None, runtime
        )
        for line in log_lines:
            await job_log(job_id, line)
        for p in paths:
            await job_manager.add_downloaded(job_id, str(p))
        if not paths:
            await job_log(
                job_id,
                "Новых вложений не найдено (все уже в реестре или писем без файлов).",
            )
    except Exception as e:
        logger.exception("mail attachments job failed")
        await job_log(job_id, f"Ошибка: {e}")
        raise
