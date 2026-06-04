from __future__ import annotations

import asyncio

from cvetopt.core.job_manager import job_log
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.auto1_pipeline import run_auto1_pipeline


async def run_auto1_pipeline_job(job_id: str, env: EnvSettings) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.auto1_pipeline

    if not cfg.enabled:
        await job_log(job_id, "auto1_pipeline отключён в config.yaml")
        return

    root = env.project_root
    wb_path = (root / cfg.workbook_path).resolve()

    async def lg(msg: str) -> None:
        await job_log(job_id, msg)

    await lg(
        "Auto1: Scan → Import → Calculate → Sort → for sklad "
        f"(лист {cfg.sheet_name!r})"
    )

    loop = asyncio.get_running_loop()

    def _thread_log(msg: str) -> None:
        asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop).result(timeout=120)

    try:
        await asyncio.to_thread(
            run_auto1_pipeline,
            wb_path,
            cfg,
            log=_thread_log,
        )
    except Exception as e:
        await job_log(job_id, f"Ошибка auto1: {e}")
        raise

    await lg("Готово: цепочка auto1 выполнена.")
