from __future__ import annotations

import asyncio

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    effective_auto_new_workbook_raw,
    effective_holland_sklad_dir_raw,
    load_runtime_settings,
    resolve_auto_new_workbook,
    resolve_biflorica_archive_dir,
    resolve_holland_sklad_dir,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.auto1_pipeline import run_auto1_pipeline
from cvetopt.invoice.holland_translate import (
    archive_stale_holland_exports,
    find_holland_export_file,
)


async def run_auto1_pipeline_job(job_id: str, env: EnvSettings) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.auto1_pipeline

    if not cfg.enabled:
        await job_log(job_id, "auto1_pipeline отключён в config.yaml")
        return

    runtime = load_runtime_settings(env)
    wb_raw = effective_auto_new_workbook_raw(
        runtime,
        yaml_auto1=cfg.workbook_path,
        yaml_balance=yaml_cfg.balance_auto.workbook_path,
    )
    wb_path = resolve_auto_new_workbook(env, wb_raw)
    await job_log(job_id, f"Книга Auto_new: {wb_path}")

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

    holland_cfg = yaml_cfg.holland_translate
    if holland_cfg.archive_previous_on_auto1:
        sklad_dir = resolve_holland_sklad_dir(
            env,
            effective_holland_sklad_dir_raw(runtime, yaml_dir=holland_cfg.sklad_output_dir),
        )
        archive_dir = resolve_biflorica_archive_dir(
            env,
            runtime.biflorica_archive_dir,
            runtime=runtime,
        )
        archive_dir.mkdir(parents=True, exist_ok=True)
        newest = find_holland_export_file(sklad_dir)
        if newest is None:
            await lg(f"Голландия: файл не найден в {sklad_dir} — архив не нужен")
        else:
            await lg(f"Голландия: оставляем {newest.name}")
            try:
                await asyncio.to_thread(
                    archive_stale_holland_exports,
                    sklad_dir,
                    archive_dir,
                    keep_path=newest,
                    log=_thread_log,
                )
            except Exception as e:
                await lg(f"Голландия: архив пропущен — {e}")

    await lg("Готово: цепочка auto1 выполнена.")
