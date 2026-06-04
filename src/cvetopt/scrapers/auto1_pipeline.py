from __future__ import annotations

import asyncio
from datetime import date

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    _resolve_dir,
    effective_auto_new_workbook_raw,
    load_runtime_settings,
    resolve_auto_new_workbook,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.auto1_pipeline import run_auto1_pipeline
from cvetopt.invoice.holland_translate import postprocess_holland_after_auto1


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

    if cfg.translate_holland_export:
        sklad_dir = _resolve_dir(env, cfg.sklad_output_dir, cfg.sklad_output_dir)
        dict_path = _resolve_dir(env, cfg.dictionary_path, cfg.dictionary_path)
        await lg(
            f"Перевод Description в выгрузке склада (словарь {dict_path.name})…"
        )
        try:
            export_path = await asyncio.to_thread(
                postprocess_holland_after_auto1,
                sklad_output_dir=sklad_dir,
                dictionary_path=dict_path,
                on_date=date.today(),
                log=_thread_log,
            )
            if export_path is None:
                await job_log(
                    job_id,
                    f"Файл Голландия_1_*.xlsx не найден в {sklad_dir} — перевод пропущен.",
                )
        except Exception as e:
            await job_log(job_id, f"Ошибка перевода Description: {e}")
            raise

    await lg("Готово: цепочка auto1 выполнена.")
