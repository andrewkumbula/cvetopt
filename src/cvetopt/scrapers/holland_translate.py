from __future__ import annotations

import asyncio
from datetime import date

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import (
    effective_holland_dictionary_raw,
    effective_holland_sklad_dir_raw,
    load_runtime_settings,
    resolve_holland_dictionary,
    resolve_holland_sklad_dir,
)
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.holland_translate import postprocess_holland_after_auto1


async def run_holland_translate_job(job_id: str, env: EnvSettings) -> None:
    yaml_cfg = env.yaml_config()
    cfg = yaml_cfg.holland_translate

    if not cfg.enabled:
        await job_log(job_id, "holland_translate отключён в config.yaml")
        return

    runtime = load_runtime_settings(env)
    dict_raw = effective_holland_dictionary_raw(
        runtime, yaml_path=cfg.dictionary_path
    )
    sklad_raw = effective_holland_sklad_dir_raw(
        runtime, yaml_dir=cfg.sklad_output_dir
    )
    sklad_dir = resolve_holland_sklad_dir(env, sklad_raw)
    dict_path = resolve_holland_dictionary(env, dict_raw)

    await job_log(job_id, f"Перевод: папка склада {sklad_dir}")
    await job_log(job_id, f"Перевод: словарь {dict_path} (B → C)")

    loop = asyncio.get_running_loop()

    def _thread_log(msg: str) -> None:
        asyncio.run_coroutine_threadsafe(job_log(job_id, msg), loop).result(timeout=120)

    export_path = await asyncio.to_thread(
        postprocess_holland_after_auto1,
        sklad_output_dir=sklad_dir,
        dictionary_path=dict_path,
        on_date=date.today(),
        log=_thread_log,
    )
    if export_path is None:
        raise FileNotFoundError(
            f"Файл Голландия_1_*.xlsx не найден в {sklad_dir}. "
            "Сначала выполните выгрузку для склада (auto1)."
        )
    await job_log(job_id, f"Готово: {export_path.name}")
