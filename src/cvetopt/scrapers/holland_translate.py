from __future__ import annotations

import asyncio
from datetime import date

from cvetopt.core.job_manager import job_log
from cvetopt.core.runtime_settings import _resolve_dir
from cvetopt.core.settings import EnvSettings
from cvetopt.invoice.holland_translate import postprocess_holland_after_auto1


async def run_holland_translate_job(job_id: str, env: EnvSettings) -> None:
    cfg = env.yaml_config().holland_translate

    if not cfg.enabled:
        await job_log(job_id, "holland_translate отключён в config.yaml")
        return

    sklad_dir = _resolve_dir(env, cfg.sklad_output_dir, cfg.sklad_output_dir)
    dict_path = _resolve_dir(env, cfg.dictionary_path, cfg.dictionary_path)

    await job_log(
        job_id,
        f"Перевод: папка {sklad_dir}, словарь {dict_path.name} (B → C).",
    )

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
